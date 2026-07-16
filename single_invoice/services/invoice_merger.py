"""Single invoice merger — main processing logic.

Generates a single Freight Invoice from:
- Invoice template (.xls)
- BL Draft (.docx)
- Shanghai manifest (舱单)
- Trade terms manifest (贸易条款清单)

Output preserves template formatting (font, size, style).
Only modifies cell values — never cells that have placeholder text like "Loading Port:*".
Instead fills data alongside the template's label-text structure.
"""

import os
import re
from datetime import date
from pathlib import Path

from single_invoice.parsers.draft_parser import (
    extract_draft_data, search_po_in_draft, get_container_qty
)
from single_invoice.parsers.trade_terms_parser import find_by_blno as find_trade_by_blno
from single_invoice.parsers.manifest_parser import find_containers_by_blno
from utils.logger import logger


def generate_single_invoice(
    template_path: str,
    draft_path: str,
    manifest_path: str,
    trade_terms_path: str,
    output_dir: str,
    progress_callback=None,
) -> dict:
    """
    Generate a single Freight Invoice from template + draft + manifests.
    """
    errors = []
    warnings = []

    def progress(msg):
        if progress_callback:
            progress_callback(msg)
        logger.info(msg)

    progress("读取Draft...")
    try:
        draft_data = extract_draft_data(draft_path)
    except Exception as e:
        errors.append(f"Draft读取失败：{e}")
        return {"output_path": "", "success": False, "message": str(e),
                "po": "", "bl_no": "", "errors": errors}

    bl_no = draft_data.get("D10", "")
    progress(f"提单号：{bl_no}")

    # Trade terms lookup
    trade_info = {}
    if trade_terms_path and bl_no:
        progress("读取贸易条款清单...")
        trade_info = find_trade_by_blno(trade_terms_path, bl_no) or {}
        if trade_info:
            progress(f"找到PO：{trade_info.get('po', '')}")
        else:
            warnings.append(f"贸易条款清单未找到提单号：{bl_no}")

    po = trade_info.get("po", "")
    business_no = trade_info.get("business_no", "")
    sailing_date = trade_info.get("sailing_date", "")

    # Manifest lookup
    containers = []
    if manifest_path and bl_no:
        progress("读取上海舱单...")
        containers = find_containers_by_blno(manifest_path, bl_no)
        if containers:
            progress(f"找到{len(containers)}个集装箱")
        else:
            warnings.append(f"舱单未找到提单号：{bl_no}")

    progress("生成Invoice...")

    output_name = f"Freight Invoice to DAI PO{po}.xlsx" if po else "Freight Invoice to DAI.xlsx"
    output_name = re.sub(r'[<>:"/\\|?*]', '_', output_name)
    output_path = str(Path(output_dir) / output_name)
    counter = 1
    while os.path.exists(output_path):
        stem = Path(output_path).stem
        stem_clean = re.sub(r'\(\d+\)$', '', stem)
        output_path = str(Path(output_dir) / f"{stem_clean}({counter}){Path(output_path).suffix}")
        counter += 1

    try:
        template_ext = Path(template_path).suffix.lower()
        if template_ext == ".xls":
            import xlrd
            import xlwt
            import xlutils.copy as cp

            # Patch None format_str
            orig_nfr = xlwt.BIFFRecords.NumberFormatRecord.__init__
            def _patched_nfr(self, fmtidx, fmtstr):
                if fmtstr is None:
                    fmtstr = ''
                orig_nfr(self, fmtidx, fmtstr)
            xlwt.BIFFRecords.NumberFormatRecord.__init__ = _patched_nfr

            rb = xlrd.open_workbook(template_path, formatting_info=True)
            wb = cp.copy(rb)
            sheet_name = rb.sheet_names()[-1]
            tmpl_fmt = 'xls'

            # Helper: read source font for a cell, return an xlwt style
            src_sheet = rb.sheet_by_index(rb.nsheets - 1)

            def _get_style(r, c):
                """Get xlwt style preserving template font for cell (1-based r,c)."""
                try:
                    xf_idx = src_sheet.cell_xf_index(r - 1, c - 1)
                    font = rb.font_list[rb.xf_list[xf_idx].font_index]
                    fnt = xlwt.Font()
                    fnt.name = font.name
                    fnt.height = font.height
                    fnt.bold = bool(font.bold)
                    style = xlwt.XFStyle()
                    style.font = fnt
                    return style
                except Exception:
                    return None

            def set_cell(r, c, val):
                """Write cell value preserving template font (1-based r,c)."""
                st = _get_style(r, c)
                ws = wb.get_sheet(sheet_name)
                if st:
                    ws.write(r - 1, c - 1, val, st)
                else:
                    ws.write(r - 1, c - 1, val)

            def save_wb(p):
                wb.save(p)

        else:
            from openpyxl import load_workbook
            from openpyxl.styles import Font as OXFont
            wb = load_workbook(template_path)
            ws = wb.active
            tmpl_fmt = 'xlsx'

            def set_cell(r, c, val):
                ws.cell(row=r, column=c).value = val

            def save_wb(p):
                from single_invoice.services.hidden_sheet_service import remove_hidden_sheets
                remove_hidden_sheets(wb)
                wb.save(p)

        # ── BUILD the cell values based on template structure ──
        # The template has prefix+placeholder like "Loading Port:*,CHINA"
        # We REPLACE the placeholder part, keeping the label prefix

        today = date.today()
        date_str = f"{today.year}/{today.month}/{today.day}"

        shipper = draft_data.get("shipper", "")
        loading_port = draft_data.get("B10_loading_port", "")
        discharge_port = draft_data.get("B6_discharge_port", "")
        container_info = draft_data.get("B7_container", "")
        vessel = draft_data.get("E10", "")
        goods_desc = draft_data.get("B12_goods_desc", "")
        danger = draft_data.get("B11_danger_class", "")

        # ── Build each cell value with proper formatting ──
        # B5: "Loading Port:{port},{country}"
        set_cell(5, 2, f"Loading Port:{loading_port},CHINA")

        # B6: "Destination: {port},{country}"
        set_cell(6, 2, f"Destination: {discharge_port},COLOMBIA")

        # B7: "Shipment of: {container_info}"
        set_cell(7, 2, f"Shipment of: {container_info}")

        # H4: "Invoice date:{date}"
        set_cell(4, 8, f"Invoice date:{date_str}")

        # H5: ETD
        if sailing_date:
            set_cell(5, 8, f"ETD:{sailing_date}")
        else:
            set_cell(5, 8, "ETD:*")

        # H6: ATD (use same date as template placeholder or sailing)
        # Keep ATD placeholder or fill with sailing
        h6_val = f"ATD:{sailing_date}" if sailing_date else "ATD:*"
        set_cell(6, 8, h6_val)

        # H7: "Invoice #:{business_no}"
        set_cell(7, 8, f"Invoice #:{business_no}" if business_no else "Invoice #:*")

        # B10: "PRODUCT: {shipper}"
        set_cell(10, 2, f"PRODUCT: {shipper}")

        # B11: "CLASS: {danger}"
        set_cell(11, 2, f"CLASS: {danger}" if danger else "CLASS: *")

        # B12: "Shipper: {goods_desc}"
        set_cell(12, 2, f"Shipper: {goods_desc}")

        # C10: "PO{po}"
        po_val = f"PO{po}" if po else "PO*"
        set_cell(10, 3, po_val)

        # D10: BL No
        set_cell(10, 4, bl_no)

        # E10: Vessel/Voyage
        set_cell(10, 5, vessel)

        # Container data rows (F=6, G=7, H=8) — same row as B10/C10/D10/E10
        for i, c_info in enumerate(containers):
            row = 10  # Container data goes on row 10, same as BL info
            set_cell(row, 6, c_info.get("container_no", ""))   # F = Container No
            set_cell(row, 7, c_info.get("container_type", "")) # G = Container Type
            if c_info.get("container_type", "").strip():
                set_cell(row, 8, "USD")                         # H = USD

        # B30 formula
        formula = '="TOTAL FREIGHT USD"&TEXT(I29,"#,##0.########")&"."'
        set_cell(30, 2, formula)

        # ── PO Validation ──
        if po:
            progress(f"校验PO：{po}")
            po_found = search_po_in_draft(draft_path, po)
            if not po_found:
                warnings.append(f"PO {po} 在Draft中未找到")

        # ── Container quantity validation ──
        expected_qty = get_container_qty(container_info)
        actual_qty = len(containers)
        progress(f"箱量校验：Draft={expected_qty}, 舱单={actual_qty}")
        if expected_qty > 0 and expected_qty != actual_qty:
            warnings.append(f"箱量不一致：Draft={expected_qty}, 舱单={actual_qty}")

        # ── Save ──
        save_wb(output_path)
        progress(f"保存完成：{output_path}")

    except Exception as e:
        errors.append(f"模板写入失败：{e}")
        import traceback
        traceback.print_exc()
        return {"output_path": "", "success": False, "message": str(e),
                "po": po, "bl_no": bl_no, "errors": errors}

    logger.info(f"成功：{output_name}")
    return {
        "output_path": output_path,
        "success": True,
        "message": "单张发票生成成功",
        "po": po,
        "bl_no": bl_no,
        "warnings": warnings,
        "errors": errors,
    }
