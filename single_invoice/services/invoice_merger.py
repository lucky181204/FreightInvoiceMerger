"""Single invoice merger — main processing logic.

Generates a single Freight Invoice from:
- Invoice template (.xls)
- BL Draft (.docx)
- Shanghai manifest (舱单)
- Trade terms manifest (贸易条款清单)

Output filename: Freight Invoice to DAI PO{PO}.xlsx
"""

import logging
import os
import re
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font
from single_invoice.parsers.draft_parser import (
    extract_draft_data, search_po_in_draft, get_container_qty
)
from single_invoice.parsers.trade_terms_parser import find_by_blno as find_trade_by_blno
from single_invoice.parsers.manifest_parser import find_containers_by_blno
from single_invoice.services.hidden_sheet_service import remove_hidden_sheets
from utils.logger import logger

# Track whether to use openpyxl or xlrd for template
_XLS_EXTENSIONS = {".xls"}


def _is_xls(path: str) -> bool:
    return Path(path).suffix.lower() in _XLS_EXTENSIONS


def _load_template(template_path: str):
    """Load a template file - supports both .xlsx and .xls."""
    if _is_xls(template_path):
        import xlrd
        import xlwt
        import xlutils.copy as cp

        # Patch None format_str bug
        orig_nfr = xlwt.BIFFRecords.NumberFormatRecord.__init__
        def _patched_nfr(self, fmtidx, fmtstr):
            if fmtstr is None:
                fmtstr = ''
            orig_nfr(self, fmtidx, fmtstr)
        xlwt.BIFFRecords.NumberFormatRecord.__init__ = _patched_nfr

        rb = xlrd.open_workbook(template_path, formatting_info=True)
        wb = cp.copy(rb)
        # Return (xlwt_workbook, sheet_name, 'xls')
        sheet_name = rb.sheet_names()[-1]  # Last sheet (data sheet)
        ws = wb.get_sheet(sheet_name)
        return wb, ws, sheet_name, 'xls'
    else:
        wb = load_workbook(template_path)
        ws = wb.active
        return wb, ws, ws.title, 'xlsx'


def _save_template(wb, fmt: str, output_path: str):
    """Save the workbook in its original format."""
    if fmt == 'xls':
        wb.save(output_path)
    else:
        wb.save(output_path)


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

    Returns dict with keys:
      output_path, success (bool), message (str)
      po (str), bl_no (str), errors (list), warnings (list)
    """
    errors = []
    warnings = []

    def progress(msg):
        if progress_callback:
            progress_callback(msg)
        logger.info(msg)

    progress("读取Draft...")

    # ── Step 1: Extract draft data ──
    try:
        draft_data = extract_draft_data(draft_path)
    except Exception as e:
        errors.append(f"Draft读取失败：{e}")
        return {"output_path": "", "success": False, "message": str(e),
                "po": "", "bl_no": "", "errors": errors}

    bl_no = draft_data.get("D10", "")
    progress(f"提单号：{bl_no}")

    # ── Step 2: Look up trade terms ──
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

    # ── Step 3: Look up manifest ──
    containers = []
    if manifest_path and bl_no:
        progress("读取上海舱单...")
        containers = find_containers_by_blno(manifest_path, bl_no)
        if containers:
            progress(f"找到{len(containers)}个集装箱")
        else:
            warnings.append(f"舱单未找到提单号：{bl_no}")

    # ── Step 4: Load template and fill ──
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
        # Load template (supports .xls and .xlsx)
        tmpl_fmt = 'xlsx'
        if _is_xls(template_path):
            import xlrd
            import xlwt
            import xlutils.copy as cp

            orig_nfr = xlwt.BIFFRecords.NumberFormatRecord.__init__
            def _patched_nfr(self, fmtidx, fmtstr):
                if fmtstr is None:
                    fmtstr = ''
                orig_nfr(self, fmtidx, fmtstr)
            xlwt.BIFFRecords.NumberFormatRecord.__init__ = _patched_nfr

            rb = xlrd.open_workbook(template_path, formatting_info=True)
            wb = cp.copy(rb)
            sheet_name = rb.sheet_names()[-1]
            ws = wb.get_sheet(sheet_name)
            tmpl_fmt = 'xls'

            # For xlwt, cell writing uses 0-indexed rows/cols
            def set_cell(r, c, val):
                """Set cell value, 1-based row and column."""
                ws.write(int(r) - 1, int(c) - 1, val)

        else:
            wb = load_workbook(template_path)
            ws = wb.active

            def set_cell(r, c, val):
                """Set cell value, 1-based row and column."""
                ws.cell(row=int(r), column=int(c)).value = val

        # ── Fill Fields ──
        today = date.today()
        set_cell(4, 8, f"{today.year}/{today.month}/{today.day}")  # H4

        set_cell(5, 2, draft_data.get("B5", ""))    # B5
        set_cell(6, 2, draft_data.get("B6_port", ""))  # B6
        set_cell(7, 2, draft_data.get("B7", ""))    # B7
        set_cell(10, 2, draft_data.get("B10", ""))  # B10
        set_cell(11, 2, draft_data.get("B11", ""))  # B11
        set_cell(12, 2, draft_data.get("B12", ""))  # B12
        set_cell(10, 4, bl_no)                       # D10
        set_cell(10, 5, draft_data.get("E10", ""))  # E10

        if po:
            set_cell(10, 3, po)                      # C10
        if business_no:
            set_cell(7, 8, business_no)               # H7
        if sailing_date:
            set_cell(5, 8, sailing_date)              # H5

        # ── Container data ──
        container_start = 12
        for i, c_info in enumerate(containers):
            row = container_start + i
            set_cell(row, 6, c_info.get("container_no", ""))   # F
            set_cell(row, 7, c_info.get("container_type", "")) # G
            if c_info.get("container_type", "").strip():
                set_cell(row, 8, "USD")                        # H

        # ── B30 formula ──
        formula = '="TOTAL FREIGHT USD "&TEXT(I29,"#,##0.########")&"."'
        if tmpl_fmt == 'xls':
            ws.write(29, 1, formula)  # 0-indexed: row 29=30, col 1=B
        else:
            ws["B30"] = formula

        # ── PO Validation ──
        if po:
            progress(f"校验PO：{po}")
            po_found = search_po_in_draft(draft_path, po)
            if not po_found and tmpl_fmt == 'xlsx':
                # Only openpyxl supports per-cell font changes
                # For xls, skip font coloring
                if tmpl_fmt == 'xlsx':
                    ws["C10"].font = Font(color="FF0000")
                warnings.append(f"PO {po} 在Draft中未找到")

        # ── Container quantity validation ──
        b7_text = draft_data.get("B7", "")
        expected_qty = get_container_qty(b7_text)
        actual_qty = len(containers)
        progress(f"箱量校验：Draft={expected_qty}, 舱单={actual_qty}")
        if expected_qty > 0 and expected_qty != actual_qty:
            warnings.append(f"箱量不一致：Draft={expected_qty}, 舱单={actual_qty}")

        # ── Remove hidden sheets (openpyxl only) ──
        if tmpl_fmt != 'xls':
            progress("检查隐藏工作表...")
            remove_hidden_sheets(wb)

        # ── Save ──
        if tmpl_fmt == 'xls':
            wb.save(output_path)
        else:
            wb.save(output_path)

        progress(f"保存完成：{output_path}")

    except Exception as e:
        errors.append(f"模板写入失败：{e}")
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
