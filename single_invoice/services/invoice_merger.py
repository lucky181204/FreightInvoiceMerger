"""Single invoice merger — main processing logic.

Generates a single Freight Invoice from:
- Invoice template (.xlsx)
- BL Draft (.xlsx)
- Shanghai manifest (舱单)
- Trade terms manifest (贸易条款清单)

Output filename: Freight Invoice to DAI PO{PO}.xlsx
"""

import logging
import os
import re
import time
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, numbers
from openpyxl.utils import column_index_from_string

from single_invoice.parsers.draft_parser import (
    extract_draft_data, search_po_in_draft, get_container_qty
)
from single_invoice.parsers.trade_terms_parser import find_by_blno as find_trade_by_blno
from single_invoice.parsers.manifest_parser import find_containers_by_blno
from single_invoice.services.hidden_sheet_service import remove_hidden_sheets
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

    Returns dict with keys:
      output_path, success (bool), message (str)
      po (str), bl_no (str), errors (list)
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
    # Sanitize filename (remove invalid chars)
    output_name = re.sub(r'[<>:"/\\|?*]', '_', output_name)
    output_path = str(Path(output_dir) / output_name)

    counter = 1
    while os.path.exists(output_path):
        stem = Path(output_path).stem
        ext = Path(output_path).suffix
        # Remove existing (N) suffix if present
        stem_clean = re.sub(r'\(\d+\)$', '', stem)
        output_path = str(Path(output_dir) / f"{stem_clean}({counter}){ext}")
        counter += 1

    try:
        wb = load_workbook(template_path)
        ws = wb.active

        # ── Fill Fields ──
        # H4 = current date
        today = date.today()
        ws["H4"] = f"{today.year}/{today.month}/{today.day}"

        # B5 = first word from draft B5
        ws["B5"] = draft_data.get("B5", "")

        # B6 = destination port from draft
        ws["B6"] = draft_data.get("B6_port", "")

        # B6 country (right side of B6 or another location)
        # Typically destination port maps to B2 or similar
        # Try to find country cell

        # B7 = container info from draft
        ws["B7"] = draft_data.get("B7", "")

        # B10 = Port of Loading
        ws["B10"] = draft_data.get("B10", "")

        # B11 = Dangerous goods class
        ws["B11"] = draft_data.get("B11", "")

        # B12 = Place of Receipt
        ws["B12"] = draft_data.get("B12", "")

        # D10 = BL No
        ws["D10"] = bl_no

        # E10 = Vessel/Voyage
        ws["E10"] = draft_data.get("E10", "")

        # C10 = Customer PO (from trade terms)
        if po:
            ws["C10"] = po

        # H7 = Business No (from trade terms)
        if business_no:
            ws["H7"] = business_no

        # H5 = Sailing date (from trade terms)
        if sailing_date:
            ws["H5"] = sailing_date

        # ── Container data (below row 10) ──
        # Start from row 12 or next available row
        container_start = 12
        for i, c in enumerate(containers):
            row = container_start + i
            ws.cell(row=row, column=6).value = c.get("container_no", "")  # F column
            ws.cell(row=row, column=7).value = c.get("container_type", "")  # G column
            # H column = USD if container_type non-empty
            if c.get("container_type", "").strip():
                ws.cell(row=row, column=8).value = "USD"  # H column

        # ── I column format ──
        # Apply number format #,##0.######## to I column data area
        for row_num in range(10, 50):
            cell = ws.cell(row=row_num, column=9)
            if cell.value is not None and cell.value != "":
                try:
                    float(cell.value)
                    cell.number_format = '#,##0.########'
                except (ValueError, TypeError):
                    pass

        # ── B30 formula ──
        ws["B30"] = f'="TOTAL FREIGHT USD "&TEXT(I29,"#,##0.########")&"."'

        # ── PO Validation ──
        if po:
            progress(f"校验PO：{po}")
            po_found = search_po_in_draft(draft_path, po)
            if not po_found:
                ws["C10"].font = Font(color="FF0000")  # Red font
                warnings.append(f"PO {po} 在Draft中未找到，C10已标红")
            else:
                # Keep original font (openpyxl preserves by default)
                pass

        # ── Container quantity validation ──
        b7_text = draft_data.get("B7", "")
        expected_qty = get_container_qty(b7_text)
        actual_qty = len(containers)
        progress(f"箱量校验：Draft={expected_qty}, 舱单={actual_qty}")
        if expected_qty > 0 and expected_qty != actual_qty:
            ws["B7"].font = Font(color="FF0000")  # Red font
            warnings.append(f"箱量不一致：Draft={expected_qty}, 舱单={actual_qty}, B7已标红")

        # ── Remove hidden sheets ──
        progress("检查隐藏工作表...")
        remove_hidden_sheets(wb)

        # ── Save ──
        wb.save(output_path)
        wb.close()
        progress(f"保存完成：{output_name}")

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
