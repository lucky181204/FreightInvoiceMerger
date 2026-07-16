"""Manifest (舱单) parser.

Matches BL No in the manifest to find container numbers and types.
Dynamically identifies column headers by content keyword matching.

Compatible column names:
  Container No.:  箱号, *箱号, 柜号, *柜号, 集装箱号, Container No.
  Container Type: 箱型, *箱型, 箱型尺寸, Container Type
"""

import logging
import re
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


def find_containers_by_blno(manifest_path: str, bl_no: str) -> list[dict]:
    """
    Search manifest for BL No and return matching container info.

    Dynamically detects column positions by header content matching.

    Returns list of dicts with:
      container_no (str), container_type (str)
    """
    if not bl_no:
        return []

    wb = load_workbook(manifest_path, data_only=True)
    results = []

    try:
        ws = wb["舱单"]
    except KeyError:
        ws = wb.active

    try:
        # Find header row by locating BL No or 提单号 column
        header_row = None
        bl_col = None
        cntr_col = None
        type_col = None
        type_next_col = None  # column to the right of type column (for combining)

        for r in range(1, min(ws.max_row + 1, 25)):
            for c in range(1, min(ws.max_column + 1, 15)):
                v = ws.cell(row=r, column=c).value
                if not v or not isinstance(v, str):
                    continue
                vc = v.strip().replace("\n", "").replace(" ", "").replace("　", "")

                # BL No / 提单号
                if "提单号" in vc:
                    bl_col = c
                    header_row = r

                # Container No: 箱号 / 柜号 / 集装箱号 / Container No
                if "箱号" in vc or "柜号" in vc:
                    cntr_col = c
                elif "集装箱号" in vc:
                    cntr_col = c
                elif "CONTAINERNO" in vc.upper().replace(" ", "").replace(".", ""):
                    cntr_col = c

                # Container Type: 箱型 / 箱型尺寸 / Container Type
                if "箱型" in vc:
                    type_col = c
                elif "CONTAINERTYPE" in vc.upper().replace(" ", "").replace(".", ""):
                    type_col = c

        if not bl_col or not header_row:
            logger.warning("舱单未找到提单号列")
            wb.close()
            return []

        if not cntr_col:
            logger.warning("舱单未找到箱号/柜号列")
            wb.close()
            return []

        # Search data rows starting after header
        current_bl = None
        in_matching_bl = False
        for r in range(header_row + 1, ws.max_row + 1):
            bl_val = ws.cell(row=r, column=bl_col).value
            if bl_val and str(bl_val).strip():
                current_bl = str(bl_val).strip()
                in_matching_bl = (current_bl == bl_no)

            cntr_val = ws.cell(row=r, column=cntr_col).value
            if in_matching_bl and cntr_val and str(cntr_val).strip():
                cntr_no = str(cntr_val).strip()
                # Container type
                type_val = ""
                if type_col:
                    tv = ws.cell(row=r, column=type_col).value
                    next_tv = ws.cell(row=r, column=type_col + 1).value if type_col + 1 <= ws.max_column else None
                    type_val = normalize_container_type(tv, next_tv)

                results.append({
                    "container_no": cntr_no,
                    "container_type": type_val,
                })

        wb.close()
        return results

    except Exception as e:
        logger.warning(f"舱单读取失败: {e}")
        try:
            wb.close()
        except Exception:
            pass
        return []


def normalize_container_type(type_value, next_value=None) -> str:
    """
    Normalize container type value.

    - If pure digits (e.g. "40", "20", "45"), combine with next column value
    - If already has letters (e.g. "40HQ", "20GP"), use as-is
    - Remove ".0" suffix from numbers
    """
    raw_type = str(type_value or "").strip()
    raw_next = str(next_value or "").strip()

    # Remove .0 suffix
    if raw_type.endswith(".0"):
        raw_type = raw_type[:-2]

    # If pure digits, combine with next column
    if re.fullmatch(r"\d+", raw_type):
        return f"{raw_type}{raw_next}"

    return raw_type
