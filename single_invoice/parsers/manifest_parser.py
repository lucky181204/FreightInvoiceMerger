"""Manifest (舱单) parser.

Matches BL No in the manifest to find container numbers and types.
Manifest has a specific template structure: header rows (1-20), then data rows.
"""

import logging
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


def find_containers_by_blno(manifest_path: str, bl_no: str) -> list[dict]:
    """
    Search manifest for BL No and return matching container info.

    Manifest template structure:
      Row 21: Header row (*提单号 | *箱号 | *封号 | *箱型 | ...)
      Row 22+: Data rows (BL No, Container No, Seal No, Container Type, ...)

    Returns list of dicts with 'container_no' and 'container_type'.
    Uses the container's own BL to match - skips the header.
    """
    if not bl_no:
        return []

    wb = load_workbook(manifest_path, data_only=True)
    results = []

    try:
        ws = wb["舱单"]
    except KeyError:
        # Try first sheet
        ws = wb.active

    try:
        # Find header row: look for *提单号 and *箱号
        header_row = None
        bl_col = None
        cntr_col = None
        type_col = None

        for r in range(1, min(ws.max_row + 1, 25)):
            for c in range(1, min(ws.max_column + 1, 12)):
                v = ws.cell(row=r, column=c).value
                if v and isinstance(v, str):
                    v_clean = v.strip()
                    if "提单号" in v_clean or v_clean == "*提单号":
                        bl_col = c
                        header_row = r
                    elif "箱号" in v_clean or v_clean == "*箱号":
                        cntr_col = c
                    elif "箱型" in v_clean or v_clean == "*箱型":
                        type_col = c

        if not bl_col or not cntr_col or not header_row:
            logger.warning("舱单未找到提单号/箱号/箱型列")
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
                # This is a container row under our BL
                cntr_no = str(cntr_val).strip()
                type_val = ""
                if type_col:
                    tv = ws.cell(row=r, column=type_col).value
                    if tv:
                        type_val = str(tv).strip()
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
