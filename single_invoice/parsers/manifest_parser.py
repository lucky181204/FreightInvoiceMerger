"""Manifest (舱单) parser.

Matches BL No (D10) in the manifest to find container numbers and types.
"""

import logging
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


def find_containers_by_blno(manifest_path: str, bl_no: str) -> list[dict]:
    """
    Search manifest for BL No and return matching container info.

    Returns list of dicts with 'container_no' and 'container_type'.
    The manifest has columns like:
      提单号 | Container No. | Container Type | ...
    """
    if not bl_no:
        return []

    wb = load_workbook(manifest_path, data_only=True)
    results = []

    try:
        for ws in wb.worksheets:
            # Find header row
            headers = {}
            for row in ws.iter_rows(min_row=1, max_row=15, values_only=False):
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        val = cell.value.strip().upper().replace(" ", "").replace("　", "")
                        headers[cell.column] = val

            bl_col = None
            cntr_col = None
            type_col = None

            for col, val in headers.items():
                if val in ("提单号", "BLNO", "BL NO", "B/L NO", "*"):
                    bl_col = col
                elif val in ("箱号", "CONTAINERNO.", "CONTAINER NO", "集装箱号"):
                    cntr_col = col
                elif val in ("箱型", "CONTAINERTYPE", "CONTAINER TYPE", "尺寸类型"):
                    type_col = col

            # Try alternate detection: look for columns with "*" marker
            if not bl_col:
                for col, val in headers.items():
                    if "提单" in val or "BL" in val:
                        bl_col = col

            if bl_col and cntr_col:
                for row in ws.iter_rows(min_row=2, values_only=True):
                    row_list = list(row)
                    cell_val = row_list[bl_col - 1] if bl_col <= len(row_list) else None
                    if cell_val and str(bl_no) in str(cell_val):
                        # Found matching row - collect containers below
                        for scan_row in ws.iter_rows(min_row=row[0].row + 1, values_only=True):
                            scan_list = list(scan_row)
                            cntr_val = scan_list[cntr_col - 1] if cntr_col <= len(scan_list) else None
                            type_val = scan_list[type_col - 1] if type_col and type_col <= len(scan_list) else ""
                            if cntr_val and str(cntr_val).strip():
                                if type_val and str(type_val).strip():
                                    results.append({
                                        "container_no": str(cntr_val).strip(),
                                        "container_type": str(type_val).strip() if type_val else "",
                                    })
                            else:
                                # Stop scanning when we hit an empty container cell
                                if cntr_val is None or str(cntr_val).strip() == "":
                                    break
                        break  # Only process first match

            if results:
                break

        wb.close()
        return results

    except Exception as e:
        logger.warning(f"舱单读取失败: {e}")
        try:
            wb.close()
        except Exception:
            pass
        return []
