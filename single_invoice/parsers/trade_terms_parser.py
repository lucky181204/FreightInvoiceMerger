"""Trade terms manifest parser.

Mathes BL No (D10) in the trade terms manifest to find:
- Customer PO → C10
- Business No → H7
- Scheduled sailing → H5, H6
"""

import logging
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


def find_by_blno(manifest_path: str, bl_no: str) -> dict | None:
    """
    Search trade terms manifest for the given BL number.
    Returns dict with matched row data or None if not found.

    The manifest contains column headers. Searches for BL No in
    the '提单号' or 'BL NO' or similar column, then reads
    corresponding PO, business no, and sailing date columns.
    """
    if not bl_no:
        return None

    wb = load_workbook(manifest_path, data_only=True)
    try:
        for ws in wb.worksheets:
            # Find header row by looking for 'PO' or '提单号' or 'BL'
            headers = {}
            for row in ws.iter_rows(min_row=1, max_row=10, values_only=False):
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        val = cell.value.strip().upper().replace(" ", "").replace("　", "")
                        headers[cell.column] = val

            # Identify column positions
            po_col = None
            bl_col = None
            biz_col = None
            sail_col = None
            sail_col2 = None

            for col, val in headers.items():
                if val in ("PO", "P/O", "客户PO", "客户订单"):
                    po_col = col
                elif val in ("提单号", "BLNO", "BL NO", "B/L NO"):
                    bl_col = col
                elif val in ("业务编号", "业务号", "订单号", "ORDERNUMBER"):
                    biz_col = col
                elif val in ("船期", "ETD", "预计船期"):
                    sail_col = col
                elif val in ("船名航次", "VESSEL"):
                    sail_col2 = col

            if bl_col and po_col:
                # Search for matching BL No
                for row in ws.iter_rows(min_row=2, values_only=True):
                    row_list = list(row)
                    cell_val = row_list[bl_col - 1] if bl_col <= len(row_list) else None
                    if cell_val and str(bl_no) in str(cell_val):
                        result = {}
                        if po_col:
                            result["po"] = str(row_list[po_col - 1]) if po_col <= len(row_list) else ""
                        if biz_col:
                            result["business_no"] = str(row_list[biz_col - 1]) if biz_col <= len(row_list) else ""
                        if sail_col:
                            result["sailing_date"] = str(row_list[sail_col - 1]) if sail_col <= len(row_list) else ""
                        if sail_col2:
                            result["vessel"] = str(row_list[sail_col2 - 1]) if sail_col2 <= len(row_list) else ""
                        wb.close()
                        return result

        wb.close()
        return None
    except Exception as e:
        logger.warning(f"贸易条款清单读取失败: {e}")
        try:
            wb.close()
        except Exception:
            pass
        return None
