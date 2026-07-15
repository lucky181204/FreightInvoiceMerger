"""Hidden sheets removal service.

Only removes sheets that are hidden or veryHidden.
Never removes the last visible sheet.
Operates on sheet state, not sheet name.
"""

import logging

logger = logging.getLogger(__name__)


def remove_hidden_sheets(wb) -> list[str]:
    """
    Remove hidden and veryHidden sheets from an openpyxl workbook.

    Rules:
    - Only removes sheets with sheet_state in ('hidden', 'veryHidden')
    - Does NOT remove visible sheets regardless of name
    - Does NOT remove the last remaining visible sheet
    - Logs all operations

    Returns list of removed sheet names.
    """
    removed = []

    for sheet_name in wb.sheetnames[:]:
        ws = wb[sheet_name]

        if ws.sheet_state in ("hidden", "veryHidden"):
            # Don't remove if this would leave zero visible sheets
            visible_count = sum(
                1 for sn in wb.sheetnames
                if sn != sheet_name and wb[sn].sheet_state == "visible"
            )
            if visible_count == 0:
                logger.warning(f"跳过删除最后一个可见Sheet: {sheet_name}")
                continue

            state_label = "隐藏" if ws.sheet_state == "hidden" else "超隐藏"
            try:
                wb.remove(ws)
                removed.append(sheet_name)
                logger.info(f"已删除{state_label}工作表：{sheet_name}")
            except Exception as e:
                logger.warning(f"隐藏工作表删除失败：{sheet_name} — {e}")

    if not removed:
        logger.info("未发现隐藏工作表，跳过删除")

    return removed


def remove_hidden_sheets_xlwings(app, workbook, target_sheet_name=None) -> list[str]:
    """
    Remove hidden/veryHidden sheets using xlwings (Excel COM).

    Returns list of removed sheet names.
    """
    removed = []
    try:
        app.display_alerts = False
        for i in range(workbook.sheets.count, 0, -1):
            ws = workbook.sheets[i]
            if ws.api.Visible in (0, 2):  # xlSheetHidden=0, xlSheetVeryHidden=2
                if workbook.sheets.count <= 1:
                    continue
                ws_name = ws.name
                ws.delete()
                removed.append(ws_name)
    finally:
        try:
            app.display_alerts = True
        except Exception:
            pass

    return removed
