"""Writer — writes extracted invoice data into the template Excel file.

Only modifies cell.value in the mapped columns. Preserves all formatting,
fonts, borders, formulas, column widths, print settings, freeze panes,
filters, merged cells.
Data is sorted BEFORE writing to avoid touching any existing cells.
"""

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from utils.logger import logger
from rules.registry import RuleRegistry


def write_to_template(
    template_path: str,
    output_path: str,
    data: list[dict],
    rule_id: str,
    auto_sort: bool = True,
):
    """
    Write extracted invoice data into the template starting at the rule's start_row.
    Sorts data in Python BEFORE writing — never re-reads or overwrites template cells.
    Only modifies cell.value for the mapped columns.
    """
    rule = RuleRegistry.get_rule(rule_id)
    if not rule:
        raise ValueError(f"Rule not found: {rule_id}")

    start_row = rule.RULE_META.get("start_row", 5)
    sort_col = rule.RULE_META.get("sort_column", "U")

    # Sort data BEFORE writing if enabled
    if auto_sort and data:
        logger.info(f"按{sort_col}列排序...")
        data = _sort_data(data, sort_col)

    # Load template (writable mode, preserve everything)
    wb = load_workbook(template_path)
    ws = wb.active
    logger.info(f"写入模板，起始行：{start_row}")

    for i, row_data in enumerate(data):
        row_num = start_row + i
        for col_letter, value in row_data.items():
            try:
                col_idx = column_index_from_string(col_letter)
                ws.cell(row=row_num, column=col_idx).value = value
            except Exception as e:
                logger.warning(f"写入失败：列{col_letter} 行{row_num} — {e}")
        logger.info(f"写入第{row_num}行")

    wb.save(output_path)
    wb.close()
    logger.info(f"保存完成：{output_path}")
    return output_path


def _sort_data(data: list[dict], sort_column: str) -> list[dict]:
    """Sort data list by sort_column (numeric first, then text)."""
    def sort_key(row):
        val = row.get(sort_column, "")
        try:
            return (0, float(val))
        except (ValueError, TypeError):
            return (1, str(val or ""))

    sorted_data = sorted(data, key=sort_key)
    logger.info(f"排序完成")
    return sorted_data
