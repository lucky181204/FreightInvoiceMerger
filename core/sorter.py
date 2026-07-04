"""Sorter — sorts extracted invoice data by a specified column."""

from utils.logger import logger

# Column letter to index mapping for sorting
COL_LETTERS = {chr(65 + i): i for i in range(26)}  # A-Z


def sort_data(data: list[dict], sort_column: str = "U") -> list[dict]:
    """
    Sort invoice data by the specified column in ascending order.
    The column value is treated as text for comparison.
    """
    col_idx = COL_LETTERS.get(sort_column.upper(), -1)
    if col_idx < 0:
        logger.warning(f"Unknown sort column: {sort_column}, skipping sort")
        return data

    def sort_key(row):
        val = row.get(sort_column, "")
        # Try numeric sort if possible
        try:
            return (0, float(val))
        except (ValueError, TypeError):
            return (1, str(val))

    sorted_data = sorted(data, key=sort_key)
    logger.info(f"按{sort_column}列升序排序完成")
    return sorted_data
