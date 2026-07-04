"""Excel utility functions."""

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def load_excel(path: str, data_only: bool = False):
    """Load an Excel workbook."""
    return load_workbook(path, data_only=data_only)


def get_cell_value(sheet, cell_ref: str):
    """Get cell value by reference like 'I29'."""
    return sheet[cell_ref].value


def get_cell_range(sheet, start_ref: str, end_ref: str):
    """Get a range of cell values and merge with '/'."""
    from openpyxl.utils import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(f"{start_ref}:{end_ref}")
    values = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            v = sheet.cell(row=row, column=col).value
            if v is not None:
                values.append(str(v))
    return values


def col_letter_to_index(col: str) -> int:
    """Convert column letter to 1-based index."""
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(col)
