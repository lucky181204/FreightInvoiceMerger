"""
Rule V1 — Freight Invoice list 2026Fareast

Field mapping rules for extracting data from invoice Excel files
and mapping to the template columns.
All string parsing uses centralized utils/str_parse.py.
"""

from utils.str_parse import after_colon, before_comma_after_colon, after_comma, before_v, replace_line_break

RULE_META = {
    "id": "rule_v1",
    "name": "Rule1",
    "description": "Freight Invoice list 2026Fareast — default extraction rules",
    "template_name": "Freight Invoice list 2026Fareast",
    "start_row": 5,
    "sort_column": "U",
}

# mode names map directly to utils/str_parse functions:
#   "value"                     — raw cell value
#   "after_colon"               — text after first colon
#   "before_comma_after_colon"  — after colon, then before first comma
#   "after_comma"               — after first comma
#   "before_v"                  — before V/v, case-insensitive
#   "concat_slash"              — after_colon(source1) + "/" + source2
#   "range_merge"               — merge cell range with "/"
FIELD_MAPPINGS = [
    # I29 → D and Q (same value, both columns)
    {"source": "I29",         "target_col": "D", "mode": "value"},
    {"source": "I29",         "target_col": "Q", "mode": "value"},
    # H7 after colon → E
    {"source": "H7",          "target_col": "E", "mode": "after_colon"},
    # B12 after colon → F
    {"source": "B12",         "target_col": "F", "mode": "after_colon"},
    # E10 before V → I
    {"source": "E10",         "target_col": "I", "mode": "before_v"},
    # B5 after colon, before comma → K
    {"source": "B5",          "target_col": "K", "mode": "before_comma_after_colon"},
    # B6 after colon, before comma → L
    {"source": "B6",          "target_col": "L", "mode": "before_comma_after_colon"},
    # B10 after colon → M
    {"source": "B10",         "target_col": "M", "mode": "after_colon"},
    # B6 after comma → O   (B6 = "Consignee: Los Angeles, USA" → "USA")
    {"source": "B6",          "target_col": "O", "mode": "after_comma"},
    # H7(after colon) + "/" + C10 → P
    {"source": "H7",          "target_col": "P", "mode": "concat_slash", "source2": "C10"},
    # D10 → S
    {"source": "D10",         "target_col": "S", "mode": "value"},
    # F10:F28 merge → T
    {"source": "F10:F28",     "target_col": "T", "mode": "range_merge"},
    # H6 after colon → U
    {"source": "H6",          "target_col": "U", "mode": "after_colon"},
]


def extract_value(sheet, mapping):
    """Extract a value from an invoice sheet according to the mapping.
    All string parsing delegates to utils/str_parse.py.

    IMPORTANT: For "value" mode we preserve the original Python type.
    For string modes we always return a string.
    """
    mode = mapping["mode"]

    if mode == "value":
        # Preserve original type (int/float for numbers, str for text)
        cell = sheet[mapping["source"]]
        v = cell.value
        if v is None:
            return ""
        return v  # Return as-is int, float, or str

    if mode == "after_colon":
        v = _get_cell(sheet, mapping["source"])
        return after_colon(v)

    if mode == "before_comma_after_colon":
        v = _get_cell(sheet, mapping["source"])
        return before_comma_after_colon(v)

    if mode == "after_comma":
        v = _get_cell(sheet, mapping["source"])
        return after_comma(v)

    if mode == "before_v":
        v = _get_cell(sheet, mapping["source"])
        return before_v(v)

    if mode == "concat_slash":
        v1 = _get_cell(sheet, mapping["source"])
        v2 = _get_cell(sheet, mapping["source2"])
        return f"{after_colon(v1)}/{v2.strip()}"

    if mode == "range_merge":
        return _range_merge(sheet, mapping["source"])

    return ""


def _get_cell(sheet, cell_ref: str) -> str:
    """Get cell value as string, returning empty string if None."""
    cell = sheet[cell_ref]
    v = cell.value
    return "" if v is None else str(v)


def _range_merge(sheet, range_ref: str) -> str:
    """Merge a cell range with '/' separator.
    Empty cells are skipped. Line breaks are replaced with '/'."""
    from openpyxl.utils import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    parts = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            v = sheet.cell(row=row, column=col).value
            if v is not None:
                text = str(v)
                if text:
                    parts.append(replace_line_break(text))
    return "/".join(parts)
