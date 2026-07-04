"""
Rule V1 — Freight Invoice list 2026Fareast

Field mapping rules for extracting data from invoice Excel files
and mapping to the template columns.
"""

RULE_META = {
    "id": "rule_v1",
    "name": "Rule1",
    "description": "Freight Invoice list 2026Fareast — default extraction rules",
    "template_name": "Freight Invoice list 2026Fareast",
    "start_row": 5,
    "sort_column": "U",
}


# Each mapping: (source_location, target_column, parse_mode)
# parse_mode:
#   "value"       — direct cell value
#   "after_colon" — take text after the first colon
#   "before_v"    — take text before 'V' (case-insensitive)
#   "after_colon_before_comma" — after colon, then before comma
#   "concat_slash" — concatenate with "/"
#   "range_merge" — merge a range of cells with "/"
#   "after_colon_comma_after" — after colon, take text after comma
FIELD_MAPPINGS = [
    # Invoice I29 → D column, and also → Q column (same value)
    {"source": "I29", "target_col": "D", "mode": "value"},
    {"source": "I29", "target_col": "Q", "mode": "value"},
    # Invoice H7, after colon → E column
    {"source": "H7", "target_col": "E", "mode": "after_colon"},
    # Invoice B12, after colon → F column
    {"source": "B12", "target_col": "F", "mode": "after_colon"},
    # Invoice E10, before V → I column
    {"source": "E10", "target_col": "I", "mode": "before_v"},
    # Invoice B5, after colon, before comma → K column
    {"source": "B5", "target_col": "K", "mode": "after_colon_before_comma"},
    # Invoice B6, after colon, before comma → L column
    {"source": "B6", "target_col": "L", "mode": "after_colon_before_comma"},
    # Invoice B10, after colon → M column
    {"source": "B10", "target_col": "M", "mode": "after_colon"},
    # Invoice B6, after comma → O column
    {"source": "B6", "target_col": "O", "mode": "after_colon_comma_after"},
    # Invoice H7 (after colon) + "/" + C10 → P column
    {"source": "H7", "target_col": "P", "mode": "concat_slash", "source2": "C10", "source_mode": "after_colon"},
    # Invoice D10 → S column
    {"source": "D10", "target_col": "S", "mode": "value"},
    # Invoice F10:F28 → T column (merge with /)
    {"source": "F10:F28", "target_col": "T", "mode": "range_merge"},
    # Invoice H6, after colon → U column
    {"source": "H6", "target_col": "U", "mode": "after_colon"},
]


def extract_value(sheet, mapping):
    """Extract a value from an invoice sheet according to the mapping."""
    mode = mapping["mode"]
    source = mapping["source"]

    if mode == "value":
        cell = sheet[source]
        v = cell.value
        return _safe_str(v)

    elif mode == "after_colon":
        v = _get_cell_value(sheet, source)
        return _after_colon(v)

    elif mode == "before_v":
        v = _get_cell_value(sheet, source)
        return _before_v(v)

    elif mode == "after_colon_before_comma":
        v = _get_cell_value(sheet, source)
        return _after_colon_before_comma(v)

    elif mode == "after_colon_comma_after":
        v = _get_cell_value(sheet, source)
        return _after_colon_comma_after(v)

    elif mode == "concat_slash":
        v1 = _get_cell_value(sheet, source)
        v2 = _get_cell_value(sheet, mapping["source2"])
        # Apply optional source_mode transform to v1
        source_mode = mapping.get("source_mode")
        if source_mode == "after_colon":
            v1 = _after_colon(v1)
        elif source_mode == "after_colon_before_comma":
            v1 = _after_colon_before_comma(v1)
        return _concat_slash(v1, v2)

    elif mode == "range_merge":
        return _range_merge(sheet, source)

    return ""


def _get_cell_value(sheet, cell_ref):
    cell = sheet[cell_ref]
    v = cell.value
    return _safe_str(v)


def _safe_str(v):
    if v is None:
        return ""
    return str(v)


def _after_colon(v: str) -> str:
    """Return text after the first colon, stripped."""
    if ":" in v:
        return v.split(":", 1)[1].strip()
    return v.strip()


def _before_v(v: str) -> str:
    """Return text before 'V' (case-insensitive), stripped."""
    import re
    match = re.split(r'[Vv]', v, maxsplit=1)
    return match[0].strip()


def _after_colon_before_comma(v: str) -> str:
    """Return text after colon but before comma."""
    after = _after_colon(v)
    if "," in after:
        return after.split(",", 1)[0].strip()
    return after


def _after_colon_comma_after(v: str) -> str:
    """Return text after colon, then text after comma."""
    after = _after_colon(v)
    if "," in after:
        return after.split(",", 1)[1].strip()
    return after.strip()


def _concat_slash(v1: str, v2: str) -> str:
    """Concatenate two values with '/'."""
    return f"{v1.strip()}/{v2.strip()}"


def _range_merge(sheet, range_ref: str) -> str:
    """Merge a cell range with '/' separator, replacing newlines with '/'."""
    from openpyxl.utils import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    parts = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            v = sheet.cell(row=row, column=col).value
            if v is not None:
                text = str(v).replace("\n", "/").replace("\r", "")
                if text:
                    parts.append(text)
    return "/".join(parts)
