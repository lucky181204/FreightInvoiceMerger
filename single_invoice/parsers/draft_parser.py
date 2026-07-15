"""Draft parser for single invoice generation.

Extracts fields from BL Draft (.xlsx) workbook following defined rules.
"""

import logging
import re
from openpyxl import load_workbook

logger = logging.getLogger(__name__)

# ── Destination port → country mapping ──
DESTINATION_COUNTRY_MAP = {
    "SANTOS": "BRAZIL",
    "BUENOS AIRES": "ARGENTINA",
    "BUENAVENTURA": "COLOMBIA",
    "MANZANILLO": "MEXICO",
    "CARTAGENA": "COLOMBIA",
    "VALPARAISO": "CHILE",
    "LONG BEACH": "USA",
    "LOS ANGELES": "USA",
    "HOUSTON": "USA",
    "NEW YORK": "USA",
    "ROTTERDAM": "NETHERLANDS",
    "HAMBURG": "GERMANY",
    "ANTWERP": "BELGIUM",
    "FELIXSTOWE": "UK",
    "SHANGHAI": "CHINA",
    "NINGBO": "CHINA",
    "SHENZHEN": "CHINA",
    "QINGDAO": "CHINA",
}


def get_first_word(text: str) -> str:
    """Return the first word of a text string."""
    if not text:
        return ""
    text = text.strip()
    parts = text.split()
    return parts[0] if parts else ""


def extract_draft_data(draft_path: str) -> dict:
    """
    Extract all fields from BL Draft workbook.

    Draft has multiple tables on sheet 1.
    Table 1 (rows 1-8):
      Row 1 (1): E10 → Vessel/Voyage
      Row 2 (2): D10 → BL No
      Row 3 (3): B10 → Port of Loading (first line)
      Row 4 (4): (unused in extraction)
      Row 5 (5): (unused)
      Row 6 (6): B5 → First word; B6 → Port of Destination
      Row 7 (7): (unused)
      Row 8 (8): B11 → Dangerous goods class; B12 → Place of Receipt

    Table 2: B7 → Container info (first row, second column)

    Returns dict with keys:
      B5, B6_port, B6_country, B7, B10, B11, B12, D10, E10
    """
    wb = load_workbook(draft_path, data_only=True)
    ws = wb.active

    def cv(coord: str) -> str:
        """Get cell value as string."""
        v = ws[coord].value
        return "" if v is None else str(v)

    result = {}

    # Table 1 - Row 6: B5 = first word of B5
    b5_val = cv("B5")
    result["B5"] = get_first_word(b5_val)

    # B6 = text before comma from cell at row 6, col 4
    b6_val = cv("B6")
    if "," in b6_val:
        result["B6_port"] = b6_val.split(",")[0].strip()
    else:
        result["B6_port"] = b6_val.strip()

    # B6_country = mapped from destination port
    port_upper = result["B6_port"].upper()
    result["B6_country"] = DESTINATION_COUNTRY_MAP.get(port_upper, "")

    # B10 = Port of Loading (first line)
    result["B10"] = cv("B10").split("\n")[0].strip() if "\n" in cv("B10") else cv("B10").strip()

    # B11 = Dangerous goods class from row 8, col 2
    result["B11"] = _extract_danger_class(cv("B11"))

    # B12 = Place of Receipt (first line)
    result["B12"] = cv("B12").split("\n")[0].strip() if "\n" in cv("B12") else cv("B12").strip()

    # D10 = BL No
    result["D10"] = cv("D10").strip()

    # E10 = Vessel/Voyage
    result["E10"] = cv("E10").strip()

    # Table 2: B7 = first row, second column value
    result["B7"] = cv("B7").strip()

    wb.close()
    return result


def _extract_danger_class(text: str) -> str:
    """Extract dangerous goods class from text like 'Class 9' or '9'."""
    if not text:
        return ""
    m = re.search(r'(?:Class\s*)?(\d+(?:\.\d+)?)', text, re.IGNORECASE)
    return m.group(0) if m else text.strip()


def search_po_in_draft(draft_path: str, po: str) -> bool:
    """
    Search for PO string throughout entire Draft workbook.
    Returns True if found, False if not.
    """
    if not po:
        return False
    wb = load_workbook(draft_path, data_only=True)
    found = False
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell is not None and str(po) in str(cell):
                    found = True
                    break
            if found:
                break
        if found:
            break
    wb.close()
    return found


def get_container_qty(text: str) -> int:
    """
    Extract container quantity from B7 text.
    Examples: "10*40HQ" → 10, "5*20GP+3*40HQ" → 8, "1*40RF" → 1
    """
    if not text:
        return 0
    total = 0
    # Match patterns like "10*40HQ", "5*20GP"
    for m in re.finditer(r'(\d+)\s*\*', text):
        total += int(m.group(1))
    return total if total > 0 else 0
