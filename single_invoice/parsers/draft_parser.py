"""Draft parser for single invoice generation.

Extracts fields from BL Draft (.docx) file following defined rules.
Draft contains two tables with shipping document data.
"""

import logging
import re
from docx import Document

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
    "BUENAVENTURA": "COLOMBIA",
}


def get_first_word(text: str) -> str:
    """Return the first word of a text string."""
    if not text:
        return ""
    text = text.strip()
    parts = text.split()
    return parts[0] if parts else ""


def get_cell(table, row_idx: int, col_idx: int) -> str:
    """Safely get cell text from a docx table. row_idx/col_idx are 0-based."""
    try:
        return table.cell(row_idx, col_idx).text.strip()
    except Exception:
        return ""


def extract_draft_data(draft_path: str) -> dict:
    """
    Extract all fields from BL Draft .docx file.

    Draft Table 1 (booking info):
      Row 0: OCEAN VESSEL VOYAGE → col 1 = E10 Vessel/Voyage
      Row 1: BOOKING NO → col 1 = Booking ref
      Row 2: SHIPPER → col 1 = Shipper name (B5 = first word of Shipper)
      Row 5: PORT OF LOADING → col 1 = B10 Loading port; col 3 = B6 Port of Discharge
      Row 7: DESCRIPTION → col 1 = B11 danger class; B12 first line

    Draft Table 2 (container info):
      Row 0: EQU → Container count/type = B7
      Row 1+: Container details

    Draft Doc paragraph 1: D10 = Booking/Ref No
    """
    doc = Document(draft_path)

    # Find D10 from paragraph text or first table
    d10_value = ""
    for p in doc.paragraphs:
        txt = p.text.strip()
        # Look for booking ref / doc number patterns
        if txt and not txt.startswith("CONFIRMATION") and not txt.startswith("OCEAN"):
            d10_value = txt
            break

    tables = doc.tables
    result = {}

    if len(tables) >= 1:
        t1 = tables[0]
        logger.info(f"Table 1: {len(t1.rows)} rows x {len(t1.columns)} cols")

        # ── E10: Vessel/Voyage (Row 0, col 1) ──
        result["E10"] = get_cell(t1, 0, 1)

        # ── D10: Booking No (Row 1, col 1) ──
        bl_no = get_cell(t1, 1, 1)
        result["D10"] = bl_no if bl_no else d10_value

        # ── B5: First word of Shipper (Row 2, col 1) ──
        shipper = get_cell(t1, 2, 1)
        result["B5"] = get_first_word(shipper)

        # ── B6 Port of Loading: Row 5, col 1 ──
        result["B10"] = get_cell(t1, 5, 1)

        # ── B6: Port of Discharge: Row 5, col 3 ──
        discharge_port = get_cell(t1, 5, 3)
        result["B6_port"] = discharge_port

        # B6_country = mapped from destination port
        port_upper = discharge_port.upper()
        result["B6_country"] = DESTINATION_COUNTRY_MAP.get(port_upper, "")

        # ── B11/B12: Row 7 (description/danger class) ──
        desc_text = get_cell(t1, 7, 1)
        result["B12"] = desc_text.split("\n")[0].strip() if desc_text else ""
        result["B11"] = _extract_danger_class(desc_text)

    if len(tables) >= 2:
        t2 = tables[1]
        logger.info(f"Table 2: {len(t2.rows)} rows x {len(t2.columns)} cols")
        # ── B7: Container info (Row 0, col 1) ──
        result["B7"] = get_cell(t2, 0, 1)

    # Fallbacks
    result.setdefault("B5", "")
    result.setdefault("B6_port", "")
    result.setdefault("B6_country", "")
    result.setdefault("B7", "")
    result.setdefault("B10", "")
    result.setdefault("B11", "")
    result.setdefault("B12", "")
    result.setdefault("D10", d10_value)
    result.setdefault("E10", "")

    return result


def _extract_danger_class(text: str) -> str:
    """Extract dangerous goods class from text like 'CLASS:9' or 'Class 9' or '9'."""
    if not text:
        return ""
    # Look for CLASS:9 or similar
    m = re.search(r'(?:CLASS\s*[:：]\s*)(\d+(?:\.\d+)?)', text, re.IGNORECASE)
    if m:
        return m.group(1)
    # Just look for a digit right after 'CLASS'
    m = re.search(r'CLASS[^\d]*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
    if m:
        return m.group(1)
    return ""


def search_po_in_draft(draft_path: str, po: str) -> bool:
    """
    Search for PO string throughout entire Draft .docx.
    Returns True if found, False if not.
    """
    if not po:
        return False
    doc = Document(draft_path)
    # Search paragraphs
    for p in doc.paragraphs:
        if p.text and po in p.text:
            return True
    # Search tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if po in cell.text:
                    return True
    return False


def get_container_qty(text: str) -> int:
    """
    Extract container quantity from B7 text.
    Examples: '1*20GP' → 1, '10*40HQ' → 10, '5*20GP+3*40HQ' → 8
    """
    if not text:
        return 0
    total = 0
    for m in re.finditer(r'(\d+)\s*\*', text):
        total += int(m.group(1))
    return total if total > 0 else 0
