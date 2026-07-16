"""Draft parser for single invoice generation.

Extracts fields from BL Draft (.docx) file following defined rules.
Draft contains two tables with shipping document data.

Expected template cells and their format:
  B5  = "Loading Port:{port},{country}"        (e.g. "Loading Port:SHANGHAI,CHINA")
  B6  = "Destination: {port},{country}"          (e.g. "Destination: BUENAVENTURA,COLOMBIA")
  B7  = "Shipment of: {qty}x{type}"             (e.g. "Shipment of: 1*40HQ")
  H4  = "Invoice date:{date}"                   (e.g. "Invoice date:2026/7/15")
  H5  = "ETD:{date}"                            (e.g. "ETD:2026/6/21")
  H6  = "ATD:{date}"                            (e.g. "ATD:2026/6/21")
  H7  = "Invoice #:{business_no}"               (e.g. "Invoice #:SE26060373")
  B10 = "PRODUCT: {shipper_name}"               (e.g. "PRODUCT: HUNAN HAILI CHEMICAL INDUSTRY CO., LTD.")
  B11 = "CLASS: {danger_class}"                 (e.g. "CLASS: 6.1 HAZ")
  B12 = "Shipper: {goods_description}"           (e.g. "Shipper: METHOMYL TECH")
  C10 = "PO{po}"                                (e.g. "PO4500019465")
  D10 = "{booking_no}"                          (e.g. "EGLV142601959259")
  E10 = "{vessel_voyage}"                       (e.g. "EVER LISSOME V.0788-074E")
  F10+ = container_no
  G10+ = container_type
  H10+ = "USD"
  B30 = TOTAL formula
"""

import logging
import re
from docx import Document

logger = logging.getLogger(__name__)


def get_cell(table, row_idx: int, col_idx: int) -> str:
    """Safely get cell text from a docx table. row_idx/col_idx are 0-based."""
    try:
        return table.cell(row_idx, col_idx).text.strip()
    except Exception:
        return ""


def extract_draft_data(draft_path: str) -> dict:
    """
    Extract all fields from BL Draft .docx file.

    Draft Table 1 (booking info) — 9 rows x 4 cols:
      Row 0: OCEAN VESSEL VOYAGE | {vessel}
      Row 1: BOOKING NO. | {booking_no}
      Row 2: SHIPPER: | {shipper_name}
      Row 5: PORT OF LOADING: | {load_port} | PORT OF DISCHARGE: | {discharge_port}
      Row 7: N/M | {goods_desc}
      Row 8: SHIPPING TERMS: | {terms} | Bill of lading type: | {bl_type}

    Draft Table 2 (container info):
      Row 0: EQU: | {container_info}  (e.g. "1*40HQ")
      Row 1: Header
      Row 2+: Container data

    Returns dict with keys matching template cells.
    """
    doc = Document(draft_path)
    tables = doc.tables
    result = {}

    if len(tables) >= 1:
        t1 = tables[0]

        # E10: Vessel/Voyage (Row 0, col 1)
        result["E10"] = get_cell(t1, 0, 1)

        # D10: Booking No (Row 1, col 1)
        result["D10"] = get_cell(t1, 1, 1)

        # Shipper name (Row 2, col 1) - used for PRODUCT field (B5 is the full org name)
        shipper_name = get_cell(t1, 2, 1)
        # Clean up multi-line: take first meaningful part
        shipper_clean = shipper_name.split("\n")[0].strip().rstrip(",")
        result["shipper"] = shipper_clean

        # First word of shipper for B5 prefix parsing
        result["B5_first"] = shipper_clean.split()[0] if shipper_clean else ""

        # B10: Loading Port (Row 5, col 1)
        result["B10_loading_port"] = get_cell(t1, 5, 1)

        # B6: Discharge Port (Row 5, col 3)
        result["B6_discharge_port"] = get_cell(t1, 5, 3)

        # B12: Goods description (Row 7, col 1) - first line
        desc_text = get_cell(t1, 7, 1)
        desc_first_line = desc_text.split("\n")[0].strip() if desc_text else ""
        result["B12_goods_desc"] = desc_first_line

        # B11: Dangerous goods class extracted from description
        result["B11_danger_class"] = _extract_danger_class_full(desc_text)

        # B7: Container info from Table 2
    else:
        result["E10"] = ""
        result["D10"] = ""
        result["shipper"] = ""
        result["B5_first"] = ""
        result["B10_loading_port"] = ""
        result["B6_discharge_port"] = ""
        result["B12_goods_desc"] = ""
        result["B11_danger_class"] = ""

    if len(tables) >= 2:
        t2 = tables[1]
        # Row 0, col 1 = container info (e.g. "1*40HQ")
        result["B7_container"] = get_cell(t2, 0, 1)
    else:
        result["B7_container"] = ""

    return result


def _extract_danger_class_full(text: str) -> str:
    """
    Extract cargo class from draft description cell text.

    Logic:
    1. Split text by lines, only consider non-empty lines
    2. Only inspect lines containing CLASS or UN NO / UN NUMBER keywords
    3. Find lines with exactly 2 numbers where:
       - One is a 4-digit integer (UN number)
       - The other is NOT 4 digits (the danger class: 9, 6.1, etc.)
    4. Return "number HAZ" or "GENERAL CARGO"

    Examples:
      "CLASS:9 UN NO:3077"           -> "9 HAZ"
      "CLASS:6.1 UN NO:2757"         -> "6.1 HAZ"
      "METHOMYL TECH\n400 DRUMS"     -> "GENERAL CARGO"
      ""                              -> "GENERAL CARGO"
    """
    if not text:
        return "GENERAL CARGO"

    lines = [line.strip() for line in text.splitlines() if line.strip()]

    candidate_lines = [
        line for line in lines
        if re.search(r'\bCLASS\b|UN\s*NO\.?|UN\s*NUMBER', line, re.IGNORECASE)
    ]

    for line in candidate_lines:
        numbers = re.findall(r'(?<![\d.])\d+(?:\.\d+)?(?![\d.])', line)

        four_digit = [n for n in numbers if re.fullmatch(r'\d{4}', n)]
        other = [n for n in numbers if not re.fullmatch(r'\d{4}', n)]

        if len(numbers) == 2 and len(four_digit) == 1 and len(other) == 1:
            return f"{other[0]} HAZ"

    return "GENERAL CARGO"


def search_po_in_draft(draft_path: str, po: str) -> bool:
    """Search for PO string throughout entire Draft .docx."""
    if not po:
        return False
    doc = Document(draft_path)
    for p in doc.paragraphs:
        if p.text and po in p.text:
            return True
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if po in cell.text:
                    return True
    return False


def get_container_qty(text: str) -> int:
    """Extract container quantity from text like '1*40HQ' → 1."""
    if not text:
        return 0
    total = 0
    for m in re.finditer(r'(\d+)\s*\*', text):
        total += int(m.group(1))
    return total if total > 0 else 0
