"""Manifest (舱单/container data) parser.

Dynamically finds container data by scanning ALL worksheets for
container number (箱号/柜号) and container type (箱型) headers.

Supports:
- Any workbook name, any sheet name, any header position
- BL No in fixed cells (e.g. SI Form header area) or per-row columns
- Numeric container type combining
- Container-like row filtering
"""

import logging
import re
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


def normalize_header(value) -> str:
    text = str(value or "")
    text = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()


def is_container_no_header(value) -> bool:
    text = normalize_header(value)
    return "箱号" in text or "柜号" in text or "container no" in text or "container number" in text or "集装箱号" in text


def is_container_type_header(value) -> bool:
    text = normalize_header(value)
    return "箱型" in text or "container type" in text or "container size" in text


def is_bl_no_header(value) -> bool:
    text = normalize_header(value)
    return ("提单号" in text or "主提单号" in text or
            "hbl no" in text or "b/l no" in text or "bl no" in text or
            "booking no" in text)


# Known container number prefixes (common patterns)
CONTAINER_PREFIXES = ('MR', 'EG', 'EM', 'BS', 'CA', 'TC', 'TX', 'MS',
                      'TL', 'FD', 'EI', 'EH', 'HA', 'MI', 'CI', 'TR',
                      'TT', 'OC', 'CM', 'CS', 'GC', 'TGH', 'TLL')


def _is_likely_container(text: str) -> bool:
    """Check if text looks like a container number (has letters + digits)."""
    if any(kw in text.upper() for kw in
           ['PAYMENT', 'FREIGHT', 'CHARGE', 'OCEAN', 'ORIGIN',
            'DESTINATION', 'SPECIAL', 'REMARK', 'TOTAL']):
        return False
    if text.upper().startswith(CONTAINER_PREFIXES):
        return True
    return bool(re.search(r'[A-Z]{2,}', text.upper()) and re.search(r'\d{4,}', text))


def find_containers_by_blno(manifest_path: str, bl_no: str, progress_callback=None) -> list[dict]:
    if not bl_no:
        return []

    def log(msg):
        logger.info(msg)
        if progress_callback:
            progress_callback(msg)

    log(f"读取舱单文件: {manifest_path}")
    wb = load_workbook(manifest_path, data_only=True)
    candidates = []

    for ws in wb.worksheets:
        scan = _scan_sheet(ws)
        if scan:
            scan["sheet_name"] = ws.title
            candidates.append(scan)

    if not candidates:
        log("未找到箱号或柜号列，请检查文件。")
        wb.close()
        return []

    best = _select_best_sheet(candidates)
    if len(candidates) > 1:
        log(f"找到多个候选工作表，已选择: {best['sheet_name']}")

    ws = wb[best["sheet_name"]]
    cntr_col = best["cntr_col"]
    type_col = best.get("type_col")
    bl_col = best.get("bl_col")
    data_start = best["data_start"]

    log(f"数据工作表: {best['sheet_name']}")
    log(f"箱号标题: {_col_letter(cntr_col)}{best['cntr_row']}")
    if type_col:
        log(f"箱型标题: {_col_letter(type_col)}{best['type_row']}")
    if bl_col:
        log(f"提单号列: {_col_letter(bl_col)}")

    # BL fixed value detection
    bl_fixed = None
    if bl_col:
        br = best.get("bl_row", 1)
        for sr in range(max(1, br - 1), br + 3):
            for sc in range(max(1, bl_col - 1), bl_col + 3):
                if sc != bl_col:
                    cv = ws.cell(row=sr, column=sc).value
                    if cv and str(cv).strip() == bl_no:
                        bl_fixed = bl_no
                        break
            if bl_fixed:
                break

    results = []
    in_matching = (not bl_col) or bool(bl_fixed)
    blank_count = 0
    max_row = ws.max_row

    for r in range(data_start, max_row + 1):
        cv = ws.cell(row=r, column=cntr_col).value
        if cv is None or str(cv).strip() == "":
            blank_count += 1
            if blank_count >= 8:
                break
            continue
        blank_count = 0

        cs = str(cv).strip()
        if cs.upper() in ("TOTAL", "合计", "总计"):
            break
        if not _is_likely_container(cs):
            continue

        if bl_col and not bl_fixed:
            bv = ws.cell(row=r, column=bl_col).value
            if bv and str(bv).strip():
                in_matching = (str(bv).strip() == bl_no)

        if not in_matching:
            continue

        tv = ""
        if type_col:
            tv_raw = ws.cell(row=r, column=type_col).value
            next_tv = ws.cell(row=r, column=type_col + 1).value if type_col + 1 <= ws.max_column else None
            tv = _build_container_type(tv_raw, next_tv)

        results.append({"container_no": cs, "container_type": tv})

    wb.close()
    logger.info(f"提取{len(results)}个集装箱")
    return results


def _scan_sheet(ws) -> dict | None:
    max_row = ws.max_row
    max_col = ws.max_column
    headers = []

    for r in range(1, min(max_row + 1, 200)):
        for c in range(1, min(max_col + 1, 100)):
            v = ws.cell(row=r, column=c).value
            if v is None or not isinstance(v, str):
                continue
            if is_container_no_header(v):
                headers.append({"row": r, "col": c, "type": "cntr"})
            if is_container_type_header(v):
                headers.append({"row": r, "col": c, "type": "type"})
            if is_bl_no_header(v):
                headers.append({"row": r, "col": c, "type": "bl"})

    cntr_h = [h for h in headers if h["type"] == "cntr"]
    type_h = [h for h in headers if h["type"] == "type"]
    bl_h = [h for h in headers if h["type"] == "bl"]

    if not cntr_h:
        return None

    best_cntr = cntr_h[0]
    best_type = None
    best_bl = bl_h[0] if bl_h else None

    for th in type_h:
        if abs(th["row"] - best_cntr["row"]) <= 2:
            best_type = th
            break
    if not best_type and type_h:
        best_type = type_h[0]

    data_start = best_cntr["row"] + 1
    if best_type:
        data_start = max(data_start, best_type["row"] + 1)

    return {
        "cntr_col": best_cntr["col"], "cntr_row": best_cntr["row"],
        "type_col": best_type["col"] if best_type else None,
        "type_row": best_type["row"] if best_type else None,
        "bl_col": best_bl["col"] if best_bl else None,
        "bl_row": best_bl["row"] if best_bl else None,
        "data_start": data_start,
    }


def _select_best_sheet(candidates: list) -> dict:
    complete = [c for c in candidates if c["type_col"]]
    if complete:
        candidates = complete
    candidates.sort(key=lambda c: c.get("data_start", 0))
    return candidates[0]


def _col_letter(idx: int) -> str:
    r = ""
    while idx > 0:
        idx -= 1
        r = chr(ord("A") + idx % 26) + r
        idx //= 26
    return r


def _normalize_excel_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    t = str(value).strip()
    if t.endswith(".0"):
        try:
            return str(int(float(t)))
        except ValueError:
            pass
    return t


def _build_container_type(type_value, next_value=None) -> str:
    raw = _normalize_excel_value(type_value)
    raw_next = _normalize_excel_value(next_value)
    if not raw:
        return ""
    if re.fullmatch(r"\d+", raw):
        return f"{raw}{raw_next}"
    return raw
