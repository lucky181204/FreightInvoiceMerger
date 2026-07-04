"""Invoice extractor — iterates all invoice files and extracts structured data."""

from pathlib import Path

from core.parser import parse_invoice
from utils.excel import load_workbook_universal, get_active_sheet
from utils.logger import logger


def extract_all(zip_path: Path, extract_dir: Path, rule_id: str) -> tuple[list[dict], int]:
    """
    Extract data from all invoice Excel files (.xlsx, .xls, etc.) in a ZIP archive.
    Returns (data_list, error_count).
    Skips files that fail to parse and logs the error.
    """
    from utils.zip_helper import extract_zip

    excel_files = extract_zip(zip_path, extract_dir)
    logger.info(f"发现{len(excel_files)}个Excel文件")

    results = []
    errors = 0

    for idx, excel_path in enumerate(excel_files, 1):
        try:
            wb, fmt = load_workbook_universal(excel_path)
            sheet = get_active_sheet(wb, fmt)
            data = parse_invoice(sheet, rule_id)
            results.append(data)
            logger.info(f"解析第{idx}个Invoice [.{fmt}]：{excel_path.name}")
        except Exception as e:
            errors += 1
            logger.error(f"第{idx}个Invoice读取失败：{excel_path.name} — {e}")
            continue

    logger.info(f"解析完成：成功 {len(results)}，失败 {errors}")
    return results, errors
