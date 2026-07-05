"""Main processor — orchestrates the entire invoice merging workflow."""

import json
import os
import subprocess
import tempfile
import time
from pathlib import Path

from utils.logger import logger
from utils.zip_helper import count_excel_in_zip
from core.validator import validate_template, validate_zip
from core.extractor import extract_all
from core.writer import write_to_template


def load_config() -> dict:
    """Load configuration from config.json."""
    config_path = Path(__file__).parent.parent / "config.json"
    defaults = {
        "output_name": "Freight Invoice list 2026Fareast_Output.xlsx",
        "open_after_finish": True,
        "auto_sort": True,
        "remember_last_path": True,
        "last_output_dir": "",
    }
    if config_path.exists():
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                return {**defaults, **json.load(f)}
        except Exception:
            return defaults
    return defaults


def get_output_path(output_dir: str, config: dict) -> str:
    """Generate output file path with auto-increment if exists."""
    base_name = config.get("output_name", "Freight Invoice list 2026Fareast_Output.xlsx")
    output_path = Path(output_dir) / base_name

    if not output_path.exists():
        return str(output_path)

    stem = output_path.stem
    suffix = output_path.suffix
    counter = 1
    while True:
        new_name = f"{stem}({counter}){suffix}"
        new_path = Path(output_dir) / new_name
        if not new_path.exists():
            return str(new_path)
        counter += 1


class ProcessingResult:
    def __init__(self):
        self.success_count = 0
        self.error_count = 0
        self.output_path = ""
        self.error_message = ""
        self.elapsed = 0.0


def open_file_location(file_path: str):
    """Open the folder containing the file and select it (Windows only)."""
    import platform
    if platform.system() == "Windows":
        try:
            os.startfile(os.path.dirname(file_path))
        except Exception:
            pass


def run_processing(
    template_path: str,
    zip_path: str,
    rule_id: str,
    output_dir: str = "",
    progress_callback=None,
) -> ProcessingResult:
    """
    Main processing workflow.
    Returns ProcessingResult with status, counts, output path, and elapsed time.
    """
    result = ProcessingResult()
    start_time = time.time()

    def progress(msg):
        if progress_callback:
            progress_callback(msg)

    config = load_config()

    # Step 1: Validate inputs
    progress("验证文件...")
    valid, msg = validate_template(template_path)
    if not valid:
        result.error_message = msg
        logger.error(msg)
        return result

    valid, msg = validate_zip(zip_path)
    if not valid:
        result.error_message = msg
        logger.error(msg)
        return result

    # Step 2: Determine output directory
    if not output_dir:
        output_dir = str(Path(template_path).parent)
    logger.info(f"输出目录：{output_dir}")

    # Step 3: Count files
    logger.info("读取ZIP...")
    try:
        file_count = count_excel_in_zip(Path(zip_path))
        logger.info(f"发现{file_count}个Excel")
    except Exception as e:
        result.error_message = f"ZIP损坏：{e}"
        logger.error(result.error_message)
        return result

    # Step 4: Extract data
    progress("解压ZIP...")
    extract_dir = Path(tempfile.mkdtemp(prefix="invoice_"))
    try:
        data, errors = extract_all(Path(zip_path), extract_dir, rule_id)
        result.success_count = len(data)
        result.error_count = errors
    except Exception as e:
        result.error_message = f"提取数据失败：{e}"
        logger.error(result.error_message)
        return result

    if not data:
        result.error_message = "未找到有效的Invoice数据"
        logger.error(result.error_message)
        return result

    # Step 5: Write to template
    output_path = get_output_path(output_dir, config)
    progress("写入模板...")
    try:
        write_to_template(
            template_path,
            output_path,
            data,
            rule_id,
            auto_sort=config.get("auto_sort", True),
        )
    except Exception as e:
        result.error_message = f"写入模板失败：{e}"
        logger.error(result.error_message)
        return result

    result.output_path = output_path
    result.elapsed = time.time() - start_time

    # Step 6: Open file location
    if config.get("open_after_finish", True):
        open_file_location(output_path)

    # Summary log
    logger.info(f"成功：{result.success_count}  失败：{result.error_count}  耗时：{result.elapsed:.1f}秒")
    return result
