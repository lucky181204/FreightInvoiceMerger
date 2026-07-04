"""Main processor — orchestrates the entire invoice merging workflow."""

import json
import os
import tempfile
from pathlib import Path
from datetime import datetime

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
    }
    if config_path.exists():
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                return {**defaults, **json.load(f)}
        except Exception:
            return defaults
    return defaults


def get_output_path(template_path: str, config: dict) -> str:
    """Generate output file path with auto-increment if exists."""
    template_dir = Path(template_path).parent
    base_name = config.get("output_name", "Freight Invoice list 2026Fareast_Output.xlsx")
    output_path = template_dir / base_name

    if not output_path.exists():
        return str(output_path)

    stem = output_path.stem
    suffix = output_path.suffix
    counter = 1
    while True:
        new_name = f"{stem}({counter}){suffix}"
        new_path = template_dir / new_name
        if not new_path.exists():
            return str(new_path)
        counter += 1


class ProcessingResult:
    def __init__(self):
        self.success_count = 0
        self.error_count = 0
        self.output_path = ""
        self.error_message = ""


def run_processing(
    template_path: str,
    zip_path: str,
    rule_id: str,
    progress_callback=None,
) -> ProcessingResult:
    """
    Main processing workflow.
    Returns ProcessingResult with status, counts, and output path.
    """
    result = ProcessingResult()

    def progress(msg):
        if progress_callback:
            progress_callback(msg)

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

    # Step 2: Count files
    logger.info("读取ZIP...")
    try:
        file_count = count_excel_in_zip(Path(zip_path))
        logger.info(f"发现{file_count}个Excel")
    except Exception as e:
        result.error_message = f"ZIP损坏：{e}"
        logger.error(result.error_message)
        return result

    # Step 3: Extract data
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

    # Step 4: Write to template
    config = load_config()
    output_path = get_output_path(template_path, config)

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

    # Step 5: Open result
    if config.get("open_after_finish", True):
        try:
            os.startfile(output_path)
        except AttributeError:
            import subprocess
            try:
                subprocess.Popen(["xdg-open", output_path])
            except Exception:
                pass  # Windows handled by os.startfile

    # Summary
    logger.info("完成")
    return result
