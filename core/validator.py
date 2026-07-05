"""Input validation utilities."""

from pathlib import Path

# Allowed template extensions
_TEMPLATE_EXTS = {".xlsx", ".xlsm", ".xltx", ".xls"}


def validate_template(path: str) -> tuple[bool, str]:
    """Validate the template Excel file exists and is valid."""
    if not path:
        return False, "请选择模板。"
    p = Path(path)
    if not p.exists():
        return False, f"模板文件不存在：{p.name}"
    if p.suffix.lower() not in _TEMPLATE_EXTS:
        return False, f"模板格式错误（需要 .xlsx 或 .xls）：{p.name}"
    return True, ""


def validate_zip(path: str) -> tuple[bool, str]:
    """Validate the ZIP file exists and is accessible."""
    if not path:
        return False, "请选择ZIP文件。"
    p = Path(path)
    if not p.exists():
        return False, f"ZIP文件不存在：{p.name}"
    if p.suffix not in (".zip",):
        return False, f"文件格式错误（需要 .zip）：{p.name}"
    import zipfile
    try:
        with zipfile.ZipFile(p, "r") as zf:
            bad = zf.testzip()
            if bad:
                return False, f"ZIP损坏：{bad}"
    except zipfile.BadZipFile:
        return False, "ZIP损坏。"
    except Exception as e:
        return False, f"ZIP读取失败：{e}"
    return True, ""
