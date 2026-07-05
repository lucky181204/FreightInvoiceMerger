"""ZIP file helper utilities — supports both .xlsx and .xls."""

import zipfile
from pathlib import Path


_EXCEL_EXTS = {".xlsx", ".xlsm", ".xltx", ".xls"}


def extract_zip(zip_path: Path, extract_dir: Path) -> list[Path]:
    """Extract a ZIP file and return paths to all Excel files found,
    sorted by filename in natural ascending order."""
    extract_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(extract_dir)

    excel_files = sorted(
        (p for p in extract_dir.rglob("*")
         if p.suffix.lower() in _EXCEL_EXTS and not p.name.startswith("~$")),
        key=lambda p: p.name.lower(),  # case-insensitive sort
    )
    return excel_files


def count_excel_in_zip(zip_path: Path) -> int:
    """Count Excel entries inside a ZIP without extracting."""
    count = 0
    with zipfile.ZipFile(zip_path, "r") as zf:
        for name in zf.namelist():
            ext = Path(name).suffix.lower()
            if ext in _EXCEL_EXTS and not name.startswith("~$"):
                count += 1
    return count
