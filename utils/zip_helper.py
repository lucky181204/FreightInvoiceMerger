"""ZIP file helper utilities."""

import zipfile
from pathlib import Path


def extract_zip(zip_path: Path, extract_dir: Path) -> list[Path]:
    """Extract a ZIP file and return paths to all .xlsx files found."""
    extract_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(extract_dir)

    excel_files = sorted(
        p for p in extract_dir.rglob("*.xlsx") if not p.name.startswith("~$")
    )
    return excel_files


def count_excel_in_zip(zip_path: Path) -> int:
    """Count .xlsx entries inside a ZIP without extracting."""
    count = 0
    with zipfile.ZipFile(zip_path, "r") as zf:
        for name in zf.namelist():
            if name.endswith(".xlsx") and not name.startswith("~$"):
                count += 1
    return count
