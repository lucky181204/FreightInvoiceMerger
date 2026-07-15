#!/usr/bin/env python
"""Build script — packages FreightInvoiceMerger.exe using PyInstaller."""

import sys
import shutil
import subprocess
from pathlib import Path


PROJECT_ROOT = Path(__file__).parent
BUILD_DIR = PROJECT_ROOT / "build"
DIST_DIR = PROJECT_ROOT / "dist"


def build():
    # Clean previous builds
    for d in [BUILD_DIR, DIST_DIR]:
        if d.exists():
            shutil.rmtree(d)
    for f in PROJECT_ROOT.glob("*.spec"):
        f.unlink()

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onefile",
        "--name", "FreightInvoiceMerger",
        "--add-data", f"resources{os.pathsep}resources",
        "--add-data", f"config.json{os.pathsep}.",
        "--exclude-module", "matplotlib",
        "--exclude-module", "PIL",
        "--hidden-import", "xlrd",
        "--hidden-import", "xlwt",
        "--hidden-import", "xlutils",
        "--hidden-import", "docx",
        "--hidden-import", "rules.rule_v1",
        "--hidden-import", "rules.rule_v2",
        "--hidden-import", "single_invoice",
        "--hidden-import", "single_invoice.parsers",
        "--hidden-import", "single_invoice.services",
        "--hidden-import", "single_invoice.parsers.draft_parser",
        "--hidden-import", "single_invoice.parsers.manifest_parser",
        "--hidden-import", "single_invoice.parsers.trade_terms_parser",
        "--hidden-import", "single_invoice.services.hidden_sheet_service",
        "--hidden-import", "single_invoice.services.invoice_merger",
        "--hidden-import", "tabs.single_invoice_tab",
        "--hidden-import", "openpyxl",
        "--hidden-import", "openpyxl.cell._writer",
        "--hidden-import", "PySide6.QtCore",
        "--hidden-import", "PySide6.QtWidgets",
        "--hidden-import", "PySide6.QtGui",
        str(PROJECT_ROOT / "main.py"),
    ]

    result = subprocess.run(cmd, cwd=PROJECT_ROOT)
    return result.returncode == 0


if __name__ == "__main__":
    ok = build()
    if ok:
        exe = DIST_DIR / "FreightInvoiceMerger.exe"
        size_mb = exe.stat().st_size / (1024 * 1024)
        print(f"✅ Build successful: {exe} ({size_mb:.1f} MB)", flush=True)
        sys.exit(0)
    else:
        print("❌ Build failed", flush=True)
        sys.exit(1)
