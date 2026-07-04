#!/usr/bin/env python
"""Build script — packages FreightInvoiceMerger.exe using PyInstaller."""

import os
import sys
import shutil
import subprocess
from pathlib import Path


PROJECT_ROOT = Path(__file__).parent
BUILD_DIR = PROJECT_ROOT / "build"
DIST_DIR = PROJECT_ROOT / "dist"


def build():
    # Clean
    for d in [BUILD_DIR, DIST_DIR]:
        if d.exists():
            shutil.rmtree(d)
    for f in PROJECT_ROOT.glob("*.spec"):
        f.unlink()

    # Build command — note: NO --collect-all PySide6 (pulls QtWebEngine = 400MB+ crash)
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onefile",
        "--name", "FreightInvoiceMerger",
        "--add-data", f"resources{os.pathsep}resources",
        "--add-data", f"config.json{os.pathsep}.",
        "--exclude-module", "xlwings",
        "--hidden-import", "xlrd",
        "--hidden-import", "rules.rule_v1",
        "--hidden-import", "openpyxl",
        "--hidden-import", "openpyxl.cell._writer",
        "--hidden-import", "PySide6.QtCore",
        "--hidden-import", "PySide6.QtWidgets",
        "--hidden-import", "PySide6.QtGui",
        str(PROJECT_ROOT / "main.py"),
    ]

    print("Building FreightInvoiceMerger.exe ...", flush=True)
    result = subprocess.run(cmd, cwd=PROJECT_ROOT, capture_output=True, text=True)

    # Always print full output so GitHub log shows it
    if result.stdout:
        print(result.stdout, flush=True)
    if result.stderr:
        print(result.stderr, flush=True)

    if result.returncode != 0:
        print(f"\n❌ PyInstaller exit code: {result.returncode}", flush=True)
        return False

    exe = DIST_DIR / "FreightInvoiceMerger.exe"
    if not exe.exists():
        print(f"\n❌ Executable not found at {exe}", flush=True)
        return False

    size_mb = exe.stat().st_size / (1024 * 1024)
    print(f"\n✅ Build successful: {exe} ({size_mb:.1f} MB)", flush=True)
    return True


if __name__ == "__main__":
    sys.exit(0 if build() else 1)
