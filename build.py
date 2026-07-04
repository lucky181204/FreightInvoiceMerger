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


def clean_build():
    """Clean previous build artifacts."""
    for d in [BUILD_DIR, DIST_DIR]:
        if d.exists():
            shutil.rmtree(d)
    spec = PROJECT_ROOT / "FreightInvoiceMerger.spec"
    if spec.exists():
        spec.unlink()


def build():
    """Run PyInstaller to create the executable."""
    clean_build()

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
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
        "--collect-all", "openpyxl",
        "--collect-all", "PySide6",
    ]

    # Add icon if available
    for icon_path in [PROJECT_ROOT / "resources" / "icon.ico",
                      PROJECT_ROOT / "resources" / "icon.png"]:
        if icon_path.exists():
            cmd.extend(["--icon", str(icon_path)])
            break

    cmd.append(str(PROJECT_ROOT / "main.py"))

    print(f"Running: {' '.join(cmd)}", flush=True)
    result = subprocess.run(cmd, cwd=PROJECT_ROOT, capture_output=False, text=True)
    return result.returncode == 0


if __name__ == "__main__":
    success = build()
    if success:
        exe = DIST_DIR / "FreightInvoiceMerger.exe"
        print(f"\n✅ Build successful: {exe}", flush=True)
        sys.exit(0)
    else:
        print("\n❌ Build failed", flush=True)
        sys.exit(1)
