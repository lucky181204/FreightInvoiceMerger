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
SPEC_FILE = PROJECT_ROOT / "FreightInvoiceMerger.spec"


def clean_build():
    """Clean previous build artifacts."""
    for d in [BUILD_DIR, DIST_DIR]:
        if d.exists():
            shutil.rmtree(d)
    spec = PROJECT_ROOT / "FreightInvoiceMerger.spec"
    if spec.exists():
        spec.unlink()


def create_icon():
    """Create a default icon (PNG) for the app if no .ico exists."""
    icons_dir = PROJECT_ROOT / "resources"
    icons_dir.mkdir(parents=True, exist_ok=True)
    # The icon can be replaced with a proper .ico file
    # PyInstaller can use .png or .ico


def build():
    """Run PyInstaller to create the executable."""
    # Ensure resources
    create_icon()

    # Build command
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
    icon_paths = [
        PROJECT_ROOT / "resources" / "icon.ico",
        PROJECT_ROOT / "resources" / "icon.png",
    ]
    for ip in icon_paths:
        if ip.exists():
            cmd.extend(["--icon", str(ip)])
            break

    cmd.append(str(PROJECT_ROOT / "main.py"))

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=PROJECT_ROOT, capture_output=True, text=True)
    if result.returncode != 0:
        print("Build failed!")
        print(result.stdout)
        print(result.stderr)
        return False

    # Locate the built executable
    exe_name = "FreightInvoiceMerger.exe" if sys.platform == "win32" else "FreightInvoiceMerger"
    exe_path = DIST_DIR / exe_name
    if exe_path.exists():
        print(f"\n✅ Build successful: {exe_path}")
        return True
    else:
        print(f"\n❌ Executable not found at expected path: {exe_path}")
        return False


if __name__ == "__main__":
    clean_build()
    build()
