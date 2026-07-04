"""Freight Invoice Merger Pro — Main entry point.

Usage:
    python main.py
    # or after building: FreightInvoiceMerger.exe
"""

import sys
from pathlib import Path

# Ensure the project root is on sys.path
project_root = Path(__file__).parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from app import run


if __name__ == "__main__":
    run()
