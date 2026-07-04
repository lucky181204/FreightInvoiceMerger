"""Freight Invoice Merger Pro — Application entry point.

This module initializes the Qt application, applies the stylesheet,
and creates the main window.
"""

import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont, QIcon

from ui.main_window import MainWindow


def create_app() -> QApplication:
    """Create and configure the QApplication instance."""
    app = QApplication(sys.argv)
    app.setApplicationName("Freight Invoice Merger Pro")
    app.setApplicationVersion("1.0")
    app.setOrganizationName("FreightInvoiceMerger")

    # Load stylesheet
    qss_path = Path(__file__).parent / "ui" / "styles.qss"
    if qss_path.exists():
        with open(qss_path, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())

    # Set default font
    font = QFont("Microsoft YaHei UI", 10)
    app.setFont(font)

    return app


def run():
    app = create_app()
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()
