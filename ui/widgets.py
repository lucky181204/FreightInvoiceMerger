"""Custom UI widgets for the Freight Invoice Merger."""

from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QVBoxLayout, QLabel, QPushButton,
    QLineEdit, QFileDialog, QProgressBar, QComboBox, QPlainTextEdit,
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont


class CardFrame(QFrame):
    """A rounded card container with shadow-like border."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setStyleSheet("")


class FileBrowseRow(QFrame):
    """Label + file path + browse button in one row."""

    file_changed = Signal(str)

    def __init__(self, label_text: str, browse_mode: str = "file",
                 file_filter: str = "", placeholder: str = "", parent=None):
        super().__init__(parent)
        self.browse_mode = browse_mode
        self.file_filter = file_filter
        self._file_path = ""

        self.setObjectName("card")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        # Section label
        label = QLabel(label_text)
        label.setObjectName("sectionLabel")
        layout.addWidget(label)

        # Path row
        path_row = QHBoxLayout()
        path_row.setSpacing(8)

        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText(placeholder)
        self.path_edit.setReadOnly(True)
        path_row.addWidget(self.path_edit, 1)

        self.browse_btn = QPushButton("浏览")
        self.browse_btn.setObjectName("browseButton")
        self.browse_btn.clicked.connect(self._on_browse)
        path_row.addWidget(self.browse_btn)

        layout.addLayout(path_row)

    def _on_browse(self):
        if self.browse_mode == "file":
            path, _ = QFileDialog.getOpenFileName(
                self, "选择文件", "", self.file_filter
            )
        else:
            path = QFileDialog.getExistingDirectory(self, "选择目录")

        if path:
            self._file_path = path
            self.path_edit.setText(path)
            self.file_changed.emit(path)

    def get_path(self) -> str:
        return self._file_path

    def set_path(self, path: str):
        self._file_path = path
        self.path_edit.setText(path)


class OutputDirRow(QFrame):
    """Label + folder path + browse button, remembers last path."""

    dir_changed = Signal(str)

    def __init__(self, label_text: str, default_dir: str = "",
                 placeholder: str = "", parent=None):
        super().__init__(parent)
        self._dir_path = default_dir

        self.setObjectName("card")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        label = QLabel(label_text)
        label.setObjectName("sectionLabel")
        layout.addWidget(label)

        path_row = QHBoxLayout()
        path_row.setSpacing(8)

        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText(placeholder)
        self.path_edit.setReadOnly(True)
        if default_dir:
            self.path_edit.setText(default_dir)
        path_row.addWidget(self.path_edit, 1)

        self.browse_btn = QPushButton("浏览")
        self.browse_btn.setObjectName("browseButton")
        self.browse_btn.clicked.connect(self._on_browse)
        path_row.addWidget(self.browse_btn)

        layout.addLayout(path_row)

    def _on_browse(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录", self._dir_path)
        if path:
            self._dir_path = path
            self.path_edit.setText(path)
            self.dir_changed.emit(path)

    def get_path(self) -> str:
        return self._dir_path

    def set_path(self, path: str):
        self._dir_path = path
        self.path_edit.setText(path)


class RuleSelector(QFrame):
    """Dropdown to select processing rule.
    Using QComboBox styled as a tab-like selection."""

    rule_changed = Signal(str)

    def __init__(self, rules: list[dict], parent=None):
        super().__init__(parent)
        self.setObjectName("card")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        label = QLabel("选择功能")
        label.setObjectName("sectionLabel")
        layout.addWidget(label)

        self.combo = QComboBox()
        display_names = {
            "rule_v1": "月发票清单 (Rule1)",
            "rule_v2": "大发票 (Rule2)",
        }
        for rule in rules:
            name = display_names.get(rule["id"], rule["name"])
            self.combo.addItem(name, rule["id"])
        self.combo.currentIndexChanged.connect(self._on_change)
        layout.addWidget(self.combo)

    def _on_change(self, idx):
        rule_id = self.combo.currentData()
        self.rule_changed.emit(rule_id)

    def current_rule_id(self) -> str:
        return self.combo.currentData()


class ProgressPanel(QFrame):
    """Progress bar with percentage and status text."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        # Percentage row
        pct_row = QHBoxLayout()
        self.percent_label = QLabel("0%")
        self.percent_label.setObjectName("progressPercent")
        pct_row.addWidget(self.percent_label)
        pct_row.addStretch()
        layout.addLayout(pct_row)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)

        # Status text
        self.status_label = QLabel("等待开始...")
        self.status_label.setObjectName("progressStatus")
        layout.addWidget(self.status_label)

    def set_progress(self, value: int, status: str = ""):
        self.progress_bar.setValue(value)
        self.percent_label.setText(f"{value}%")
        if status:
            self.status_label.setText(status)

    def reset(self):
        self.progress_bar.setValue(0)
        self.percent_label.setText("0%")
        self.status_label.setText("等待开始...")


class LogPanel(QFrame):
    """Log display panel with auto-scroll."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        label = QLabel("日志")
        label.setObjectName("sectionLabel")
        layout.addWidget(label)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        layout.addWidget(self.log_view, 1)

    def append_log(self, message: str):
        self.log_view.appendPlainText(message)
        # Auto-scroll to bottom
        scrollbar = self.log_view.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def clear_log(self):
        self.log_view.clear()
