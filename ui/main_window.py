"""Main window for Freight Invoice Merger Pro.

Contains two tabs:
  Tab 1: 发票整合 — Rule1 (月发票清单) and Rule2 (大发票)
  Tab 2: 单张发票 — single invoice generation from draft + manifests

Tab 1 contains the existing functionality untouched.
Tab 2 is entirely new for single invoice generation.
"""

from pathlib import Path
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QLabel, QPushButton, QMessageBox,
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QFont, QIcon

from ui.widgets import FileBrowseRow, OutputDirRow, RuleSelector, ProgressPanel, LogPanel
from core.processor import run_processing, load_config
from rules.registry import RuleRegistry
from utils.logger import log_signal
from tabs.single_invoice_tab import SingleInvoiceTab


class ProcessingThread(QThread):
    """Background thread for Tab1 (发票整合) processing."""

    progress_update = Signal(str)
    finished_signal = Signal(object)

    def __init__(self, template_path, zip_path, output_dir, rule_id,
                 sort_file_path="", parent=None):
        super().__init__(parent)
        self.template_path = template_path
        self.zip_path = zip_path
        self.output_dir = output_dir
        self.rule_id = rule_id
        self.sort_file_path = sort_file_path

    def run(self):
        def progress_callback(msg):
            self.progress_update.emit(msg)

        result = run_processing(
            self.template_path,
            self.zip_path,
            self.rule_id,
            output_dir=self.output_dir,
            sort_file_path=self.sort_file_path,
            progress_callback=progress_callback,
        )
        self.finished_signal.emit(result)


class Tab1MergeWidget(QWidget):
    """Widget for Tab 1 — 发票整合 (existing Rule1 + Rule2 functionality)."""

    def __init__(self, config: dict, parent=None):
        super().__init__(parent)
        self.config = config

        # Track state
        self.template_path = ""
        self.zip_path = ""
        self.sort_file_path = ""
        self.output_dir = ""
        self.current_rule_id = "rule_v1"
        self.is_processing = False

        self._setup_ui()
        log_signal.connect(self._on_log_message)
        self._load_rules()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 24, 32, 24)
        layout.setSpacing(16)

        # ── Header ──
        header = QWidget()
        header.setObjectName("card")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(20, 16, 20, 16)
        title_row = QHBoxLayout()
        title = QLabel("发票整合")
        title.setObjectName("titleLabel")
        title_row.addWidget(title)
        title_row.addStretch()
        header_layout.addLayout(title_row)
        layout.addWidget(header)

        # ── Rule selector ──
        all_rules = RuleRegistry.list_rules() or [
            {"id": "rule_v1", "name": "Rule1"},
            {"id": "rule_v2", "name": "Rule2"},
        ]
        self.rule_selector = RuleSelector(all_rules)
        self.rule_selector.rule_changed.connect(self._on_rule_changed)
        layout.addWidget(self.rule_selector)

        # ── Template row ──
        self.template_row = FileBrowseRow(
            label_text="模板",
            browse_mode="file",
            file_filter="Excel 文件 (*.xlsx *.xlsm *.xltx *.xls);;所有文件 (*)",
            placeholder="请选择模板文件...",
        )
        self.template_row.file_changed.connect(lambda p: setattr(self, 'template_path', p))
        layout.addWidget(self.template_row)

        # ── Sort file row (Rule2 only) ──
        self.sort_row = FileBrowseRow(
            label_text="排序模板（请选择 Freight Invoice list 2026Fareast_merged.xlsx）",
            browse_mode="file",
            file_filter="Excel 文件 (*.xlsx *.xlsm *.xltx);;所有文件 (*)",
            placeholder="用于确定大发票生成顺序...",
        )
        self.sort_row.file_changed.connect(lambda p: setattr(self, 'sort_file_path', p))
        self.sort_row.setVisible(False)
        layout.addWidget(self.sort_row)

        # ── ZIP row ──
        self.zip_row = FileBrowseRow(
            label_text="发票ZIP",
            browse_mode="file",
            file_filter="ZIP 压缩包 (*.zip);;所有文件 (*)",
            placeholder="请选择发票ZIP压缩包...",
        )
        self.zip_row.file_changed.connect(lambda p: setattr(self, 'zip_path', p))
        layout.addWidget(self.zip_row)

        # ── Output directory ──
        last_output = self.config.get("last_output_dir", "")
        self.output_dir_row = OutputDirRow(
            label_text="输出路径",
            default_dir=last_output,
            placeholder="选择输出文件保存目录...",
        )
        self.output_dir_row.dir_changed.connect(lambda p: setattr(self, 'output_dir', p))
        layout.addWidget(self.output_dir_row)

        # ── Progress ──
        self.progress_panel = ProgressPanel()
        layout.addWidget(self.progress_panel)

        # ── Start button ──
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.start_btn = QPushButton("开始整合")
        self.start_btn.setObjectName("startButton")
        self.start_btn.setMinimumWidth(200)
        self.start_btn.clicked.connect(self._on_start)
        btn_row.addWidget(self.start_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # ── Log panel ──
        self.log_panel = LogPanel()
        layout.addWidget(self.log_panel, 1)

    def _load_rules(self):
        rules = RuleRegistry.list_rules()
        if rules:
            self.current_rule_id = rules[0]["id"]
            self._on_rule_changed(rules[0]["id"])

    def _on_rule_changed(self, rule_id: str):
        self.current_rule_id = rule_id
        self.template_path = ""
        self.template_row.set_path("")
        self.zip_path = ""
        self.zip_row.set_path("")
        self.sort_file_path = ""
        self.sort_row.set_path("")
        self.progress_panel.reset()
        self.log_panel.clear_log()
        self.sort_row.setVisible(rule_id == "rule_v2")

    def _on_log_message(self, message: str):
        self.log_panel.append_log(message)

    def _on_start(self):
        if self.is_processing:
            return

        if not self.template_path:
            QMessageBox.warning(self, "提示", "请选择模板。")
            return
        if not self.zip_path:
            QMessageBox.warning(self, "提示", "请选择发票ZIP。")
            return
        if self.current_rule_id == "rule_v2" and not self.sort_file_path:
            QMessageBox.warning(self, "提示",
                "Rule2需要选择排序模板\n（Freight Invoice list 2026Fareast_merged.xlsx）")
            return

        output_dir = self.output_dir or str(Path(self.template_path).parent)
        self.config["last_output_dir"] = output_dir
        try:
            import json
            config_path = Path(__file__).parent.parent / "config.json"
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

        self.progress_panel.reset()
        self.log_panel.clear_log()
        self.is_processing = True
        self.start_btn.setEnabled(False)
        self.start_btn.setText("处理中...")

        self._thread = ProcessingThread(
            self.template_path, self.zip_path, output_dir,
            self.current_rule_id, sort_file_path=self.sort_file_path,
        )
        self._thread.progress_update.connect(lambda m: self.progress_panel.set_progress(50, m))
        self._thread.finished_signal.connect(self._on_done)
        self._thread.start()

    def _on_done(self, result):
        self.is_processing = False
        self.start_btn.setEnabled(True)
        self.start_btn.setText("开始整合")

        if result.error_message:
            self.progress_panel.set_progress(0, "处理失败")
            QMessageBox.critical(self, "错误", result.error_message)
            return

        self.progress_panel.set_progress(100, "处理完成 ✓")
        summary = (
            f"数据整合完成！\n\n"
            f"成功：{result.success_count}\n"
            f"失败：{result.error_count}\n"
            f"耗时：{result.elapsed:.1f}秒\n\n"
            f"输出文件：{result.output_path}"
        )
        QMessageBox.information(self, "完成", summary)


class MainWindow(QMainWindow):
    """Main window with tabbed interface."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Freight Invoice Merger Pro")
        self.setMinimumSize(1100, 760)
        self.resize(1100, 760)

        self.config = load_config()

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)

        # ── Header ──
        header = QWidget()
        header.setObjectName("card")
        hl = QVBoxLayout(header)
        hl.setContentsMargins(24, 12, 24, 12)
        tr = QHBoxLayout()
        t = QLabel("Freight Invoice Merger Pro")
        t.setObjectName("titleLabel")
        tr.addWidget(t)
        tr.addStretch()
        v = QLabel("Version 2.0")
        v.setObjectName("versionLabel")
        tr.addWidget(v)
        hl.addLayout(tr)
        layout.addWidget(header)

        # ── Tab Widget ──
        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.TabPosition.North)

        # Tab 1: Existing merge functionality
        self.tab1 = Tab1MergeWidget(self.config)
        self.tabs.addTab(self.tab1, "月发票清单/大发票")

        # Tab 2: Single invoice generation
        self.tab2 = SingleInvoiceTab(self.config)
        self.tabs.addTab(self.tab2, "单张发票")

        layout.addWidget(self.tabs, 1)
