"""Main window for Freight Invoice Merger Pro."""

from pathlib import Path
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QMessageBox,
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QMetaObject
from PySide6.QtGui import QFont, QIcon

from ui.widgets import FileBrowseRow, OutputDirRow, RuleSelector, ProgressPanel, LogPanel
from core.processor import run_processing, load_config
from rules.registry import RuleRegistry
from utils.logger import log_signal


class ProcessingThread(QThread):
    """Background thread for processing to keep UI responsive."""

    progress_update = Signal(str)
    finished_signal = Signal(object)

    def __init__(self, template_path, zip_path, output_dir, rule_id, parent=None):
        super().__init__(parent)
        self.template_path = template_path
        self.zip_path = zip_path
        self.output_dir = output_dir
        self.rule_id = rule_id

    def run(self):
        def progress_callback(msg):
            self.progress_update.emit(msg)

        result = run_processing(
            self.template_path,
            self.zip_path,
            self.rule_id,
            output_dir=self.output_dir,
            progress_callback=progress_callback,
        )
        self.finished_signal.emit(result)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Freight Invoice Merger Pro  v1.0")
        self.setMinimumSize(1100, 760)
        self.resize(1100, 760)

        # Load config
        self.config = load_config()

        # Track state
        self.template_path = ""
        self.zip_path = ""
        self.output_dir = ""
        self.current_rule_id = "rule_v1"
        self.is_processing = False

        # Setup UI
        self._setup_ui()

        # Connect log signal
        log_signal.connect(self._on_log_message)

        # Load rules
        self._load_rules()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(32, 24, 32, 24)
        main_layout.setSpacing(16)

        # ── Header ──
        header = QWidget()
        header.setObjectName("card")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(20, 16, 20, 16)

        title_row = QHBoxLayout()
        title = QLabel("Freight Invoice Merger Pro")
        title.setObjectName("titleLabel")
        title_row.addWidget(title)
        title_row.addStretch()
        ver = QLabel("Version 1.0")
        ver.setObjectName("versionLabel")
        title_row.addWidget(ver)
        header_layout.addLayout(title_row)
        main_layout.addWidget(header)

        # ── Template row ──
        self.template_row = FileBrowseRow(
            label_text="模板",
            browse_mode="file",
            file_filter="Excel 模板 (*.xlsx *.xlsm *.xltx);;所有文件 (*)",
            placeholder="请选择模板文件...",
        )
        self.template_row.file_changed.connect(self._on_template_changed)
        main_layout.addWidget(self.template_row)

        # ── ZIP row ──
        self.zip_row = FileBrowseRow(
            label_text="ZIP",
            browse_mode="file",
            file_filter="ZIP 压缩包 (*.zip);;所有文件 (*)",
            placeholder="请选择发票ZIP压缩包...",
        )
        self.zip_row.file_changed.connect(self._on_zip_changed)
        main_layout.addWidget(self.zip_row)

        # ── Output directory row ──
        last_output = self.config.get("last_output_dir", "")
        self.output_dir_row = OutputDirRow(
            label_text="输出路径",
            default_dir=last_output,
            placeholder="选择输出文件保存目录...",
        )
        self.output_dir_row.dir_changed.connect(self._on_output_dir_changed)
        main_layout.addWidget(self.output_dir_row)

        # ── Rule selector ──
        all_rules = RuleRegistry.list_rules() or [{"id": "rule_v1", "name": "Rule1"}]
        self.rule_selector = RuleSelector(all_rules)
        self.rule_selector.rule_changed.connect(self._on_rule_changed)
        main_layout.addWidget(self.rule_selector)

        # ── Progress panel ──
        self.progress_panel = ProgressPanel()
        main_layout.addWidget(self.progress_panel)

        # ── Start button ──
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.start_btn = QPushButton("开始整合")
        self.start_btn.setObjectName("startButton")
        self.start_btn.setMinimumWidth(200)
        self.start_btn.clicked.connect(self._on_start)
        btn_row.addWidget(self.start_btn)
        btn_row.addStretch()
        main_layout.addLayout(btn_row)

        # ── Log panel ──
        self.log_panel = LogPanel()
        main_layout.addWidget(self.log_panel, 1)

    def _load_rules(self):
        """Discover and load all rule plugins."""
        rules = RuleRegistry.list_rules()
        if rules:
            self.current_rule_id = rules[0]["id"]

    def _on_template_changed(self, path: str):
        self.template_path = path

    def _on_zip_changed(self, path: str):
        self.zip_path = path

    def _on_output_dir_changed(self, path: str):
        self.output_dir = path

    def _on_rule_changed(self, rule_id: str):
        self.current_rule_id = rule_id

    def _on_log_message(self, message: str):
        """Append log message from any thread (thread-safe via QMetaObject)."""
        self.log_panel.append_log(message)

    def _on_start(self):
        if self.is_processing:
            return

        # Validate
        if not self.template_path:
            QMessageBox.warning(self, "提示", "请选择模板。")
            return
        if not self.zip_path:
            QMessageBox.warning(self, "提示", "请选择ZIP文件。")
            return

        # Use output_dir if set, otherwise fall back to template dir
        output_dir = self.output_dir
        if not output_dir:
            from pathlib import Path
            output_dir = str(Path(self.template_path).parent)
        # Save for next launch
        self.config["last_output_dir"] = output_dir
        try:
            import json
            from pathlib import Path
            config_path = Path(__file__).parent.parent / "config.json"
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

        # Reset progress
        self.progress_panel.reset()
        self.log_panel.clear_log()
        self.is_processing = True
        self.start_btn.setEnabled(False)
        self.start_btn.setText("处理中...")

        # Start background thread
        self._processing_thread = ProcessingThread(
            self.template_path,
            self.zip_path,
            output_dir,
            self.current_rule_id,
        )
        self._processing_thread.progress_update.connect(self._on_progress_update)
        self._processing_thread.finished_signal.connect(self._on_processing_done)
        self._processing_thread.start()

    def _on_progress_update(self, msg: str):
        self.progress_panel.set_progress(50, msg)

    def _on_processing_done(self, result):
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
