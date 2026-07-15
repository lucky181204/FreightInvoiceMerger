"""Single Invoice generation tab for Freight Invoice Merger Pro."""

import os
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QMessageBox, QFrame,
)
from PySide6.QtCore import QThread, Signal, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl

from ui.widgets import FileBrowseRow, OutputDirRow, ProgressPanel, LogPanel
from single_invoice.services.invoice_merger import generate_single_invoice
from utils.logger import log_signal


class SingleInvoiceWorker(QThread):
    """Background worker for single invoice generation."""

    progress_update = Signal(str)
    finished_signal = Signal(object)

    def __init__(self, template_path, draft_path, manifest_path,
                 trade_terms_path, output_dir, parent=None):
        super().__init__(parent)
        self.template_path = template_path
        self.draft_path = draft_path
        self.manifest_path = manifest_path
        self.trade_terms_path = trade_terms_path
        self.output_dir = output_dir

    def run(self):
        def progress_cb(msg):
            self.progress_update.emit(msg)

        result = generate_single_invoice(
            self.template_path,
            self.draft_path,
            self.manifest_path,
            self.trade_terms_path,
            self.output_dir,
            progress_callback=progress_cb,
        )
        self.finished_signal.emit(result)


class SingleInvoiceTab(QWidget):
    """Tab page for single invoice generation."""

    def __init__(self, config: dict, parent=None):
        super().__init__(parent)
        self.config = config
        self.template_path = ""
        self.draft_path = ""
        self.manifest_path = ""
        self.trade_terms_path = ""
        self.output_dir = ""
        self._is_running = False

        self._setup_ui()
        log_signal.connect(self._on_log)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 24, 32, 24)
        layout.setSpacing(16)

        # ── Title ──
        title = QLabel("单张发票生成")
        title.setObjectName("titleLabel")
        layout.addWidget(title)

        # ── File Selection rows ──
        self.template_row = FileBrowseRow(
            label_text="选择 Freight Invoice 模板",
            browse_mode="file",
            file_filter="Excel 文件 (*.xlsx *.xlsm *.xltx);;所有文件 (*)",
            placeholder="请选择Invoice模板...",
        )
        self.template_row.file_changed.connect(lambda p: setattr(self, 'template_path', p))
        layout.addWidget(self.template_row)

        self.draft_row = FileBrowseRow(
            label_text="选择 BL Draft",
            browse_mode="file",
            file_filter="Excel 文件 (*.xlsx *.xlsm *.xltx);;所有文件 (*)",
            placeholder="请选择BL Draft...",
        )
        self.draft_row.file_changed.connect(lambda p: setattr(self, 'draft_path', p))
        layout.addWidget(self.draft_row)

        self.manifest_row = FileBrowseRow(
            label_text="选择 上海舱单",
            browse_mode="file",
            file_filter="Excel 文件 (*.xlsx *.xlsm *.xltx);;所有文件 (*)",
            placeholder="请选择上海舱单...",
        )
        self.manifest_row.file_changed.connect(lambda p: setattr(self, 'manifest_path', p))
        layout.addWidget(self.manifest_row)

        self.trade_terms_row = FileBrowseRow(
            label_text="选择 开票贸易条款清单",
            browse_mode="file",
            file_filter="Excel 文件 (*.xlsx *.xlsm *.xltx);;所有文件 (*)",
            placeholder="请选择贸易条款清单...",
        )
        self.trade_terms_row.file_changed.connect(lambda p: setattr(self, 'trade_terms_path', p))
        layout.addWidget(self.trade_terms_row)

        # ── Output directory ──
        last_output = self.config.get("last_output_dir", "")
        self.output_dir_row = OutputDirRow(
            label_text="输出目录",
            default_dir=last_output,
            placeholder="选择输出文件保存目录...",
        )
        self.output_dir_row.dir_changed.connect(lambda p: setattr(self, 'output_dir', p))
        layout.addWidget(self.output_dir_row)

        # ── Buttons row ──
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.generate_btn = QPushButton("开始生成单张发票")
        self.generate_btn.setObjectName("startButton")
        self.generate_btn.clicked.connect(self._on_generate)
        btn_row.addWidget(self.generate_btn)

        self.open_file_btn = QPushButton("打开结果文件")
        self.open_file_btn.setObjectName("browseButton")
        self.open_file_btn.clicked.connect(self._open_result_file)
        self.open_file_btn.setEnabled(False)
        btn_row.addWidget(self.open_file_btn)

        self.open_dir_btn = QPushButton("打开输出目录")
        self.open_dir_btn.setObjectName("browseButton")
        self.open_dir_btn.clicked.connect(self._open_output_dir)
        self.open_dir_btn.setEnabled(False)
        btn_row.addWidget(self.open_dir_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        # ── Progress panel ──
        self.progress_panel = ProgressPanel()
        layout.addWidget(self.progress_panel)

        # ── Log panel ──
        self.log_panel = LogPanel()
        layout.addWidget(self.log_panel, 1)

        self._last_output_path = ""

    def _on_log(self, message: str):
        self.log_panel.append_log(message)

    def _on_generate(self):
        if self._is_running:
            return

        # Validate
        if not self.template_path:
            QMessageBox.warning(self, "提示", "请选择Invoice模板。")
            return
        if not self.draft_path:
            QMessageBox.warning(self, "提示", "请选择BL Draft。")
            return
        if not self.output_dir:
            self.output_dir = str(Path(self.template_path).parent)

        # Save output dir
        self.config["last_output_dir"] = self.output_dir
        try:
            import json
            config_path = Path(__file__).parent.parent / "config.json"
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

        self.progress_panel.reset()
        self.log_panel.clear_log()
        self._is_running = True
        self.generate_btn.setEnabled(False)
        self.generate_btn.setText("生成中...")
        self.open_file_btn.setEnabled(False)
        self.open_dir_btn.setEnabled(False)

        self._worker = SingleInvoiceWorker(
            self.template_path,
            self.draft_path,
            self.manifest_path,
            self.trade_terms_path,
            self.output_dir,
        )
        self._worker.progress_update.connect(self._on_progress)
        self._worker.finished_signal.connect(self._on_done)
        self._worker.start()

    def _on_progress(self, msg: str):
        self.progress_panel.set_progress(50, msg)

    def _on_done(self, result):
        self._is_running = False
        self.generate_btn.setEnabled(True)
        self.generate_btn.setText("开始生成单张发票")

        if result.get("success"):
            self.progress_panel.set_progress(100, "生成完成 ✓")
            self._last_output_path = result.get("output_path", "")
            self.open_file_btn.setEnabled(True)
            self.open_dir_btn.setEnabled(True)

            warnings = result.get("warnings", [])
            warn_text = "\n".join(f"⚠ {w}" for w in warnings) if warnings else ""
            summary = (
                f"单张发票生成成功！\n\n"
                f"PO：{result.get('po', '')}\n"
                f"提单号：{result.get('bl_no', '')}\n"
                f"输出文件：{self._last_output_path}\n"
            )
            if warn_text:
                summary += f"\n警告：\n{warn_text}"
            QMessageBox.information(self, "完成", summary)
        else:
            self.progress_panel.set_progress(0, "生成失败")
            QMessageBox.critical(self, "错误", result.get("message", "未知错误"))

    def _open_result_file(self):
        if self._last_output_path and os.path.exists(self._last_output_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self._last_output_path))

    def _open_output_dir(self):
        if self._last_output_path:
            dir_path = os.path.dirname(self._last_output_path)
            if os.path.exists(dir_path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(dir_path))
