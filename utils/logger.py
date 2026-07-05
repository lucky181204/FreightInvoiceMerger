"""Logging utility - unified logger with real-time signal support.

Uses a QTimer-based queue to ensure log messages are always delivered
on the main (GUI) thread — preventing Qt thread-safety crashes.
"""

import logging
from datetime import datetime
from PySide6.QtCore import QTimer
from collections import deque


_log_queue: deque = deque()
_log_timer = None


class LogSignal:
    """Thread-safe log signal: queues messages and dispatches on main thread."""
    def __init__(self):
        self.listeners = []

    def connect(self, callback):
        self.listeners.append(callback)
        # Start the dispatch timer on first connect
        global _log_timer
        if _log_timer is None:
            _log_timer = QTimer()
            _log_timer.setInterval(50)  # 50ms poll interval
            _log_timer.timeout.connect(self._dispatch)
            _log_timer.start()

    def disconnect(self, callback):
        if callback in self.listeners:
            self.listeners.remove(callback)

    def emit(self, message):
        """Called from any thread — safely queues the message."""
        _log_queue.append(message)

    def _dispatch(self):
        """Called on main thread by QTimer — dispatches all queued messages."""
        while _log_queue:
            msg = _log_queue.popleft()
            for cb in self.listeners:
                try:
                    cb(msg)
                except Exception:
                    pass


log_signal = LogSignal()


class SignalHandler(logging.Handler):
    def emit(self, record):
        msg = self.format(record)
        log_signal.emit(msg)


def setup_logger(name="FreightInvoiceMerger"):
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # File handler
    fh = logging.FileHandler(f"merger_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log", encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(fh)

    # Signal handler for UI (safe: queues to main thread via QTimer)
    sh = SignalHandler()
    sh.setLevel(logging.INFO)
    sh.setFormatter(logging.Formatter("%(asctime)s  %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(sh)

    return logger


logger = setup_logger()
