"""Logging utility - unified logger with real-time signal support."""

import logging
import sys
from datetime import datetime


class LogSignal:
    """Emits log messages via a callable for UI real-time display."""
    def __init__(self):
        self.listeners = []

    def connect(self, callback):
        self.listeners.append(callback)

    def disconnect(self, callback):
        if callback in self.listeners:
            self.listeners.remove(callback)

    def emit(self, message):
        for cb in self.listeners:
            try:
                cb(message)
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

    # Signal handler for UI
    sh = SignalHandler()
    sh.setLevel(logging.INFO)
    sh.setFormatter(logging.Formatter("%(asctime)s  %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(sh)

    return logger


logger = setup_logger()
