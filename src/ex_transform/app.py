"""PySide6 entry point for the intentionally minimal first window."""

from __future__ import annotations

import sys
from typing import Any

from .config import AppConfig
from .diagnostics import collect_environment_report, format_report
from .logging_utils import configure_logging


def run(argv: list[str] | None = None) -> int:
    try:
        from PySide6.QtWidgets import QApplication, QLabel, QMainWindow, QStatusBar
    except ImportError:
        print("PySide6 is required to start the desktop application", file=sys.stderr)
        return 2

    config = AppConfig.from_environment()
    configure_logging(config.log_level)
    qt_args = [sys.argv[0], *(argv or [])]
    application = QApplication(qt_args)

    window = QMainWindow()
    window.setWindowTitle(config.app_name)
    window.resize(720, 420)
    window.setCentralWidget(QLabel("Select Excel files to begin a search."))
    status = QStatusBar()
    status.showMessage(format_report(collect_environment_report()))
    window.setStatusBar(status)
    window.show()
    return application.exec()
