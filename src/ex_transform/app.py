"""PySide6 entry point for the minimal Excel search window."""

from __future__ import annotations

import sys

from .config import AppConfig
from .diagnostics import collect_environment_report, format_report
from .exporter import cleanup_stale_exports
from .logging_utils import configure_logging
from .session_index import SessionIndex
from .ui import MainWindow


def run(argv: list[str] | None = None) -> int:
    try:
        from PySide6.QtWidgets import QApplication, QStatusBar
    except ImportError:
        print("PySide6 is required to start the desktop application", file=sys.stderr)
        return 2

    config = AppConfig.from_environment()
    configure_logging(config.log_level)
    cleanup_stale_exports(config.temp_root or __import__("tempfile").gettempdir())
    application = QApplication([sys.argv[0], *(argv or [])])
    index = SessionIndex(temp_root=config.temp_root)
    window = MainWindow(index)
    window.setStatusBar(QStatusBar(window))
    window.statusBar().showMessage(format_report(collect_environment_report(probe_excel=False)))
    window.show()
    return application.exec()
