"""Desktop application entry point."""

from __future__ import annotations

import logging
import logging.handlers
import multiprocessing
import sys
from pathlib import Path


def _configure_logging() -> None:
    from docgen.infrastructure.storage import paths

    handler = logging.handlers.RotatingFileHandler(
        paths.log_dir() / "docgen.log", maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(handler)


def main(argv: list[str] | None = None) -> int:
    multiprocessing.freeze_support()
    from PySide6.QtCore import QThreadPool
    from PySide6.QtWidgets import QApplication

    from docgen.i18n.catalog import set_language
    from docgen.ui.main_window import MainWindow
    from docgen.ui.state import AppState

    _configure_logging()
    app = QApplication(argv if argv is not None else sys.argv)
    app.setApplicationName("DocGen")
    app.setOrganizationName("DocGen")
    app.setStyle("Fusion")
    from docgen.ui.gc_guard import MainThreadGarbageCollector

    collector = MainThreadGarbageCollector(app)  # noqa: F841 - keeps GC on the GUI thread
    # keep I/O work bounded so the UI stays responsive on modest machines
    QThreadPool.globalInstance().setMaxThreadCount(
        max(2, min(4, QThreadPool.globalInstance().maxThreadCount()))
    )
    state = AppState()
    set_language(state.settings.language)
    window = MainWindow(state)
    window.show()
    args = (argv if argv is not None else sys.argv)[1:]
    if args and Path(args[0]).suffix.lower() in (".xlsx", ".xlsm", ".xls", ".csv"):
        excel = window.pages["excel"]
        excel.set_excel(Path(args[0]))  # type: ignore[attr-defined]
        window.go("excel")
    return int(app.exec())


if __name__ == "__main__":
    sys.exit(main())
