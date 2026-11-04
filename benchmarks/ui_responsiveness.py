"""Measure GUI event-loop responsiveness while the Generate page produces documents.

A 10 ms QTimer heartbeat runs on the GUI thread during generation; the gaps between beats
show how long the UI could not react. Usage::

    QT_QPA_PLATFORM=offscreen uv run python benchmarks/ui_responsiveness.py --employees 100
"""

from __future__ import annotations

import argparse
import json
import shutil
import statistics
import sys
import tempfile
import time
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--employees", type=int, default=100)
    parser.add_argument("--workers", type=int, default=2)
    args = parser.parse_args()

    from PySide6.QtCore import QTimer
    from PySide6.QtWidgets import QApplication

    from docgen.i18n.catalog import set_language
    from docgen.testing.samples import write_customer_a
    from docgen.ui.gc_guard import MainThreadGarbageCollector
    from docgen.ui.main_window import MainWindow
    from docgen.ui.state import AppState

    app = QApplication([])
    gc_guard = MainThreadGarbageCollector(app)  # noqa: F841
    set_language("en")
    root = Path(tempfile.mkdtemp(prefix="docgen-ui-bench-"))
    state = AppState(root / "home")
    profile, workbook = write_customer_a(root / "src", n_employees=args.employees)
    state.repo.import_bundle(profile)
    shutil.copytree(
        profile.parent / "templates", state.repo.directory("customer_a") / "templates", dirs_exist_ok=True
    )
    state.open_profile("customer_a")
    window = MainWindow(state)
    window.show()
    window.pages["excel"].set_excel(workbook)  # type: ignore[attr-defined]
    while state.workbook is None:
        app.processEvents()
        time.sleep(0.01)
    page = window.pages["generate"]
    page.output.set_path(root / "out")  # type: ignore[attr-defined]
    page.workers.setValue(args.workers)  # type: ignore[attr-defined]

    gaps: list[float] = []
    last = [time.perf_counter()]

    def beat() -> None:
        now = time.perf_counter()
        gaps.append(now - last[0])
        last[0] = now

    timer = QTimer()
    timer.timeout.connect(beat)
    timer.start(10)
    start = time.perf_counter()
    page._start(dry_run=False)  # type: ignore[attr-defined]
    while page.worker is not None or page.result is None:  # type: ignore[attr-defined]
        app.processEvents()
        time.sleep(0.002)
    total = time.perf_counter() - start
    timer.stop()
    result = page.result  # type: ignore[attr-defined]
    beats = sorted(gaps[1:])
    report = {
        "documents": len(result.documents),
        "status": result.status,
        "seconds": round(total, 2),
        "heartbeat_interval_ms": 10,
        "max_gap_ms": round(max(beats) * 1000, 1),
        "p95_gap_ms": round(beats[int(len(beats) * 0.95)] * 1000, 1),
        "median_gap_ms": round(statistics.median(beats) * 1000, 1),
        "beats": len(beats),
    }
    print(json.dumps(report, indent=2))
    shutil.rmtree(root, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
