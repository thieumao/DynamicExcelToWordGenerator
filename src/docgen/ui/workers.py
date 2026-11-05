"""Background execution for the UI.

Heavy work (reading workbooks, previews, generation) runs in ``QThreadPool``
workers; results come back through queued signals so the UI thread never
blocks. Progress updates are throttled to keep the event loop responsive.
"""

from __future__ import annotations

import time
import traceback
from collections.abc import Callable
from typing import Any

import shiboken6
from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal

from docgen.application.generation.job import CancellationToken
from docgen.domain.diagnostics import Issue


class WorkerSignals(QObject):
    finished = Signal(object)
    failed = Signal(str, str)
    progress = Signal(int, int, str)
    stage = Signal(str)
    issue = Signal(object)


class QtReporter:
    """Adapts the core ProgressReporter protocol to Qt signals (thread-safe, throttled)."""

    def __init__(self, signals: WorkerSignals, min_interval: float = 0.1) -> None:
        self.signals = signals
        self.min_interval = min_interval
        self._last = 0.0
        self._issue_count = 0

    def stage(self, name: str) -> None:
        self.signals.stage.emit(name)

    def progress(self, done: int, total: int, message: str = "") -> None:
        now = time.monotonic()
        if done >= total or now - self._last >= self.min_interval:
            self._last = now
            self.signals.progress.emit(done, total, message)

    def issue(self, issue: Issue) -> None:
        if self._issue_count < 5000:
            self._issue_count += 1
            self.signals.issue.emit(issue)


class Worker(QRunnable):
    def __init__(
        self, fn: Callable[..., Any], *args: Any, with_reporter: bool = False, **kwargs: Any
    ) -> None:
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        self.with_reporter = with_reporter
        self.cancel = CancellationToken()
        self.setAutoDelete(True)

    def run(self) -> None:
        try:
            if self.with_reporter:
                result = self.fn(
                    *self.args, reporter=QtReporter(self.signals), cancel=self.cancel, **self.kwargs
                )
            else:
                result = self.fn(*self.args, **self.kwargs)
        except Exception as exc:  # report to the UI instead of crashing the app
            self.signals.failed.emit(f"{type(exc).__name__}: {exc}", traceback.format_exc())
            return
        self.signals.finished.emit(result)


_ACTIVE: set[Worker] = set()


def run_in_background(
    fn: Callable[..., Any],
    *args: Any,
    on_done: Callable[[Any], None] | None = None,
    on_error: Callable[[str, str], None] | None = None,
    on_progress: Callable[[int, int, str], None] | None = None,
    on_stage: Callable[[str], None] | None = None,
    on_issue: Callable[[Issue], None] | None = None,
    with_reporter: bool = False,
    owner: QObject | None = None,
    **kwargs: Any,
) -> Worker:
    """Run *fn* on the thread pool; callbacks run on the GUI thread.

    Callbacks are skipped once *owner* (the page that started the work) has been destroyed, e.g.
    when the window is closed while a scan is still running.
    """
    worker = Worker(fn, *args, with_reporter=with_reporter, **kwargs)
    _ACTIVE.add(worker)

    def guard(callback: Callable[..., None]) -> Callable[..., None]:
        if owner is None:
            return callback

        def call(*values: Any) -> None:
            if shiboken6.isValid(owner):
                callback(*values)

        return call

    on_done, on_error = (guard(on_done) if on_done else None), (guard(on_error) if on_error else None)
    on_progress = guard(on_progress) if on_progress else None
    on_stage = guard(on_stage) if on_stage else None
    on_issue = guard(on_issue) if on_issue else None

    def _cleanup(*_: Any) -> None:
        _ACTIVE.discard(worker)

    if on_done:
        worker.signals.finished.connect(on_done)
    if on_error:
        worker.signals.failed.connect(on_error)
    if on_progress:
        worker.signals.progress.connect(on_progress)
    if on_stage:
        worker.signals.stage.connect(on_stage)
    if on_issue:
        worker.signals.issue.connect(on_issue)
    worker.signals.finished.connect(_cleanup)
    worker.signals.failed.connect(_cleanup)
    QThreadPool.globalInstance().start(worker)
    return worker
