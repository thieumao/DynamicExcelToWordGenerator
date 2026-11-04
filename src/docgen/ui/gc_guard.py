"""Run Python's cyclic garbage collector only on the GUI thread.

Python may trigger cyclic GC in whichever thread allocates; if that is a
QThreadPool worker and the garbage contains Qt widgets (e.g. a page with
signal-connected lambdas), Qt objects get destroyed off the GUI thread and the
process crashes. Automatic GC is therefore disabled and replaced by a timer
on the GUI thread that performs the same generational collections.
"""

from __future__ import annotations

import gc

from PySide6.QtCore import QObject, QTimer


class MainThreadGarbageCollector(QObject):
    def __init__(self, parent: QObject | None = None, interval_ms: int = 500) -> None:
        super().__init__(parent)
        self.threshold = gc.get_threshold()
        gc.disable()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check)
        self.timer.start(interval_ms)

    def check(self) -> None:
        count0, count1, count2 = gc.get_count()
        if count0 > self.threshold[0]:
            gc.collect(0)
            if count1 > self.threshold[1]:
                gc.collect(1)
                if count2 > self.threshold[2]:
                    gc.collect(2)

    def stop(self) -> None:
        self.timer.stop()
        gc.enable()
