"""Background task helper for the Qt GUI."""
from __future__ import annotations

import traceback
from typing import Any, Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, QTimer, Signal, Slot

# Strong references to running workers. Without this, Python may garbage-collect a Worker (and its
# signals object) before the queued "finished" signal reaches the GUI thread, and the callback is
# silently lost (e.g. the repetitive-notes scan finished but the badges never updated).
_ACTIVE: set = set()


class WorkerSignals(QObject):
    finished = Signal(object)        # result
    error = Signal(str)              # message
    progress = Signal(str)           # text
    progress_pct = Signal(int, int)  # done, total


class Worker(QRunnable):
    """Run `fn(*args, progress=cb, **kwargs)` on the global thread pool."""

    def __init__(self, fn: Callable[..., Any], *args, pass_progress: bool = True, **kwargs):
        super().__init__()
        self.fn, self.args, self.kwargs = fn, args, kwargs
        self.pass_progress = pass_progress
        self.signals = WorkerSignals()
        self.cancelled = False
        self.setAutoDelete(False)      # lifetime is managed by _ACTIVE

    def cancel(self):
        self.cancelled = True

    @Slot()
    def run(self):
        try:
            if self.pass_progress:
                self.kwargs["progress"] = self._progress
            result = self.fn(*self.args, **self.kwargs)
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc(limit=3)
            self.signals.error.emit(f"{exc}\n\n{tb}")
            return
        self.signals.finished.emit(result)

    def _progress(self, msg: str, done: int | None = None, total: int | None = None):
        self.signals.progress.emit(str(msg))
        if done is not None and total is not None:
            self.signals.progress_pct.emit(int(done), int(total))


def run_in_background(fn: Callable[..., Any], *args, on_done=None, on_error=None,
                      on_progress=None, on_progress_pct=None, pass_progress: bool = True, **kwargs) -> Worker:
    w = Worker(fn, *args, pass_progress=pass_progress, **kwargs)
    if on_done:
        w.signals.finished.connect(on_done)
    if on_error:
        w.signals.error.connect(on_error)
    if on_progress:
        w.signals.progress.connect(on_progress)
    if on_progress_pct:
        w.signals.progress_pct.connect(on_progress_pct)

    def _release(*_):
        # defer: never drop the last reference while its own signal is being delivered
        QTimer.singleShot(0, lambda: _ACTIVE.discard(w))

    w.signals.finished.connect(_release)
    w.signals.error.connect(_release)
    _ACTIVE.add(w)
    QThreadPool.globalInstance().start(w)
    return w
