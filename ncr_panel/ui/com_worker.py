"""A long-lived worker thread for COM automation (Outlook, CATIA).

Why: Outlook / CATIA object models are single-threaded. Calling them from many short-lived
threads at once (one per click) queues heavy requests inside the target application and freezes
its UI. A single thread per application serialises every call, keeps one COM connection alive and
lets urgent requests (a click) jump ahead of a running bulk scan via `yield_now()`.
"""
from __future__ import annotations

import collections
import threading
import traceback
from typing import Any, Callable, Optional

from PySide6.QtCore import QObject, Signal, Slot

from ..utils import IS_WINDOWS


class _Task:
    __slots__ = ("fn", "args", "kwargs", "on_done", "on_error", "on_progress", "tag")

    def __init__(self, fn, args, kwargs, on_done, on_error, on_progress, tag=None):
        self.fn, self.args, self.kwargs, self.tag = fn, args, kwargs, tag
        self.on_done, self.on_error, self.on_progress = on_done, on_error, on_progress


class ComWorker(QObject):
    _call = Signal(object, object)      # (callback, value) -> executed on the GUI thread

    def __init__(self, name: str, parent: Optional[QObject] = None):
        super().__init__(parent)
        self.name = name
        self.state: dict[str, Any] = {}           # per-thread objects (COM sessions); touch only on worker
        self._normal: collections.deque[_Task] = collections.deque()
        self._urgent: collections.deque[_Task] = collections.deque()
        self._cv = threading.Condition()
        self._stop = False
        self._busy = False
        self._call.connect(self._dispatch)
        self._thread = threading.Thread(target=self._loop, name=f"com-{name}", daemon=True)
        self._thread.start()

    # ---- public (GUI thread) --------------------------------------------------
    def submit(self, fn: Callable[..., Any], *args, on_done=None, on_error=None, on_progress=None,
               urgent: bool = False, tag: Optional[str] = None, **kwargs) -> None:
        """Run fn(worker, *args, **kwargs) on the worker thread. `worker` gives access to state/progress."""
        task = _Task(fn, args, kwargs, on_done, on_error, on_progress, tag)
        with self._cv:
            (self._urgent if urgent else self._normal).append(task)
            self._cv.notify()

    def clear_pending(self, tag: Optional[str] = None) -> list:
        """Drop queued (not running) tasks, optionally only those with `tag`. Returns dropped tasks."""
        dropped = []
        with self._cv:
            for q in (self._urgent, self._normal):
                keep = [t for t in q if tag is not None and t.tag != tag]
                dropped += [t for t in q if t not in keep]
                q.clear()
                q.extend(keep)
        return dropped

    def stop(self) -> None:
        with self._cv:
            self._stop = True
            self._normal.clear()
            self._urgent.clear()
            self._cv.notify()

    @property
    def busy(self) -> bool:
        return self._busy

    # ---- called from inside a running task (worker thread) ----------------------
    def yield_now(self) -> None:
        """Run queued urgent tasks now (used by long jobs between steps)."""
        while True:
            with self._cv:
                if not self._urgent or self._stop:
                    return
                task = self._urgent.popleft()
            self._run(task)

    def emit(self, callback: Optional[Callable[[Any], None]], value: Any) -> None:
        if callback is not None:
            self._call.emit(callback, value)

    # ---- internals --------------------------------------------------------------
    @Slot(object, object)
    def _dispatch(self, cb, value):
        try:
            cb(value)
        except Exception:  # noqa: BLE001
            traceback.print_exc()

    def _loop(self):
        com = None
        if IS_WINDOWS:
            try:
                import pythoncom  # type: ignore
                pythoncom.CoInitialize()
                com = pythoncom
            except ImportError:
                com = None
        try:
            while True:
                with self._cv:
                    while not self._stop and not self._urgent and not self._normal:
                        self._cv.wait()
                    if self._stop:
                        break
                    task = self._urgent.popleft() if self._urgent else self._normal.popleft()
                self._busy = True
                self._run(task)
                self._busy = False
        finally:
            # release COM proxies BEFORE uninitialising, otherwise the target app can be left hanging
            for v in list(self.state.values()):
                try:
                    v.close()
                except Exception:
                    pass
            self.state.clear()
            if com is not None:
                try:
                    com.CoUninitialize()
                except Exception:
                    pass

    def _run(self, task: _Task):
        def progress(msg, *_):
            self.emit(task.on_progress, str(msg))
        try:
            if task.on_progress is not None:
                task.kwargs.setdefault("progress", progress)
            result = task.fn(self, *task.args, **task.kwargs)
        except Exception as exc:  # noqa: BLE001
            self.emit(task.on_error, f"{exc}\n\n{traceback.format_exc(limit=3)}")
            return
        self.emit(task.on_done, result)
