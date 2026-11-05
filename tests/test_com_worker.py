"""The COM worker serialises jobs and lets urgent ones (a click) run inside a long bulk job."""
import os
import threading
import time

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
QtCore = pytest.importorskip("PySide6.QtCore")

from ncr_panel.ui.com_worker import ComWorker  # noqa: E402


@pytest.fixture(scope="module")
def app():
    return QtCore.QCoreApplication.instance() or QtCore.QCoreApplication([])


def _wait(app, cond, timeout=5.0):
    end = time.time() + timeout
    while not cond() and time.time() < end:
        app.processEvents()
        time.sleep(0.01)
    return cond()


def test_urgent_task_runs_during_bulk(app):
    w = ComWorker("test")
    order, threads = [], set()
    started = threading.Event()

    def bulk(worker):
        started.set()
        for i in range(30):
            order.append(f"bulk{i}")
            threads.add(threading.get_ident())
            worker.yield_now()
            time.sleep(0.01)
        return "bulk-done"

    def click(worker, ncr):
        order.append(f"click:{ncr}")
        threads.add(threading.get_ident())
        return ncr

    results = []
    w.submit(bulk, on_done=results.append)
    assert started.wait(2)
    w.submit(click, "NCR-1", urgent=True, tag="one", on_done=results.append)
    assert _wait(app, lambda: len(results) == 2)
    assert results == ["NCR-1", "bulk-done"]            # the click did not wait for the bulk job
    assert order.index("click:NCR-1") < order.index("bulk29")
    assert len(threads) == 1                             # everything on ONE thread
    w.stop()


def test_clear_pending_by_tag_keeps_others(app):
    w = ComWorker("test2")
    gate = threading.Event()
    w.submit(lambda wk: gate.wait(2))                    # occupy the thread
    time.sleep(0.05)
    w.submit(lambda wk: "a", urgent=True, tag="one")
    w.submit(lambda wk: "open", urgent=True)
    dropped = w.clear_pending(tag="one")
    assert len(dropped) == 1
    got = []
    w.submit(lambda wk: "b", urgent=True, tag="one", on_done=got.append)
    gate.set()
    assert _wait(app, lambda: got == ["b"])
    w.stop()


def test_error_is_reported(app):
    w = ComWorker("test3")
    errs = []
    w.submit(lambda wk: 1 / 0, on_error=errs.append)
    assert _wait(app, lambda: errs)
    assert "division" in errs[0]
    w.stop()
