"""GUI-side mail service: owns the cache, schedules bulk / incremental Outlook searches."""
from __future__ import annotations

import datetime as _dt
import threading

from PySide6.QtCore import QObject, Signal

from ..config import AppConfig
from ..outlook_mail import (MailCache, MailHit, OutlookNotRunning, OutlookSession, display_mail,
                            outlook_available, plan_since, search, settings_signature)
from .com_worker import ComWorker


def _session(worker: ComWorker, settings) -> OutlookSession:
    s = worker.state.get("outlook")
    if s is None:
        s = OutlookSession(settings)
        worker.state["outlook"] = s
    return s


def _drop_session(worker: ComWorker):
    s = worker.state.pop("outlook", None)
    if s is not None:
        s.close()


def _job_search(worker: ComWorker, groups, settings, cancel, partial_cb=None, progress=None):
    try:
        session = _session(worker, settings)
        out: dict[str, list[MailHit]] = {}
        for since, ncrs in groups:
            def partial(res, chunk, _since=since):
                worker.emit(partial_cb, (res, chunk))
            res = search(session, ncrs, since, settings, cancel=cancel, yield_cb=worker.yield_now,
                         progress=progress, partial=partial if partial_cb else None)
            for k, v in res.items():
                out.setdefault(k, []).extend(v)
            if cancel is not None and cancel.is_set():
                break
        return out
    except OutlookNotRunning:
        raise
    except Exception:
        _drop_session(worker)       # Outlook restarted / RPC server unavailable -> reconnect next time
        raise


def _job_open(worker: ComWorker, entry_id, store_id, settings):
    try:
        display_mail(_session(worker, settings), entry_id, store_id)
    except Exception:
        _drop_session(worker)
        display_mail(_session(worker, settings), entry_id, store_id)


class MailService(QObject):
    updated = Signal(object)          # set[str] of NCR numbers whose mail list changed
    checked = Signal(str)             # single NCR finished an incremental check
    status = Signal(str)
    bulkRunning = Signal(bool)
    error = Signal(str)

    def __init__(self, cfg: AppConfig, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.cache = MailCache.load(settings_signature(cfg.outlook))
        self.worker = ComWorker("outlook", self)
        self._bulk_cancel: threading.Event | None = None
        self._pending_one: set[str] = set()
        self.bulk_active = False
        self.last_error = ""

    # ---- state ---------------------------------------------------------------
    def enabled(self) -> bool:
        return self.cfg.outlook.enabled and outlook_available()

    def hits(self, ncr: str) -> list[MailHit]:
        return self.cache.hits_for(ncr)

    def approved(self, ncr: str) -> bool:
        return self.cache.approved(ncr)

    def checked_at(self, ncr: str):
        return self.cache.checked_at(ncr)

    def is_pending(self, ncr: str) -> bool:
        return ncr in self._pending_one

    def settings_changed(self):
        sig = settings_signature(self.cfg.outlook)
        if sig != self.cache.signature:
            self.cancel_bulk()
            self.cache = MailCache(sig)
            self.cache.save()
            self.worker.submit(lambda w: _drop_session(w))

    # ---- bulk ----------------------------------------------------------------
    def bulk(self, ncrs: list[str]):
        """Search all given NCRs: new ones over the whole window, known ones only since last check."""
        if not self.enabled() or not ncrs:
            return
        self.cancel_bulk()
        cancel = threading.Event()
        self._bulk_cancel = cancel
        start = _dt.datetime.now()
        groups = plan_since(list(dict.fromkeys(ncrs)), {n: self.cache.checked_at(n) for n in ncrs}, self.cfg.outlook)
        n_new = sum(len(g) for s, g in groups if self.cache.checked_at(g[0]) is None) if groups else 0
        self.bulk_active = True
        self.bulkRunning.emit(True)
        self.status.emit(f"Outlook taraması başladı: {len(set(ncrs))} NCR ({n_new} yeni).")

        def partial(payload):
            if cancel.is_set():
                return
            res, chunk = payload
            changed = self.cache.merge(res)
            self.cache.mark_checked(chunk, start)
            if changed:
                self.updated.emit(changed)

        def done(_res):
            if self._bulk_cancel is cancel:
                self._bulk_cancel = None
                self.bulk_active = False
                self.bulkRunning.emit(False)
            self.cache.save()
            if not cancel.is_set():
                total = sum(self.cache.count(n) for n in set(ncrs))
                self.status.emit(f"Outlook taraması bitti: {total} e-posta eşleşmesi.")

        def err(msg):
            if self._bulk_cancel is cancel:
                self._bulk_cancel = None
                self.bulk_active = False
                self.bulkRunning.emit(False)
            self.cache.save()
            self.last_error = msg.splitlines()[0]
            self.status.emit("Outlook taraması yapılamadı: " + self.last_error)

        self.worker.submit(_job_search, groups, self.cfg.outlook, cancel, partial_cb=partial,
                           on_done=done, on_error=err, on_progress=self.status.emit)

    def cancel_bulk(self):
        if self._bulk_cancel is not None:
            self._bulk_cancel.set()
            self._bulk_cancel = None
        if self.bulk_active:
            self.bulk_active = False
            self.bulkRunning.emit(False)

    # ---- single NCR (on click) -------------------------------------------------
    def check_one(self, ncr: str, force: bool = False, full: bool = False) -> bool:
        """Incrementally search one NCR (mail newer than its last check). Returns True if queued."""
        if not self.enabled() or not ncr or ncr in self._pending_one:
            return False
        last = None if full else self.cache.checked_at(ncr)
        if not force and last is not None and \
                _dt.datetime.now() - last < _dt.timedelta(minutes=max(0, self.cfg.outlook.recheck_minutes)):
            return False
        groups = plan_since([ncr], {ncr: last}, self.cfg.outlook)
        start = _dt.datetime.now()
        # a newer click makes older queued (not yet started) single-NCR checks pointless
        for t in self.worker.clear_pending(tag="one"):
            dropped_groups = t.args[0]                  # [(since, [ncr])]
            self._pending_one.discard(dropped_groups[0][1][0])
        self._pending_one.add(ncr)

        def done(res):
            self._pending_one.discard(ncr)
            changed = self.cache.merge(res)
            self.cache.mark_checked([ncr], start)
            self.cache.save()
            if changed:
                self.updated.emit(changed)
            self.checked.emit(ncr)

        def err(msg):
            self._pending_one.discard(ncr)
            self.last_error = msg.splitlines()[0]
            self.checked.emit(ncr)
            self.status.emit("Outlook: " + self.last_error)

        self.worker.submit(_job_search, groups, self.cfg.outlook, None, on_done=done, on_error=err,
                           urgent=True, tag="one")
        return True

    def open(self, hit: MailHit):
        self.worker.submit(_job_open, hit.entry_id, hit.store_id, self.cfg.outlook, urgent=True,
                           on_error=lambda m: self.error.emit("E-posta açılamadı:\n" + m))

    def shutdown(self):
        self.cancel_bulk()
        self.cache.save()
        self.worker.stop()
