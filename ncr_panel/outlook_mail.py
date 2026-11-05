"""Outlook integration (Windows, COM) with a persistent per-NCR mail cache.

Search strategy (fast and gentle on Outlook):
  * Bulk scan when the NCR list is (re)loaded: NCR numbers are grouped into chunks and each chunk
    is ONE Outlook query per folder (``A OR B OR C ...``) instead of one query per NCR.
    NCRs that were already scanned before are only searched in mail received since their last check.
  * Clicking an NCR shows the cached result immediately and then searches only mail received since
    that NCR's last check (with a one day overlap for time-zone safety).
  * All Outlook calls run on ONE dedicated COM thread (see ui/com_worker.py), never in parallel,
    and the scan yields regularly so Outlook's own UI stays responsive.
"""
from __future__ import annotations

import datetime as _dt
import json
import re
import time
from dataclasses import asdict, dataclass, field
from typing import Callable, Iterable, Optional

from .config import CACHE_DIR, OutlookSettings
from .utils import IS_WINDOWS, process_running

ProgressCb = Optional[Callable[[str], None]]
CACHE_FILE = CACHE_DIR / "mail_cache.json"
OVERLAP = _dt.timedelta(days=1)
_CACHE_VERSION = 2


class OutlookNotRunning(RuntimeError):
    pass


@dataclass
class MailHit:
    entry_id: str
    store_id: str
    subject: str
    sender: str
    received: Optional[_dt.datetime]
    folder: str
    approved: bool = False
    rejected: bool = False
    snippet: str = ""
    has_attachments: bool = False
    matched_keywords: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        d = asdict(self)
        d["received"] = self.received.isoformat() if self.received else None
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "MailHit":
        d = dict(d)
        r = d.get("received")
        d["received"] = _dt.datetime.fromisoformat(r) if r else None
        known = {f for f in cls.__dataclass_fields__}
        return cls(**{k: v for k, v in d.items() if k in known})


def outlook_available() -> bool:
    if not IS_WINDOWS:
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def settings_signature(s: OutlookSettings) -> str:
    """Cache is only valid for the same search scope / keyword setup."""
    return json.dumps([sorted(x.casefold() for x in s.folders), s.search_all_folders, s.days_back,
                       s.search_body, sorted(s.approval_keywords), sorted(s.rejection_keywords)],
                      ensure_ascii=False)


# ---------------------------------------------------------------------------
# text helpers
# ---------------------------------------------------------------------------
_PATTERNS: dict[str, re.Pattern] = {}


def ncr_pattern(ncr: str) -> re.Pattern:
    """Regex for an NCR number that does not match inside a longer number (NCR-1 vs NCR-12)."""
    p = _PATTERNS.get(ncr)
    if p is None:
        p = re.compile(r"(?<![0-9A-Za-z])" + re.escape(ncr.strip()) + r"(?![0-9A-Za-z])", re.IGNORECASE)
        _PATTERNS[ncr] = p
    return p


def classify(text: str, settings: OutlookSettings) -> tuple[bool, bool, list[str]]:
    """(approved, rejected, matched keywords). A rejection phrase ('not approved') wins over 'approved'."""
    low = (text or "").casefold()
    rejected = [k for k in settings.rejection_keywords if k and k.casefold() in low]
    approved = [k for k in settings.approval_keywords if k and k.casefold() in low]
    if rejected:
        # drop approval words that only occur as part of a rejection phrase
        rest = low
        for k in rejected:
            rest = rest.replace(k.casefold(), " ")
        approved = [k for k in approved if k.casefold() in rest]
    return bool(approved), bool(rejected), approved + rejected


def _dasl(text: str) -> str:
    return text.replace("'", "''")


def build_filter(ncrs: list[str], since_utc: Optional[_dt.datetime], search_body: bool) -> str:
    conds = []
    for n in ncrs:
        t = _dasl(n.strip())
        conds.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{t}%'")
        if search_body:
            conds.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{t}%'")
    f = "(" + " OR ".join(conds) + ")"
    if since_utc is not None:
        f = f"\"urn:schemas:httpmail:datereceived\" >= '{since_utc:%Y-%m-%d %H:%M}' AND {f}"
    return "@SQL=" + f


def _to_naive(value) -> Optional[_dt.datetime]:
    try:
        return _dt.datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
    except Exception:
        return None


def _local_to_utc(d: _dt.datetime) -> _dt.datetime:
    try:
        return d.astimezone(_dt.timezone.utc).replace(tzinfo=None)
    except (OverflowError, OSError, ValueError):
        return d


# ---------------------------------------------------------------------------
# cache (owned by the GUI thread; the COM thread only returns results)
# ---------------------------------------------------------------------------
class MailCache:
    def __init__(self, signature: str = ""):
        self.signature = signature
        self.hits: dict[str, dict[str, MailHit]] = {}
        self.checked: dict[str, _dt.datetime] = {}

    @classmethod
    def load(cls, signature: str) -> "MailCache":
        c = cls(signature)
        try:
            data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
            if data.get("version") != _CACHE_VERSION or data.get("signature") != signature:
                return c
            for ncr, items in data.get("hits", {}).items():
                c.hits[ncr] = {d["entry_id"]: MailHit.from_dict(d) for d in items}
            c.checked = {k: _dt.datetime.fromisoformat(v) for k, v in data.get("checked", {}).items()}
        except Exception:
            pass
        return c

    def save(self) -> None:
        try:
            CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
            data = {"version": _CACHE_VERSION, "signature": self.signature,
                    "hits": {n: [h.to_dict() for h in m.values()] for n, m in self.hits.items() if m},
                    "checked": {n: t.isoformat() for n, t in self.checked.items()}}
            tmp = CACHE_FILE.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            tmp.replace(CACHE_FILE)
        except Exception:
            pass

    def merge(self, results: dict[str, list[MailHit]]) -> set[str]:
        changed = set()
        for ncr, hits in results.items():
            bucket = self.hits.setdefault(ncr, {})
            for h in hits:
                old = bucket.get(h.entry_id)
                if old is None or old.approved != h.approved or old.rejected != h.rejected:
                    changed.add(ncr)
                bucket[h.entry_id] = h
        return changed

    def mark_checked(self, ncrs: Iterable[str], when: _dt.datetime) -> None:
        for n in ncrs:
            self.checked[n] = when

    def checked_at(self, ncr: str) -> Optional[_dt.datetime]:
        return self.checked.get(ncr)

    def hits_for(self, ncr: str) -> list[MailHit]:
        return sorted(self.hits.get(ncr, {}).values(), key=lambda h: h.received or _dt.datetime.min, reverse=True)

    def approved(self, ncr: str) -> bool:
        return any(h.approved for h in self.hits.get(ncr, {}).values())

    def count(self, ncr: str) -> int:
        return len(self.hits.get(ncr, {}))

    def latest(self, ncr: str) -> Optional[_dt.datetime]:
        dates = [h.received for h in self.hits.get(ncr, {}).values() if h.received]
        return max(dates) if dates else None


def plan_since(ncrs: list[str], cache_checked: dict[str, Optional[_dt.datetime]],
               settings: OutlookSettings, now: Optional[_dt.datetime] = None) -> list[tuple[Optional[_dt.datetime], list[str]]]:
    """Group NCRs by search start: never-checked NCRs search the whole window (days_back),
    already-checked NCRs only since their oldest last check (minus overlap)."""
    now = now or _dt.datetime.now()
    floor = now - _dt.timedelta(days=settings.days_back) if settings.days_back > 0 else None
    new = [n for n in ncrs if cache_checked.get(n) is None]
    old = [n for n in ncrs if cache_checked.get(n) is not None]
    groups: list[tuple[Optional[_dt.datetime], list[str]]] = []
    if old:
        since = min(cache_checked[n] for n in old) - OVERLAP
        if floor is not None and since < floor:
            since = floor
        groups.append((since, old))
    if new:
        groups.append((floor, new))
    return groups


# ---------------------------------------------------------------------------
# Outlook session (lives on the COM thread only)
# ---------------------------------------------------------------------------
class OutlookSession:
    def __init__(self, settings: OutlookSettings):
        import win32com.client  # type: ignore
        if not settings.start_outlook_if_closed and not process_running("OUTLOOK.EXE"):
            raise OutlookNotRunning("Outlook açık değil. Outlook'u açın (veya Ayarlar → Outlook → "
                                    "'Kapalıysa Outlook'u başlat' seçeneğini işaretleyin).")
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        self._folders: Optional[list[tuple[str, object]]] = None
        self._folder_sig = ""

    def folders(self, settings: OutlookSettings) -> list[tuple[str, object]]:
        sig = settings_signature(settings)
        if self._folders is None or sig != self._folder_sig:
            self._folders = _select_folders(self.ns, settings)
            self._folder_sig = sig
        return self._folders

    def get_item(self, entry_id: str, store_id: str):
        return self.ns.GetItemFromID(entry_id, store_id) if store_id else self.ns.GetItemFromID(entry_id)

    def close(self):
        self._folders = None
        self.ns = None
        self.app = None


def _walk_folders(folder, path: str, out: list, depth: int = 0):
    if depth > 6:
        return
    try:
        subs = list(folder.Folders)
    except Exception:
        return
    for sub in subs:
        try:
            name = sub.Name
        except Exception:
            continue
        p = f"{path}/{name}" if path else name
        out.append((p, sub))
        _walk_folders(sub, p, out, depth + 1)


def _select_folders(ns, settings: OutlookSettings) -> list[tuple[str, object]]:
    all_folders: list[tuple[str, object]] = []
    try:
        stores = list(ns.Stores)
    except Exception:
        stores = []
    for store in stores:
        try:
            if store.ExchangeStoreType == 2:      # public folders: huge and slow, never scan
                continue
            _walk_folders(store.GetRootFolder(), "", all_folders)
        except Exception:
            continue
    wanted = [w.strip().casefold() for w in settings.folders if w.strip()]
    chosen: list[tuple[str, object]] = []
    for path, folder in all_folders:
        try:
            if folder.DefaultItemType != 0:        # olMailItem
                continue
        except Exception:
            continue
        if settings.search_all_folders:
            chosen.append((path, folder))
            continue
        name = path.rsplit("/", 1)[-1].casefold()
        p = path.casefold()
        if any(w == name or w == p or p.endswith("/" + w) for w in wanted):
            chosen.append((path, folder))
    if not chosen:
        for const in (6, 5):   # olFolderInbox, olFolderSentMail
            try:
                f = ns.GetDefaultFolder(const)
                chosen.append((f.Name, f))
            except Exception:
                pass
    return chosen


def _iter_rows(folder, filt: str, since_local: Optional[_dt.datetime]):
    """Yield (entry_id, subject, received, sender, message_class) for items matching `filt`.
    Uses Folder.GetTable (fast, no item loading); falls back to Items.Restrict."""
    try:
        try:
            table = folder.GetTable(filt, 0)
        except Exception:
            # some Outlook builds reject the date literal -> filter dates in Python instead
            table = folder.GetTable(_strip_date(filt), 0)
        cols = table.Columns
        cols.RemoveAll()
        for c in ("EntryID", "Subject", "ReceivedTime", "SenderName", "MessageClass"):
            cols.Add(c)
        while not table.EndOfTable:
            row = table.GetNextRow()
            if row is None:
                break
            v = row.GetValues()
            yield v[0], str(v[1] or ""), _to_naive(v[2]), str(v[3] or ""), str(v[4] or "")
        return
    except Exception:
        pass
    items = folder.Items
    try:
        found = items.Restrict(filt)
    except Exception:
        found = items.Restrict(_strip_date(filt))
    for item in found:
        try:
            yield (item.EntryID, str(item.Subject or ""), _to_naive(item.ReceivedTime),
                   str(getattr(item, "SenderName", "") or ""), str(getattr(item, "MessageClass", "") or ""))
        except Exception:
            continue


def _strip_date(filt: str) -> str:
    return re.sub(r"\"urn:schemas:httpmail:datereceived\" >= '[^']*' AND ", "", filt)


def search(session, ncrs: list[str], since_local: Optional[_dt.datetime], settings: OutlookSettings,
           cancel=None, yield_cb: Optional[Callable[[], None]] = None, progress: ProgressCb = None,
           partial: Optional[Callable[[dict, list], None]] = None) -> dict[str, list[MailHit]]:
    """Search mail for the given NCR numbers received since `since_local` (None = everything).

    `partial(results_of_chunk, chunk_ncrs)` is called after every chunk is finished in all folders.
    Returns all results {ncr: [MailHit]} (NCRs without hits are omitted)."""
    ncrs = [n for n in dict.fromkeys(x.strip() for x in ncrs) if n]
    results: dict[str, list[MailHit]] = {}
    if not ncrs:
        return results
    folders = session.folders(settings)
    since_utc = _local_to_utc(since_local) if since_local else None
    size = max(1, int(getattr(settings, "chunk_size", 40) or 40))
    chunks = [ncrs[i:i + size] for i in range(0, len(ncrs), size)]
    seen: dict[str, tuple[str, str, MailHit]] = {}      # entry_id -> (subject, body, base hit)
    pump = _pumper()
    n_rows = 0
    for ci, chunk in enumerate(chunks):
        chunk_res: dict[str, list[MailHit]] = {}
        patterns = {n: ncr_pattern(n) for n in chunk}
        filt = build_filter(chunk, since_utc, settings.search_body)
        for path, folder in folders:
            if cancel is not None and cancel.is_set():
                return results
            if progress:
                progress(f"Outlook: {path} ({ci + 1}/{len(chunks)})")
            count = 0
            try:
                rows = list(_iter_rows(folder, filt, since_local))
            except Exception:
                continue
            for entry_id, subject, received, sender, mclass in rows:
                if cancel is not None and cancel.is_set():
                    return results
                n_rows += 1
                if n_rows % 15 == 0:
                    pump()
                    if yield_cb:
                        yield_cb()
                if mclass and not mclass.upper().startswith("IPM.NOTE"):
                    continue
                if since_local and received and received < since_local:
                    continue
                cached = seen.get(entry_id)
                if cached is None:
                    try:
                        item = session.get_item(entry_id, _store_id(folder))
                        body = str(item.Body or "")
                        try:
                            has_att = item.Attachments.Count > 0
                        except Exception:
                            has_att = False
                        del item
                    except Exception:
                        body, has_att = "", False
                    approved, rejected, kws = classify(subject + "\n" + body, settings)
                    base = MailHit(entry_id=entry_id, store_id=_store_id(folder), subject=subject, sender=sender,
                                   received=received, folder=path, approved=approved, rejected=rejected,
                                   snippet=" ".join(body.split())[:300], has_attachments=has_att,
                                   matched_keywords=kws)
                    cached = (subject, body, base)
                    seen[entry_id] = cached
                subj, body, base = cached
                text = subj + "\n" + body
                for n, pat in patterns.items():
                    if count >= settings.max_results * len(chunk):
                        break
                    if pat.search(text):
                        chunk_res.setdefault(n, []).append(base)
                        count += 1
            time.sleep(0.03)        # give Outlook a breather between folders
            if yield_cb:
                yield_cb()
        for n, hits in chunk_res.items():
            results.setdefault(n, []).extend(hits)
        if partial:
            partial(chunk_res, chunk)
    return results


def _store_id(folder) -> str:
    try:
        return folder.StoreID
    except Exception:
        return ""


def _pumper():
    try:
        import pythoncom  # type: ignore

        def pump():
            pythoncom.PumpWaitingMessages()
            time.sleep(0.005)
        return pump
    except ImportError:
        return lambda: None


def display_mail(session, entry_id: str, store_id: str = "") -> None:
    session.get_item(entry_id, store_id).Display()


def new_mail(subject: str, body: str = "") -> None:
    if not outlook_available():
        raise RuntimeError("Outlook erişimi yok.")
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body
    mail.Display()
