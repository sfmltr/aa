"""Outlook integration via COM (Windows). Searches mails mentioning an NCR number."""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Callable, Optional

from .config import OutlookSettings
from .utils import IS_WINDOWS

ProgressCb = Optional[Callable[[str], None]]


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


def outlook_available() -> bool:
    if not IS_WINDOWS:
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def _dasl_escape(text: str) -> str:
    return text.replace("'", "''").replace("%", "[%]")


def _walk_folders(folder, path: str, out: list):
    try:
        for sub in folder.Folders:
            name = sub.Name
            p = f"{path}/{name}" if path else name
            out.append((p, sub))
            _walk_folders(sub, p, out)
    except Exception:
        pass


def _select_folders(ns, settings: OutlookSettings) -> list[tuple[str, object]]:
    all_folders: list[tuple[str, object]] = []
    try:
        for store in ns.Stores:
            try:
                root = store.GetRootFolder()
                _walk_folders(root, "", all_folders)
            except Exception:
                continue
    except Exception:
        pass
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
        name = folder.Name.casefold()
        p = path.casefold()
        if any(w == name or w == p or p.endswith("/" + w) for w in wanted):
            chosen.append((path, folder))
    if not chosen and not settings.search_all_folders:
        for const in (6, 5):   # olFolderInbox, olFolderSentMail
            try:
                f = ns.GetDefaultFolder(const)
                chosen.append((f.Name, f))
            except Exception:
                pass
    return chosen


def _classify(text: str, settings: OutlookSettings) -> tuple[bool, bool, list[str]]:
    low = (text or "").casefold()
    rejected = [k for k in settings.rejection_keywords if k and k.casefold() in low]
    approved = [k for k in settings.approval_keywords if k and k.casefold() in low]
    return bool(approved), bool(rejected), approved + rejected


def search_mails(ncr_no: str, settings: OutlookSettings, progress: ProgressCb = None) -> list[MailHit]:
    """Search mails whose subject/body contain `ncr_no`. Must be called from a thread that owns COM."""
    if not ncr_no:
        return []
    if not outlook_available():
        raise RuntimeError("Outlook erişimi yok (Windows + pywin32 + Outlook gerekir).")
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    hits: list[MailHit] = []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        folders = _select_folders(ns, settings)
        term = _dasl_escape(ncr_no)
        parts = [f"\"urn:schemas:httpmail:subject\" LIKE '%{term}%'"]
        if settings.search_body:
            parts.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{term}%'")
        filt = "@SQL=(" + " OR ".join(parts) + ")"
        if settings.days_back and settings.days_back > 0:
            since = (_dt.datetime.now() - _dt.timedelta(days=settings.days_back)).strftime("%Y-%m-%d %H:%M")
            filt_dated = filt[:-1] + f" AND \"urn:schemas:httpmail:datereceived\" >= '{since}')"
        else:
            filt_dated = filt
        for path, folder in folders:
            if progress:
                progress(f"Outlook aranıyor: {path}")
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                try:
                    found = items.Restrict(filt_dated)
                except Exception:
                    found = items.Restrict(filt)
            except Exception:
                continue
            count = 0
            for item in found:
                if count >= settings.max_results:
                    break
                try:
                    if getattr(item, "Class", 43) != 43:   # olMail
                        continue
                    subject = str(item.Subject or "")
                    try:
                        body = str(item.Body or "")
                    except Exception:
                        body = ""
                    approved, rejected, kws = _classify(subject + "\n" + body, settings)
                    received = None
                    try:
                        rt = item.ReceivedTime
                        received = _dt.datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)
                    except Exception:
                        pass
                    snippet = " ".join(body.split())[:300]
                    try:
                        store_id = item.Parent.StoreID
                    except Exception:
                        store_id = ""
                    hits.append(MailHit(
                        entry_id=item.EntryID, store_id=store_id, subject=subject,
                        sender=str(getattr(item, "SenderName", "") or ""), received=received,
                        folder=path, approved=approved, rejected=rejected, snippet=snippet,
                        has_attachments=bool(getattr(item, "Attachments", None) and item.Attachments.Count > 0),
                        matched_keywords=kws))
                    count += 1
                except Exception:
                    continue
    finally:
        pythoncom.CoUninitialize()
    seen = set()
    unique = []
    for h in hits:
        if h.entry_id in seen:
            continue
        seen.add(h.entry_id)
        unique.append(h)
    unique.sort(key=lambda h: h.received or _dt.datetime.min, reverse=True)
    return unique


def open_mail(entry_id: str, store_id: str = "") -> None:
    """Display a mail in Outlook (call from GUI thread)."""
    if not outlook_available():
        raise RuntimeError("Outlook erişimi yok.")
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    item = ns.GetItemFromID(entry_id, store_id) if store_id else ns.GetItemFromID(entry_id)
    item.Display()


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
