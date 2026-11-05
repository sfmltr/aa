"""Index mails / PDFs / documents in user-chosen folders and match them to NCR records
by the part numbers that appear in the *Discrepancy* text (never by the Part No column)."""
from __future__ import annotations

import email
import email.policy
import gzip
import os
import pickle
import re
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Optional

from .config import CACHE_DIR, FileSearchSettings
from .excel_data import NcrRecord

ProgressCb = Optional[Callable[[str, int, int], None]]   # message, done, total
INDEX_FILE = CACHE_DIR / "file_index.pkl.gz"
_MAX_TEXT = 400_000


@dataclass
class FileEntry:
    path: str
    mtime: float
    size: int
    kind: str
    text: str = ""              # lower-cased text (capped)
    tokens: set[str] = field(default_factory=set)   # part-number-like tokens (upper)
    error: str = ""

    @property
    def name(self) -> str:
        return os.path.basename(self.path)


@dataclass
class FileHit:
    entry: FileEntry
    reasons: list[str]          # e.g. ["Parça no 123A4567-1", "NCR NO"]
    by_part: bool
    by_ncr: bool


# ---------------------------------------------------------------------------
# text extraction
# ---------------------------------------------------------------------------
def _read_txt(path: Path) -> str:
    data = path.read_bytes()
    for enc in ("utf-8-sig", "utf-16", "cp1254", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def _read_pdf(path: Path, max_pages: int) -> str:
    from pypdf import PdfReader
    reader = PdfReader(str(path))
    parts = []
    for i, page in enumerate(reader.pages):
        if i >= max_pages:
            break
        try:
            parts.append(page.extract_text() or "")
        except Exception:
            continue
    return "\n".join(parts)


def _read_docx(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
    xml = re.sub(r"</w:p>", "\n", xml)
    return re.sub(r"<[^>]+>", " ", xml)


def _read_eml(path: Path) -> str:
    with open(path, "rb") as fh:
        msg = email.message_from_binary_file(fh, policy=email.policy.default)
    parts = [str(msg.get("Subject", "")), str(msg.get("From", "")), str(msg.get("To", ""))]
    for part in msg.walk():
        ctype = part.get_content_type()
        if ctype == "text/plain":
            try:
                parts.append(part.get_content())
            except Exception:
                pass
        elif ctype == "text/html":
            try:
                parts.append(re.sub(r"<[^>]+>", " ", part.get_content()))
            except Exception:
                pass
        fname = part.get_filename()
        if fname:
            parts.append(fname)
    return "\n".join(parts)


def _read_msg(path: Path) -> str:
    try:
        import extract_msg  # type: ignore
    except ImportError:
        extract_msg = None
    if extract_msg is not None:
        msg = extract_msg.Message(str(path))
        try:
            parts = [msg.subject or "", msg.sender or "", msg.to or "", msg.body or ""]
            for att in msg.attachments:
                try:
                    parts.append(att.longFilename or att.shortFilename or "")
                except Exception:
                    pass
            return "\n".join(str(p) for p in parts)
        finally:
            try:
                msg.close()
            except Exception:
                pass
    # fallback: minimal reader with olefile (subject, body, display names)
    import olefile
    ole = olefile.OleFileIO(str(path))
    try:
        parts = []
        for stream in ole.listdir(streams=True, storages=False):
            name = stream[-1]
            if not name.startswith("__substg1.0_"):
                continue
            tag = name[12:16].upper()
            if tag in ("0037", "1000", "0C1A", "0E04", "3001", "3707"):   # subject, body, sender, to, display, att name
                typ = name[16:20].upper()
                raw = ole.openstream(stream).read()
                if typ == "001F":
                    parts.append(raw.decode("utf-16-le", errors="ignore"))
                elif typ == "001E":
                    parts.append(raw.decode("cp1254", errors="ignore"))
        return "\n".join(parts)
    finally:
        ole.close()


def extract_text(path: Path, settings: FileSearchSettings) -> str:
    ext = path.suffix.lower()
    if ext == ".pdf":
        return _read_pdf(path, settings.max_pdf_pages)
    if ext == ".docx":
        return _read_docx(path)
    if ext == ".eml":
        return _read_eml(path)
    if ext == ".msg":
        return _read_msg(path)
    return _read_txt(path)


# ---------------------------------------------------------------------------
# index
# ---------------------------------------------------------------------------
class FileIndex:
    def __init__(self):
        self.entries: dict[str, FileEntry] = {}
        self.built_at: float = 0.0

    # persistence ---------------------------------------------------------
    @classmethod
    def load(cls) -> "FileIndex":
        idx = cls()
        try:
            if INDEX_FILE.exists():
                with gzip.open(INDEX_FILE, "rb") as fh:
                    data = pickle.load(fh)
                if isinstance(data, dict):
                    idx.entries = {k: v for k, v in data.get("entries", {}).items() if isinstance(v, FileEntry)}
                    idx.built_at = float(data.get("built_at", 0))
        except Exception:
            idx.entries = {}
        return idx

    def save(self) -> None:
        try:
            INDEX_FILE.parent.mkdir(parents=True, exist_ok=True)
            with gzip.open(INDEX_FILE, "wb") as fh:
                pickle.dump({"entries": self.entries, "built_at": self.built_at}, fh)
        except Exception:
            pass

    # building ------------------------------------------------------------
    @staticmethod
    def iter_files(folders: Iterable[str], extensions: Iterable[str]) -> list[Path]:
        exts = {e.lower() if e.startswith(".") else "." + e.lower() for e in extensions}
        out: list[Path] = []
        for folder in folders:
            root = Path(folder)
            if not root.is_dir():
                continue
            for dirpath, _dirs, files in os.walk(root):
                for fn in files:
                    if Path(fn).suffix.lower() in exts:
                        out.append(Path(dirpath) / fn)
        return out

    def rebuild(self, settings: FileSearchSettings, progress: ProgressCb = None,
                cancel: Optional[Callable[[], bool]] = None) -> int:
        """Scan folders; re-extract only changed files. Returns number of (re)read files."""
        files = self.iter_files(settings.folders, settings.extensions)
        try:
            rx = re.compile(settings.part_regex, re.IGNORECASE)
        except re.error:
            rx = re.compile(r"\b\d{3}[A-Z]\d{4}(?:-\d+)*\b", re.IGNORECASE)
        keep: dict[str, FileEntry] = {}
        changed = 0
        total = len(files)
        max_bytes = settings.max_file_mb * 1024 * 1024
        for i, path in enumerate(files):
            if cancel and cancel():
                break
            key = str(path)
            try:
                st = path.stat()
            except OSError:
                continue
            old = self.entries.get(key)
            if old and old.mtime == st.st_mtime and old.size == st.st_size and not old.error:
                keep[key] = old
                continue
            if progress and (i % 10 == 0 or i == total - 1):
                progress(f"Dosya okunuyor: {path.name}", i + 1, total)
            entry = FileEntry(path=key, mtime=st.st_mtime, size=st.st_size, kind=path.suffix.lower())
            if st.st_size > max_bytes:
                entry.error = "çok büyük"
            else:
                try:
                    text = extract_text(path, settings)
                except Exception as exc:
                    entry.error = f"{type(exc).__name__}: {exc}"[:200]
                    text = ""
                full = (path.name + "\n" + text)
                # underscores are common separators in file names: 456D7890-3_report.pdf
                entry.tokens = {m.group(0).upper() for m in rx.finditer(full.upper().replace("_", " "))}
                entry.text = full.casefold()[:_MAX_TEXT]
            keep[key] = entry
            changed += 1
            time.sleep(0.002)       # release the GIL so the GUI stays fluid during indexing
        self.entries = keep
        self.built_at = time.time()
        return changed

    # matching ------------------------------------------------------------
    def match_record(self, rec: NcrRecord, settings: FileSearchSettings) -> list[FileHit]:
        parts = [p.upper() for p in rec.part_numbers]
        bases = set()
        if settings.match_base_number:
            for p in parts:
                bases.add(p.split("-")[0])
        ncr = rec.ncr_no.casefold() if settings.include_ncr_no and rec.ncr_no else ""
        hits: list[FileHit] = []
        for entry in self.entries.values():
            reasons = []
            by_part = False
            for p in parts:
                if p in entry.tokens:
                    reasons.append(f"Parça no {p}")
                    by_part = True
            if bases:
                for b in bases:
                    if any(t == b or t.startswith(b + "-") for t in entry.tokens):
                        if f"Parça no {b} (kök)" not in reasons:
                            reasons.append(f"Parça no {b} (kök)")
                            by_part = True
            by_ncr = False
            if ncr and len(ncr) >= 3 and ncr in entry.text:
                reasons.append(f"NCR {rec.ncr_no}")
                by_ncr = True
            if reasons:
                hits.append(FileHit(entry=entry, reasons=reasons, by_part=by_part, by_ncr=by_ncr))
        hits.sort(key=lambda h: (not h.by_part, -h.entry.mtime))
        return hits

    def annotate_records(self, records: list[NcrRecord], settings: FileSearchSettings) -> None:
        for rec in records:
            hits = self.match_record(rec, settings)
            rec.file_match_count = sum(1 for h in hits if h.by_part)
            rec.ncr_file_match_count = sum(1 for h in hits if h.by_ncr)
