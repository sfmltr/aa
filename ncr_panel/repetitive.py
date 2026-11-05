"""Repetitive NCR detection: each sub-folder of the configured roots holds a note file
(e.g. notepad .txt) describing a known discrepancy; match it against the NCR's discrepancy."""
from __future__ import annotations

import fnmatch
import os
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Callable, Optional

from .config import RepetitiveSettings
from .excel_data import NcrRecord
from .utils import normalize_text

ProgressCb = Optional[Callable[[str], None]]


@dataclass
class RepEntry:
    folder: str
    note_path: str
    text: str
    phrases: list[str] = field(default_factory=list)   # normalized non-empty lines
    norm_text: str = ""

    @property
    def name(self) -> str:
        return os.path.basename(self.folder.rstrip("\\/")) or self.folder


@dataclass
class RepHit:
    entry: RepEntry
    score: float          # 1.0 exact phrase, else similarity ratio
    phrase: str           # the phrase that matched
    exact: bool


def _read(path: Path) -> str:
    data = path.read_bytes()
    for enc in ("utf-8-sig", "utf-16", "cp1254", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def scan_notes(settings: RepetitiveSettings, progress: ProgressCb = None) -> list[RepEntry]:
    entries: list[RepEntry] = []
    patterns = [p for p in settings.note_patterns if p.strip()] or ["*.txt"]
    for root in settings.folders:
        root_p = Path(root)
        if not root_p.is_dir():
            continue
        root_depth = len(root_p.parts)
        for dirpath, dirs, files in os.walk(root_p):
            depth = len(Path(dirpath).parts) - root_depth
            if depth >= settings.scan_depth:
                dirs[:] = []
            for fn in files:
                if any(fnmatch.fnmatch(fn.lower(), pat.lower()) for pat in patterns):
                    p = Path(dirpath) / fn
                    try:
                        text = _read(p)
                    except OSError:
                        continue
                    phrases = [normalize_text(line) for line in text.splitlines()]
                    phrases = [ph for ph in phrases if len(ph) >= settings.min_phrase_len]
                    entries.append(RepEntry(folder=dirpath, note_path=str(p), text=text,
                                            phrases=phrases, norm_text=normalize_text(text)))
            if progress:
                progress(f"Notlar taranıyor: {dirpath}")
    return entries


def match_record(rec: NcrRecord, entries: list[RepEntry], settings: RepetitiveSettings) -> list[RepHit]:
    disc = normalize_text(rec.discrepancy)
    if len(disc) < 3:
        return []
    hits: list[RepHit] = []
    for e in entries:
        best: Optional[RepHit] = None
        # exact: whole note or any line contained in the discrepancy (or vice versa)
        candidates = ([e.norm_text] if len(e.norm_text) >= settings.min_phrase_len else []) + e.phrases
        for ph in candidates:
            if ph and (ph in disc or (len(disc) >= settings.min_phrase_len and disc in ph)):
                best = RepHit(entry=e, score=1.0, phrase=ph, exact=True)
                break
        if best is None and settings.similarity_threshold < 1.0:
            for ph in candidates:
                sm = SequenceMatcher(None, disc, ph, autojunk=False)
                if sm.real_quick_ratio() < settings.similarity_threshold:
                    continue
                if sm.quick_ratio() < settings.similarity_threshold:
                    continue
                r = sm.ratio()
                if r >= settings.similarity_threshold and (best is None or r > best.score):
                    best = RepHit(entry=e, score=r, phrase=ph, exact=False)
        if best:
            hits.append(best)
    hits.sort(key=lambda h: -h.score)
    return hits


def annotate_records(records: list[NcrRecord], entries: list[RepEntry], settings: RepetitiveSettings) -> None:
    for rec in records:
        rec.repetitive_count = len(match_record(rec, entries, settings))
