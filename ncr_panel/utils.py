"""Small cross-platform helpers."""
from __future__ import annotations

import os
import re
import subprocess
import sys
import unicodedata
from pathlib import Path

IS_WINDOWS = sys.platform == "win32"


def open_path(path: str | os.PathLike) -> None:
    """Open a file or folder with the OS default handler."""
    path = str(path)
    if IS_WINDOWS:
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def reveal_in_folder(path: str | os.PathLike) -> None:
    """Open the containing folder and select the file where supported."""
    p = Path(path)
    if IS_WINDOWS and p.is_file():
        subprocess.Popen(["explorer", "/select,", str(p)])
    else:
        open_path(p if p.is_dir() else p.parent)


def normalize_text(text: str) -> str:
    """Lower-case, strip accents/punctuation and collapse whitespace for fuzzy comparisons."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.casefold()
    text = re.sub(r"[^\w\s\-]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def col_letter_to_index(letter: str) -> int:
    """'A' -> 1, 'Q' -> 17."""
    letter = (letter or "").strip().upper()
    n = 0
    for ch in letter:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Geçersiz kolon harfi: {letter!r}")
        n = n * 26 + (ord(ch) - 64)
    if n == 0:
        raise ValueError("Kolon harfi boş")
    return n


def contrast_text_color(hex_color: str) -> str:
    """Return '#000000' or '#FFFFFF' depending on background luminance."""
    try:
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except (ValueError, IndexError):
        return "#000000"
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    return "#000000" if lum > 150 else "#FFFFFF"


def human_size(num: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024:
            return f"{num:.0f} {unit}"
        num /= 1024
    return f"{num:.1f} TB"


def safe_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()
