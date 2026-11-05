"""Small cross-platform helpers."""
from __future__ import annotations

import os
import re
import subprocess
import sys
import unicodedata
from pathlib import Path

IS_WINDOWS = sys.platform == "win32"


def norm_path(path: str | os.PathLike) -> str:
    """Normalise a user path. On Windows this turns '//vds/share/x' (what Qt's file dialogs
    return for network shares) into '\\\\vds\\share\\x', which Explorer/ShellExecute understand."""
    s = str(path or "").strip().strip('"')
    if not s:
        return ""
    if IS_WINDOWS:
        import ntpath
        s = ntpath.normpath(s.replace("/", "\\"))
    return s


def open_path(path: str | os.PathLike) -> None:
    """Open a file or folder with the OS default handler (UNC paths supported)."""
    path = norm_path(path)
    if IS_WINDOWS:
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except OSError:
            # ShellExecute sometimes refuses UNC folders; Explorer itself never does
            if os.path.isdir(path) or path.startswith("\\\\"):
                subprocess.Popen(["explorer", path])
            else:
                raise
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def reveal_in_folder(path: str | os.PathLike) -> None:
    """Open the containing folder and select the file where supported."""
    p = Path(norm_path(path))
    if IS_WINDOWS and p.is_file():
        subprocess.Popen(["explorer", "/select,", str(p)])
    else:
        open_path(p if p.is_dir() else p.parent)


def process_running(image_name: str) -> bool:
    """True if a process with this executable name runs (Windows only)."""
    if not IS_WINDOWS or not image_name:
        return False
    try:
        out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/NH"],
                             capture_output=True, text=True, timeout=10,
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
        return image_name.lower() in out.lower()
    except Exception:
        return False


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
