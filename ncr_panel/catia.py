"""Launch CATIA V5 and open a model in it."""
from __future__ import annotations

import os
import subprocess
import time
from pathlib import Path
from typing import Callable, Optional

from .config import CatiaSettings
from .utils import IS_WINDOWS, open_path

ProgressCb = Optional[Callable[[str], None]]


def list_models(settings: CatiaSettings) -> list[Path]:
    root = Path(settings.models_folder)
    if not root.is_dir():
        return []
    exts = {e.lower() if e.startswith(".") else "." + e.lower() for e in settings.extensions}
    out: list[Path] = []
    if settings.recursive:
        for dirpath, _d, files in os.walk(root):
            for fn in files:
                if Path(fn).suffix.lower() in exts:
                    out.append(Path(dirpath) / fn)
    else:
        for p in root.iterdir():
            if p.is_file() and p.suffix.lower() in exts:
                out.append(p)
    out.sort(key=lambda p: (str(p.parent).casefold(), p.name.casefold()))
    return out


def catia_running(exe_name: str = "CNEXT.exe") -> bool:
    if not IS_WINDOWS:
        return False
    try:
        out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {exe_name}", "/NH"],
                             capture_output=True, text=True, timeout=10,
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
        return exe_name.lower() in out.lower()
    except Exception:
        return False


def start_catia(settings: CatiaSettings, progress: ProgressCb = None) -> bool:
    """Start CATIA if it is not running. Returns True if it was started now."""
    exe = Path(settings.exe_path)
    name = exe.name or "CNEXT.exe"
    if catia_running(name):
        return False
    if not exe.is_file():
        raise FileNotFoundError(f"CATIA çalıştırılabilir dosyası bulunamadı: {exe}")
    if progress:
        progress("CATIA V5 başlatılıyor…")
    subprocess.Popen([str(exe)], cwd=str(exe.parent))
    deadline = time.time() + max(5, settings.startup_wait_s)
    while time.time() < deadline:
        if catia_running(name):
            break
        time.sleep(1)
    # give the UI a few more seconds to become responsive
    time.sleep(min(8, max(2, settings.startup_wait_s // 3)))
    return True


def open_model(path: str, settings: CatiaSettings, progress: ProgressCb = None) -> str:
    """Open `path` in CATIA V5 (starting CATIA first when needed). Returns a status message."""
    model = Path(path)
    if not model.is_file():
        raise FileNotFoundError(f"Model bulunamadı: {model}")
    if not IS_WINDOWS:
        open_path(model)
        return "Windows dışı ortam: dosya varsayılan uygulamayla açıldı."
    started = start_catia(settings, progress)
    if settings.use_com:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
            pythoncom.CoInitialize()
            try:
                catia = win32com.client.Dispatch("CATIA.Application")
                catia.Visible = True
                if progress:
                    progress(f"Model açılıyor: {model.name}")
                catia.Documents.Open(str(model))
                return f"{model.name} CATIA'da açıldı." + (" (CATIA başlatıldı)" if started else "")
            finally:
                pythoncom.CoUninitialize()
        except Exception as exc:
            if progress:
                progress(f"COM ile açılamadı ({exc}); dosya ilişkilendirmesi deneniyor…")
    open_path(model)
    return f"{model.name} açılması için CATIA'ya gönderildi."
