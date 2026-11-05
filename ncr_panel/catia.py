"""CATIA V5 helpers: list models, match them to part numbers, start CATIA, open a model.

Stability rules (CATIA froze / broke before):
  * Opening a model NEVER starts CATIA. It only attaches to the running instance with
    GetActiveObject. `Dispatch("CATIA.Application")` is not used because it can launch a second,
    environment-less CATIA when the running one is not found.
  * A model that is already open is activated instead of opened again (re-opening shows a modal
    dialog that blocks the COM call and makes CATIA look frozen).
  * File alerts are switched off during Open so no modal dialog can block the call.
  * CATIA is started with the user's own launch arguments (-env / -direnv) or via their shortcut.
"""
from __future__ import annotations

import os
import shlex
import subprocess
from pathlib import Path
from typing import Callable, Iterable, Optional

from .config import CatiaSettings
from .utils import IS_WINDOWS, norm_path, open_path, process_running

ProgressCb = Optional[Callable[[str], None]]


def list_models(settings: CatiaSettings, progress: ProgressCb = None) -> list[Path]:
    root = Path(norm_path(settings.models_folder))
    if not settings.models_folder or not root.is_dir():
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


def match_models(models: Iterable[Path], part_numbers: Iterable[str], root: str = "") -> dict[str, list[str]]:
    """{model path: [matched part numbers]} for models whose name (or sub-folder) contains a part
    number from the discrepancy. Dash suffixes are ignored (123A4567-1 matches 123A4567.CATPart and
    123A4567-1.CATPart) because model files are often named after the base number."""
    keys: dict[str, str] = {}
    for p in part_numbers:
        p = p.upper().strip()
        if p:
            keys.setdefault(p.split("-")[0], p)
    if not keys:
        return {}
    root_p = Path(root) if root else None
    out: dict[str, list[str]] = {}
    for m in models:
        try:
            rel = m.relative_to(root_p) if root_p else Path(m.name)
        except ValueError:
            rel = Path(m.name)
        hay = str(rel).upper().replace("_", " ")
        hits = [full for base, full in keys.items() if base in hay]
        if hits:
            out[str(m)] = hits
    return out


def catia_exe_name(settings: CatiaSettings) -> str:
    name = Path(settings.exe_path).name
    return name if name.lower().endswith(".exe") else "CNEXT.exe"


def catia_running(settings: CatiaSettings) -> bool:
    return process_running(catia_exe_name(settings))


def start_catia(settings: CatiaSettings, progress: ProgressCb = None) -> bool:
    """Start CATIA once. Returns False if it was already running."""
    if catia_running(settings):
        return False
    exe = Path(norm_path(settings.exe_path))
    if not exe.is_file():
        raise FileNotFoundError(f"CATIA bulunamadı: {exe}\nAyarlar → CATIA'dan CNEXT.exe veya CATIA kısayolunu (.lnk) seçin.")
    if progress:
        progress("CATIA V5 başlatılıyor…")
    if exe.suffix.lower() in (".lnk", ".bat", ".cmd"):
        open_path(exe)                       # the shortcut already carries -env / -direnv
    else:
        args = shlex.split(settings.launch_args, posix=False) if settings.launch_args.strip() else []
        subprocess.Popen([str(exe), *args], cwd=str(exe.parent))
    return True


def _norm(p: str) -> str:
    return os.path.normcase(os.path.normpath(p))


def open_model(path: str, settings: CatiaSettings, progress: ProgressCb = None) -> str:
    """Open `path` in the RUNNING CATIA. Never starts CATIA. Returns a status message."""
    model = Path(norm_path(path))
    if not model.is_file():
        raise FileNotFoundError(f"Model bulunamadı: {model}")
    if not IS_WINDOWS:
        open_path(model)
        return "Windows dışı ortam: dosya varsayılan uygulamayla açıldı."
    import win32com.client  # type: ignore
    try:
        catia = win32com.client.GetActiveObject("CATIA.Application")
    except Exception:
        if catia_running(settings):
            raise RuntimeError("CATIA açık ama bağlanılamadı. CATIA tamamen açılana kadar bekleyin; CATIA "
                               "yönetici olarak çalışıyorsa paneli de aynı şekilde çalıştırın.")
        raise RuntimeError("CATIA açık değil. Önce 'CATIA'yı başlat' ile açın, açıldıktan sonra modeli seçin.")
    target = _norm(str(model))
    # already open? -> just bring its window to front
    try:
        for i in range(1, catia.Windows.Count + 1):
            win = catia.Windows.Item(i)
            try:
                if _norm(win.Parent.FullName) == target:
                    win.Activate()
                    return f"{model.name} zaten açık, penceresi öne getirildi."
            except Exception:
                continue
    except Exception:
        pass
    if progress:
        progress(f"Model açılıyor: {model.name}")
    old_alerts = None
    try:
        old_alerts = catia.DisplayFileAlerts
        catia.DisplayFileAlerts = False
    except Exception:
        pass
    try:
        catia.Documents.Open(str(model))
    finally:
        if old_alerts is not None:
            try:
                catia.DisplayFileAlerts = old_alerts
            except Exception:
                pass
        del catia
    return f"{model.name} CATIA'da açıldı."
