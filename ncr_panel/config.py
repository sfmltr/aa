"""Application settings: dataclasses + JSON persistence.

All user-tunable behaviour lives here so the settings dialog can expose it.
Unknown keys in an old config file are ignored; missing keys fall back to defaults.
"""
from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass, field, fields, is_dataclass
from pathlib import Path
from typing import Any

APP_NAME = "NCR Panel"


def _config_dir() -> Path:
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home()))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / "NCRPanel"


CONFIG_DIR = _config_dir()
CONFIG_FILE = CONFIG_DIR / "config.json"
CACHE_DIR = CONFIG_DIR / "cache"


@dataclass
class ExcelSettings:
    path: str = ""
    sheet: str = ""                 # "" -> active sheet
    header_row: int = 0             # 0 -> auto detect
    col_note: str = "A"
    col_engineer: str = "C"
    col_ncr: str = "D"
    col_discrepancy: str = "G"
    col_status: str = "H"
    col_part: str = "J"
    col_cncsi: str = "Q"
    status_mrb: str = "MRB-ENG"
    status_external: str = "EXTERNAL APPROVER"
    macro_name: str = "B03_Update_NCRs"
    # auto: COM when workbook uses conditional formatting (Windows), else openpyxl
    color_source: str = "auto"      # auto | openpyxl | com
    color_columns: str = "D,G,J"    # cells whose fill decides the row colour (first coloured wins)
    reload_after_macro: bool = True


@dataclass
class OutlookSettings:
    enabled: bool = True
    folders: list[str] = field(default_factory=lambda: ["Inbox", "Sent Items", "Gelen Kutusu", "Gönderilmiş Öğeler"])
    search_all_folders: bool = False
    days_back: int = 730
    search_body: bool = True
    max_results: int = 150
    approval_keywords: list[str] = field(default_factory=lambda: [
        "approved", "concurred", "concur", "approve", "accepted", "agreed", "onaylandı", "uygundur"])
    rejection_keywords: list[str] = field(default_factory=lambda: [
        "not approved", "rejected", "disapproved", "not concurred", "reddedildi"])
    bulk_on_refresh: bool = True        # search every listed NCR when the Excel list is (re)loaded
    recheck_minutes: int = 3            # clicking an NCR re-checks new mail only if older than this
    start_outlook_if_closed: bool = False
    chunk_size: int = 40                # NCR numbers per Outlook query during the bulk scan


@dataclass
class FileSearchSettings:
    folders: list[str] = field(default_factory=list)
    extensions: list[str] = field(default_factory=lambda: [".msg", ".eml", ".pdf", ".txt", ".docx"])
    part_regex: str = r"\b\d{3}[A-Z]\d{4}(?:-\d+)*\b"
    match_base_number: bool = False     # also match 123A4567 when discrepancy says 123A4567-1
    include_ncr_no: bool = True         # also list files mentioning the NCR NO itself
    max_file_mb: int = 60
    max_pdf_pages: int = 150
    auto_index_on_start: bool = True


@dataclass
class CatiaSettings:
    exe_path: str = r"C:\Program Files\Dassault Systemes\B28\win_b64\code\bin\CNEXT.exe"
    models_folder: str = ""
    extensions: list[str] = field(default_factory=lambda: [
        ".CATPart", ".CATProduct", ".CATDrawing", ".cgr", ".stp", ".step", ".igs", ".model", ".CATAnalysis"])
    recursive: bool = True
    startup_wait_s: int = 25
    use_com: bool = True
    launch_args: str = ""               # e.g. -env B28 -direnv "C:\\ProgramData\\DassaultSystemes\\CATEnv"
    only_matching_models: bool = False  # model list shows only models matching the NCR's part numbers


@dataclass
class RepetitiveSettings:
    folders: list[str] = field(default_factory=list)
    note_patterns: list[str] = field(default_factory=lambda: ["*.txt"])
    similarity_threshold: float = 0.65
    min_phrase_len: int = 12
    scan_depth: int = 2


@dataclass
class UISettings:
    theme: str = "light"            # light | dark
    font_size: int = 10
    selected_engineer: str = ""
    auto_refresh_min: int = 0
    window_geometry: str = ""       # hex of saved geometry
    splitter_sizes: dict[str, list[int]] = field(default_factory=dict)
    quick_folders: list[str] = field(default_factory=list)
    categories: dict[str, list[str]] = field(default_factory=lambda: {
        "Consultation": ["consultation", "consult", "danışma"],
        "Stress": ["stress", "stres"],
        "SNN": ["snn"],
    })
    color_legend: dict[str, str] = field(default_factory=dict)   # "#RRGGBB" -> label
    show_other_status: bool = False  # show rows whose status is neither MRB-ENG nor EXTERNAL APPROVER


@dataclass
class AppConfig:
    excel: ExcelSettings = field(default_factory=ExcelSettings)
    outlook: OutlookSettings = field(default_factory=OutlookSettings)
    files: FileSearchSettings = field(default_factory=FileSearchSettings)
    catia: CatiaSettings = field(default_factory=CatiaSettings)
    repetitive: RepetitiveSettings = field(default_factory=RepetitiveSettings)
    ui: UISettings = field(default_factory=UISettings)

    # ---- persistence -------------------------------------------------
    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "AppConfig":
        cfg = cls()
        for f in fields(cfg):
            section = data.get(f.name)
            if isinstance(section, dict):
                setattr(cfg, f.name, _merge_dataclass(getattr(cfg, f.name), section))
        return cfg

    def save(self, path: Path = CONFIG_FILE) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
        tmp.replace(path)

    @classmethod
    def load(cls, path: Path = CONFIG_FILE) -> "AppConfig":
        if not path.exists():
            return cls()
        try:
            return cls.from_dict(json.loads(path.read_text(encoding="utf-8"))).normalize_paths()
        except (OSError, ValueError):
            return cls()

    def copy(self) -> "AppConfig":
        return AppConfig.from_dict(self.to_dict())

    def normalize_paths(self) -> "AppConfig":
        """Convert stored paths to native form (UNC '//server/x' -> '\\\\server\\x' on Windows)."""
        from .utils import norm_path
        self.excel.path = norm_path(self.excel.path)
        self.catia.models_folder = norm_path(self.catia.models_folder)
        self.catia.exe_path = norm_path(self.catia.exe_path)
        self.files.folders = [norm_path(p) for p in self.files.folders if p]
        self.repetitive.folders = [norm_path(p) for p in self.repetitive.folders if p]
        self.ui.quick_folders = [norm_path(p) for p in self.ui.quick_folders if p]
        return self


def _merge_dataclass(instance: Any, data: dict[str, Any]) -> Any:
    """Return a new dataclass instance with values from `data` overriding defaults."""
    if not is_dataclass(instance):
        return instance
    kwargs = {}
    for f in fields(instance):
        default = getattr(instance, f.name)
        if f.name in data and data[f.name] is not None:
            value = data[f.name]
            # keep the type of the default where it matters
            if isinstance(default, bool):
                value = bool(value)
            elif isinstance(default, int) and not isinstance(default, bool):
                try:
                    value = int(value)
                except (TypeError, ValueError):
                    value = default
            elif isinstance(default, float):
                try:
                    value = float(value)
                except (TypeError, ValueError):
                    value = default
            elif isinstance(default, list) and not isinstance(value, list):
                value = default
            elif isinstance(default, dict) and not isinstance(value, dict):
                value = default
            kwargs[f.name] = value
        else:
            kwargs[f.name] = default
    return type(instance)(**kwargs)
