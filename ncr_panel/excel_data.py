"""Read the NCR tracking workbook: values, row colours, engineers, categories; run macros.

Two loaders:
  * openpyxl  - works everywhere, reads static cell fills (incl. theme/indexed colours)
  * Excel COM - Windows only, reads *displayed* colours (conditional formatting included)
"""
from __future__ import annotations

import datetime as _dt
import re
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from .config import AppConfig, ExcelSettings
from .utils import IS_WINDOWS, col_letter_to_index, safe_str

ProgressCb = Optional[Callable[[str], None]]


@dataclass
class NcrRecord:
    row: int
    ncr_no: str
    part_no: str
    discrepancy: str
    cn_csi: str
    note: str
    engineer: str
    status: str
    color: Optional[str]                      # "#RRGGBB" or None
    categories: list[str] = field(default_factory=list)
    part_numbers: list[str] = field(default_factory=list)
    extra: dict[str, str] = field(default_factory=dict)   # header -> value (all columns)
    file_match_count: int = 0
    ncr_file_match_count: int = 0
    repetitive_count: int = 0
    mail_count: int = 0
    mail_approved: bool = False

    @property
    def key(self) -> str:
        return f"{self.row}:{self.ncr_no}"

    def kind(self, cfg: ExcelSettings) -> str:
        s = self.status.casefold()
        if cfg.status_mrb.casefold() in s:
            return "mrb"
        if cfg.status_external.casefold() in s:
            return "external"
        return "other"


@dataclass
class ExcelData:
    records: list[NcrRecord]
    engineers: list[str]
    headers: list[str]
    sheet_name: str
    header_row: int
    loaded_at: _dt.datetime
    source: str                                  # "openpyxl" | "com"
    color_counts: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    def by_kind(self, cfg: ExcelSettings, kind: str) -> list[NcrRecord]:
        return [r for r in self.records if r.kind(cfg) == kind]


# ---------------------------------------------------------------------------
# colour helpers
# ---------------------------------------------------------------------------
_THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                "accent4", "accent5", "accent6", "hlink", "folHlink"]


def _load_theme_colors(wb) -> list[str]:
    """Return the 12 theme colours (RRGGBB) of the workbook, Office defaults if missing."""
    defaults = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
                "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"]
    try:
        from openpyxl.xml.functions import fromstring
        xml = wb.loaded_theme
        if not xml:
            return defaults
        root = fromstring(xml)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = root.find(".//a:clrScheme", ns)
        if scheme is None:
            return defaults
        colors = []
        for name in _THEME_ORDER:
            el = scheme.find(f"a:{name}", ns)
            val = None
            if el is not None:
                srgb = el.find("a:srgbClr", ns)
                sysc = el.find("a:sysClr", ns)
                if srgb is not None:
                    val = srgb.get("val")
                elif sysc is not None:
                    val = sysc.get("lastClr")
            colors.append((val or defaults[len(colors)]).upper()[-6:])
        return colors
    except Exception:  # pragma: no cover - defensive
        return defaults


def _apply_tint(rgb: str, tint: float) -> str:
    try:
        r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    except ValueError:
        return rgb
    out = []
    for c in (r, g, b):
        if tint < 0:
            c = c * (1 + tint)
        elif tint > 0:
            c = c + (255 - c) * tint
        out.append(max(0, min(255, int(round(c)))))
    return "".join(f"{c:02X}" for c in out)


def cell_fill_hex(cell, theme: list[str]) -> Optional[str]:
    """Static fill colour of an openpyxl cell as '#RRGGBB' (None for no fill / white)."""
    try:
        fill = cell.fill
        if fill is None or fill.fill_type is None:
            return None
        color = fill.fgColor if fill.fill_type == "solid" else fill.fgColor
        if color is None:
            return None
        rgb: Optional[str] = None
        if color.type == "rgb":
            val = color.rgb
            if not isinstance(val, str) or val in ("00000000",):
                return None
            rgb = val[-6:]
        elif color.type == "theme":
            idx = int(color.theme or 0)
            base = theme[idx] if 0 <= idx < len(theme) else "FFFFFF"
            rgb = _apply_tint(base, float(color.tint or 0.0))
        elif color.type == "indexed":
            from openpyxl.styles.colors import COLOR_INDEX
            idx = int(color.indexed or 0)
            if idx in (64, 65) or idx >= len(COLOR_INDEX):
                return None
            rgb = COLOR_INDEX[idx][-6:]
        if not rgb:
            return None
        rgb = rgb.upper()
        if rgb == "FFFFFF":
            return None
        return "#" + rgb
    except Exception:
        return None


def bgr_int_to_hex(value) -> Optional[str]:
    """Excel COM Interior.Color (BGR integer) -> '#RRGGBB'."""
    try:
        v = int(value)
    except (TypeError, ValueError):
        return None
    if v < 0:
        return None
    r, g, b = v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF
    if (r, g, b) == (255, 255, 255):
        return None
    return f"#{r:02X}{g:02X}{b:02X}"


# ---------------------------------------------------------------------------
# record construction (shared by both loaders)
# ---------------------------------------------------------------------------
def extract_part_numbers(text: str, pattern: str) -> list[str]:
    if not text:
        return []
    try:
        rx = re.compile(pattern, re.IGNORECASE)
    except re.error:
        return []
    seen: list[str] = []
    for m in rx.finditer(str(text).upper().replace("_", " ")):
        token = m.group(0).upper()
        if token not in seen:
            seen.append(token)
    return seen


def categorize(text: str, categories: dict[str, list[str]]) -> list[str]:
    low = (text or "").casefold()
    hits = []
    for name, keywords in categories.items():
        for kw in keywords:
            if kw and kw.casefold() in low:
                hits.append(name)
                break
    return hits


def _detect_header_row(rows: list[list], ncr_idx: int, max_scan: int = 30) -> int:
    """1-based header row: first row whose NCR column mentions 'NCR'. Fallback 1."""
    for i, row in enumerate(rows[:max_scan], start=1):
        if ncr_idx - 1 < len(row):
            v = safe_str(row[ncr_idx - 1]).upper()
            if "NCR" in v:
                return i
    return 1


def _build_records(rows: list[list], colors: list[Optional[str]], cfg: AppConfig,
                   header_row: int) -> tuple[list[NcrRecord], list[str]]:
    ex = cfg.excel
    idx = {
        "note": col_letter_to_index(ex.col_note),
        "eng": col_letter_to_index(ex.col_engineer),
        "ncr": col_letter_to_index(ex.col_ncr),
        "disc": col_letter_to_index(ex.col_discrepancy),
        "status": col_letter_to_index(ex.col_status),
        "part": col_letter_to_index(ex.col_part),
        "cncsi": col_letter_to_index(ex.col_cncsi),
    }
    header_vals = rows[header_row - 1] if header_row - 1 < len(rows) else []
    headers = [safe_str(h) for h in header_vals]

    def get(row: list, i: int) -> str:
        return safe_str(row[i - 1]) if i - 1 < len(row) else ""

    records: list[NcrRecord] = []
    for rnum in range(header_row + 1, len(rows) + 1):
        row = rows[rnum - 1]
        ncr = get(row, idx["ncr"])
        if not ncr:
            continue
        disc = get(row, idx["disc"])
        note = get(row, idx["note"])
        extra = {}
        for ci, val in enumerate(row):
            name = headers[ci] if ci < len(headers) and headers[ci] else f"Kolon {ci + 1}"
            s = safe_str(val)
            if s:
                extra[name] = s
        records.append(NcrRecord(
            row=rnum,
            ncr_no=ncr,
            part_no=get(row, idx["part"]),
            discrepancy=disc,
            cn_csi=get(row, idx["cncsi"]),
            note=note,
            engineer=get(row, idx["eng"]),
            status=get(row, idx["status"]),
            color=colors[rnum - 1] if rnum - 1 < len(colors) else None,
            categories=categorize(note, cfg.ui.categories),
            part_numbers=extract_part_numbers(disc, cfg.files.part_regex),
            extra=extra,
        ))
    return records, headers


def _finish(records: list[NcrRecord], headers: list[str], sheet: str, header_row: int,
            source: str, warnings: list[str]) -> ExcelData:
    engineers = sorted({r.engineer for r in records if r.engineer}, key=str.casefold)
    counts: dict[str, int] = {}
    for r in records:
        if r.color:
            counts[r.color] = counts.get(r.color, 0) + 1
    return ExcelData(records=records, engineers=engineers, headers=headers, sheet_name=sheet,
                     header_row=header_row, loaded_at=_dt.datetime.now(), source=source,
                     color_counts=counts, warnings=warnings)


# ---------------------------------------------------------------------------
# openpyxl loader
# ---------------------------------------------------------------------------
def _color_column_indices(cfg: ExcelSettings) -> list[int]:
    out = []
    for part in (cfg.color_columns or "").split(","):
        part = part.strip()
        if part:
            try:
                out.append(col_letter_to_index(part))
            except ValueError:
                pass
    return out or [col_letter_to_index(cfg.col_ncr)]


def load_with_openpyxl(cfg: AppConfig, progress: ProgressCb = None) -> ExcelData:
    import openpyxl

    ex = cfg.excel
    src = Path(ex.path)
    if not src.is_file():
        raise FileNotFoundError(f"Excel dosyası bulunamadı: {src}")
    # copy first so an open/locked workbook can still be read
    tmpdir = Path(tempfile.mkdtemp(prefix="ncrpanel_"))
    tmp = tmpdir / src.name
    try:
        shutil.copy2(src, tmp)
        if progress:
            progress("Excel okunuyor…")
        wb = openpyxl.load_workbook(tmp, data_only=True, read_only=False, keep_vba=False)
        try:
            ws = wb[ex.sheet] if ex.sheet and ex.sheet in wb.sheetnames else wb.active
            theme = _load_theme_colors(wb)
            warnings: list[str] = []
            if ex.sheet and ex.sheet not in wb.sheetnames:
                warnings.append(f"Sayfa '{ex.sheet}' bulunamadı, aktif sayfa kullanıldı: {ws.title}")
            rows: list[list] = []
            colors: list[Optional[str]] = []
            color_cols = _color_column_indices(ex)
            max_col = ws.max_column
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=max_col):
                rows.append([c.value for c in row])
                col = None
                for ci in color_cols:
                    if ci - 1 < len(row):
                        col = cell_fill_hex(row[ci - 1], theme)
                        if col:
                            break
                colors.append(col)
            header_row = ex.header_row or _detect_header_row(rows, col_letter_to_index(ex.col_ncr))
            records, headers = _build_records(rows, colors, cfg, header_row)
            try:
                has_cf = bool(ws.conditional_formatting) and len(list(ws.conditional_formatting)) > 0
            except Exception:
                has_cf = False
            if has_cf:
                warnings.append("Sayfada koşullu biçimlendirme var; görüntülenen renkler için "
                                "Ayarlar → Excel → Renk kaynağı: 'Excel COM' seçin (Windows).")
            data = _finish(records, headers, ws.title, header_row, "openpyxl", warnings)
            data.warnings = warnings
            return data
        finally:
            wb.close()
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def workbook_has_conditional_formatting(path: str, sheet: str = "") -> bool:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
        try:
            ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
            return len(list(ws.conditional_formatting)) > 0
        finally:
            wb.close()
    except Exception:
        return False


def list_sheets(path: str) -> list[str]:
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Excel COM loader (Windows)
# ---------------------------------------------------------------------------
class _ExcelSession:
    """isolated=False: attach to the user's running Excel (needed to run the macro on the real file).
    isolated=True : always a separate hidden Excel instance, so reading never freezes the user's Excel."""

    def __init__(self, isolated: bool = False):
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
        pythoncom.CoInitialize()
        self._pythoncom = pythoncom
        self.owned = False
        self.app = None
        if not isolated:
            try:
                self.app = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                self.app = None
        if self.app is None:
            self.app = win32com.client.DispatchEx("Excel.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = False
            try:
                self.app.ScreenUpdating = False
                self.app.EnableEvents = False
            except Exception:
                pass
            self.owned = True
        self.opened_by_us = None

    def workbook(self, path: str, read_only: bool = True):
        target = Path(path)
        for wb in self.app.Workbooks:
            try:
                if Path(wb.FullName).resolve() == target.resolve() or wb.Name.lower() == target.name.lower():
                    return wb
            except Exception:
                continue
        wb = self.app.Workbooks.Open(str(target), ReadOnly=read_only, UpdateLinks=0)
        self.opened_by_us = wb
        return wb

    def close(self, save: bool = False):
        try:
            if self.opened_by_us is not None:
                self.opened_by_us.Close(SaveChanges=save)
            if self.owned:
                self.app.Quit()
        except Exception:
            pass
        finally:
            # drop every proxy before CoUninitialize, otherwise Excel can stay alive / hang
            self.opened_by_us = None
            self.app = None
            import gc
            gc.collect()
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass


def load_with_com(cfg: AppConfig, progress: ProgressCb = None) -> ExcelData:
    if not IS_WINDOWS:
        raise RuntimeError("Excel COM yalnızca Windows'ta kullanılabilir.")
    ex = cfg.excel
    if not Path(ex.path).is_file():
        raise FileNotFoundError(f"Excel dosyası bulunamadı: {ex.path}")
    tmpdir = Path(tempfile.mkdtemp(prefix="ncrpanel_com_"))
    copy = tmpdir / Path(ex.path).name
    shutil.copy2(ex.path, copy)
    session = _ExcelSession(isolated=True)
    try:
        if progress:
            progress("Excel (ayrı gizli Excel) ile renkler okunuyor…")
        wb = session.workbook(str(copy), read_only=True)
        ws = None
        if ex.sheet:
            try:
                ws = wb.Worksheets(ex.sheet)
            except Exception:
                ws = None
        if ws is None:
            ws = wb.ActiveSheet
        used = ws.UsedRange
        values = used.Value2
        first_row = used.Row
        first_col = used.Column
        if values is None:
            values = ()
        elif not isinstance(values, (tuple, list)):
            values = ((values,),)
        rows: list[list] = [[None] * (first_col - 1 + len(values[0])) for _ in range(first_row - 1)]
        for r in values:
            rows.append([None] * (first_col - 1) + list(r))
        header_row = ex.header_row or _detect_header_row(rows, col_letter_to_index(ex.col_ncr))
        color_cols = _color_column_indices(ex)
        colors: list[Optional[str]] = [None] * len(rows)
        total = len(rows)
        for rnum in range(header_row + 1, total + 1):
            if progress and rnum % 200 == 0:
                progress(f"Renkler okunuyor… {rnum}/{total}")
            for ci in color_cols:
                try:
                    cell = ws.Cells(rnum, ci)
                    if cell.DisplayFormat.Interior.ColorIndex == -4142:   # xlNone
                        continue
                    col = bgr_int_to_hex(cell.DisplayFormat.Interior.Color)
                except Exception:
                    col = None
                if col:
                    colors[rnum - 1] = col
                    break
        records, headers = _build_records(rows, colors, cfg, header_row)
        sheet_name = ws.Name
        del ws, used, wb
        return _finish(records, headers, sheet_name, header_row, "com", [])
    finally:
        session.close(save=False)
        shutil.rmtree(tmpdir, ignore_errors=True)


def load_excel(cfg: AppConfig, progress: ProgressCb = None) -> ExcelData:
    mode = (cfg.excel.color_source or "auto").lower()
    if mode == "com":
        return load_with_com(cfg, progress)
    if mode == "auto" and IS_WINDOWS:
        try:
            import win32com.client  # noqa: F401
            if workbook_has_conditional_formatting(cfg.excel.path, cfg.excel.sheet):
                try:
                    return load_with_com(cfg, progress)
                except Exception as exc:  # fall back silently
                    data = load_with_openpyxl(cfg, progress)
                    data.warnings.append(f"Excel COM okuması başarısız, openpyxl kullanıldı: {exc}")
                    return data
        except ImportError:
            pass
    return load_with_openpyxl(cfg, progress)


# ---------------------------------------------------------------------------
# macro
# ---------------------------------------------------------------------------
def run_macro(cfg: ExcelSettings, progress: ProgressCb = None) -> str:
    """Run `cfg.macro_name` inside the workbook and save it. Returns a status message."""
    if not IS_WINDOWS:
        raise RuntimeError("Makro çalıştırma yalnızca Windows + Excel ile mümkündür.")
    try:
        import win32com.client  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("pywin32 kurulu değil: pip install pywin32") from exc
    if not Path(cfg.path).is_file():
        raise FileNotFoundError(f"Excel dosyası bulunamadı: {cfg.path}")
    session = _ExcelSession()
    try:
        if progress:
            progress("Excel açılıyor…")
        wb = session.workbook(cfg.path, read_only=False)
        if wb.ReadOnly and session.opened_by_us is None:
            # the user has it open read-only; still try, macro may not need to save
            pass
        if progress:
            progress(f"Makro çalışıyor: {cfg.macro_name}")
        name = cfg.macro_name.strip()
        if "!" not in name:
            name = f"'{wb.Name}'!{name}"
        session.app.Run(name)
        try:
            wb.Save()
            saved = "kaydedildi"
        except Exception as exc:
            saved = f"kaydedilemedi ({exc})"
        return f"{cfg.macro_name} çalıştırıldı, çalışma kitabı {saved}."
    finally:
        session.close(save=False)
