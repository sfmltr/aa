"""Main window: three-column dashboard, toolbar, background loading orchestration."""
from __future__ import annotations

import threading
from pathlib import Path

from PySide6.QtCore import QByteArray, Qt, QTimer
from PySide6.QtGui import QAction, QKeySequence
from PySide6.QtWidgets import (QApplication, QLabel, QMainWindow, QMessageBox, QProgressBar, QSplitter,
                               QStatusBar, QToolBar, QWidget, QHBoxLayout)

from .. import __version__
from ..config import APP_NAME, AppConfig
from ..excel_data import ExcelData, NcrRecord, load_excel, run_macro
from ..file_search import FileIndex
from ..repetitive import annotate_records as annotate_repetitive, scan_notes
from ..utils import IS_WINDOWS, open_path
from .center_panel import CenterPanel
from .left_panel import LeftPanel
from .right_panel import RightPanel
from .settings_dialog import SettingsDialog
from .theme import build_stylesheet
from .workers import run_in_background


class MainWindow(QMainWindow):
    def __init__(self, cfg: AppConfig):
        super().__init__()
        self.cfg = cfg
        self.data: ExcelData | None = None
        self.index = FileIndex.load()
        self.notes = []
        self._index_worker = None
        self._index_cancel = threading.Event()
        self._excel_loading = False
        self.setWindowTitle(f"{APP_NAME} {__version__}")
        self.resize(1700, 950)

        root = QWidget()
        root.setObjectName("root")
        lay = QHBoxLayout(root)
        lay.setContentsMargins(8, 8, 8, 8)
        self.splitter = QSplitter(Qt.Horizontal)
        lay.addWidget(self.splitter)
        self.setCentralWidget(root)

        self.left = LeftPanel(cfg)
        self.center = CenterPanel(cfg, lambda: self.index)
        self.right = RightPanel(cfg)
        for w in (self.left, self.center, self.right):
            self.splitter.addWidget(w)
        self.splitter.setSizes([700, 600, 400])
        self.splitter.setStretchFactor(0, 4)
        self.splitter.setStretchFactor(1, 3)
        self.splitter.setStretchFactor(2, 2)

        self._build_toolbar()
        self._build_statusbar()
        self._connect()
        self._restore_layout()
        self.apply_theme()
        self._timer = QTimer(self)
        self._timer.timeout.connect(self.reload_excel)
        self._setup_timer()
        QTimer.singleShot(50, self.initial_load)

    # ---- UI chrome -----------------------------------------------------------
    def _build_toolbar(self):
        tb = QToolBar("Ana")
        tb.setMovable(False)
        tb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.addToolBar(tb)

        def act(text, slot, tip="", shortcut=None):
            a = QAction(text, self)
            a.triggered.connect(slot)
            if tip:
                a.setToolTip(tip)
            if shortcut:
                a.setShortcut(QKeySequence(shortcut))
            tb.addAction(a)
            return a

        self.a_reload = act("⟳ Excel'i yenile", self.reload_excel, "NCR listesini Excel'den yeniden oku", "F5")
        self.a_macro = act("▶ Makro çalıştır", self.run_macro, "Excel içindeki güncelleme makrosunu çalıştır", "F6")
        self.a_excel = act("📗 Excel'i aç", self.open_excel, "Çalışma kitabını Excel'de aç")
        tb.addSeparator()
        self.a_index = act("🗂 Dosya indeksi", self.rebuild_index, "Mail/PDF klasörlerini yeniden tara", "F7")
        self.a_notes = act("🔁 Notları tara", self.rescan_notes, "Tekrarlayan NCR notlarını yeniden oku", "F8")
        tb.addSeparator()
        self.a_theme = act("◐ Tema", self.toggle_theme, "Açık / koyu tema")
        self.a_settings = act("⚙ Ayarlar", self.open_settings, "Ayarlar", "Ctrl+,")
        spacer = QWidget()
        spacer.setSizePolicy(spacer.sizePolicy().horizontalPolicy().Expanding, spacer.sizePolicy().verticalPolicy())
        tb.addWidget(spacer)
        self.lbl_source = QLabel("")
        self.lbl_source.setObjectName("muted")
        tb.addWidget(self.lbl_source)

    def _build_statusbar(self):
        sb = QStatusBar()
        self.setStatusBar(sb)
        self.lbl_status = QLabel("Hazır")
        sb.addWidget(self.lbl_status, 1)
        self.progress = QProgressBar()
        self.progress.setMaximumWidth(220)
        self.progress.setTextVisible(False)
        self.progress.hide()
        sb.addPermanentWidget(self.progress)
        self.lbl_counts = QLabel("")
        self.lbl_counts.setObjectName("muted")
        sb.addPermanentWidget(self.lbl_counts)

    def _connect(self):
        self.left.recordSelected.connect(self.on_record_selected)
        self.left.refreshRequested.connect(self.reload_excel)
        self.left.runMacroRequested.connect(self.run_macro)
        self.left.openInExcelRequested.connect(self.open_excel_row)
        self.left.engineerChanged.connect(lambda _e: self.cfg.save())
        for p in (self.center, self.right):
            p.status.connect(self.set_status)
            p.error.connect(self.show_error)
        self.right.configChanged.connect(self.cfg.save)
        self.right.rescanNotesRequested.connect(self.rescan_notes)

    def _restore_layout(self):
        g = self.cfg.ui.window_geometry
        if g:
            try:
                self.restoreGeometry(QByteArray.fromHex(g.encode()))
            except Exception:
                pass
        sizes = self.cfg.ui.splitter_sizes
        for name, sp in (("main", self.splitter), ("left", self.left.splitter),
                         ("center", self.center.splitter), ("right", self.right.splitter)):
            if name in sizes and sizes[name]:
                sp.setSizes([int(x) for x in sizes[name]])

    def _save_layout(self):
        self.cfg.ui.window_geometry = bytes(self.saveGeometry().toHex()).decode()
        self.cfg.ui.splitter_sizes = {
            "main": self.splitter.sizes(), "left": self.left.splitter.sizes(),
            "center": self.center.splitter.sizes(), "right": self.right.splitter.sizes()}

    def closeEvent(self, event):
        self._save_layout()
        self.cfg.save()
        self._index_cancel.set()
        super().closeEvent(event)

    # ---- theme ---------------------------------------------------------------
    def apply_theme(self):
        QApplication.instance().setStyleSheet(build_stylesheet(self.cfg.ui.theme, self.cfg.ui.font_size))

    def toggle_theme(self):
        self.cfg.ui.theme = "dark" if self.cfg.ui.theme == "light" else "light"
        self.apply_theme()
        self.cfg.save()
        if self.data:
            self.left.set_data(self.data)

    def _setup_timer(self):
        self._timer.stop()
        mins = self.cfg.ui.auto_refresh_min
        if mins and mins > 0:
            self._timer.start(mins * 60 * 1000)

    # ---- status --------------------------------------------------------------
    def set_status(self, msg: str):
        self.lbl_status.setText(msg)

    def set_progress(self, done: int, total: int):
        if total <= 0:
            self.progress.hide()
            return
        self.progress.show()
        self.progress.setRange(0, total)
        self.progress.setValue(done)
        if done >= total:
            QTimer.singleShot(1500, self.progress.hide)

    def show_error(self, msg: str):
        self.set_status("Hata: " + msg.splitlines()[0][:120])
        QMessageBox.warning(self, APP_NAME, msg[:4000])

    # ---- loading -------------------------------------------------------------
    def initial_load(self):
        if not self.cfg.excel.path:
            self.set_status("Hoş geldiniz — önce Ayarlar'dan Excel dosyasını seçin.")
            QTimer.singleShot(300, self.open_settings)
        else:
            self.reload_excel()
        self.rescan_notes()
        if self.cfg.files.auto_index_on_start and self.cfg.files.folders:
            self.rebuild_index()

    def reload_excel(self):
        if self._excel_loading or not self.cfg.excel.path:
            return
        self._excel_loading = True
        self.a_reload.setEnabled(False)
        self.set_status("Excel okunuyor…")
        current = self.center.record.key if self.center.record else None

        def done(data: ExcelData):
            self._excel_loading = False
            self.a_reload.setEnabled(True)
            self.data = data
            self._annotate_all()
            self.left.set_data(data)
            n_mrb = len(data.by_kind(self.cfg.excel, "mrb"))
            n_ext = len(data.by_kind(self.cfg.excel, "external"))
            self.lbl_counts.setText(f"{len(data.records)} NCR  •  MRB-ENG {n_mrb}  •  EXTERNAL {n_ext}")
            self.lbl_source.setText(f"{Path(self.cfg.excel.path).name} / {data.sheet_name}  •  "
                                    f"{data.loaded_at:%H:%M:%S}  •  renk: {data.source}")
            msg = f"Excel yüklendi ({len(data.records)} kayıt)."
            if data.warnings:
                msg += "  ⚠ " + " | ".join(data.warnings)
            self.set_status(msg)
            if current:
                self.left.reselect(current)
            self.center.invalidate_mail_cache()

        def err(msg):
            self._excel_loading = False
            self.a_reload.setEnabled(True)
            self.show_error("Excel okunamadı:\n" + msg)

        run_in_background(load_excel, self.cfg, on_done=done, on_error=err, on_progress=self.set_status)

    def _annotate_all(self):
        if not self.data:
            return
        if self.index.entries:
            self.index.annotate_records(self.data.records, self.cfg.files)
        if self.notes:
            annotate_repetitive(self.data.records, self.notes, self.cfg.repetitive)
        self.left.refresh_badges()

    def rebuild_index(self):
        if not self.cfg.files.folders:
            self.set_status("Dosya arama klasörü tanımlı değil (Ayarlar → Dosya arama).")
            return
        if self._index_worker is not None:
            self.set_status("Dosya indeksi zaten güncelleniyor…")
            return
        self.a_index.setEnabled(False)
        self.set_status("Dosya indeksi güncelleniyor…")
        idx = self.index
        settings = self.cfg.files
        cancel = self._index_cancel
        cancel.clear()

        def job(progress=None):
            changed = idx.rebuild(settings, progress=progress, cancel=cancel.is_set)
            idx.save()
            return changed

        def done(changed):
            self._index_worker = None
            self.a_index.setEnabled(True)
            self.set_progress(1, 1)
            self._annotate_all()
            self.center.refresh_files()
            self.set_status(f"Dosya indeksi hazır: {len(idx.entries)} dosya ({changed} yeni/değişmiş okundu).")

        def err(msg):
            self._index_worker = None
            self.a_index.setEnabled(True)
            self.show_error("Dosya indeksi hatası:\n" + msg)

        self._index_worker = run_in_background(job, on_done=done, on_error=err, on_progress=self.set_status,
                                               on_progress_pct=self.set_progress)

    def rescan_notes(self):
        if not self.cfg.repetitive.folders:
            self.right.set_entries([])
            return
        settings = self.cfg.repetitive

        def done(entries):
            self.notes = entries
            self.right.set_entries(entries)
            self._annotate_all()
            if self.center.record:
                self.right.show_record(self.center.record)
            self.set_status(f"Tekrarlayan NCR notları okundu: {len(entries)} not.")

        run_in_background(scan_notes, settings, on_done=done,
                          on_error=lambda m: self.show_error("Not taraması hatası:\n" + m),
                          on_progress=self.set_status)

    def run_macro(self):
        if not self.cfg.excel.path:
            self.show_error("Önce Ayarlar'dan Excel dosyasını seçin.")
            return
        if not IS_WINDOWS:
            self.show_error("Makro çalıştırma yalnızca Windows'ta (Excel ile) mümkündür.")
            return
        self.a_macro.setEnabled(False)
        self.left.btn_macro.setEnabled(False)
        self.set_status(f"Makro çalıştırılıyor: {self.cfg.excel.macro_name}…")

        def done(msg):
            self.a_macro.setEnabled(True)
            self.left.btn_macro.setEnabled(True)
            self.set_status(msg)
            if self.cfg.excel.reload_after_macro:
                self.reload_excel()

        def err(msg):
            self.a_macro.setEnabled(True)
            self.left.btn_macro.setEnabled(True)
            self.show_error("Makro çalıştırılamadı:\n" + msg)

        run_in_background(run_macro, self.cfg.excel, on_done=done, on_error=err, on_progress=self.set_status)

    # ---- actions ---------------------------------------------------------------
    def on_record_selected(self, rec: NcrRecord):
        self.center.show_record(rec)
        self.right.show_record(rec)
        legend = self.cfg.ui.color_legend.get(rec.color or "", "")
        self.set_status(f"Seçili: {rec.ncr_no}" + (f"  •  renk: {legend}" if legend else ""))

    def open_excel(self):
        if not self.cfg.excel.path:
            self.show_error("Excel dosyası seçilmedi.")
            return
        try:
            open_path(self.cfg.excel.path)
        except Exception as exc:  # noqa: BLE001
            self.show_error(f"Excel açılamadı: {exc}")

    def open_excel_row(self, rec: NcrRecord):
        """Activate the workbook in Excel and select the record's NCR cell (Windows), else just open it."""
        if not IS_WINDOWS:
            self.open_excel()
            return
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                app = win32com.client.Dispatch("Excel.Application")
            app.Visible = True
            target = Path(self.cfg.excel.path)
            wb = None
            for w in app.Workbooks:
                if w.Name.lower() == target.name.lower():
                    wb = w
                    break
            if wb is None:
                wb = app.Workbooks.Open(str(target))
            ws = wb.Worksheets(self.data.sheet_name) if self.data else wb.ActiveSheet
            ws.Activate()
            ws.Range(f"{self.cfg.excel.col_ncr}{rec.row}").Select()
            try:
                import ctypes
                ctypes.windll.user32.SetForegroundWindow(app.Hwnd)
            except Exception:
                pass
        except Exception as exc:  # noqa: BLE001
            self.set_status(f"Excel'e gidilemedi ({exc}); dosya açılıyor.")
            self.open_excel()

    def open_settings(self):
        dlg = SettingsDialog(self.cfg, self.data, self)
        if dlg.exec() != SettingsDialog.Accepted:
            return
        new = dlg.result_config()
        old_excel = (self.cfg.excel.path, self.cfg.excel.sheet, self.cfg.excel.header_row,
                     self.cfg.excel.color_source, self.cfg.excel.color_columns)
        old_files = (tuple(self.cfg.files.folders), tuple(self.cfg.files.extensions), self.cfg.files.part_regex)
        old_notes = (tuple(self.cfg.repetitive.folders), tuple(self.cfg.repetitive.note_patterns),
                     self.cfg.repetitive.scan_depth, self.cfg.repetitive.min_phrase_len)
        # mutate the shared config object in place so every panel sees the change
        for name in ("excel", "outlook", "files", "catia", "repetitive", "ui"):
            setattr(self.cfg, name, getattr(new, name))
        self._save_layout()
        self.cfg.save()
        self.apply_theme()
        self._setup_timer()
        self.right.reload_config()
        self.center.invalidate_mail_cache()
        new_files = (tuple(self.cfg.files.folders), tuple(self.cfg.files.extensions), self.cfg.files.part_regex)
        if new_files != old_files:
            if self.cfg.files.folders:
                self.rebuild_index()
            else:
                self.index.entries = {}
        new_notes = (tuple(self.cfg.repetitive.folders), tuple(self.cfg.repetitive.note_patterns),
                     self.cfg.repetitive.scan_depth, self.cfg.repetitive.min_phrase_len)
        if new_notes != old_notes:
            self.rescan_notes()
        new_excel = (self.cfg.excel.path, self.cfg.excel.sheet, self.cfg.excel.header_row,
                     self.cfg.excel.color_source, self.cfg.excel.color_columns)
        # columns / categories / regex may have changed too: always reload
        del new_excel, old_excel
        self.reload_excel()
        if self.center.record:
            self.right.show_record(self.center.record)
        self.set_status("Ayarlar kaydedildi.")
