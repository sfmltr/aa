"""Right column: CATIA models, quick folders, repetitive NCR documents."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QAction, QGuiApplication
from PySide6.QtWidgets import (QDialog, QDialogButtonBox, QFileDialog, QHBoxLayout, QLabel, QLineEdit,
                               QListWidget, QListWidgetItem, QMenu, QPlainTextEdit, QPushButton, QSplitter,
                               QVBoxLayout, QWidget)

from ..catia import list_models, open_model, start_catia
from ..config import AppConfig
from ..excel_data import NcrRecord
from ..repetitive import RepEntry, RepHit, match_record
from ..utils import open_path, reveal_in_folder
from .widgets import Card
from .workers import run_in_background

ROLE = Qt.UserRole + 1


class RightPanel(QWidget):
    status = Signal(str)
    error = Signal(str)
    configChanged = Signal()
    rescanNotesRequested = Signal()

    def __init__(self, cfg: AppConfig, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self._models: list[Path] = []
        self._entries: list[RepEntry] = []
        self._hits: list[RepHit] = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.splitter = QSplitter(Qt.Vertical)
        lay.addWidget(self.splitter)

        # ---- CATIA models ---------------------------------------------------
        self.card_models = Card("CATIA V5 Modelleri")
        self.btn_catia = QPushButton("CATIA'yı başlat")
        self.btn_catia.clicked.connect(self._start_catia)
        self.card_models.add_header_widget(self.btn_catia)
        self.btn_models_folder = QPushButton("📁")
        self.btn_models_folder.setToolTip("Model klasörünü aç")
        self.btn_models_folder.clicked.connect(lambda: self._open(self.cfg.catia.models_folder))
        self.card_models.add_header_widget(self.btn_models_folder)
        self.model_search = QLineEdit()
        self.model_search.setPlaceholderText("Model ara…")
        self.model_search.setClearButtonEnabled(True)
        self.model_search.textChanged.connect(self._fill_models)
        self.card_models.add(self.model_search)
        self.models = QListWidget()
        self.models.itemDoubleClicked.connect(self._open_model)
        self.models.setContextMenuPolicy(Qt.CustomContextMenu)
        self.models.customContextMenuRequested.connect(self._model_menu)
        self.card_models.add(self.models, 1)
        self.models_status = QLabel("")
        self.models_status.setObjectName("muted")
        self.card_models.add(self.models_status)
        self.splitter.addWidget(self.card_models)

        # ---- quick folders ----------------------------------------------------
        self.card_folders = Card("Hızlı Klasörler")
        self.btn_add_folder = QPushButton("＋")
        self.btn_add_folder.setToolTip("Klasör ekle")
        self.btn_add_folder.clicked.connect(self._add_folder)
        self.btn_del_folder = QPushButton("－")
        self.btn_del_folder.setToolTip("Seçili klasörü listeden kaldır")
        self.btn_del_folder.clicked.connect(self._remove_folder)
        self.card_folders.add_header_widget(self.btn_add_folder)
        self.card_folders.add_header_widget(self.btn_del_folder)
        self.folders = QListWidget()
        self.folders.itemDoubleClicked.connect(lambda it: self._open(it.data(ROLE)))
        self.folders.setContextMenuPolicy(Qt.CustomContextMenu)
        self.folders.customContextMenuRequested.connect(self._folder_menu)
        self.card_folders.add(self.folders, 1)
        self.splitter.addWidget(self.card_folders)

        # ---- repetitive -------------------------------------------------------
        self.card_rep = Card("Tekrarlayan NCR Dokümanları")
        self.btn_rescan = QPushButton("⟳ Tara")
        self.btn_rescan.setToolTip("Not dosyalarını yeniden oku")
        self.btn_rescan.clicked.connect(self.rescanNotesRequested)
        self.card_rep.add_header_widget(self.btn_rescan)
        self.rep_banner = QLabel("")
        self.rep_banner.setObjectName("banner")
        self.rep_banner.setWordWrap(True)
        self.card_rep.add(self.rep_banner)
        self.rep_list = QListWidget()
        self.rep_list.itemDoubleClicked.connect(self._open_rep)
        self.rep_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.rep_list.customContextMenuRequested.connect(self._rep_menu)
        self.card_rep.add(self.rep_list, 1)
        self.rep_status = QLabel("")
        self.rep_status.setObjectName("muted")
        self.card_rep.add(self.rep_status)
        self.splitter.addWidget(self.card_rep)
        self.splitter.setSizes([330, 250, 400])
        self.reload_config()

    # ------------------------------------------------------------------
    def reload_config(self):
        self.refresh_models()
        self._fill_folders()
        self._show_rep_hits([])

    # ---- models ---------------------------------------------------------
    def refresh_models(self):
        try:
            self._models = list_models(self.cfg.catia)
        except Exception as exc:  # noqa: BLE001
            self._models = []
            self.error.emit(f"Model listesi okunamadı: {exc}")
        self._fill_models()

    def _fill_models(self, *_):
        q = self.model_search.text().strip().casefold()
        self.models.clear()
        root = Path(self.cfg.catia.models_folder) if self.cfg.catia.models_folder else None
        shown = 0
        for p in self._models:
            if q and q not in p.name.casefold() and q not in str(p).casefold():
                continue
            try:
                rel = p.relative_to(root) if root else p
            except ValueError:
                rel = p
            it = QListWidgetItem(f"{p.name}" + (f"   ({rel.parent})" if str(rel.parent) not in (".", "") else ""))
            it.setData(ROLE, str(p))
            it.setToolTip(str(p))
            self.models.addItem(it)
            shown += 1
        if not self.cfg.catia.models_folder:
            self.models_status.setText("Ayarlar → CATIA'dan model klasörü seçin")
        else:
            self.models_status.setText(f"{shown} model" + (f" / {len(self._models)}" if shown != len(self._models) else ""))

    def _open_model(self, item: QListWidgetItem):
        path = item.data(ROLE)
        self.status.emit(f"CATIA'da açılıyor: {Path(path).name}")
        self.models.setEnabled(False)

        def done(msg):
            self.models.setEnabled(True)
            self.status.emit(msg)

        def err(msg):
            self.models.setEnabled(True)
            self.error.emit("Model açılamadı:\n" + msg)

        run_in_background(open_model, path, self.cfg.catia, on_done=done, on_error=err, on_progress=self.status.emit)

    def _start_catia(self):
        self.btn_catia.setEnabled(False)

        def done(started):
            self.btn_catia.setEnabled(True)
            self.status.emit("CATIA başlatıldı." if started else "CATIA zaten çalışıyor.")

        def err(msg):
            self.btn_catia.setEnabled(True)
            self.error.emit("CATIA başlatılamadı:\n" + msg)

        run_in_background(start_catia, self.cfg.catia, on_done=done, on_error=err, on_progress=self.status.emit)

    def _model_menu(self, pos):
        it = self.models.itemAt(pos)
        if not it:
            return
        path = it.data(ROLE)
        m = QMenu(self)
        a = QAction("CATIA'da aç", self)
        a.triggered.connect(lambda: self._open_model(it))
        b = QAction("Klasörde göster", self)
        b.triggered.connect(lambda: reveal_in_folder(path))
        c = QAction("Yolu kopyala", self)
        c.triggered.connect(lambda: QGuiApplication.clipboard().setText(path))
        m.addActions([a, b, c])
        m.exec(self.models.viewport().mapToGlobal(pos))

    # ---- quick folders ----------------------------------------------------
    def _fill_folders(self):
        self.folders.clear()
        for f in self.cfg.ui.quick_folders:
            p = Path(f)
            it = QListWidgetItem(("📁 " if p.is_dir() else "⚠ ") + (p.name or f))
            it.setData(ROLE, f)
            it.setToolTip(f if p.is_dir() else f"{f}\n(klasör bulunamadı)")
            self.folders.addItem(it)

    def _add_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Hızlı klasör ekle")
        if path and path not in self.cfg.ui.quick_folders:
            self.cfg.ui.quick_folders.append(path)
            self._fill_folders()
            self.configChanged.emit()

    def _remove_folder(self):
        it = self.folders.currentItem()
        if it:
            try:
                self.cfg.ui.quick_folders.remove(it.data(ROLE))
            except ValueError:
                pass
            self._fill_folders()
            self.configChanged.emit()

    def _folder_menu(self, pos):
        it = self.folders.itemAt(pos)
        if not it:
            return
        m = QMenu(self)
        a = QAction("Aç", self)
        a.triggered.connect(lambda: self._open(it.data(ROLE)))
        b = QAction("Yolu kopyala", self)
        b.triggered.connect(lambda: QGuiApplication.clipboard().setText(it.data(ROLE)))
        c = QAction("Listeden kaldır", self)
        c.triggered.connect(self._remove_folder)
        m.addActions([a, b, c])
        m.exec(self.folders.viewport().mapToGlobal(pos))

    def _open(self, path: str):
        if not path:
            return
        try:
            open_path(path)
        except Exception as exc:  # noqa: BLE001
            self.error.emit(f"Açılamadı: {path}\n{exc}")

    # ---- repetitive -------------------------------------------------------
    def set_entries(self, entries: list[RepEntry]):
        self._entries = entries
        n = len(entries)
        if not self.cfg.repetitive.folders:
            self.rep_status.setText("Ayarlar → Tekrarlayan NCR'dan klasör ekleyin")
        else:
            self.rep_status.setText(f"{n} not dosyası okundu")

    def show_record(self, rec: NcrRecord | None):
        if rec is None:
            self._show_rep_hits([])
            return
        hits = match_record(rec, self._entries, self.cfg.repetitive)
        self._show_rep_hits(hits)

    def _show_rep_hits(self, hits: list[RepHit]):
        self._hits = hits
        self.rep_list.clear()
        if hits:
            self.rep_banner.setObjectName("bannerWarn")
            self.rep_banner.setText(f"🔁 REPETITIVE!  {len(hits)} benzer doküman bulundu")
        else:
            self.rep_banner.setObjectName("bannerOk" if self._entries else "banner")
            self.rep_banner.setText("Tekrarlayan kayıt yok" if self._entries else "")
        self.rep_banner.style().unpolish(self.rep_banner)
        self.rep_banner.style().polish(self.rep_banner)
        for h in hits:
            score = "birebir" if h.exact else f"%{h.score * 100:.0f} benzer"
            it = QListWidgetItem(f"📁 {h.entry.name}   —   {score}")
            it.setData(ROLE, h.entry.folder)
            it.setToolTip(f"{h.entry.folder}\nNot: {h.entry.note_path}\nEşleşen ifade: {h.phrase}")
            self.rep_list.addItem(it)

    def _open_rep(self, item: QListWidgetItem):
        self._open(item.data(ROLE))

    def _rep_menu(self, pos):
        it = self.rep_list.itemAt(pos)
        if not it:
            return
        hit = self._hits[self.rep_list.row(it)]
        m = QMenu(self)
        a = QAction("Klasörü aç", self)
        a.triggered.connect(lambda: self._open(hit.entry.folder))
        b = QAction("Notu göster", self)
        b.triggered.connect(lambda: self._show_note(hit))
        c = QAction("Not dosyasını aç", self)
        c.triggered.connect(lambda: self._open(hit.entry.note_path))
        m.addActions([a, b, c])
        m.exec(self.rep_list.viewport().mapToGlobal(pos))

    def _show_note(self, hit: RepHit):
        dlg = QDialog(self)
        dlg.setWindowTitle(hit.entry.name)
        dlg.resize(600, 400)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(f"<b>{hit.entry.note_path}</b><br>Eşleşen ifade: <i>{hit.phrase}</i>"))
        txt = QPlainTextEdit(hit.entry.text)
        txt.setReadOnly(True)
        lay.addWidget(txt, 1)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(dlg.reject)
        bb.accepted.connect(dlg.accept)
        lay.addWidget(bb)
        dlg.exec()
