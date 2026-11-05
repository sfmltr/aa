"""Centre column: details of the selected NCR, Outlook mails, related files."""
from __future__ import annotations

import datetime as _dt
from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QAction, QBrush, QColor, QGuiApplication
from PySide6.QtWidgets import (QAbstractItemView, QGridLayout, QHeaderView, QLabel, QMenu, QPlainTextEdit,
                               QPushButton, QSplitter, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from ..config import AppConfig
from ..excel_data import NcrRecord
from ..file_search import FileHit, FileIndex
from ..outlook_mail import MailHit, open_mail, outlook_available, search_mails
from ..utils import human_size, open_path, reveal_in_folder
from .widgets import Card, Chip, FlowRow, elide, set_tooltip_rich
from .workers import run_in_background

ROLE = Qt.UserRole + 1


class CenterPanel(QWidget):
    status = Signal(str)
    error = Signal(str)

    def __init__(self, cfg: AppConfig, index_getter: Callable[[], FileIndex], parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self._index = index_getter
        self.record: NcrRecord | None = None
        self._mail_cache: dict[str, list[MailHit]] = {}
        self._search_token = 0

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.splitter = QSplitter(Qt.Vertical)
        lay.addWidget(self.splitter)

        # ---- details ------------------------------------------------------
        self.card_info = Card("Rapor Detayı")
        self.btn_copy = QPushButton("NCR NO kopyala")
        self.btn_copy.clicked.connect(self._copy_ncr)
        self.card_info.add_header_widget(self.btn_copy)
        grid = QGridLayout()
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(4)
        self.lbl = {}
        fields = [("NCR NO", "ncr_no"), ("Part No", "part_no"), ("CN/CSI", "cn_csi"), ("Durum", "status"),
                  ("MRB ENG.", "engineer"), ("Kategori", "categories")]
        for i, (title, key) in enumerate(fields):
            t = QLabel(title + ":")
            t.setObjectName("muted")
            v = QLabel("—")
            v.setTextInteractionFlags(Qt.TextSelectableByMouse)
            v.setWordWrap(True)
            grid.addWidget(t, i // 3, (i % 3) * 2)
            grid.addWidget(v, i // 3, (i % 3) * 2 + 1)
            self.lbl[key] = v
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)
        grid.setColumnStretch(5, 1)
        self.card_info.add_layout(grid)
        disc_title = QLabel("Discrepancy:")
        disc_title.setObjectName("muted")
        self.card_info.add(disc_title)
        self.disc = QPlainTextEdit()
        self.disc.setReadOnly(True)
        self.disc.setMaximumHeight(110)
        self.card_info.add(self.disc)
        pn_title = QLabel("Discrepancy'de geçen parça numaraları:")
        pn_title.setObjectName("muted")
        self.card_info.add(pn_title)
        self.chips = FlowRow()
        self.card_info.add(self.chips)
        note_title = QLabel("Note:")
        note_title.setObjectName("muted")
        self.note_title = note_title
        self.card_info.add(note_title)
        self.note = QLabel("")
        self.note.setWordWrap(True)
        self.note.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.card_info.add(self.note)
        self.splitter.addWidget(self.card_info)

        # ---- mails --------------------------------------------------------
        self.card_mail = Card("E-postalar (Outlook)")
        self.mail_status = QLabel("")
        self.mail_status.setObjectName("muted")
        self.card_mail.add_header_widget(self.mail_status)
        self.btn_mail_refresh = QPushButton("⟳ Ara")
        self.btn_mail_refresh.setToolTip("Outlook'ta bu NCR NO'yu yeniden ara")
        self.btn_mail_refresh.clicked.connect(lambda: self._search_mails(force=True))
        self.card_mail.add_header_widget(self.btn_mail_refresh)
        self.mails = QTableWidget(0, 5)
        self.mails.setHorizontalHeaderLabels(["Tarih", "Gönderen", "Konu", "Klasör", "Durum"])
        self._setup_table(self.mails)
        self.mails.itemDoubleClicked.connect(self._open_mail)
        self.mails.setContextMenuPolicy(Qt.CustomContextMenu)
        self.mails.customContextMenuRequested.connect(self._mail_menu)
        self.card_mail.add(self.mails, 1)
        self.splitter.addWidget(self.card_mail)

        # ---- files --------------------------------------------------------
        self.card_files = Card("İlgili Dosyalar (mail / PDF)")
        self.files_status = QLabel("")
        self.files_status.setObjectName("muted")
        self.card_files.add_header_widget(self.files_status)
        self.files = QTableWidget(0, 5)
        self.files.setHorizontalHeaderLabels(["Dosya", "Tür", "Eşleşme", "Klasör", "Tarih"])
        self._setup_table(self.files)
        self.files.itemDoubleClicked.connect(self._open_file)
        self.files.setContextMenuPolicy(Qt.CustomContextMenu)
        self.files.customContextMenuRequested.connect(self._file_menu)
        self.card_files.add(self.files, 1)
        self.splitter.addWidget(self.card_files)
        self.splitter.setSizes([320, 330, 330])
        self.clear()

    @staticmethod
    def _setup_table(t: QTableWidget):
        t.verticalHeader().setVisible(False)
        t.setSelectionBehavior(QAbstractItemView.SelectRows)
        t.setSelectionMode(QAbstractItemView.SingleSelection)
        t.setEditTriggers(QAbstractItemView.NoEditTriggers)
        t.setSortingEnabled(True)
        t.setWordWrap(False)
        t.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        t.horizontalHeader().setStretchLastSection(True)
        t.horizontalHeader().setHighlightSections(False)

    # ------------------------------------------------------------------
    def clear(self):
        self.record = None
        for v in self.lbl.values():
            v.setText("—")
        self.disc.setPlainText("")
        self.chips.clear()
        self.note.setText("")
        self.mails.setRowCount(0)
        self.files.setRowCount(0)
        self.mail_status.setText("Soldan bir rapor seçin")
        self.files_status.setText("")

    def show_record(self, rec: NcrRecord):
        self.record = rec
        self.lbl["ncr_no"].setText(f"<b>{rec.ncr_no}</b>")
        self.lbl["part_no"].setText(rec.part_no or "—")
        self.lbl["cn_csi"].setText(rec.cn_csi or "—")
        self.lbl["status"].setText(rec.status or "—")
        self.lbl["engineer"].setText(rec.engineer or "—")
        self.lbl["categories"].setText(", ".join(rec.categories) or "—")
        self.disc.setPlainText(rec.discrepancy)
        self.chips.clear()
        if rec.part_numbers:
            for p in rec.part_numbers:
                self.chips.add(Chip(p, "Bu numara dosya aramasında kullanılır"))
        else:
            self.chips.add(Chip("parça numarası bulunamadı", "Ayarlar → Dosya arama → Parça no deseni"))
        self.note.setText(rec.note or "—")
        tip = "\n".join(f"{k}: {v}" for k, v in rec.extra.items())
        set_tooltip_rich(self.card_info.title_label, tip)
        self._show_files()
        self._search_mails(force=False)

    # ---- mails ---------------------------------------------------------
    def _search_mails(self, force: bool):
        rec = self.record
        if rec is None:
            return
        self.mails.setRowCount(0)
        if not self.cfg.outlook.enabled:
            self.mail_status.setText("Outlook araması ayarlardan kapalı")
            return
        if not outlook_available():
            self.mail_status.setText("Outlook erişimi yok (Windows + pywin32 gerekir)")
            return
        if not force and rec.ncr_no in self._mail_cache:
            self._fill_mails(self._mail_cache[rec.ncr_no])
            return
        self._search_token += 1
        token = self._search_token
        self.mail_status.setText("Outlook aranıyor…")
        self.btn_mail_refresh.setEnabled(False)

        def done(hits):
            self.btn_mail_refresh.setEnabled(True)
            if token != self._search_token:
                return
            self._mail_cache[rec.ncr_no] = hits
            self._fill_mails(hits)

        def err(msg):
            self.btn_mail_refresh.setEnabled(True)
            self.mail_status.setText("Outlook hatası")
            self.error.emit("Outlook araması başarısız:\n" + msg)

        run_in_background(search_mails, rec.ncr_no, self.cfg.outlook, on_done=done, on_error=err,
                          on_progress=self.status.emit)

    def _fill_mails(self, hits: list[MailHit]):
        self.mails.setSortingEnabled(False)
        self.mails.setRowCount(len(hits))
        approved = 0
        for row, h in enumerate(hits):
            date = h.received.strftime("%Y-%m-%d %H:%M") if h.received else ""
            state = "✔ Onay" if h.approved else ("✖ Red" if h.rejected else "")
            vals = [date, h.sender, h.subject, h.folder, state]
            for col, v in enumerate(vals):
                it = QTableWidgetItem(v)
                it.setData(ROLE, row)
                if h.approved:
                    it.setBackground(QBrush(QColor("#BBF7D0")))
                    it.setForeground(QBrush(QColor("#14532D")))
                elif h.rejected:
                    it.setBackground(QBrush(QColor("#FECACA")))
                    it.setForeground(QBrush(QColor("#7F1D1D")))
                self.mails.setItem(row, col, it)
            tip = h.snippet
            if h.matched_keywords:
                tip = "Anahtar kelimeler: " + ", ".join(h.matched_keywords) + "\n\n" + tip
            if h.has_attachments:
                tip = "📎 Ek var\n" + tip
            for col in range(5):
                set_tooltip_rich(self.mails.item(row, col), tip)
            approved += int(h.approved)
        self._hits = hits
        self.mails.setSortingEnabled(True)
        self.mails.resizeColumnsToContents()
        if self.mails.columnWidth(2) > 420:
            self.mails.setColumnWidth(2, 420)
        self.mail_status.setText(f"{len(hits)} e-posta" + (f", {approved} onaylı" if approved else ""))

    def _mail_at(self, item) -> MailHit | None:
        if item is None:
            return None
        idx = item.data(ROLE)
        try:
            return self._hits[idx]
        except (AttributeError, IndexError, TypeError):
            return None

    def _open_mail(self, item):
        h = self._mail_at(item)
        if not h:
            return
        try:
            open_mail(h.entry_id, h.store_id)
        except Exception as exc:  # noqa: BLE001
            self.error.emit(f"E-posta açılamadı: {exc}")

    def _mail_menu(self, pos):
        h = self._mail_at(self.mails.itemAt(pos))
        if not h:
            return
        m = QMenu(self)
        a = QAction("Outlook'ta aç", self)
        a.triggered.connect(lambda: self._open_mail(self.mails.itemAt(pos)))
        b = QAction("Konuyu kopyala", self)
        b.triggered.connect(lambda: QGuiApplication.clipboard().setText(h.subject))
        m.addAction(a)
        m.addAction(b)
        m.exec(self.mails.viewport().mapToGlobal(pos))

    # ---- files ---------------------------------------------------------
    def _show_files(self):
        rec = self.record
        self.files.setRowCount(0)
        if rec is None:
            return
        index = self._index()
        if not self.cfg.files.folders:
            self.files_status.setText("Ayarlar → Dosya arama'dan klasör ekleyin")
            return
        if not index.entries:
            self.files_status.setText("Dosya indeksi boş / hazırlanıyor…")
            return
        hits = index.match_record(rec, self.cfg.files)
        self._file_hits = hits
        self.files.setSortingEnabled(False)
        self.files.setRowCount(len(hits))
        for row, h in enumerate(hits):
            p = Path(h.entry.path)
            date = _dt.datetime.fromtimestamp(h.entry.mtime).strftime("%Y-%m-%d") if h.entry.mtime else ""
            vals = [p.name, p.suffix.lower().lstrip("."), "; ".join(h.reasons), str(p.parent), date]
            for col, v in enumerate(vals):
                it = QTableWidgetItem(v)
                it.setData(ROLE, row)
                if h.by_part:
                    it.setBackground(QBrush(QColor("#FEF3C7")))
                    it.setForeground(QBrush(QColor("#4A3B00")))
                set_tooltip_rich(it, f"{h.entry.path}\n{human_size(h.entry.size)}" + (f"\nHata: {h.entry.error}" if h.entry.error else ""))
                self.files.setItem(row, col, it)
        self.files.setSortingEnabled(True)
        self.files.resizeColumnsToContents()
        if self.files.columnWidth(0) > 320:
            self.files.setColumnWidth(0, 320)
        n_part = sum(1 for h in hits if h.by_part)
        self.files_status.setText(f"{len(hits)} dosya ({n_part} parça no eşleşmesi)")

    def refresh_files(self):
        if self.record is not None:
            self._show_files()

    def _file_at(self, item) -> FileHit | None:
        if item is None:
            return None
        try:
            return self._file_hits[item.data(ROLE)]
        except (AttributeError, IndexError, TypeError):
            return None

    def _open_file(self, item):
        h = self._file_at(item)
        if h:
            try:
                open_path(h.entry.path)
            except Exception as exc:  # noqa: BLE001
                self.error.emit(f"Dosya açılamadı: {exc}")

    def _file_menu(self, pos):
        h = self._file_at(self.files.itemAt(pos))
        if not h:
            return
        m = QMenu(self)
        a = QAction("Aç", self)
        a.triggered.connect(lambda: open_path(h.entry.path))
        b = QAction("Klasörde göster", self)
        b.triggered.connect(lambda: reveal_in_folder(h.entry.path))
        c = QAction("Yolu kopyala", self)
        c.triggered.connect(lambda: QGuiApplication.clipboard().setText(h.entry.path))
        m.addActions([a, b, c])
        m.exec(self.files.viewport().mapToGlobal(pos))

    def _copy_ncr(self):
        if self.record:
            QGuiApplication.clipboard().setText(self.record.ncr_no)
            self.status.emit(f"Kopyalandı: {self.record.ncr_no}")

    def invalidate_mail_cache(self):
        self._mail_cache.clear()
