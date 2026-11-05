"""Left column: MRB-ENG list (top) and EXTERNAL APPROVER list (bottom)."""
from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QAction, QBrush, QColor, QGuiApplication
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
                               QMenu, QPushButton, QSplitter, QTableWidget, QTableWidgetItem, QVBoxLayout,
                               QWidget)

from ..config import AppConfig
from ..excel_data import ExcelData, NcrRecord
from ..utils import contrast_text_color
from .widgets import Card, Chip, FlowRow, elide, set_tooltip_rich

ALL = "— Tümü —"
KEY_ROLE = Qt.UserRole + 1


class NcrTable(QTableWidget):
    """Table of NCR records coloured like the Excel sheet, with a badge column."""
    recordActivated = Signal(object)
    recordSelected = Signal(object)
    openInExcelRequested = Signal(object)

    def __init__(self, columns: list[str], parent=None):
        super().__init__(0, len(columns), parent)
        self.columns = columns
        self.setHorizontalHeaderLabels(columns)
        self.verticalHeader().setVisible(False)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setAlternatingRowColors(False)
        self.setSortingEnabled(True)
        self.setWordWrap(False)
        hh = self.horizontalHeader()
        hh.setStretchLastSection(False)
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setHighlightSections(False)
        self.setColumnWidth(0, 72)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._menu)
        self.itemSelectionChanged.connect(self._on_sel)
        self.itemDoubleClicked.connect(lambda _it: self._emit(self.recordActivated))
        self._records: dict[str, NcrRecord] = {}
        self._dark = False

    # population -----------------------------------------------------------
    def populate(self, records: list[NcrRecord], dark: bool, extra_cols: bool, legend: dict[str, str] | None = None):
        self._dark = dark
        legend = legend or {}
        self.setSortingEnabled(False)
        self.setRowCount(0)
        self._records = {r.key: r for r in records}
        self.setRowCount(len(records))
        for row, rec in enumerate(records):
            values = [self._badge_text(rec), rec.ncr_no, rec.part_no, elide(rec.discrepancy, 120), rec.cn_csi]
            if extra_cols:
                values += [", ".join(rec.categories), elide(rec.note, 100)]
            bg = QColor(rec.color) if rec.color else None
            fg = QColor(contrast_text_color(rec.color)) if rec.color else None
            for col, val in enumerate(values):
                item = QTableWidgetItem(val)
                item.setData(KEY_ROLE, rec.key)
                if col == 0:
                    item.setTextAlignment(Qt.AlignCenter)
                    self._style_badge(item, rec)
                else:
                    if bg is not None:
                        item.setBackground(QBrush(bg))
                        item.setForeground(QBrush(fg))
                if col == 1 and rec.color:
                    label = legend.get(rec.color)
                    item.setToolTip(f"Renk: {label} ({rec.color})" if label else f"Renk: {rec.color}")
                if col == 3:
                    set_tooltip_rich(item, rec.discrepancy)
                elif col == len(values) - 1 and extra_cols:
                    set_tooltip_rich(item, rec.note)
                self.setItem(row, col, item)
        self.setSortingEnabled(True)
        self.resizeRowsToContents()

    def refresh_badges(self, keys: set[str] | None = None):
        sorting = self.isSortingEnabled()
        self.setSortingEnabled(False)
        for row in range(self.rowCount()):
            it = self.item(row, 0)
            if it is None:
                continue
            key = it.data(KEY_ROLE)
            if keys is not None and key not in keys:
                continue
            rec = self._records.get(key)
            if rec:
                it.setText(self._badge_text(rec))
                self._style_badge(it, rec)
        self.setSortingEnabled(sorting)
        self.fit_badge_column()

    def fit_badge_column(self):
        self.resizeColumnToContents(0)
        if self.columnWidth(0) < 72:
            self.setColumnWidth(0, 72)

    def records(self) -> list[NcrRecord]:
        return list(self._records.values())

    @staticmethod
    def _badge_text(rec: NcrRecord) -> str:
        """Same badges for both lists: ⚠ related files, 🔁 repetitive, ✅ approval mail."""
        parts = []
        if rec.file_match_count:
            parts.append("⚠")
        if rec.repetitive_count:
            parts.append("🔁")
        if rec.mail_approved:
            parts.append("✅")
        return " ".join(parts)

    def _style_badge(self, item: QTableWidgetItem, rec: NcrRecord):
        tips = []
        if rec.file_match_count:
            tips.append(f"⚠ Discrepancy'deki parça no ile eşleşen {rec.file_match_count} dosya/mail var")
            item.setBackground(QBrush(QColor("#FACC15")))
            item.setForeground(QBrush(QColor("#3B2F00")))
        else:
            item.setBackground(QBrush())
            item.setForeground(QBrush())
        if rec.repetitive_count:
            tips.append(f"🔁 Tekrarlayan NCR: {rec.repetitive_count} eşleşme")
        if rec.mail_approved:
            tips.append("✅ Onay içeren e-posta var")
        if rec.mail_count and not rec.mail_approved:
            tips.append(f"✉ {rec.mail_count} ilgili e-posta")
        item.setToolTip("\n".join(tips))

    # selection --------------------------------------------------------------
    def current_record(self) -> NcrRecord | None:
        items = self.selectedItems()
        if not items:
            return None
        return self._records.get(items[0].data(KEY_ROLE))

    def _on_sel(self):
        rec = self.current_record()
        if rec:
            self.recordSelected.emit(rec)

    def _emit(self, sig):
        rec = self.current_record()
        if rec:
            sig.emit(rec)

    def select_key(self, key: str) -> bool:
        for row in range(self.rowCount()):
            it = self.item(row, 1)
            if it and it.data(KEY_ROLE) == key:
                self.selectRow(row)
                self.scrollToItem(it)
                return True
        return False

    def _menu(self, pos):
        rec = self.current_record()
        if not rec:
            return
        menu = QMenu(self)
        a_copy = QAction(f"NCR NO kopyala: {rec.ncr_no}", self)
        a_copy.triggered.connect(lambda: QGuiApplication.clipboard().setText(rec.ncr_no))
        a_copy_disc = QAction("Discrepancy metnini kopyala", self)
        a_copy_disc.triggered.connect(lambda: QGuiApplication.clipboard().setText(rec.discrepancy))
        a_excel = QAction(f"Excel'de satıra git (satır {rec.row})", self)
        a_excel.triggered.connect(lambda: self.openInExcelRequested.emit(rec))
        menu.addAction(a_copy)
        menu.addAction(a_copy_disc)
        menu.addSeparator()
        menu.addAction(a_excel)
        menu.exec(self.viewport().mapToGlobal(pos))


class LeftPanel(QWidget):
    recordSelected = Signal(object)
    runMacroRequested = Signal()
    refreshRequested = Signal()
    engineerChanged = Signal(str)
    openInExcelRequested = Signal(object)

    def __init__(self, cfg: AppConfig, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.data: ExcelData | None = None
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.splitter = QSplitter(Qt.Vertical)
        lay.addWidget(self.splitter)

        # ---- top: MRB-ENG --------------------------------------------------
        self.card_top = Card("MRB-ENG Raporları")
        self.engineer = QComboBox()
        self.engineer.setMinimumWidth(160)
        self.engineer.setToolTip("MRB ENG. (C kolonu) kullanıcısına göre filtrele")
        self.engineer.currentTextChanged.connect(self._on_engineer)
        self.card_top.add_header_widget(QLabel("Kullanıcı:"))
        self.card_top.add_header_widget(self.engineer)
        self.btn_refresh = QPushButton("⟳ Yenile")
        self.btn_refresh.setToolTip("Excel'i yeniden oku")
        self.btn_refresh.clicked.connect(self.refreshRequested)
        self.card_top.add_header_widget(self.btn_refresh)
        row = QHBoxLayout()
        self.search_top = QLineEdit()
        self.search_top.setPlaceholderText("Ara: NCR NO, parça no, discrepancy…")
        self.search_top.setClearButtonEnabled(True)
        self.search_top.textChanged.connect(self._apply_filters)
        self.count_top = QLabel("")
        self.count_top.setObjectName("muted")
        row.addWidget(self.search_top, 1)
        row.addWidget(self.count_top)
        self.card_top.add_layout(row)
        self.legend = FlowRow()
        self.legend.setVisible(False)
        self.card_top.add(self.legend)
        self.table_top = NcrTable(["", "NCR NO", "Part No", "Discrepancy", "CN/CSI"])
        self.card_top.add(self.table_top, 1)
        self.splitter.addWidget(self.card_top)

        # ---- bottom: EXTERNAL APPROVER --------------------------------------
        self.card_bottom = Card("EXTERNAL APPROVER")
        self.category = QComboBox()
        self.category.setToolTip("Note (A kolonu) içeriğine göre kategori")
        self.category.currentTextChanged.connect(self._apply_filters)
        self.card_bottom.add_header_widget(QLabel("Kategori:"))
        self.card_bottom.add_header_widget(self.category)
        self.btn_macro = QPushButton("▶ Makro ile güncelle")
        self.btn_macro.setObjectName("primary")
        self.btn_macro.setToolTip(f"Excel içindeki '{cfg.excel.macro_name}' makrosunu çalıştır ve listeyi yenile")
        self.btn_macro.clicked.connect(self.runMacroRequested)
        self.card_bottom.add_header_widget(self.btn_macro)
        row2 = QHBoxLayout()
        self.search_bottom = QLineEdit()
        self.search_bottom.setPlaceholderText("Ara…")
        self.search_bottom.setClearButtonEnabled(True)
        self.search_bottom.textChanged.connect(self._apply_filters)
        self.count_bottom = QLabel("")
        self.count_bottom.setObjectName("muted")
        row2.addWidget(self.search_bottom, 1)
        row2.addWidget(self.count_bottom)
        self.card_bottom.add_layout(row2)
        self.table_bottom = NcrTable(["", "NCR NO", "Part No", "Discrepancy", "CN/CSI", "Kategori", "Note"])
        self.card_bottom.add(self.table_bottom, 1)
        self.splitter.addWidget(self.card_bottom)
        self.splitter.setSizes([500, 400])

        for t in (self.table_top, self.table_bottom):
            t.recordSelected.connect(self._on_selected)
            t.openInExcelRequested.connect(self.openInExcelRequested)
        self._updating = False

    # ---------------------------------------------------------------------
    def set_data(self, data: ExcelData):
        self.data = data
        self._updating = True
        current = self.cfg.ui.selected_engineer or self.engineer.currentText()
        self.engineer.clear()
        self.engineer.addItem(ALL)
        self.engineer.addItems(data.engineers)
        if current and current in data.engineers:
            self.engineer.setCurrentText(current)
        cats = sorted({c for r in data.records for c in r.categories})
        cur_cat = self.category.currentText()
        self.category.clear()
        self.category.addItem(ALL)
        self.category.addItems(cats)
        self.category.addItem("(Kategorisiz)")
        if cur_cat:
            self.category.setCurrentText(cur_cat)
        self._updating = False
        self._fill_legend(data)
        self.btn_macro.setToolTip(f"Excel içindeki '{self.cfg.excel.macro_name}' makrosunu çalıştır ve listeyi yenile")
        self._apply_filters()

    def _fill_legend(self, data: ExcelData):
        self.legend.clear()
        legend = self.cfg.ui.color_legend
        shown = 0
        for hexv, n in sorted(data.color_counts.items(), key=lambda kv: -kv[1]):
            label = legend.get(hexv)
            chip = Chip(f"  {label}  " if label else f"  {hexv}  ", f"{hexv} — {n} kayıt" + (f": {label}" if label else "\n(Ayarlar → Renkler'den etiket verin)"))
            chip.setStyleSheet(f"background:{hexv}; color:{contrast_text_color(hexv)}; border-radius:9px; padding:2px 6px;")
            self.legend.add(chip)
            shown += 1
            if shown >= 12:
                break
        self.legend.setVisible(shown > 0)

    def current_engineer(self) -> str:
        t = self.engineer.currentText()
        return "" if t == ALL else t

    def _on_engineer(self, text: str):
        if self._updating:
            return
        self.cfg.ui.selected_engineer = "" if text == ALL else text
        self.engineerChanged.emit(self.cfg.ui.selected_engineer)
        self._apply_filters()

    def _apply_filters(self, *_):
        if self.data is None or self._updating:
            return
        eng = self.current_engineer()
        dark = self.cfg.ui.theme == "dark"
        ex = self.cfg.excel

        def eng_ok(r: NcrRecord) -> bool:
            return not eng or r.engineer.casefold() == eng.casefold()

        def text_ok(r: NcrRecord, q: str) -> bool:
            if not q:
                return True
            hay = " ".join([r.ncr_no, r.part_no, r.discrepancy, r.cn_csi, r.note, " ".join(r.part_numbers)]).casefold()
            return all(tok in hay for tok in q.casefold().split())

        q_top = self.search_top.text().strip()
        top = [r for r in self.data.records
               if (r.kind(ex) == "mrb" or (self.cfg.ui.show_other_status and r.kind(ex) == "other"))
               and eng_ok(r) and text_ok(r, q_top)]
        self.table_top.populate(top, dark, extra_cols=False, legend=self.cfg.ui.color_legend)
        self.count_top.setText(self._summary(top))

        cat = self.category.currentText()
        q_bot = self.search_bottom.text().strip()
        bottom = []
        for r in self.data.records:
            if r.kind(ex) != "external" or not eng_ok(r) or not text_ok(r, q_bot):
                continue
            if cat == "(Kategorisiz)" and r.categories:
                continue
            if cat not in (ALL, "(Kategorisiz)", "") and cat not in r.categories:
                continue
            bottom.append(r)
        self.table_bottom.populate(bottom, dark, extra_cols=True, legend=self.cfg.ui.color_legend)
        self.count_bottom.setText(self._summary(bottom))
        self._autosize()

    def _autosize(self):
        for t in (self.table_top, self.table_bottom):
            t.resizeColumnsToContents()
            t.fit_badge_column()
            if t.columnWidth(3) > 320:
                t.setColumnWidth(3, 320)
            if t.columnCount() > 6 and t.columnWidth(6) > 240:
                t.setColumnWidth(6, 240)
            t.horizontalHeader().setStretchLastSection(True)

    @staticmethod
    def _summary(recs: list[NcrRecord]) -> str:
        parts = [f"{len(recs)} kayıt"]
        w = sum(1 for r in recs if r.file_match_count)
        rp = sum(1 for r in recs if r.repetitive_count)
        ok = sum(1 for r in recs if r.mail_approved)
        if w:
            parts.append(f"⚠ {w}")
        if rp:
            parts.append(f"🔁 {rp}")
        if ok:
            parts.append(f"✅ {ok}")
        return "  ·  ".join(parts)

    def refresh_badges(self, keys: set[str] | None = None):
        self.table_top.refresh_badges(keys)
        self.table_bottom.refresh_badges(keys)
        self.count_top.setText(self._summary(self.table_top.records()))
        self.count_bottom.setText(self._summary(self.table_bottom.records()))

    def _on_selected(self, rec: NcrRecord):
        sender = self.sender()
        other = self.table_bottom if sender is self.table_top else self.table_top
        other.blockSignals(True)
        other.clearSelection()
        other.blockSignals(False)
        self.recordSelected.emit(rec)

    def reselect(self, key: str):
        if not self.table_top.select_key(key):
            self.table_bottom.select_key(key)
