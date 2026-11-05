"""Settings dialog: every tunable from config.py, grouped in tabs."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDialog, QDialogButtonBox, QDoubleSpinBox, QFileDialog,
                               QFormLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox,
                               QPlainTextEdit, QPushButton, QSpinBox, QTableWidget, QTableWidgetItem,
                               QTabWidget, QVBoxLayout, QWidget)

from ..config import AppConfig
from ..excel_data import ExcelData, list_sheets
from .widgets import PathListEditor


def _csv(items: list[str]) -> str:
    return ", ".join(items)


def _parse_csv(text: str) -> list[str]:
    return [t.strip() for t in text.split(",") if t.strip()]


class SettingsDialog(QDialog):
    def __init__(self, cfg: AppConfig, data: ExcelData | None = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ayarlar")
        self.resize(820, 640)
        self.cfg = cfg.copy()
        self.data = data
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs, 1)
        self._build_excel()
        self._build_outlook()
        self._build_files()
        self._build_catia()
        self._build_folders()
        self._build_repetitive()
        self._build_colors()
        self._build_general()
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Save).setText("Kaydet")
        bb.button(QDialogButtonBox.Cancel).setText("Vazgeç")
        bb.accepted.connect(self._accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    # ---- helpers ------------------------------------------------------------
    def _file_row(self, line: QLineEdit, filt: str, folder: bool = False) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(line, 1)
        btn = QPushButton("Gözat…")

        def pick():
            if folder:
                p = QFileDialog.getExistingDirectory(self, "Klasör seç", line.text())
            else:
                p, _ = QFileDialog.getOpenFileName(self, "Dosya seç", line.text(), filt)
            if p:
                line.setText(p)
        btn.clicked.connect(pick)
        h.addWidget(btn)
        return w

    def _tab(self, title: str) -> QFormLayout:
        w = QWidget()
        form = QFormLayout(w)
        form.setLabelAlignment(Qt.AlignRight)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        self.tabs.addTab(w, title)
        return form

    # ---- Excel ----------------------------------------------------------------
    def _build_excel(self):
        ex = self.cfg.excel
        form = self._tab("Excel")
        self.ex_path = QLineEdit(ex.path)
        form.addRow("Excel dosyası:", self._file_row(self.ex_path, "Excel (*.xlsm *.xlsx *.xls);;Tümü (*)"))
        self.ex_sheet = QComboBox()
        self.ex_sheet.setEditable(True)
        self.ex_sheet.addItem("")
        self.ex_sheet.setCurrentText(ex.sheet)
        btn_sheets = QPushButton("Sayfaları yükle")
        btn_sheets.clicked.connect(self._load_sheets)
        row = QWidget()
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(self.ex_sheet, 1)
        h.addWidget(btn_sheets)
        form.addRow("Sayfa (boş = aktif):", row)
        self.ex_header = QSpinBox()
        self.ex_header.setRange(0, 500)
        self.ex_header.setValue(ex.header_row)
        self.ex_header.setSpecialValueText("Otomatik")
        form.addRow("Başlık satırı:", self.ex_header)
        self.cols = {}
        for key, label in [("col_note", "Note kolonu"), ("col_engineer", "MRB ENG. (kullanıcı) kolonu"),
                           ("col_ncr", "NCR NO kolonu"), ("col_discrepancy", "Discrepancy kolonu"),
                           ("col_status", "Durum (MRB-ENG / EXTERNAL) kolonu"), ("col_part", "Part No kolonu"),
                           ("col_cncsi", "CN/CSI kolonu")]:
            e = QLineEdit(getattr(ex, key))
            e.setMaximumWidth(80)
            self.cols[key] = e
            form.addRow(label + ":", e)
        self.ex_mrb = QLineEdit(ex.status_mrb)
        self.ex_ext = QLineEdit(ex.status_external)
        form.addRow("MRB-ENG durum metni:", self.ex_mrb)
        form.addRow("EXTERNAL APPROVER durum metni:", self.ex_ext)
        self.ex_macro = QLineEdit(ex.macro_name)
        form.addRow("Makro adı:", self.ex_macro)
        self.ex_reload = QCheckBox("Makro sonrası Excel'i yeniden oku")
        self.ex_reload.setChecked(ex.reload_after_macro)
        form.addRow("", self.ex_reload)
        self.ex_color_src = QComboBox()
        self.ex_color_src.addItems(["auto", "openpyxl", "com"])
        self.ex_color_src.setCurrentText(ex.color_source)
        self.ex_color_src.setToolTip("auto: koşullu biçimlendirme varsa Excel COM (Windows), yoksa openpyxl\n"
                                     "openpyxl: hızlı, yalnızca sabit dolgu renkleri\n"
                                     "com: Excel üzerinden görüntülenen renkler (koşullu biçimlendirme dahil)")
        form.addRow("Renk kaynağı:", self.ex_color_src)
        self.ex_color_cols = QLineEdit(ex.color_columns)
        self.ex_color_cols.setToolTip("Satır rengi için bakılacak hücreler, virgülle (ilk renkli olan kullanılır)")
        form.addRow("Renk için kolonlar:", self.ex_color_cols)

    def _load_sheets(self):
        p = self.ex_path.text().strip()
        if not Path(p).is_file():
            QMessageBox.warning(self, "Excel", "Önce geçerli bir Excel dosyası seçin.")
            return
        try:
            names = list_sheets(p)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Excel", f"Sayfalar okunamadı:\n{exc}")
            return
        cur = self.ex_sheet.currentText()
        self.ex_sheet.clear()
        self.ex_sheet.addItem("")
        self.ex_sheet.addItems(names)
        self.ex_sheet.setCurrentText(cur if cur in names else (names[0] if names else ""))

    # ---- Outlook -------------------------------------------------------------
    def _build_outlook(self):
        o = self.cfg.outlook
        form = self._tab("Outlook")
        self.ol_enabled = QCheckBox("Outlook'ta NCR NO ara")
        self.ol_enabled.setChecked(o.enabled)
        form.addRow("", self.ol_enabled)
        self.ol_all = QCheckBox("Tüm posta klasörlerinde ara (yavaş olabilir)")
        self.ol_all.setChecked(o.search_all_folders)
        form.addRow("", self.ol_all)
        self.ol_folders = PathListEditor(mode="text", placeholder="Klasör adı (örn. Inbox veya Inbox/Projeler)")
        self.ol_folders.set_items(o.folders)
        self.ol_folders.setMinimumHeight(120)
        form.addRow("Aranacak klasörler:", self.ol_folders)
        self.ol_days = QSpinBox()
        self.ol_days.setRange(0, 5000)
        self.ol_days.setValue(o.days_back)
        self.ol_days.setSpecialValueText("Sınırsız")
        form.addRow("Kaç gün geriye:", self.ol_days)
        self.ol_body = QCheckBox("Mail gövdesinde de ara (yalnızca konu için kapatın)")
        self.ol_body.setChecked(o.search_body)
        form.addRow("", self.ol_body)
        self.ol_max = QSpinBox()
        self.ol_max.setRange(10, 2000)
        self.ol_max.setValue(o.max_results)
        form.addRow("Klasör başına en fazla sonuç:", self.ol_max)
        self.ol_approve = QLineEdit(_csv(o.approval_keywords))
        form.addRow("Onay kelimeleri (yeşil):", self.ol_approve)
        self.ol_reject = QLineEdit(_csv(o.rejection_keywords))
        form.addRow("Red kelimeleri (kırmızı):", self.ol_reject)

    # ---- Files ---------------------------------------------------------------
    def _build_files(self):
        f = self.cfg.files
        form = self._tab("Dosya arama")
        self.fs_folders = PathListEditor(mode="folder")
        self.fs_folders.set_items(f.folders)
        self.fs_folders.setMinimumHeight(140)
        form.addRow("Mail / PDF klasörleri:", self.fs_folders)
        self.fs_ext = QLineEdit(_csv(f.extensions))
        form.addRow("Uzantılar:", self.fs_ext)
        self.fs_regex = QLineEdit(f.part_regex)
        self.fs_regex.setToolTip("Discrepancy içindeki parça numarası deseni (regex). Varsayılan: 3 rakam, 1 harf, 4 rakam, isteğe bağlı -rakam ekleri")
        form.addRow("Parça no deseni (regex):", self.fs_regex)
        self.fs_base = QCheckBox("Kök numarayı da eşleştir (123A4567-1 → 123A4567)")
        self.fs_base.setChecked(f.match_base_number)
        form.addRow("", self.fs_base)
        self.fs_ncr = QCheckBox("NCR NO geçen dosyaları da listele (sarı ünlem için sayılmaz)")
        self.fs_ncr.setChecked(f.include_ncr_no)
        form.addRow("", self.fs_ncr)
        self.fs_auto = QCheckBox("Açılışta indeksi arka planda güncelle")
        self.fs_auto.setChecked(f.auto_index_on_start)
        form.addRow("", self.fs_auto)
        self.fs_maxmb = QSpinBox()
        self.fs_maxmb.setRange(1, 2000)
        self.fs_maxmb.setValue(f.max_file_mb)
        form.addRow("En büyük dosya (MB):", self.fs_maxmb)
        self.fs_pages = QSpinBox()
        self.fs_pages.setRange(1, 5000)
        self.fs_pages.setValue(f.max_pdf_pages)
        form.addRow("PDF başına en fazla sayfa:", self.fs_pages)

    # ---- CATIA ---------------------------------------------------------------
    def _build_catia(self):
        c = self.cfg.catia
        form = self._tab("CATIA")
        self.ca_exe = QLineEdit(c.exe_path)
        form.addRow("CNEXT.exe yolu:", self._file_row(self.ca_exe, "CATIA (CNEXT.exe);;Tümü (*)"))
        self.ca_folder = QLineEdit(c.models_folder)
        form.addRow("Model klasörü:", self._file_row(self.ca_folder, "", folder=True))
        self.ca_ext = QLineEdit(_csv(c.extensions))
        form.addRow("Model uzantıları:", self.ca_ext)
        self.ca_rec = QCheckBox("Alt klasörleri de tara")
        self.ca_rec.setChecked(c.recursive)
        form.addRow("", self.ca_rec)
        self.ca_com = QCheckBox("Modeli COM ile aç (CATIA.Application) — önerilir")
        self.ca_com.setChecked(c.use_com)
        form.addRow("", self.ca_com)
        self.ca_wait = QSpinBox()
        self.ca_wait.setRange(5, 300)
        self.ca_wait.setValue(c.startup_wait_s)
        form.addRow("CATIA açılış beklemesi (sn):", self.ca_wait)

    # ---- quick folders ---------------------------------------------------------
    def _build_folders(self):
        form = self._tab("Hızlı klasörler")
        self.qf = PathListEditor(mode="folder")
        self.qf.set_items(self.cfg.ui.quick_folders)
        form.addRow("Sağ orta bölümde listelenecek klasörler:", self.qf)

    # ---- repetitive ------------------------------------------------------------
    def _build_repetitive(self):
        r = self.cfg.repetitive
        form = self._tab("Tekrarlayan NCR")
        self.rp_folders = PathListEditor(mode="folder")
        self.rp_folders.set_items(r.folders)
        self.rp_folders.setMinimumHeight(140)
        form.addRow("Kök klasörler (her alt klasörde bir not dosyası):", self.rp_folders)
        self.rp_patterns = QLineEdit(_csv(r.note_patterns))
        form.addRow("Not dosyası deseni:", self.rp_patterns)
        self.rp_depth = QSpinBox()
        self.rp_depth.setRange(1, 10)
        self.rp_depth.setValue(r.scan_depth)
        form.addRow("Tarama derinliği:", self.rp_depth)
        self.rp_thr = QDoubleSpinBox()
        self.rp_thr.setRange(0.3, 1.0)
        self.rp_thr.setSingleStep(0.05)
        self.rp_thr.setValue(r.similarity_threshold)
        self.rp_thr.setToolTip("1.00 = yalnızca birebir ifade eşleşmesi; daha düşük = bulanık benzerlik")
        form.addRow("Benzerlik eşiği:", self.rp_thr)
        self.rp_min = QSpinBox()
        self.rp_min.setRange(3, 200)
        self.rp_min.setValue(r.min_phrase_len)
        form.addRow("En kısa ifade uzunluğu:", self.rp_min)

    # ---- colours ---------------------------------------------------------------
    def _build_colors(self):
        form = self._tab("Renkler")
        info = QLabel("Excel'de kullanılan satır renkleri ve anlamları. Etiketleri siz yazın (ör. proje adı, "
                      "'onaydan döndü'); etiketler listelerde araç ipucu ve açıklama olarak gösterilir.")
        info.setWordWrap(True)
        form.addRow(info)
        self.color_table = QTableWidget(0, 4)
        self.color_table.setHorizontalHeaderLabels(["Renk", "Hex", "Kayıt", "Etiket / Anlam"])
        self.color_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.color_table.verticalHeader().setVisible(False)
        form.addRow(self.color_table)
        btn = QPushButton("Yüklü Excel'den renkleri tara")
        btn.clicked.connect(self._scan_colors)
        form.addRow(btn)
        self._scan_colors()

    def _scan_colors(self):
        counts = dict(self.data.color_counts) if self.data else {}
        for hexv in self.cfg.ui.color_legend:
            counts.setdefault(hexv, 0)
        self.color_table.setRowCount(0)
        for hexv, n in sorted(counts.items(), key=lambda kv: -kv[1]):
            row = self.color_table.rowCount()
            self.color_table.insertRow(row)
            sw = QTableWidgetItem("")
            sw.setBackground(QColor(hexv))
            sw.setFlags(Qt.ItemIsEnabled)
            self.color_table.setItem(row, 0, sw)
            h = QTableWidgetItem(hexv)
            h.setFlags(Qt.ItemIsEnabled)
            self.color_table.setItem(row, 1, h)
            c = QTableWidgetItem(str(n))
            c.setFlags(Qt.ItemIsEnabled)
            self.color_table.setItem(row, 2, c)
            self.color_table.setItem(row, 3, QTableWidgetItem(self.cfg.ui.color_legend.get(hexv, "")))

    # ---- general -------------------------------------------------------------
    def _build_general(self):
        u = self.cfg.ui
        form = self._tab("Genel")
        self.ui_theme = QComboBox()
        self.ui_theme.addItems(["light", "dark"])
        self.ui_theme.setCurrentText(u.theme)
        form.addRow("Tema:", self.ui_theme)
        self.ui_font = QSpinBox()
        self.ui_font.setRange(8, 16)
        self.ui_font.setValue(u.font_size)
        form.addRow("Yazı boyutu (pt):", self.ui_font)
        self.ui_refresh = QSpinBox()
        self.ui_refresh.setRange(0, 720)
        self.ui_refresh.setValue(u.auto_refresh_min)
        self.ui_refresh.setSpecialValueText("Kapalı")
        form.addRow("Excel'i otomatik yenile (dk):", self.ui_refresh)
        self.ui_other = QCheckBox("Durumu MRB-ENG / EXTERNAL olmayan satırları da üst listede göster")
        self.ui_other.setChecked(u.show_other_status)
        form.addRow("", self.ui_other)
        self.ui_cats = QPlainTextEdit()
        self.ui_cats.setPlaceholderText("Kategori: anahtar1, anahtar2\nStress: stress, stres")
        self.ui_cats.setPlainText("\n".join(f"{k}: {', '.join(v)}" for k, v in u.categories.items()))
        self.ui_cats.setMaximumHeight(120)
        form.addRow("Note kategorileri:", self.ui_cats)

    # ---- accept --------------------------------------------------------------
    def _accept(self):
        c = self.cfg
        ex = c.excel
        ex.path = self.ex_path.text().strip()
        ex.sheet = self.ex_sheet.currentText().strip()
        ex.header_row = self.ex_header.value()
        for key, e in self.cols.items():
            v = e.text().strip().upper()
            if not v.isalpha():
                QMessageBox.warning(self, "Ayarlar", f"Geçersiz kolon harfi: {v!r}")
                return
            setattr(ex, key, v)
        ex.status_mrb = self.ex_mrb.text().strip() or "MRB-ENG"
        ex.status_external = self.ex_ext.text().strip() or "EXTERNAL APPROVER"
        ex.macro_name = self.ex_macro.text().strip() or "B03_Update_NCRs"
        ex.reload_after_macro = self.ex_reload.isChecked()
        ex.color_source = self.ex_color_src.currentText()
        ex.color_columns = self.ex_color_cols.text().strip() or ex.col_ncr

        o = c.outlook
        o.enabled = self.ol_enabled.isChecked()
        o.search_all_folders = self.ol_all.isChecked()
        o.folders = self.ol_folders.items()
        o.days_back = self.ol_days.value()
        o.search_body = self.ol_body.isChecked()
        o.max_results = self.ol_max.value()
        o.approval_keywords = _parse_csv(self.ol_approve.text())
        o.rejection_keywords = _parse_csv(self.ol_reject.text())

        f = c.files
        f.folders = self.fs_folders.items()
        f.extensions = [e if e.startswith(".") else "." + e for e in _parse_csv(self.fs_ext.text())]
        import re
        try:
            re.compile(self.fs_regex.text())
        except re.error as exc:
            QMessageBox.warning(self, "Ayarlar", f"Parça no deseni geçersiz: {exc}")
            return
        f.part_regex = self.fs_regex.text().strip() or f.part_regex
        f.match_base_number = self.fs_base.isChecked()
        f.include_ncr_no = self.fs_ncr.isChecked()
        f.auto_index_on_start = self.fs_auto.isChecked()
        f.max_file_mb = self.fs_maxmb.value()
        f.max_pdf_pages = self.fs_pages.value()

        ca = c.catia
        ca.exe_path = self.ca_exe.text().strip()
        ca.models_folder = self.ca_folder.text().strip()
        ca.extensions = [e if e.startswith(".") else "." + e for e in _parse_csv(self.ca_ext.text())]
        ca.recursive = self.ca_rec.isChecked()
        ca.use_com = self.ca_com.isChecked()
        ca.startup_wait_s = self.ca_wait.value()

        c.ui.quick_folders = self.qf.items()

        r = c.repetitive
        r.folders = self.rp_folders.items()
        r.note_patterns = _parse_csv(self.rp_patterns.text()) or ["*.txt"]
        r.scan_depth = self.rp_depth.value()
        r.similarity_threshold = round(self.rp_thr.value(), 2)
        r.min_phrase_len = self.rp_min.value()

        legend = {}
        for row in range(self.color_table.rowCount()):
            hexv = self.color_table.item(row, 1).text()
            label = self.color_table.item(row, 3).text().strip() if self.color_table.item(row, 3) else ""
            if label:
                legend[hexv] = label
        c.ui.color_legend = legend

        u = c.ui
        u.theme = self.ui_theme.currentText()
        u.font_size = self.ui_font.value()
        u.auto_refresh_min = self.ui_refresh.value()
        u.show_other_status = self.ui_other.isChecked()
        cats: dict[str, list[str]] = {}
        for line in self.ui_cats.toPlainText().splitlines():
            if ":" in line:
                name, kws = line.split(":", 1)
                if name.strip():
                    cats[name.strip()] = _parse_csv(kws) or [name.strip()]
        u.categories = cats or u.categories
        self.accept()

    def result_config(self) -> AppConfig:
        return self.cfg
