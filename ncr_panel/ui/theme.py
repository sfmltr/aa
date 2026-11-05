"""Light / dark Qt stylesheets with a single accent colour."""
from __future__ import annotations

PALETTES = {
    "light": dict(
        bg="#F3F5F8", panel="#FFFFFF", panel2="#F8FAFC", border="#D9DEE6", text="#1F2937",
        muted="#6B7280", accent="#2563EB", accent_text="#FFFFFF", hover="#E8EEF9", sel="#DBE7FB",
        header="#EEF2F7", warn="#FACC15", warn_text="#3B2F00", ok="#BBF7D0", ok_text="#14532D",
        bad="#FECACA", bad_text="#7F1D1D", input="#FFFFFF",
    ),
    "dark": dict(
        bg="#111418", panel="#1A1F26", panel2="#20262E", border="#2E3640", text="#E5E7EB",
        muted="#9CA3AF", accent="#3B82F6", accent_text="#FFFFFF", hover="#243041", sel="#2B3F5F",
        header="#232A33", warn="#EAB308", warn_text="#1F1700", ok="#166534", ok_text="#DCFCE7",
        bad="#7F1D1D", bad_text="#FEE2E2", input="#151A20",
    ),
}


def build_stylesheet(theme: str, font_size: int = 10) -> str:
    p = PALETTES.get(theme, PALETTES["light"])
    fs = max(8, font_size)
    return f"""
    * {{ font-size: {fs}pt; }}
    QMainWindow, QDialog, QWidget#root {{ background: {p['bg']}; color: {p['text']}; }}
    QWidget {{ color: {p['text']}; }}
    QFrame#card {{ background: {p['panel']}; border: 1px solid {p['border']}; border-radius: 10px; }}
    QLabel#cardTitle {{ font-weight: 600; font-size: {fs + 1}pt; color: {p['text']}; }}
    QLabel#muted {{ color: {p['muted']}; }}
    QLabel#banner {{ border-radius: 8px; padding: 8px 12px; font-weight: 600; }}
    QLabel#bannerWarn {{ background: {p['warn']}; color: {p['warn_text']}; border-radius: 8px; padding: 8px 12px; font-weight: 700; }}
    QLabel#bannerOk {{ background: {p['ok']}; color: {p['ok_text']}; border-radius: 8px; padding: 8px 12px; font-weight: 600; }}
    QLabel#chip {{ background: {p['sel']}; color: {p['text']}; border-radius: 9px; padding: 2px 8px; }}
    QToolBar {{ background: {p['panel']}; border-bottom: 1px solid {p['border']}; spacing: 6px; padding: 4px; }}
    QToolButton {{ background: transparent; border: 1px solid transparent; border-radius: 6px; padding: 5px 9px; }}
    QToolButton:hover {{ background: {p['hover']}; border-color: {p['border']}; }}
    QToolButton:pressed {{ background: {p['sel']}; }}
    QStatusBar {{ background: {p['panel']}; border-top: 1px solid {p['border']}; color: {p['muted']}; }}
    QPushButton {{ background: {p['panel2']}; border: 1px solid {p['border']}; border-radius: 6px; padding: 6px 12px; }}
    QPushButton:hover {{ background: {p['hover']}; }}
    QPushButton:pressed {{ background: {p['sel']}; }}
    QPushButton:disabled {{ color: {p['muted']}; }}
    QPushButton#primary {{ background: {p['accent']}; color: {p['accent_text']}; border: none; font-weight: 600; }}
    QPushButton#primary:hover {{ background: #1D4ED8; }}
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QPlainTextEdit, QTextEdit, QListWidget, QTableWidget, QTreeWidget {{
        background: {p['input']}; border: 1px solid {p['border']}; border-radius: 6px; padding: 4px 6px;
        selection-background-color: {p['accent']}; selection-color: {p['accent_text']};
    }}
    QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus {{ border-color: {p['accent']}; }}
    QComboBox::drop-down {{ border: none; width: 22px; }}
    QComboBox QAbstractItemView {{ background: {p['panel']}; border: 1px solid {p['border']}; selection-background-color: {p['sel']}; selection-color: {p['text']}; }}
    QTableWidget, QTreeWidget, QListWidget {{ padding: 0; gridline-color: {p['border']}; alternate-background-color: {p['panel2']}; outline: 0; }}
    QTableWidget::item {{ padding: 3px 6px; }}
    QTableWidget::item:selected, QListWidget::item:selected, QTreeWidget::item:selected {{ background: {p['accent']}; color: {p['accent_text']}; }}
    QHeaderView::section {{ background: {p['header']}; color: {p['muted']}; border: none; border-bottom: 1px solid {p['border']}; border-right: 1px solid {p['border']}; padding: 5px 6px; font-weight: 600; }}
    QTableCornerButton::section {{ background: {p['header']}; border: none; }}
    QSplitter::handle {{ background: transparent; }}
    QSplitter::handle:horizontal {{ width: 6px; }}
    QSplitter::handle:vertical {{ height: 6px; }}
    QTabWidget::pane {{ border: 1px solid {p['border']}; border-radius: 8px; background: {p['panel']}; top: -1px; }}
    QTabBar::tab {{ background: transparent; padding: 7px 14px; border: 1px solid transparent; border-bottom: none; border-top-left-radius: 6px; border-top-right-radius: 6px; color: {p['muted']}; }}
    QTabBar::tab:selected {{ background: {p['panel']}; border-color: {p['border']}; color: {p['text']}; font-weight: 600; }}
    QTabBar::tab:hover {{ color: {p['text']}; }}
    QGroupBox {{ border: 1px solid {p['border']}; border-radius: 8px; margin-top: 10px; padding-top: 6px; }}
    QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 4px; color: {p['muted']}; }}
    QScrollBar:vertical {{ background: transparent; width: 10px; margin: 0; }}
    QScrollBar::handle:vertical {{ background: {p['border']}; border-radius: 5px; min-height: 24px; }}
    QScrollBar::handle:vertical:hover {{ background: {p['muted']}; }}
    QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 0; }}
    QScrollBar::handle:horizontal {{ background: {p['border']}; border-radius: 5px; min-width: 24px; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
    QProgressBar {{ border: 1px solid {p['border']}; border-radius: 5px; background: {p['input']}; text-align: center; height: 12px; }}
    QProgressBar::chunk {{ background: {p['accent']}; border-radius: 4px; }}
    QToolTip {{ background: {p['panel']}; color: {p['text']}; border: 1px solid {p['border']}; padding: 4px; }}
    QMenu {{ background: {p['panel']}; border: 1px solid {p['border']}; }}
    QMenu::item {{ padding: 6px 18px; }}
    QMenu::item:selected {{ background: {p['sel']}; }}
    QCheckBox {{ spacing: 6px; }}
    """
