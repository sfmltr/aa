"""Reusable widgets."""
from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QFileDialog, QFrame, QHBoxLayout, QLabel, QListWidget, QPushButton,
                               QSizePolicy, QVBoxLayout, QWidget, QInputDialog)


class Card(QFrame):
    """A rounded panel with a title row and optional right-side widgets."""

    def __init__(self, title: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("card")
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(10, 8, 10, 10)
        self._layout.setSpacing(6)
        self.header = QHBoxLayout()
        self.header.setSpacing(6)
        self.title_label = QLabel(title)
        self.title_label.setObjectName("cardTitle")
        self.header.addWidget(self.title_label)
        self.header.addStretch(1)
        self._layout.addLayout(self.header)

    def add_header_widget(self, w: QWidget):
        self.header.addWidget(w)

    def add(self, w: QWidget, stretch: int = 0):
        self._layout.addWidget(w, stretch)

    def add_layout(self, lay):
        self._layout.addLayout(lay)


class Chip(QLabel):
    def __init__(self, text: str, tooltip: str = ""):
        super().__init__(text)
        self.setObjectName("chip")
        if tooltip:
            self.setToolTip(tooltip)
        self.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)


class FlowRow(QWidget):
    """A simple horizontal row of chips that wraps by re-laying out on resize."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._items: list[QWidget] = []
        self._lay = QHBoxLayout(self)
        self._lay.setContentsMargins(0, 0, 0, 0)
        self._lay.setSpacing(4)
        self._lay.addStretch(1)

    def clear(self):
        for w in self._items:
            self._lay.removeWidget(w)
            w.deleteLater()
        self._items.clear()

    def add(self, w: QWidget):
        self._lay.insertWidget(self._lay.count() - 1, w)
        self._items.append(w)


class PathListEditor(QWidget):
    """List of folders (or files) with add/remove buttons."""
    changed = Signal()

    def __init__(self, mode: str = "folder", parent=None, placeholder: str = ""):
        super().__init__(parent)
        self.mode = mode
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.list = QListWidget()
        self.list.setSelectionMode(QListWidget.ExtendedSelection)
        lay.addWidget(self.list, 1)
        row = QHBoxLayout()
        self.btn_add = QPushButton("＋ Ekle")
        self.btn_remove = QPushButton("－ Kaldır")
        self.btn_add.clicked.connect(self.add_item)
        self.btn_remove.clicked.connect(self.remove_selected)
        row.addWidget(self.btn_add)
        row.addWidget(self.btn_remove)
        row.addStretch(1)
        lay.addLayout(row)
        self._placeholder = placeholder

    def items(self) -> list[str]:
        return [self.list.item(i).text() for i in range(self.list.count())]

    def set_items(self, items: list[str]):
        self.list.clear()
        for it in items:
            if it:
                self.list.addItem(it)

    def add_item(self):
        if self.mode == "folder":
            path = QFileDialog.getExistingDirectory(self, "Klasör seç")
        elif self.mode == "file":
            path, _ = QFileDialog.getOpenFileName(self, "Dosya seç")
        else:
            path, ok = QInputDialog.getText(self, "Ekle", self._placeholder or "Değer:")
            if not ok:
                path = ""
        if path and path not in self.items():
            self.list.addItem(path)
            self.changed.emit()

    def remove_selected(self):
        for it in self.list.selectedItems():
            self.list.takeItem(self.list.row(it))
        self.changed.emit()


def qcolor(hex_color: str | None) -> QColor | None:
    if not hex_color:
        return None
    c = QColor(hex_color)
    return c if c.isValid() else None


def elide(text: str, n: int = 90) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= n else text[: n - 1] + "…"


def set_tooltip_rich(widget: QWidget, text: str):
    widget.setToolTip("<div style='max-width:480px'>" + (text or "").replace("\n", "<br>") + "</div>")


ALIGN_CENTER = Qt.AlignCenter
