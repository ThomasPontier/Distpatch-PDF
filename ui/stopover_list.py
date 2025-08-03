from __future__ import annotations

from typing import List, Optional, Dict, Set
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QWidget, QVBoxLayout, QScrollArea, QSizePolicy

from models.stopover import Stopover
from .stopover_item import StopoverItem


class StopoverList(QWidget):
    """
    Scrollable vertically stacked list of StopoverItem widgets.

    Signals:
      - selectionChanged(codes: list[str])
      - itemSendClicked(code: str)
      - itemOverrideChanged(code: str, subject: str, body: str)
    """
    selectionChanged = Signal(list)
    itemSendClicked = Signal(str)
    itemOverrideChanged = Signal(str, str, str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._items: Dict[str, StopoverItem] = {}
        self._selected: Set[str] = set()
        self._pdf_path: Optional[str] = None
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)

        self.container = QWidget()
        self.container_layout = QVBoxLayout(self.container)
        self.container_layout.setContentsMargins(8, 8, 8, 8)
        self.container_layout.setSpacing(8)
        self.container_layout.addStretch(1)

        self.scroll.setWidget(self.container)
        root.addWidget(self.scroll)

        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    def setStopovers(self, stopovers: List[Stopover]):
        # Clear existing
        for code, item in list(self._items.items()):
            item.setParent(None)
        self._items.clear()
        # Remove stretch and re-add at end
        self._take_container_stretch()

        # Add new items
        for s in stopovers:
            item = StopoverItem(s, parent=self.container)
            item.sendClicked.connect(self.itemSendClicked.emit)
            item.overrideChanged.connect(self.itemOverrideChanged.emit)
            item.validationChanged.connect(self._on_item_validation_changed)
            self.container_layout.addWidget(item)
            self._items[s.code] = item
            if self._pdf_path:
                item.setPdf(self._pdf_path)

        self._add_container_stretch()

    def setPdfPath(self, pdf_path: str):
        self._pdf_path = pdf_path
        for item in self._items.values():
            item.setPdf(pdf_path)

    def setTemplateForSelected(self, subject: str, body: str):
        for code in self._selected:
            item = self._items.get(code)
            if item:
                item.setTemplateValues(subject, body, is_override=False)

    def setItemOverride(self, code: str, subject: str, body: str):
        item = self._items.get(code)
        if item:
            item.setTemplateValues(subject, body, is_override=True)

    def select(self, codes: List[str]):
        self._selected = set(codes)
        self.selectionChanged.emit(list(self._selected))

    def filter(self, text: str, status: str):
        t = (text or "").lower().strip()
        st = (status or "Tous").lower()
        for code, item in self._items.items():
            visible = True
            if t and t not in code.lower():
                visible = False
            if st != "tous":
                # status comes from item.meta.status label text
                visible = visible and (item.statusLabel.text().lower() == st)
            item.setVisible(visible)

    def updateSendStatus(self, code: str, status: str, last_sent: Optional[str] = None):
        item = self._items.get(code)
        if item:
            item.updateStatus(status, last_sent)

    def setSendEnabled(self, code: str, enabled: bool, reasons: Optional[List[str]] = None):
        item = self._items.get(code)
        if item:
            item.setSendEnabled(enabled, reasons)

    def _on_item_validation_changed(self, code: str, valid: bool, reasons: List[str]):
        # Hook for controllers; for now no-op.
        pass

    def _take_container_stretch(self):
        # Remove final stretch if present
        count = self.container_layout.count()
        if count > 0:
            last = self.container_layout.itemAt(count - 1)
            if last and last.spacerItem():
                self.container_layout.takeAt(count - 1)

    def _add_container_stretch(self):
        self.container_layout.addStretch(1)