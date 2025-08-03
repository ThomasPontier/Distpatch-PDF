from __future__ import annotations

from typing import Optional, List, Callable
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit, QGroupBox,
    QSplitter, QSizePolicy, QStyle
)
from models.template import StopoverMeta, SendStatus
from models.stopover import Stopover  # existing model
from .pdf_preview import PdfPreview


class StopoverItem(QWidget):
    """
    Per-stopover widget with left email editor and right PDF preview.

    Signals:
      - sendClicked(code: str)
      - overrideChanged(code: str, subject: str, body: str)
      - validationChanged(code: str, valid: bool, reasons: List[str])
    """
    sendClicked = Signal(str)
    overrideChanged = Signal(str, str, str)
    validationChanged = Signal(str, bool, list)

    def __init__(self, stopover: Stopover, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.stopover = stopover
        self.meta = StopoverMeta(stopover_code=stopover.code)
        self._pdf_path: Optional[str] = None

        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(8)

        box = QGroupBox(f"Escale {self.stopover.code}")
        box_layout = QVBoxLayout(box)
        box_layout.setContentsMargins(8, 8, 8, 8)
        box_layout.setSpacing(8)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left editor panel
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        # Metadata row
        meta_row = QHBoxLayout()
        self.statusLabel = QLabel(self.meta.status)
        self.statusLabel.setObjectName("Badge")
        self.lastSentLabel = QLabel(f"Dernier envoi : {self.meta.to_display_dict()['last_sent']}")
        self.lastSentLabel.setObjectName("Badge")
        self.overrideLabel = QLabel("Remplacement")
        self.overrideLabel.setObjectName("Badge")
        self.overrideLabel.setVisible(False)

        meta_row.addWidget(self.statusLabel)
        meta_row.addWidget(self.lastSentLabel)
        meta_row.addWidget(self.overrideLabel)
        meta_row.addStretch(1)
        left_layout.addLayout(meta_row)

        # Subject
        self.subjectEdit = QTextEdit()
        self.subjectEdit.setPlaceholderText("Objet pour cette escale")
        self.subjectEdit.setFixedHeight(40)
        self.subjectEdit.textChanged.connect(self._on_text_changed)
        left_layout.addWidget(self.subjectEdit)

        # Body
        self.bodyEdit = QTextEdit()
        self.bodyEdit.setPlaceholderText("Corps pour cette escale")
        self.bodyEdit.textChanged.connect(self._on_text_changed)
        left_layout.addWidget(self.bodyEdit, 1)

        # Actions row
        actions = QHBoxLayout()
        self.sendBtn = QPushButton("Envoyer individuellement")
        try:
            self.sendBtn.setIcon(self.style().standardIcon(QStyle.SP_ArrowForward))
            self.sendBtn.setIconSize(QSize(18, 18))  # compact per-item action
        except Exception:
            pass
        self.sendBtn.clicked.connect(lambda: self.sendClicked.emit(self.stopover.code))
        actions.addStretch(1)
        actions.addWidget(self.sendBtn)
        left_layout.addLayout(actions)

        # Right preview panel
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        self.preview = PdfPreview()
        right_layout.addWidget(self.preview)

        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setSizes([600, 600])

        box_layout.addWidget(splitter)
        root.addWidget(box)

        # Make the whole item expand horizontally
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)

    def setPdf(self, pdf_path: str):
        """Set PDF path and trigger preview render."""
        self._pdf_path = pdf_path
        # Page number comes from Stopover model attribute page_number
        page = getattr(self.stopover, "page_number", 1) or 1
        self.preview.setDocument(pdf_path, page)

    def setTemplateValues(self, subject: str, body: str, is_override: bool = False):
        """Set initial template text. If override, show badge."""
        self.subjectEdit.blockSignals(True)
        self.bodyEdit.blockSignals(True)
        try:
            self.subjectEdit.setPlainText(subject or "")
            self.bodyEdit.setPlainText(body or "")
        finally:
            self.subjectEdit.blockSignals(False)
            self.bodyEdit.blockSignals(False)
        self.overrideLabel.setVisible(bool(is_override))
        self._validate()

    def _on_text_changed(self):
        self.overrideLabel.setVisible(True)
        self.overrideChanged.emit(
            self.stopover.code,
            self.subjectEdit.toPlainText(),
            self.bodyEdit.toPlainText(),
        )
        self._validate()

    def setSendEnabled(self, enabled: bool, reasons: Optional[List[str]] = None):
        """Control send button availability and status tooltip."""
        self.sendBtn.setEnabled(enabled)
        if not enabled and reasons:
            self.sendBtn.setToolTip("\n".join(reasons))
        else:
            self.sendBtn.setToolTip("")

    def updateStatus(self, status: str, last_sent: Optional[str] = None):
        """Update status and last sent labels."""
        self.statusLabel.setText(status)
        if last_sent:
            self.lastSentLabel.setText(f"Dernier envoi : {last_sent}")

    def _validate(self):
        """Basic validation for subject/body non-empty."""
        reasons: List[str] = []
        if not self.subjectEdit.toPlainText().strip():
            reasons.append("Objet vide")
        if not self.bodyEdit.toPlainText().strip():
            reasons.append("Corps vide")
        valid = len(reasons) == 0
        self.setSendEnabled(valid, reasons if not valid else None)
        self.validationChanged.emit(self.stopover.code, valid, reasons)