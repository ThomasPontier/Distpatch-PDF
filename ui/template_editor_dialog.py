from __future__ import annotations

from typing import Optional, List
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit, QLineEdit,
    QPushButton, QDialogButtonBox, QGroupBox, QWidget, QStyle
)
from models.template import Template


class TemplateEditorDialog(QDialog):
    """
    Global template editor dialog.

    Allows editing subject and body templates with placeholders preview.
    Emits:
      - templateSaved(subject: str, body: str)
    """
    templateSaved = Signal(str, str)

    def __init__(self, parent: Optional[QWidget] = None, current: Optional[Template] = None):
        super().__init__(parent)
        self.setWindowTitle("Global Template Editor")
        self.setModal(True)
        self.resize(700, 520)
        self.current = current or Template()
        self._build_ui()
        self._load_current()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # Placeholders preview
        ph_box = QGroupBox("Supported placeholders")
        ph_layout = QVBoxLayout(ph_box)
        self.placeholdersLabel = QLabel(", ".join(self.current.placeholders))
        self.placeholdersLabel.setObjectName("Badge")
        ph_layout.addWidget(self.placeholdersLabel)
        root.addWidget(ph_box)

        # Subject
        subj_box = QGroupBox("Subject template")
        subj_layout = QVBoxLayout(subj_box)
        self.subjectEdit = QLineEdit()
        subj_layout.addWidget(self.subjectEdit)
        root.addWidget(subj_box)

        # Body
        body_box = QGroupBox("Body template")
        body_layout = QVBoxLayout(body_box)
        self.bodyEdit = QTextEdit()
        self.bodyEdit.setPlaceholderText("Email body template with placeholders")
        body_layout.addWidget(self.bodyEdit)
        root.addWidget(body_box, 1)

        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._on_save)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _load_current(self):
        self.subjectEdit.setText(self.current.subject or "")
        self.bodyEdit.setPlainText(self.current.body or "")

    def _on_save(self):
        subject = self.subjectEdit.text().strip()
        body = self.bodyEdit.toPlainText().strip()
        self.templateSaved.emit(subject, body)
        self.accept()