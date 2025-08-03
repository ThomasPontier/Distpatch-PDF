from __future__ import annotations

from typing import Optional, List
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit, QLineEdit,
    QPushButton, QDialogButtonBox, QGroupBox, QWidget, QStyle
)
from services.config_manager import get_config_manager


class TemplateEditorDialog(QDialog):
    """
    Global template editor dialog.

    Allows editing subject and body templates with placeholders preview.
    Emits:
      - templateSaved(subject: str, body: str)

    Note: Source of truth is ConfigManager; this dialog edits those values.
    """
    templateSaved = Signal(str, str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Éditeur de modèle global")
        self.setModal(True)
        self.resize(700, 520)
        self._build_ui()
        self._load_current()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # Placeholders preview
        ph_box = QGroupBox("Paramètres pris en charge")
        ph_layout = QVBoxLayout(ph_box)
        placeholders = ["{{stopover_code}}"]
        self.placeholdersLabel = QLabel(", ".join(placeholders))
        self.placeholdersLabel.setObjectName("Badge")
        ph_layout.addWidget(self.placeholdersLabel)
        root.addWidget(ph_box)

        # Subject
        subj_box = QGroupBox("Sujet")
        subj_layout = QVBoxLayout(subj_box)
        self.subjectEdit = QLineEdit()
        subj_layout.addWidget(self.subjectEdit)
        root.addWidget(subj_box)

        # Body
        body_box = QGroupBox("Corps")
        body_layout = QVBoxLayout(body_box)
        self.bodyEdit = QTextEdit()
        self.bodyEdit.setPlaceholderText("Corps de l’email avec paramètres dynamiques")
        body_layout.addWidget(self.bodyEdit)
        root.addWidget(body_box, 1)

        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._on_save)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _load_current(self):
        try:
            t = get_config_manager().get_templates()
            self.subjectEdit.setText(t.get("subject") or "")
            self.bodyEdit.setPlainText(t.get("body") or "")
        except Exception:
            self.subjectEdit.setText("Rapport d’escale – {{stopover_code}}")
            self.bodyEdit.setPlainText("Bonsoir,\n\nVoici le Bilan de satisfaction de l’escale {{stopover_code}}.\n\nCordialement,")

    def _on_save(self):
        subject = self.subjectEdit.text().strip()
        body = self.bodyEdit.toPlainText().strip()
        # Persist directly via ConfigManager (single source of truth)
        try:
            get_config_manager().set_templates(subject, body)
        except Exception:
            pass
        self.templateSaved.emit(subject, body)
        self.accept()