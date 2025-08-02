from typing import Optional, List
from PySide6 import QtWidgets, QtCore
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLineEdit,
    QComboBox,
    QCheckBox,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QMessageBox,
)
from services.stopover_email_service import StopoverEmailService, StopoverEmailConfig
from services.mapping_service import MappingService


class StopoverEmailSettingsDialog(QtWidgets.QDialog):
    """
    Native PySide6 dialog to configure stopover email settings.

    Fields:
      - To, Cc, Bcc (comma/semicolon separated)
      - Subject (QLineEdit)
      - Template selection (QComboBox) and read-only preview (QPlainTextEdit)
      - Include PDF attachment (QCheckBox) - mapped locally; service does not currently persist this
    """

    def __init__(self, parent: QtWidgets.QWidget, stopover_code: str, email_service: StopoverEmailService):
        super().__init__(parent)
        self.setWindowTitle(f"Email Settings - {str(stopover_code).upper()}")
        self.resize(700, 520)

        self._parent = parent
        self._stopover_code = str(stopover_code).upper()
        self._svc: StopoverEmailService = email_service
        self._mapping = MappingService()

        # Load settings/config from service
        self._config: StopoverEmailConfig = self._svc.get_config(self._stopover_code)

        # Local-only option (since service has no attachment flags); default True for parity
        self._include_pdf = True

        # Build UI
        self._build_ui()
        self._load_into_ui()

    # ----- UI construction -----
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)

        form = QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight)

        # Recipients
        self.to_edit = QLineEdit(self)
        self.cc_edit = QLineEdit(self)
        self.bcc_edit = QLineEdit(self)

        form.addRow(QLabel("To:"), self.to_edit)
        form.addRow(QLabel("Cc:"), self.cc_edit)
        form.addRow(QLabel("Bcc:"), self.bcc_edit)

        # Subject
        self.subject_edit = QLineEdit(self)
        form.addRow(QLabel("Subject:"), self.subject_edit)

        # Template select and preview
        tmpl_row = QHBoxLayout()
        self.template_combo = QComboBox(self)
        self.template_combo.setMinimumWidth(220)
        tmpl_row.addWidget(QLabel("Template:"))
        tmpl_row.addWidget(self.template_combo, 1)

        self.preview = QPlainTextEdit(self)
        self.preview.setReadOnly(True)
        self.preview.setPlaceholderText("Template preview")

        root.addLayout(form)
        root.addLayout(tmpl_row)
        root.addWidget(self.preview, 1)

        # Attachment option
        self.attach_pdf_chk = QCheckBox("Include PDF attachment", self)
        self.attach_pdf_chk.setChecked(True)
        root.addWidget(self.attach_pdf_chk)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        self.ok_btn = QPushButton("OK", self)
        self.cancel_btn = QPushButton("Cancel", self)
        self.ok_btn.clicked.connect(self._on_ok)
        self.cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(self.ok_btn)
        btn_row.addWidget(self.cancel_btn)
        root.addLayout(btn_row)

        # Signals
        self.template_combo.currentIndexChanged.connect(self._update_preview)
        self.subject_edit.textChanged.connect(self._update_preview)
        self.to_edit.textChanged.connect(self._update_preview)

    # ----- Data load/populate -----
    def _load_into_ui(self) -> None:
        # Subject
        self.subject_edit.setText(self._config.subject_template or "")

        # Recipients: To from config; fall back to mapping if empty
        to_list: List[str] = list(self._config.recipients or [])
        if not to_list:
            mapped = self._mapping.get_emails_for_stopover(self._stopover_code)
            if mapped:
                to_list = mapped
        self.to_edit.setText(self._join_emails(to_list))
        self.cc_edit.setText(self._join_emails(self._config.cc_recipients or []))
        self.bcc_edit.setText(self._join_emails(self._config.bcc_recipients or []))

        # Templates: service exposes subject/body via get_config; there is no list API.
        # Provide a minimal selection to satisfy UI requirements: "Default" only.
        self.template_combo.clear()
        self.template_combo.addItem("Default")

        # Attachment checkbox
        self.attach_pdf_chk.setChecked(self._include_pdf)

        # Initial preview
        self._update_preview()

    # ----- Helpers -----
    @staticmethod
    def _join_emails(emails: List[str]) -> str:
        return ", ".join(e for e in (emails or []) if e)

    @staticmethod
    def _split_emails(text: str) -> List[str]:
        # Split by comma/semicolon and filter empties/spaces
        parts = [p.strip() for p in re_split(r"[;,]", text or "")]
        return [p for p in parts if p]

    def _render_subject(self, subject_template: str) -> str:
        # Simple token replacement for {{stopover_code}}
        return (subject_template or "").replace("{{stopover_code}}", self._stopover_code)

    def _render_body(self) -> str:
        # Use body_template from config and render stopover_code token
        bt = self._config.body_template or ""
        return bt.replace("{{stopover_code}}", self._stopover_code)

    def _update_preview(self) -> None:
        # Preview subject and body
        subj = self._render_subject(self.subject_edit.text())
        body = self._render_body()
        lines = []
        lines.append(f"Subject: {subj}")
        lines.append("")
        lines.append(body)
        self.preview.setPlainText("\n".join(lines))

    # ----- Validation/Persistence -----
    def _on_ok(self) -> None:
        # Basic validation: To and Subject non-empty
        to_list = self._split_emails(self.to_edit.text())
        subject_text = (self.subject_edit.text() or "").strip()

        if not to_list:
            QMessageBox.critical(self, "Validation Error", "The 'To' field must contain at least one email.")
            return
        if not subject_text:
            QMessageBox.critical(self, "Validation Error", "Subject must not be empty.")
            return

        # Persist into config
        self._config.stopover_code = self._stopover_code
        self._config.subject_template = subject_text
        # Keep existing body_template; dialog is a settings surface, not a template editor
        self._config.recipients = to_list
        self._config.cc_recipients = self._split_emails(self.cc_edit.text())
        self._config.bcc_recipients = self._split_emails(self.bcc_edit.text())
        # is_enabled stays based on service rules (global enablement); keep True for parity
        self._config.is_enabled = True

        try:
            self._svc.save_config(self._config)
            # Attachment choice is currently local only; no service hook available.
            self._include_pdf = self.attach_pdf_chk.isChecked()
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save settings: {str(e)}")
            return

        self.accept()


# Local, minimal regex split util (avoid importing 're' at module level users might not expect)
import re
def re_split(pattern: str, text: str) -> List[str]:
    return re.split(pattern, text or "")