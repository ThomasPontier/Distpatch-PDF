from typing import List, Tuple
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QLabel, QPushButton
)


class RecipientEditorDialog(QDialog):
    """
    Factorized dialog to edit recipients for a stopover: To / CC / CCI.
    Reused by MappingTab for both 'Ajouter' and 'Modifier'.

    API:
      - set_initial(to_list, cc_list, bcc_list)
      - get_values() -> (to_list, cc_list, bcc_list)
      - static helpers:
          open(parent, code, to_list, cc_list, bcc_list) -> (accepted: bool, to, cc, bcc)
    """

    def __init__(self, parent=None, code: str = ""):
        super().__init__(parent)
        self._code = (code or "").upper()
        self.setWindowTitle(f"Éditer les emails — {self._code}" if self._code else "Éditer les emails")
        try:
            self.resize(640, 260)
        except Exception:
            pass

        v = QVBoxLayout(self)

        # Helper note
        hint = QLabel("Utilisez une virgule pour séparer plusieurs adresses.")
        hint.setStyleSheet("color: #555; font-size: 12px;")
        hint.setWordWrap(True)
        v.addWidget(hint)

        form = QFormLayout()
        self.to_edit = QLineEdit(self)
        self.to_edit.setPlaceholderText("ex: a@ex.com, b@ex.com")
        self.cc_edit = QLineEdit(self)
        self.cc_edit.setPlaceholderText("ex: c@ex.com, d@ex.com")
        self.bcc_edit = QLineEdit(self)
        self.bcc_edit.setPlaceholderText("ex: e@ex.com")

        form.addRow(QLabel("À (To):"), self.to_edit)
        form.addRow(QLabel("CC:"), self.cc_edit)
        form.addRow(QLabel("CCI:"), self.bcc_edit)
        v.addLayout(form)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        ok_btn = QPushButton("OK", self)
        cancel_btn = QPushButton("Annuler", self)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        v.addLayout(btn_row)

        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

    def set_initial(self, to_list: List[str], cc_list: List[str], bcc_list: List[str]) -> None:
        self.to_edit.setText(", ".join(to_list or []))
        self.cc_edit.setText(", ".join(cc_list or []))
        self.bcc_edit.setText(", ".join(bcc_list or []))

    @staticmethod
    def _split_emails(text: str) -> List[str]:
        raw = text or ""
        return [e.strip() for e in raw.split(",") if e.strip()]

    def get_values(self) -> Tuple[List[str], List[str], List[str]]:
        to_vals = self._split_emails(self.to_edit.text())
        cc_vals = self._split_emails(self.cc_edit.text())
        bcc_vals = self._split_emails(self.bcc_edit.text())
        return to_vals, cc_vals, bcc_vals

    @staticmethod
    def open(parent, code: str, to_list: List[str], cc_list: List[str], bcc_list: List[str]) -> Tuple[bool, List[str], List[str], List[str]]:
        dlg = RecipientEditorDialog(parent, code)
        dlg.set_initial(to_list, cc_list, bcc_list)
        ok = dlg.exec() == QDialog.Accepted
        if not ok:
            return False, to_list, cc_list, bcc_list
        return True, *dlg.get_values()