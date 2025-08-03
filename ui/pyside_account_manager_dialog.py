from __future__ import annotations
from typing import Optional, List
from PySide6.QtCore import Qt, QSize
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QGroupBox, QMessageBox, QWidget, QFormLayout, QDialogButtonBox, QComboBox, QStyle, QPushButton
)
from services.email_service import EmailService

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
    _WIN32_AVAILABLE = True
except Exception:
    win32com = None  # type: ignore
    pythoncom = None  # type: ignore
    _WIN32_AVAILABLE = False


class AccountManagerDialog(QDialog):
    """
    Simplified: choose which Outlook account to use. No custom sender management.
    """

    def __init__(self, email_service: EmailService, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Compte Outlook")
        self.setModal(True)
        self.resize(480, 260)

        self._email_service = email_service

        self._build_ui()
        self._load_outlook_accounts()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)

        # Detected Outlook identity (transparency)
        eff = self._email_service.get_effective_sender() if hasattr(self._email_service, "get_effective_sender") else {}
        eff_name = eff.get("name") if isinstance(eff, dict) else None
        eff_mail = eff.get("email") if isinstance(eff, dict) else None
        self._detected_label = QLabel(f"Outlook détecté : {eff_name or 'Inconnu'} <{eff_mail or 'Inconnu'}>", self)
        root.addWidget(self._detected_label)

        # Outlook account selection (machine profile)
        grp_ol = QGroupBox("Choisir le compte Outlook", self)
        form_ol = QFormLayout(grp_ol)
        self._ol_combo = QComboBox(grp_ol)
        self._ol_combo.addItem("Par défaut (laisser Outlook choisir)", userData=None)
        self._ol_combo.currentIndexChanged.connect(self._on_outlook_account_changed)
        form_ol.addRow("Compte :", self._ol_combo)
        help_lbl = QLabel("Si Outlook ne peut pas appliquer le compte choisi, il utilisera son compte par défaut. Aucun email n’est bloqué.", grp_ol)
        help_lbl.setWordWrap(True)
        form_ol.addRow(help_lbl)
        root.addWidget(grp_ol)
 
        # Quick refresh of detected identity
        refresh_note = QLabel("Astuce : ouvrez Outlook et changez de profil/compte si nécessaire, puis cliquez sur Actualiser.", self)
        refresh_note.setWordWrap(True)
        root.addWidget(refresh_note)
 
        # Optional refresh button if present in future; keep dialog buttons
        # Dialog buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Close, self)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _load_outlook_accounts(self) -> None:
        try:
            accounts = self._email_service.list_outlook_accounts()
        except Exception:
            accounts = []
        # Refill combo while preserving the first default entry
        self._ol_combo.blockSignals(True)
        while self._ol_combo.count() > 1:
            self._ol_combo.removeItem(1)
        for acc in accounts:
            label = acc.get("display_name") or "Compte Outlook"
            smtp = acc.get("smtp_address")
            if smtp:
                label = f"{label} — {smtp}"
            self._ol_combo.addItem(label, userData=acc.get("id"))
        # Preselect preferred if any
        try:
            pref = self._email_service.get_preferred_outlook_account()
            if pref is None:
                self._ol_combo.setCurrentIndex(0)
            else:
                for i in range(1, self._ol_combo.count()):
                    if str(self._ol_combo.itemData(i)) == str(pref):
                        self._ol_combo.setCurrentIndex(i)
                        break
        except Exception:
            pass
        self._ol_combo.blockSignals(False)

        # Refresh detected label
        eff = self._email_service.get_effective_sender() if hasattr(self._email_service, "get_effective_sender") else {}
        eff_name = eff.get("name") if isinstance(eff, dict) else None
        eff_mail = eff.get("email") if isinstance(eff, dict) else None
        self._detected_label.setText(f"Outlook détecté : {eff_name or 'Inconnu'} <{eff_mail or 'Inconnu'}>")

    def _on_outlook_account_changed(self) -> None:
        idx = self._ol_combo.currentIndex()
        acc_id = self._ol_combo.itemData(idx) if idx >= 0 else None
        try:
            self._email_service.set_preferred_outlook_account(acc_id)
        except Exception:
            pass
