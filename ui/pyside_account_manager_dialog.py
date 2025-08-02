from __future__ import annotations
from typing import Optional, List
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QListWidgetItem, QLineEdit,
    QPushButton, QGroupBox, QMessageBox, QWidget, QFormLayout, QDialogButtonBox
)
from services.accounts_service import AccountsService, SenderAccount
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
    Native PySide6 dialog to manage sender accounts stored in config/accounts.json via AccountsService.
    - List existing senders (Outlook/custom)
    - Add custom sender (email)
    - Connect Outlook to add the currently active Outlook identity
    - Remove selected sender
    - Choose selected sender (persist selected_sender_id)
    It does NOT force SendUsingAccount; EmailService logs selected sender for debug and lets Outlook decide.
    """

    def __init__(self, email_service: EmailService, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Manage Sender Accounts")
        self.setModal(True)
        self.resize(500, 420)

        self._accounts = AccountsService()
        self._email_service = email_service

        self._build_ui()
        self._load_list()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)

        # Current selection info
        self._current_label = QLabel("", self)
        root.addWidget(self._current_label)

        # List of senders
        grp_list = QGroupBox("Configured senders", self)
        vlist = QVBoxLayout(grp_list)
        self._list = QListWidget(grp_list)
        self._list.itemSelectionChanged.connect(self._on_selection_changed)
        vlist.addWidget(self._list)

        row_btns = QHBoxLayout()
        self._btn_remove = QPushButton("Remove", grp_list)
        self._btn_remove.clicked.connect(self._remove_selected)
        self._btn_remove.setEnabled(False)
        row_btns.addWidget(self._btn_remove)

        self._btn_choose = QPushButton("Use selected", grp_list)
        self._btn_choose.clicked.connect(self._choose_selected)
        self._btn_choose.setEnabled(False)
        row_btns.addWidget(self._btn_choose)
        vlist.addLayout(row_btns)

        root.addWidget(grp_list)

        # Add custom sender
        grp_custom = QGroupBox("Add custom sender", self)
        form = QFormLayout(grp_custom)
        self._custom_email = QLineEdit(grp_custom)
        form.addRow("Email:", self._custom_email)
        btn_add_custom = QPushButton("Add", grp_custom)
        btn_add_custom.clicked.connect(self._add_custom_sender)
        form.addRow(QWidget(), btn_add_custom)
        root.addWidget(grp_custom)

        # Outlook section
        grp_outlook = QGroupBox("Add current Outlook account", self)
        vb_out = QVBoxLayout(grp_outlook)
        self._lbl_out_info = QLabel("Attempt to read the active Outlook user via COM.", grp_outlook)
        vb_out.addWidget(self._lbl_out_info)
        btn_connect = QPushButton("Connect Outlook (optional)", grp_outlook)
        btn_connect.clicked.connect(self._connect_outlook_safe)
        vb_out.addWidget(btn_connect)
        root.addWidget(grp_outlook)

        # Close
        buttons = QDialogButtonBox(QDialogButtonBox.Close, self)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _load_list(self) -> None:
        self._list.clear()
        selected = self._accounts.get_selected_sender()
        for s in self._accounts.get_senders():
            text = f"{s.name} <{s.email}> [{s.type}]"
            if selected and s.id == selected.id:
                text = "✓ " + text
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, s.id)
            self._list.addItem(item)
        sel_text = selected.email if selected else "None"
        self._current_label.setText(f"Currently selected: {sel_text}")

    def _on_selection_changed(self) -> None:
        has = bool(self._list.selectedItems())
        self._btn_remove.setEnabled(has)
        self._btn_choose.setEnabled(has)

    def _add_custom_sender(self) -> None:
        email = (self._custom_email.text() or "").strip()
        if not email:
            QMessageBox.information(self, "Info", "Enter an email to add.")
            return
        try:
            acc = self._accounts.ensure_custom_sender(email)
            # also set as selected
            self._accounts.set_selected_sender(acc.id)
            # propagate to EmailService
            self._email_service.set_current_account("custom", {"type": "custom", "name": acc.name, "email": acc.email})
            self._custom_email.clear()
            self._load_list()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to add custom sender: {e}")

    def _connect_outlook_safe(self) -> None:
        """
        Best-effort Outlook detection without bloquer l'UI par une erreur.
        Si Outlook n'est pas accessible, on affiche une info non bloquante et on continue.
        """
        if not _WIN32_AVAILABLE:
            QMessageBox.information(self, "Outlook", "pywin32 non disponible. Vous pouvez continuer avec un expéditeur personnalisé.")
            return
        try:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
            # Certaines installations retournent une erreur COM quand le profil n'est pas prêt.
            # On encapsule chaque étape et on bascule en info non bloquante si ça échoue.
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception:
                QMessageBox.information(self, "Outlook", "Outlook n'est pas disponible pour le moment. Utilisez un expéditeur personnalisé.")
                return
            try:
                ns = outlook.GetNamespace("MAPI")
                current_user = ns.CurrentUser
            except Exception:
                QMessageBox.information(self, "Outlook", "Impossible de lire le profil Outlook actif. Utilisez un expéditeur personnalisé.")
                return

            display_name = getattr(current_user, "Name", None) or "Outlook User"
            email = getattr(current_user, "Address", None) or ""
            acc = self._accounts.ensure_outlook_sender(display_name, email)
            self._accounts.set_selected_sender(acc.id)
            # update EmailService selection (no forcing on send)
            self._email_service.set_current_account(acc.id, {"type": "outlook", "name": acc.name, "email": acc.email})
            self._load_list()
            QMessageBox.information(self, "Outlook", f"Compte Outlook ajouté: {display_name} ({email})")
        except Exception:
            # En cas d'échec inattendu, ne pas afficher une erreur bloquante
            QMessageBox.information(self, "Outlook", "Outlook n'est pas prêt. Vous pouvez continuer avec un expéditeur personnalisé.")

    def _remove_selected(self) -> None:
        items = self._list.selectedItems()
        if not items:
            return
        sender_id = items[0].data(Qt.UserRole)
        confirm = QMessageBox.question(self, "Confirm", "Remove selected sender?")
        if confirm != QMessageBox.Yes:
            return
        try:
            self._accounts.remove_sender(sender_id)
            # if removed selected one, EmailService keeps last selection; we do not clear explicitly
            self._load_list()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to remove sender: {e}")

    def _choose_selected(self) -> None:
        items = self._list.selectedItems()
        if not items:
            return
        sender_id = items[0].data(Qt.UserRole)
        try:
            self._accounts.set_selected_sender(sender_id)
            sel = self._accounts.get_selected_sender()
            # propagate selection to EmailService
            if sel:
                self._email_service.set_current_account(sel.id, {"type": sel.type, "name": sel.name, "email": sel.email})
            self._load_list()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to choose sender: {e}")
