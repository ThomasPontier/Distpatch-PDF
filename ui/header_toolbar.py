from PySide6.QtCore import Signal, Qt
from PySide6.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QLabel, QPushButton, QLineEdit, QComboBox, QSizePolicy, QStyle
from typing import Optional


class HeaderToolbar(QWidget):
    """
    Fixed global header/toolbar.

    Emits:
      - templateChanged(subject: str, body: str)
      - filterChanged(text: str, status: str)
      - sendAllClicked()
      - settingsClicked()
      - outlookConnectClicked()  # New: request Outlook connection
    """
    templateChanged = Signal(str, str)
    filterChanged = Signal(str, str)
    sendAllClicked = Signal()
    settingsClicked = Signal()
    outlookConnectClicked = Signal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(12)
        try:
            # Compact header -> uniform 18x18 icons
            self.setIconSize(QSize(18, 18))  # type: ignore[attr-defined]
        except Exception:
            pass

        # Account status (no email shown)
        self.accountLabel = QLabel("Outlook : non connecté")
        self.accountLabel.setObjectName("Badge")
        self.accountLabel.setToolTip("Statut de connexion Outlook")
        root.addWidget(self.accountLabel, 0, Qt.AlignVCenter)

        # Outlook connect/disconnect
        self.connectBtn = QPushButton("Connecter Outlook")
        try:
            # Initial state = not connected -> "Connecter Outlook"
            self.connectBtn.setIcon(self.style().standardIcon(QStyle.SP_DialogYesButton))
        except Exception:
            pass
        self.connectBtn.clicked.connect(self.outlookConnectClicked.emit)
        root.addWidget(self.connectBtn, 0)

        # Search/filter
        self.searchEdit = QLineEdit()
        self.searchEdit.setPlaceholderText("Rechercher des escales…")
        self.searchEdit.textChanged.connect(self._emit_filter_changed)

        self.statusCombo = QComboBox()
        self.statusCombo.addItems(["Tous", "En attente", "Envoi", "Réussi", "Échec", "En file"])
        self.statusCombo.currentTextChanged.connect(self._emit_filter_changed)

        root.addWidget(self.searchEdit, 1)
        root.addWidget(self.statusCombo, 0)

        # Template mini editor (subject only inline; body via dialog typically, but keep a quick body line)
        self.subjectEdit = QLineEdit()
        self.subjectEdit.setPlaceholderText("Modèle d’objet global")
        self.subjectEdit.textChanged.connect(self._emit_template_changed)

        self.bodyEdit = QLineEdit()
        self.bodyEdit.setPlaceholderText("Modèle de corps global (rapide)")
        self.bodyEdit.textChanged.connect(self._emit_template_changed)

        root.addWidget(self.subjectEdit, 2)
        root.addWidget(self.bodyEdit, 3)

        # Actions
        self.sendAllBtn = QPushButton("Envoyer à toutes les escales")
        try:
            self.sendAllBtn.setIcon(self.style().standardIcon(QStyle.SP_ArrowForward))
        except Exception:
            pass
        self.sendAllBtn.clicked.connect(self.sendAllClicked.emit)
        self.sendAllBtn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        root.addWidget(self.sendAllBtn, 0)

        self.settingsBtn = QPushButton("Paramètres")
        try:
            self.settingsBtn.setIcon(self.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        except Exception:
            pass
        self.settingsBtn.clicked.connect(self.settingsClicked.emit)
        root.addWidget(self.settingsBtn, 0)

        # Stretch at end
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        root.addWidget(spacer)

    def setAccountEmail(self, email: Optional[str]):
        # Deprecated: we no longer display the connected email address in UI.
        # Keep method for backward compatibility; update status text generically.
        self.accountLabel.setText("Outlook : connecté" if email else "Outlook : non connecté")

    def setOutlookConnected(self, connected: bool, email: Optional[str]):
        """Update connect button and label state (without showing email)."""
        self.connectBtn.setText("Reconnecter Outlook" if connected else "Connecter Outlook")
        try:
            # Update icon according to current semantics
            if connected:
                # Reconnect/refresh semantics
                self.connectBtn.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
            else:
                self.connectBtn.setIcon(self.style().standardIcon(QStyle.SP_DialogYesButton))
        except Exception:
            pass
        # Do not display the email address; show generic status only.
        self.accountLabel.setText("Outlook : connecté" if connected else "Outlook : non connecté")

    def setGlobalTemplate(self, subject: str, body: str):
        # Avoid signal storms when programmatically setting
        try:
            self.subjectEdit.blockSignals(True)
            self.bodyEdit.blockSignals(True)
            self.subjectEdit.setText(subject or "")
            self.bodyEdit.setText(body or "")
        finally:
            self.subjectEdit.blockSignals(False)
            self.bodyEdit.blockSignals(False)

    def _emit_template_changed(self):
        self.templateChanged.emit(self.subjectEdit.text(), self.bodyEdit.text())

    def _emit_filter_changed(self):
        self.filterChanged.emit(self.searchEdit.text(), self.statusCombo.currentText())
