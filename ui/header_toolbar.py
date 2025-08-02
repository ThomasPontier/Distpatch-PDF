from PySide6.QtCore import Signal, Qt
from PySide6.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QLabel, QPushButton, QLineEdit, QComboBox, QSizePolicy
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

        # Account status (no email shown)
        self.accountLabel = QLabel("Outlook: Disconnected")
        self.accountLabel.setObjectName("Badge")
        self.accountLabel.setToolTip("Outlook connection status")
        root.addWidget(self.accountLabel, 0, Qt.AlignVCenter)

        # Outlook connect/disconnect
        self.connectBtn = QPushButton("Connect Outlook")
        self.connectBtn.clicked.connect(self.outlookConnectClicked.emit)
        root.addWidget(self.connectBtn, 0)

        # Search/filter
        self.searchEdit = QLineEdit()
        self.searchEdit.setPlaceholderText("Search stopovers...")
        self.searchEdit.textChanged.connect(self._emit_filter_changed)

        self.statusCombo = QComboBox()
        self.statusCombo.addItems(["All", "Pending", "Sending", "Success", "Failed", "Queued"])
        self.statusCombo.currentTextChanged.connect(self._emit_filter_changed)

        root.addWidget(self.searchEdit, 1)
        root.addWidget(self.statusCombo, 0)

        # Template mini editor (subject only inline; body via dialog typically, but keep a quick body line)
        self.subjectEdit = QLineEdit()
        self.subjectEdit.setPlaceholderText("Global subject template")
        self.subjectEdit.textChanged.connect(self._emit_template_changed)

        self.bodyEdit = QLineEdit()
        self.bodyEdit.setPlaceholderText("Global body template (quick)")
        self.bodyEdit.textChanged.connect(self._emit_template_changed)

        root.addWidget(self.subjectEdit, 2)
        root.addWidget(self.bodyEdit, 3)

        # Actions
        self.sendAllBtn = QPushButton("Send to All Stopovers")
        self.sendAllBtn.clicked.connect(self.sendAllClicked.emit)
        self.sendAllBtn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        root.addWidget(self.sendAllBtn, 0)

        self.settingsBtn = QPushButton("Settings")
        self.settingsBtn.clicked.connect(self.settingsClicked.emit)
        root.addWidget(self.settingsBtn, 0)

        # Stretch at end
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        root.addWidget(spacer)

    def setAccountEmail(self, email: Optional[str]):
        # Deprecated: we no longer display the connected email address in UI.
        # Keep method for backward compatibility; update status text generically.
        self.accountLabel.setText("Outlook: Connected" if email else "Outlook: Disconnected")

    def setOutlookConnected(self, connected: bool, email: Optional[str]):
        """Update connect button and label state (without showing email)."""
        self.connectBtn.setText("Reconnect Outlook" if connected else "Connect Outlook")
        # Do not display the email address; show generic status only.
        self.accountLabel.setText("Outlook: Connected" if connected else "Outlook: Disconnected")

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
