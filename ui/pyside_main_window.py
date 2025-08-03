"""PySide6 main window preserving existing business logic and behavior.

This file provides a behavior-parity UI with the former Tkinter MainWindow:
- Tabs: Stopover Pages, Stopover Mappings, Email Preview
- File selection row with "Select PDF", Outlook connect/disconnect, account status
- Status bar with message and indeterminate progress bar
- All controller callbacks and flows preserved
- Uses centralized QSS and design tokens for modern, accessible styling
"""

from typing import List, Optional, Set, Callable
import os
import sys

from PySide6.QtCore import Qt, QSize, QTimer, Slot
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox, QLabel, QPushButton,
    QHBoxLayout, QVBoxLayout, QStatusBar, QProgressBar, QTabWidget, QFrame, QSpacerItem,
    QSizePolicy, QComboBox, QStyle
)

from controllers.app_controller import AppController
from services.config_service import ConfigService
from utils.file_utils import resource_path
from ui.pyside_tokens import apply_palette, load_qss
from ui.pyside_stopover_tab import StopoverTabWidget
from ui.pyside_mapping_tab import MappingTabWidget
from ui.pyside_email_preview_tab import EmailPreviewTabWidget


class MainWindowQt(QMainWindow):
    """Main application window using PySide6 components."""

    def __init__(self, parent=None):
        super().__init__(parent)
        # Services and controller
        self.config_service = ConfigService()
        self.controller = AppController()

        # Window setup
        self._setup_window()

        # Core UI
        self._setup_central()
        self._setup_top_file_bar()
        self._setup_tabs()
        self._setup_status_bar()

        # Controller callbacks and initial state
        self._setup_controller_callbacks()
        self._initialize_state()

    def _setup_window(self):
        cfg = self.config_service.get_window_config()
        self.setWindowTitle("Analyseur d’escales PDF")
        try:
            icon_path = resource_path("assets/app.ico")
            self.setWindowIcon(QIcon(icon_path))
        except Exception:
            pass

        # Set sane minimum size first
        self.setMinimumSize(cfg.get("min_width", 1000), cfg.get("min_height", 700))

        # Robust initial geometry: center a reasonable default on the primary screen.
        # Avoid forcing maximized at startup to prevent top-left "glitchy" drag behavior on some systems.
        try:
            screen = QApplication.primaryScreen()
            if screen:
                ag = screen.availableGeometry()
                # Use 80% of available space but not smaller than a sane default
                w = max(1200, int(ag.width() * 0.8))
                h = max(800, int(ag.height() * 0.8))
                x = ag.x() + (ag.width() - w) // 2
                y = ag.y() + (ag.height() - h) // 2
                self.setGeometry(x, y, w, h)
        except Exception as _e:
            # Fallback: do nothing, Qt will choose a default geometry
            pass

        # Optional: allow explicit start maximized via environment variable
        # Set START_MAXIMIZED=1 to force maximize at startup.
        if os.environ.get("START_MAXIMIZED", "0").lower() in ("1", "true", "yes"):
            try:
                self.showMaximized()
            except Exception:
                pass


    def _setup_central(self):
        self._central = QWidget(self)
        self.setCentralWidget(self._central)
        self._root_layout = QVBoxLayout(self._central)
        self._root_layout.setContentsMargins(10, 10, 10, 10)
        self._root_layout.setSpacing(8)

    def _setup_top_file_bar(self):
        # Container
        self._file_bar = QFrame(self._central)
        self._file_bar.setObjectName("FileBar")
        file_bar_layout = QHBoxLayout(self._file_bar)
        file_bar_layout.setContentsMargins(10, 10, 10, 10)
        file_bar_layout.setSpacing(8)

        # Left side: file label + actions
        self.file_label = QLabel("Aucun PDF sélectionné", self._file_bar)
        self.file_label.setObjectName("FileLabel")

        self.select_button = QPushButton("Sélectionner un PDF", self._file_bar)
        try:
            self.select_button.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
            self.select_button.setIconSize(QSize(20, 20))
        except Exception:
            pass
        self.select_button.clicked.connect(self._select_pdf)

        # Right side: Outlook account selector + status as a compact group
        right_spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)

        # Compact container for status + dropdown
        status_row = QWidget(self._file_bar)
        status_layout = QHBoxLayout(status_row)
        status_layout.setContentsMargins(0, 0, 0, 0)
        status_layout.setSpacing(6)

        # Outlook accounts dropdown (populated from EmailService)
        self.outlook_accounts_combo = QComboBox(status_row)
        self.outlook_accounts_combo.setMinimumWidth(260)
        self.outlook_accounts_combo.setObjectName("OutlookAccountsCombo")
        # No explicit "Default (Outlook)" entry anymore; first enumerated account will be selected by default
        self.outlook_accounts_combo.currentIndexChanged.connect(self._on_outlook_account_changed)
        self.outlook_accounts_combo.setToolTip("Choisir le compte Outlook à utiliser")



        status_layout.addWidget(self.outlook_accounts_combo)

        # Layout assembly
        file_bar_layout.addWidget(self.file_label, 0, Qt.AlignLeft)
        file_bar_layout.addWidget(self.select_button, 0, Qt.AlignLeft)

        file_bar_layout.addItem(right_spacer)
        file_bar_layout.addWidget(status_row, 0, Qt.AlignRight)

        self._root_layout.addWidget(self._file_bar)

    def _setup_tabs(self):
        self.tabs = QTabWidget(self._central)

        # Stopover tab
        self.stopover_tab = StopoverTabWidget(
            on_stopover_select=self._on_stopover_select,
            controller=self.controller,
            parent=self.tabs,
        )
        self.tabs.addTab(self.stopover_tab, "Pages d’escale")

        # Mapping tab
        self.mapping_tab = MappingTabWidget(
            on_mappings_change=self._on_mappings_change,
            parent=self.tabs,
        )
        self.tabs.addTab(self.mapping_tab, "Correspondances d’escales")

        # Email preview tab
        self.email_preview_tab = EmailPreviewTabWidget(
            controller=self.controller,
            parent=self.tabs,
        )
        self.tabs.addTab(self.email_preview_tab, "Aperçu des emails")

        # Trigger a refresh when switching to the Email Preview tab to ensure default previews render
        try:
            self.tabs.currentChanged.connect(self._on_tab_changed)
        except Exception:
            pass

        self._root_layout.addWidget(self.tabs, 1)

    def _setup_status_bar(self):
        sb = QStatusBar(self)
        self.setStatusBar(sb)

        self.status_label = QLabel("Prêt", self)
        sb.addWidget(self.status_label, 1)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMaximum(0)  # indeterminate
        self.progress_bar.setMinimum(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(200)
        sb.addPermanentWidget(self.progress_bar)

    def _setup_controller_callbacks(self):
        self.controller.on_status_update = self._enqueue_status_update
        self.controller.on_progress_start = self._start_progress
        self.controller.on_progress_stop = self._stop_progress
        self.controller.on_analysis_complete = self._on_analysis_complete
        self.controller.on_outlook_connection_change = self._on_outlook_connection_change

        # Populate Outlook accounts once controller is available
        try:
            self._refresh_outlook_accounts()
        except Exception:
            pass

    def _initialize_state(self):
        # Load initial mappings
        self.mapping_tab.load_mappings()

    # ========== Behavior parity methods ==========

    def _select_pdf(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Sélectionner un fichier PDF",
            "",
            "Fichiers PDF (*.pdf);;Tous les fichiers (*)",
        )
        if filename:
            if self.controller.set_pdf_path(filename):
                self.file_label.setText(f"Sélectionné : {os.path.basename(filename)}")
                

                # Clear tabs before starting analysis
                self.stopover_tab.clear()
                self.email_preview_tab.clear()

                # Set PDF path on stopover tab
                self.stopover_tab.set_pdf_path(filename)

                # Update mapping tab
                self.mapping_tab.set_found_stopovers(set())

                # Start analysis automatically
                self.status_label.setText("Analyse du PDF…")
                self._start_progress()
                self.controller.analyze_pdf()
            else:
                QMessageBox.critical(self, "Erreur", "Veuillez sélectionner un fichier PDF valide.")

    

    def _on_analysis_complete(self, stopovers: List):
        # Update stopover tab
        self.stopover_tab.set_stopovers(stopovers)

        # Mapping tab with found stopovers
        found_codes = {s.code for s in stopovers}
        self.mapping_tab.set_found_stopovers(found_codes)

        # Email preview tab
        self.email_preview_tab.set_stopovers(stopovers)
        self.email_preview_tab.set_pdf_path(self.controller.current_pdf_path)

        # Status and progress
        self.status_label.setText(f"Escales détectées : {len(stopovers)}")
        self._stop_progress()

    def _update_status(self, message: str):
        # Immediate UI-thread safe setter (used internally after queueing)
        self.status_label.setText(message)

    def _enqueue_status_update(self, message: str):
        # Ensure status label updates occur on the GUI thread
        QTimer.singleShot(0, lambda: self._update_status(message))

    def _start_progress(self):
        self.progress_bar.setVisible(True)

    def _stop_progress(self):
        self.progress_bar.setVisible(False)
        

    def _on_mappings_change(self):
        # Refresh mapping display
        self.mapping_tab.load_mappings()
        # Reflect changes in Email Preview
        try:
            self.email_preview_tab.refresh_recipients_from_configs()
        except Exception:
            try:
                self.email_preview_tab.set_stopovers(self.controller.stopovers or [])
                if self.controller.current_pdf_path:
                    self.email_preview_tab.set_pdf_path(self.controller.current_pdf_path)
            except Exception:
                pass
        # Additionally, when configs are updated elsewhere (e.g., StopoverEmailSettingsDialog),
        # ensure the Mapping tab reflects unified mappings immediately.
        try:
            self.mapping_tab.load_mappings()
        except Exception:
            pass

    def _on_stopover_select(self, stopover):
        # Delegate to stopover tab: it handles rendering preview and status updates
        self.stopover_tab.load_page_preview(stopover, self._update_status)

    def _on_outlook_connection_change(self, connected: bool, user: Optional[str]):
        # Refresh detected sender and accounts list whenever Outlook state changes
        try:
            self._refresh_outlook_accounts()
        except Exception:
            pass

        # No separate label to update anymore
        pass

    # Removed connect/manage buttons; using a single dropdown instead
    def _toggle_outlook_connection(self):
        # No-op: connection is managed implicitly by EmailService when needed
        pass

    def _on_tab_changed(self, index: int):
        """Ensure Email Preview is rebuilt when the tab becomes active."""
        try:
            widget = self.tabs.widget(index)
            if widget is self.email_preview_tab:
                # Rebuild items to respect current filters and ensure default shows all
                self.email_preview_tab._rebuild_items_async()
        except Exception:
            pass

    def _open_account_dialog(self):
        # No-op: simplified UX has no separate dialog
        pass

    def _refresh_outlook_accounts(self):
        """Populate the Outlook accounts dropdown and update status label."""
        try:
            accounts = self.controller.email_service.list_outlook_accounts()
        except Exception:
            accounts = []
        try:
            combo = self.outlook_accounts_combo
            combo.blockSignals(True)
            # Rebuild the list entirely (no "Default" entry)
            while combo.count() > 0:
                combo.removeItem(0)
            for acc in accounts:
                # Build compact label without duplicate email
                display = acc.get("display_name") or ""
                smtp = acc.get("smtp_address") or ""
                if display and smtp:
                    # Avoid duplication like "email — email"
                    if display.strip().lower() == smtp.strip().lower():
                        label = smtp
                    else:
                        label = f"{display} — {smtp}"
                else:
                    label = display or smtp or "Compte Outlook"
                combo.addItem(label, userData=acc.get("id"))
            # Selection behavior:
            # - If a preferred account is known and still present, select it
            # - Else select the first enumerated account (index 0) if any
            try:
                pref = self.controller.email_service.get_preferred_outlook_account()
                if pref is not None:
                    found = False
                    for i in range(combo.count()):
                        if str(combo.itemData(i)) == str(pref):
                            combo.setCurrentIndex(i)
                            found = True
                            break
                    if not found and combo.count() > 0:
                        combo.setCurrentIndex(0)
                else:
                    if combo.count() > 0:
                        combo.setCurrentIndex(0)
            except Exception:
                if combo.count() > 0:
                    combo.setCurrentIndex(0)
            combo.blockSignals(False)
        except Exception as e:
            print(f"[MainWindowQt] Failed to refresh Outlook accounts: {e}")

    def _on_outlook_account_changed(self):
        try:
            idx = self.outlook_accounts_combo.currentIndex()
            acc_id = self.outlook_accounts_combo.itemData(idx) if idx >= 0 else None
            self.controller.email_service.set_preferred_outlook_account(acc_id)
        except Exception as e:
            print(f"[MainWindowQt] Failed to set preferred account: {e}")


def run_app():
    app = QApplication(sys.argv)
    apply_palette(app)
    # Load QSS
    try:
        from utils.file_utils import resource_path as rp
        qss_path = rp("ui/style_pyside.qss")
        with open(qss_path, "r", encoding="utf-8") as f:
            load_qss(app, f.read())
    except Exception:
        pass

    win = MainWindowQt()
    win.show()
    sys.exit(app.exec())
