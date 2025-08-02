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
    QSizePolicy
)

from controllers.app_controller import AppController
from services.config_service import ConfigService
from utils.file_utils import resource_path
from ui.pyside_tokens import apply_palette, load_qss
from ui.pyside_stopover_tab import StopoverTabWidget
from ui.pyside_mapping_tab import MappingTabWidget
from ui.pyside_email_preview_tab import EmailPreviewTabWidget
from ui.stopover_email_dialog import StopoverEmailDialog  # kept for parity if needed
from ui.email_dialog import EmailDispatchDialog
from ui.account_dialog import AccountDialog
from ui.pyside_account_manager_dialog import AccountManagerDialog


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
        self.setWindowTitle("PDF Stopover Analyzer")
        try:
            icon_path = resource_path("assets/app.ico")
            self.setWindowIcon(QIcon(icon_path))
        except Exception:
            pass

        # Maximize if possible; otherwise set reasonable geometry and min size
        self.showMaximized()
        self.setMinimumSize(cfg.get("min_width", 1000), cfg.get("min_height", 700))

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
        self.file_label = QLabel("No PDF selected", self._file_bar)
        self.file_label.setObjectName("FileLabel")

        self.select_button = QPushButton("Select PDF", self._file_bar)
        self.select_button.clicked.connect(self._select_pdf)

        
        

        # Right side: account / outlook controls
        right_spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)

        self.account_status_label = QLabel("Account: None", self._file_bar)
        self.account_button = QPushButton("Manage Accounts", self._file_bar)
        self.account_button.clicked.connect(self._open_account_dialog)

        self.outlook_button = QPushButton("Connect Outlook", self._file_bar)
        self.outlook_button.clicked.connect(self._toggle_outlook_connection)

        # Layout assembly
        file_bar_layout.addWidget(self.file_label, 0, Qt.AlignLeft)
        file_bar_layout.addWidget(self.select_button, 0, Qt.AlignLeft)
        
        file_bar_layout.addItem(right_spacer)
        file_bar_layout.addWidget(self.account_status_label, 0, Qt.AlignRight)
        file_bar_layout.addWidget(self.account_button, 0, Qt.AlignRight)
        file_bar_layout.addWidget(self.outlook_button, 0, Qt.AlignRight)

        self._root_layout.addWidget(self._file_bar)

    def _setup_tabs(self):
        self.tabs = QTabWidget(self._central)

        # Stopover tab
        self.stopover_tab = StopoverTabWidget(
            on_stopover_select=self._on_stopover_select,
            controller=self.controller,
            parent=self.tabs,
        )
        self.tabs.addTab(self.stopover_tab, "Stopover Pages")

        # Mapping tab
        self.mapping_tab = MappingTabWidget(
            on_mappings_change=self._on_mappings_change,
            parent=self.tabs,
        )
        self.tabs.addTab(self.mapping_tab, "Stopover Mappings")

        # Email preview tab
        self.email_preview_tab = EmailPreviewTabWidget(
            controller=self.controller,
            parent=self.tabs,
        )
        self.tabs.addTab(self.email_preview_tab, "Email Preview")

        self._root_layout.addWidget(self.tabs, 1)

    def _setup_status_bar(self):
        sb = QStatusBar(self)
        self.setStatusBar(sb)

        self.status_label = QLabel("Ready", self)
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

    def _initialize_state(self):
        self.account_status_label.setText("Account: None")
        self.outlook_button.setText("Connect Outlook")
        # Load initial mappings
        self.mapping_tab.load_mappings()

    # ========== Behavior parity methods ==========

    def _select_pdf(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select PDF file",
            "",
            "PDF files (*.pdf);;All files (*)",
        )
        if filename:
            if self.controller.set_pdf_path(filename):
                self.file_label.setText(f"Selected: {os.path.basename(filename)}")
                

                # Clear tabs before starting analysis
                self.stopover_tab.clear()
                self.email_preview_tab.clear()

                # Set PDF path on stopover tab
                self.stopover_tab.set_pdf_path(filename)

                # Update mapping tab
                self.mapping_tab.set_found_stopovers(set())

                # Start analysis automatically
                self.status_label.setText("Analyzing PDF...")
                self._start_progress()
                self.controller.analyze_pdf()
            else:
                QMessageBox.critical(self, "Error", "Please select a valid PDF file.")

    

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
        self.status_label.setText(f"Found {len(stopovers)} stopover(s)")
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

    def _on_stopover_select(self, stopover):
        # Delegate to stopover tab: it handles rendering preview and status updates
        self.stopover_tab.load_page_preview(stopover, self._update_status)

    def _on_outlook_connection_change(self, connected: bool, user: Optional[str]):
        if connected:
            self.account_status_label.setText("Account: Connected")
            self.outlook_button.setText("Disconnect Outlook")
        else:
            self.account_status_label.setText("Account: None")
            self.outlook_button.setText("Connect Outlook")

    def _toggle_outlook_connection(self):
        self.controller.toggle_outlook_connection()

    def _open_account_dialog(self):
        """
        Open the (new) native PySide6 account manager dialog.
        """
        try:
            dlg = AccountManagerDialog(email_service=self.controller.email_service, parent=self)
            res = dlg.exec()
            # After dialog closes, update status label from controller
            try:
                account_name = self.controller.email_service.get_current_account_name()
                if account_name:
                    self.account_status_label.setText(f"Account: {account_name}")
                else:
                    self.account_status_label.setText("Account: None")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open account dialog: {str(e)}")


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
