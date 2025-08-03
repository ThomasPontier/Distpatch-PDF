"""PySide6 Mapping tab widget preserving behavior from Tkinter MappingTabComponent."""

from typing import Optional, Callable, Set, List, Dict
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QListWidget, QListWidgetItem,
    QPushButton, QMessageBox, QInputDialog, QLineEdit, QTableWidget, QTableWidgetItem, QHeaderView
)

from services.mapping_service import MappingService


class MappingTabWidget(QWidget):
    """
    PySide6 mapping tab.

    Public methods to keep parity:
      - load_mappings()
      - set_found_stopovers(codes: Set[str])
    Emits callback on_mappings_change when edits occur (wired externally).
    """

    def __init__(self, on_mappings_change: Optional[Callable[[], None]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.on_mappings_change = on_mappings_change
        self.mapping_service = MappingService()

        self._found_codes: Set[str] = set()

        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)

        # Unique table view required by spec:
        # Columns: Stopover Code | Last Sent | Emails | Status (Found in PDF)
        self.table = QTableWidget(self)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Stopover Code", "Last Sent", "Emails", "Status"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(self.table.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(self.table.EditTrigger.NoEditTriggers)
        # Enable interactive sorting by clicking on headers
        self.table.setSortingEnabled(True)
        # Default sort by Stopover Code ascending
        self.table.sortItems(0, Qt.SortOrder.AscendingOrder)

        # Actions
        actions = QHBoxLayout()

        self.add_btn = QPushButton("Add/Update", self)
        self.add_btn.setToolTip("Add a new stopover or update emails for an existing one")
        self.add_btn.clicked.connect(self._add_or_update_mapping)

        self.edit_btn = QPushButton("Edit Selected", self)
        self.edit_btn.setToolTip("Edit emails of the selected stopover")
        self.edit_btn.clicked.connect(self._edit_selected_mapping)

        self.remove_btn = QPushButton("Remove Selected", self)
        self.remove_btn.setToolTip("Remove mapping for the selected stopover")
        self.remove_btn.clicked.connect(self._remove_selected_mapping)

        actions.addStretch(1)
        actions.addWidget(self.add_btn)
        actions.addWidget(self.edit_btn)
        actions.addWidget(self.remove_btn)

        # Layout
        root.addLayout(actions)
        root.addWidget(self.table, 1)

        # UX: double-click a row to edit its emails
        self.table.itemDoubleClicked.connect(lambda _: self._edit_selected_mapping())

    # -------- Public API (parity) --------

    def load_mappings(self):
        """Reload the table with stopover mappings + last sent + found status."""
        try:
            # Data sources
            mappings: Dict[str, List[str]] = self.mapping_service.get_all_mappings()  # {CODE: [emails]}
            try:
                from services.stopover_email_service import StopoverEmailService
                ses = StopoverEmailService()
                last_sent_map: Dict[str, str] = ses._manager.get_last_sent()  # raw dict access from manager
            except Exception:
                last_sent_map = {}

            # Merge codes from mappings and found in current PDF
            all_codes = sorted(set((self._found_codes or set())) | set(mappings.keys()))

            # Build table
            self.table.setRowCount(len(all_codes))
            for row, code in enumerate(all_codes):
                emails = mappings.get(code, [])
                emails_str = ", ".join(emails) if emails else ""
                last_raw = last_sent_map.get(code) or last_sent_map.get(str(code).upper()) or ""
                status = "✓ Found" if code in (self._found_codes or set()) else "○ Not Found"

                self.table.setItem(row, 0, QTableWidgetItem(code))
                self.table.setItem(row, 1, QTableWidgetItem(last_raw))
                self.table.setItem(row, 2, QTableWidgetItem(emails_str))
                self.table.setItem(row, 3, QTableWidgetItem(status))

            # After populating, preserve current sort or default to code asc
            header = self.table.horizontalHeader()
            if self.table.isSortingEnabled():
                try:
                    section = header.sortIndicatorSection()
                    order = header.sortIndicatorOrder()
                except Exception:
                    section = 0
                    order = Qt.SortOrder.AscendingOrder
                self.table.blockSignals(True)
                try:
                    self.table.sortItems(section, order)
                finally:
                    self.table.blockSignals(False)
            else:
                self.table.sortItems(0, Qt.SortOrder.AscendingOrder)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load mappings: {str(e)}")

    def set_found_stopovers(self, codes: Set[str]):
        """Set codes found by analysis to display and refresh the table."""
        self._found_codes = set(codes or [])
        self.load_mappings()

    # -------- Internal behavior --------

    def _save_mappings(self):
        """Emit change callback to allow dependents to refresh."""
        try:
            if self.on_mappings_change:
                self.on_mappings_change()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save mappings: {str(e)}")

    # -------- Editing actions --------

    def _prompt_code_and_emails(self, initial_code: str = "", initial_emails: str = "") -> Optional[tuple]:
        # Ask for stopover code
        code, ok = QInputDialog.getText(self, "Stopover Code", "Enter stopover code (e.g., ABJ):", QLineEdit.Normal, initial_code)
        if not ok:
            return None
        code = str(code).strip().upper()
        if not code:
            QMessageBox.information(self, "Info", "Stopover code is required.")
            return None
        # Ask for emails
        emails_str, ok2 = QInputDialog.getText(
            self,
            "Email Addresses",
            "Enter email addresses separated by commas or semicolons:",
            QLineEdit.Normal,
            initial_emails
        )
        if not ok2:
            return None
        # Normalize separators and strip
        raw = emails_str.replace(";", ",")
        emails = [e.strip() for e in raw.split(",") if e.strip()]
        return (code, emails)

    def _add_or_update_mapping(self):
        """Add a new stopover or update emails for an existing one."""
        # Prefill code with first selected item if any
        initial_code = ""
        # Get selected row code from table if any
        selected = self.table.currentRow()
        if selected is not None and selected >= 0:
            code_item = self.table.item(selected, 0)
            if code_item:
                initial_code = code_item.text().strip().upper()

        result = self._prompt_code_and_emails(initial_code=initial_code, initial_emails="")
        if not result:
            return
        code, emails = result
        try:
            # For add/update, set complete list via MappingService/ConfigManager
            # Use service facade to ensure normalization and signals
            from services.config_manager import get_config_manager
            mgr = get_config_manager()
            mgr.set_mapping(code, emails)
            mgr.add_stopover(code)  # ensure enabled/visible
            self.load_mappings()
            if self.on_mappings_change:
                self.on_mappings_change()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to add/update mapping: {str(e)}")

    def _edit_selected_mapping(self):
        """Edit emails for the selected mapping row without asking for the code again."""
        row = self.table.currentRow()
        if row is None or row < 0:
            QMessageBox.information(self, "Info", "Please select a mapping to edit.")
            return
        code_item = self.table.item(row, 0)
        emails_item = self.table.item(row, 2)
        code = code_item.text().strip().upper() if code_item else ""
        current = emails_item.text().strip() if emails_item else ""
        # Only prompt for emails; keep the selected code unchanged
        emails_str, ok = QInputDialog.getText(
            self,
            "Edit Email Addresses",
            f"Enter email addresses for {code} (separated by commas or semicolons):",
            QLineEdit.Normal,
            current
        )
        if not ok:
            return
        raw = emails_str.replace(";", ",")
        emails = [e.strip() for e in raw.split(",") if e.strip()]
        try:
            from services.config_manager import get_config_manager
            mgr = get_config_manager()
            mgr.set_mapping(code, emails)
            mgr.add_stopover(code)
            self.load_mappings()
            if self.on_mappings_change:
                self.on_mappings_change()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to edit mapping: {str(e)}")

    def _remove_selected_mapping(self):
        """Remove mapping for the selected stopover."""
        row = self.table.currentRow()
        if row is None or row < 0:
            QMessageBox.information(self, "Info", "Please select a mapping to remove.")
            return
        code_item = self.table.item(row, 0)
        code = code_item.text().strip().upper() if code_item else ""
        confirm = QMessageBox.question(
            self,
            "Confirm Removal",
            f"Remove all email mappings for {code}?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if confirm != QMessageBox.Yes:
            return
        try:
            from services.config_manager import get_config_manager
            mgr = get_config_manager()
            mgr.remove_mapping(code)
            self.load_mappings()
            if self.on_mappings_change:
                self.on_mappings_change()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to remove mapping: {str(e)}")
