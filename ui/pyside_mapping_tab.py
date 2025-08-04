"""PySide6 Mapping tab widget preserving behavior from Tkinter MappingTabComponent."""

from typing import Optional, Callable, Set, List, Dict
from PySide6.QtCore import Qt, QSize
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QListWidget, QListWidgetItem,
    QPushButton, QMessageBox, QInputDialog, QLineEdit, QTableWidget, QTableWidgetItem, QHeaderView, QStyle,
    QDialog, QFormLayout
)
from PySide6.QtGui import QIcon

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
        # Single source of truth: read via StopoverEmailService to ensure parity with dialog and preview
        from services.stopover_email_service import StopoverEmailService
        self._email_service = StopoverEmailService()

        self._found_codes: Set[str] = set()

        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)

        # Unique table view required by spec:
        # Columns: Stopover Code | Last Sent | Emails | Cc/Bcc | Status (Found in PDF)
        self.table = QTableWidget(self)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Code escale", "Dernier envoi", "Emails", "CC/CCI", "Statut"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(self.table.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(self.table.EditTrigger.NoEditTriggers)
        # Enable interactive sorting by clicking on headers
        self.table.setSortingEnabled(True)
        # Default sort by Stopover Code ascending
        self.table.sortItems(0, Qt.SortOrder.AscendingOrder)

        # Actions
        actions = QHBoxLayout()

        self.add_btn = QPushButton("Ajouter", self)
        self.add_btn.setToolTip("Ajouter une escale et configurer ses emails (À/CC/CCI)")
        try:
            self.add_btn.setIcon(QIcon())
            self.add_btn.setIconSize(QSize(20, 20))
        except Exception:
            pass
        self.add_btn.clicked.connect(self._add_or_update_mapping)

        self.edit_btn = QPushButton("Modifier la sélection", self)
        self.edit_btn.setToolTip("Modifier les emails de l’escale sélectionnée")
        try:
            # Remove icon as per spec
            self.edit_btn.setIcon(QIcon())
            self.edit_btn.setIconSize(QSize(20, 20))
        except Exception:
            pass
        self.edit_btn.clicked.connect(self._edit_selected_mapping)
        # Use the dedicated recipients editor dialog for both Ajouter and Modifier
        try:
            from ui.components.recipient_editor_dialog import RecipientEditorDialog
            self._RecipientEditorDialog = RecipientEditorDialog
        except Exception:
            self._RecipientEditorDialog = None

        self.remove_btn = QPushButton("Supprimer la sélection", self)
        self.remove_btn.setToolTip("Supprimer la correspondance pour l’escale sélectionnée")
        try:
            # Remove icon as per spec
            self.remove_btn.setIcon(QIcon())
            self.remove_btn.setIconSize(QSize(20, 20))
        except Exception:
            pass
        # Apply inline red styling for this destructive action button
        self.remove_btn.setStyleSheet(
            "QPushButton { background-color: #D32F2F; color: white; }"
            "QPushButton:hover { background-color: #B71C1C; }"
            "QPushButton:disabled { background-color: #BDBDBD; color: #EEEEEE; }"
        )
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
        # Live refresh when config changes anywhere
        try:
            from services.config_manager import get_config_manager
            get_config_manager().on_mappings_changed(lambda _: self.load_mappings())
        except Exception:
            pass

    # -------- Public API (parity) --------

    def load_mappings(self):
        """Reload the table with stopover mappings + last sent + found status."""
        try:
            # Data sources
            # Unify with StopoverEmailService so edits from the email dialog are always reflected here.
            from services.stopover_email_service import StopoverEmailService
            ses = StopoverEmailService()
            all_cfgs = ses.get_all_configs()  # { CODE: StopoverEmailConfig }
            last_sent_map: Dict[str, str] = ses._manager.get_last_sent()  # raw dict access from manager

            # Merge codes from unified configs (may include codes not in legacy mappings) and found in current PDF
            all_codes = sorted(set((self._found_codes or set())) | set(all_cfgs.keys()))

            # Build table
            self.table.setRowCount(len(all_codes))
            for row, code in enumerate(all_codes):
                cfg = all_cfgs.get(code) or all_cfgs.get(str(code).upper())
                to_list = list((cfg.recipients if cfg else []) or [])
                cc_list = list((cfg.cc_recipients if cfg else []) or [])
                bcc_list = list((cfg.bcc_recipients if cfg else []) or [])
                emails_str = ", ".join(to_list) if to_list else ""
                ccbcc_parts = []
                if cc_list:
                    ccbcc_parts.append(f"CC: {', '.join(cc_list)}")
                if bcc_list:
                    ccbcc_parts.append(f"CCI: {', '.join(bcc_list)}")
                ccbcc_str = " | ".join(ccbcc_parts)
                last_raw = last_sent_map.get(code) or last_sent_map.get(str(code).upper()) or ""

                # Format 'Dernier envoi' in local France time (UTC+2) as 'YYYY-MM-DD HH:MM'
                def _format_last_sent(iso_s: str) -> str:
                    if not iso_s:
                        return ""
                    try:
                        from datetime import datetime, timedelta, timezone
                        s = iso_s.strip()
                        # Accept both with/without trailing 'Z' and with fractional seconds
                        if s.endswith("Z"):
                            s = s[:-1]
                            tzinfo = timezone.utc
                        else:
                            tzinfo = timezone.utc
                        # Try multiple patterns
                        dt = None
                        for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
                            try:
                                dt = datetime.strptime(s, fmt)
                                break
                            except Exception:
                                continue
                        if dt is None:
                            return iso_s  # fallback: raw
                        dt = dt.replace(tzinfo=tzinfo)
                        # Convert to France time (fixed offset +2h as requested)
                        dt_fr = dt.astimezone(timezone(timedelta(hours=2)))
                        return dt_fr.strftime("%Y-%m-%d %H:%M")
                    except Exception:
                        return iso_s

                last_display = _format_last_sent(last_raw)
                status = "✓ Présente" if code in (self._found_codes or set()) else "○ Absente"

                self.table.setItem(row, 0, QTableWidgetItem(code))
                self.table.setItem(row, 1, QTableWidgetItem(last_display))
                self.table.setItem(row, 2, QTableWidgetItem(emails_str))
                self.table.setItem(row, 3, QTableWidgetItem(ccbcc_str))
                self.table.setItem(row, 4, QTableWidgetItem(status))

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
            QMessageBox.critical(self, "Erreur", f"Échec du chargement des correspondances : {str(e)}")

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
            QMessageBox.critical(self, "Erreur", f"Échec de l’enregistrement des correspondances : {str(e)}")

    # -------- Editing actions --------

    def _add_or_update_mapping(self):
        """Ajouter une escale et configurer ses emails (À/CC/CCI) via le même dialogue factorisé que 'Modifier'."""
        # Demander le code d’escale
        code, ok = QInputDialog.getText(
            self,
            "Code d’escale",
            "Saisir le code d’escale (ex. : ABJ) :",
            QLineEdit.Normal,
            ""
        )
        if not ok:
            return
        code = str(code).strip().upper()
        if not code:
            QMessageBox.information(self, "Info", "Le code d’escale est requis.")
            return

        # Charger la config existante (ou défaut)
        try:
            cfg = self._email_service.get_config(code)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Échec de lecture de la configuration : {str(e)}")
            return

        # Ouvrir exactement le même éditeur de destinataires (composant factorisé)
        try:
            if not self._RecipientEditorDialog:
                from ui.components.recipient_editor_dialog import RecipientEditorDialog
                self._RecipientEditorDialog = RecipientEditorDialog
            accepted, to_vals, cc_vals, bcc_vals = self._RecipientEditorDialog.open(
                self,
                code,
                list(cfg.recipients or []),
                list(cfg.cc_recipients or []),
                list(cfg.bcc_recipients or []),
            )
            if not accepted:
                return
            cfg.recipients = list(to_vals or [])
            cfg.cc_recipients = list(cc_vals or [])
            cfg.bcc_recipients = list(bcc_vals or [])
            cfg.is_enabled = True
            self._email_service.save_config(cfg)
            self.load_mappings()
            if self.on_mappings_change:
                self.on_mappings_change()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Échec de l’ajout : {str(e)}")

    def _edit_selected_mapping(self):
        """Modifier les emails d’une escale via le même dialogue factorisé (RecipientEditorDialog)."""
        row = self.table.currentRow()
        if row is None or row < 0:
            QMessageBox.information(self, "Info", "Veuillez sélectionner une correspondance à modifier.")
            return
        code_item = self.table.item(row, 0)
        code = code_item.text().strip().upper() if code_item else ""
        if not code:
            QMessageBox.information(self, "Info", "Code d’escale introuvable.")
            return
        try:
            cfg = self._email_service.get_config(code)
            from ui.components.recipient_editor_dialog import RecipientEditorDialog
            accepted, to_vals, cc_vals, bcc_vals = RecipientEditorDialog.open(
                self,
                code,
                list(cfg.recipients or []),
                list(cfg.cc_recipients or []),
                list(cfg.bcc_recipients or []),
            )
            if not accepted:
                return
            cfg.recipients = list(to_vals or [])
            cfg.cc_recipients = list(cc_vals or [])
            cfg.bcc_recipients = list(bcc_vals or [])
            cfg.is_enabled = True
            self._email_service.save_config(cfg)
            self.load_mappings()
            if self.on_mappings_change:
                self.on_mappings_change()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Échec de la modification : {str(e)}")

    def _remove_selected_mapping(self):
        """Remove mapping for the selected stopover (clears To/Cc/Bcc and 'Dernier envoi')."""
        row = self.table.currentRow()
        if row is None or row < 0:
            QMessageBox.information(self, "Info", "Veuillez sélectionner une correspondance à supprimer.")
            return
        code_item = self.table.item(row, 0)
        code = code_item.text().strip().upper() if code_item else ""
        confirm = QMessageBox.question(
            self,
            "Confirmer la suppression",
            f"Supprimer toutes les correspondances email pour {code} ?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if confirm != QMessageBox.Yes:
            return
        try:
            # Clear all recipients and last sent, then remove the stopover entirely
            cfg = self._email_service.get_config(code)
            cfg.recipients = []
            cfg.cc_recipients = []
            cfg.bcc_recipients = []
            cfg.last_sent_at = None
            self._email_service.save_config(cfg)

            # Remove mapping + stopover + last_sent from unified config
            try:
                from services.stopover_email_service import StopoverEmailService
                StopoverEmailService().delete_config(code)
            except Exception:
                # Fallback direct removal to be extra safe
                try:
                    from services.config_manager import get_config_manager
                    mgr = get_config_manager()
                    mgr.remove_mapping(code)
                    mgr.clear_last_sent_normalized(code)
                    mgr.remove_stopover(code)
                except Exception:
                    pass

            # Refresh UI so the row disappears and 'Dernier envoi' is cleared
            self.load_mappings()
            if self.on_mappings_change:
                self.on_mappings_change()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Échec de la suppression : {str(e)}")
