"""PySide6 Email Preview tab preserving behavior from Tkinter EmailPreviewTabComponent."""

from typing import List, Optional, Dict, Tuple
from PySide6.QtCore import Qt, QTimer, QSize
from PySide6.QtGui import QFont, QTextOption
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QTextEdit, QPushButton, QMessageBox,
    QScrollArea, QSizePolicy, QFrame, QLineEdit, QComboBox
)
from models.stopover import Stopover
from services.mapping_service import MappingService
from services.email_service import EmailService
from services.stopover_email_service import StopoverEmailService
from services.config_manager import get_config_manager
from ui.pdf_preview import PdfPreview

# Simple persistence for per-stopover overrides (subject/body).
# Stored in config manager under a dedicated key; cleared on template change.
OVERRIDES_KEY = "email_overrides"


class StopoverEmailPreviewItem(QWidget):
    """
    One stopover row: left = subject/body text, right = page preview sized to stopover page.
    The right preview embeds the PDF preview adapted to the page size (via PDFPreviewWidget) if available.
    Falls back to a placeholder frame sized proportionally to A4/page size.
    """
    def __init__(self, stopover: Stopover, recipients: List[str], subject: str, body: str, page_size_mm: Tuple[float, float] = (210.0, 297.0), parent: Optional[QWidget] = None, pdf_path: Optional[str] = None, on_send_one=None):
        super().__init__(parent)
        self.stopover = stopover
        self.recipients = recipients
        self.subject = subject
        self.body = body
        self.page_size_mm = page_size_mm  # width, height in mm (default A4 portrait)
        self.pdf_path = pdf_path
        self._on_send_one = on_send_one

        root = QHBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(12)
        # Enlarge overall item to give more room for PDF preview
        # Increase item height to give more room to the preview
        self.setMinimumHeight(850)

        # Left panel: subject + body (read-only)
        left = QVBoxLayout()
        title = QLabel(f"Escale {stopover.code} — Page {getattr(stopover, 'page_number', '')}")
        font = QFont()
        font.setBold(True)
        title.setFont(font)
        left.addWidget(title)

        to_line = ", ".join(recipients) if recipients else "Aucune adresse email"
        header = QLabel(f"Objet: {subject}\nÀ: {to_line}")
        header.setTextInteractionFlags(Qt.TextSelectableByMouse)
        left.addWidget(header)

        self.body_view = QTextEdit()
        # Rendre le corps éditable et persister automatiquement
        self.body_view.setReadOnly(False)
        self.body_view.setPlainText(body)
        self.body_view.setWordWrapMode(QTextOption.WordWrap)
        # Persistance automatique à chaque modification (debounce léger via QTimer du parent)
        self.body_view.textChanged.connect(lambda: QTimer.singleShot(200, self._auto_persist_body))
        left.addWidget(self.body_view, 1)

        # Actions row
        actions_row = QHBoxLayout()
        actions_row.addStretch(1)
        self.send_one_btn = QPushButton("Envoyer individuellement")
        self.send_one_btn.clicked.connect(self._emit_send_one)
        actions_row.addWidget(self.send_one_btn)


        # Info if no recipients
        if not recipients:
            warn = QLabel("Aucun destinataire configuré pour cette escale")
            warn.setStyleSheet("color: #b58900;")
            actions_row.addWidget(warn)

        left.addLayout(actions_row)

        # Allocate 30% to the left (text) and 70% to the right (PDF preview)
        root.addLayout(left, 3)

        # Right panel: preview placeholder scaled to page size
        right = QVBoxLayout()
        right.setContentsMargins(0, 0, 0, 0)

        preview_label = QLabel("Aperçu PDF (adapté à la taille de page)")
        preview_label.setAlignment(Qt.AlignCenter)
        right.addWidget(preview_label)
        # Donner plus d'espace vertical au panneau de droite
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        right.addWidget(spacer)

        # Try to embed a PDF preview widget; if unavailable, show placeholder
        self.pdf_preview = None
        try:
            if self.pdf_path:
                self.pdf_preview = PdfPreview()
                page_num = getattr(stopover, "page_number", 1)
                if not isinstance(page_num, int) or page_num <= 0:
                    page_num = 1
                # Set the document and target page
                self.pdf_preview.setDocument(self.pdf_path, page_num)
                right.addWidget(self.pdf_preview, 1)
        except Exception:
            self.pdf_preview = None

        if not self.pdf_preview:
            self.preview_frame = QFrame()
            self.preview_frame.setFrameShape(QFrame.Box)
            self.preview_frame.setStyleSheet("QFrame { background: white; }")
            # Make placeholder expand to use available space nicely
            self.preview_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            right.addWidget(self.preview_frame, 2)

        # Give more space to the right panel (80%)
        root.addLayout(right, 8)

        # Initial size for placeholder if used
        if not self.pdf_preview:
            self._update_preview_size()

    def _mm_to_pixels(self, mm: float, dpi: float = 96.0) -> float:
        # 1 inch = 25.4 mm
        return (mm / 25.4) * dpi

    def _update_preview_size(self):
        # Compute a target size that fits within a reasonable thumbnail while preserving aspect ratio
        page_w_mm, page_h_mm = self.page_size_mm
        page_w_px = self._mm_to_pixels(page_w_mm, dpi=96.0)
        page_h_px = self._mm_to_pixels(page_h_mm, dpi=96.0)

        # Constrain to larger max preview size to make thumbnails bigger
        max_w = 1200
        max_h = 900
        scale = min(max_w / page_w_px, max_h / page_h_px)
        if scale <= 0:
            scale = 1.0
        target_w = int(page_w_px * scale)
        target_h = int(page_h_px * scale)

        # Guard if placeholder exists
        if hasattr(self, "preview_frame") and self.preview_frame is not None:
            self.preview_frame.setMinimumSize(QSize(target_w, target_h))
            self.preview_frame.setMaximumSize(QSize(target_w, target_h))
            self.preview_frame.update()


    def _emit_send_one(self):
        if callable(self._on_send_one):
            try:
                self._on_send_one(self.stopover, self.subject, self.body, self.recipients)
            except Exception as e:
                print(f"[StopoverEmailPreviewItem] send one failed: {e}")

    def _auto_persist_body(self):
        """Persist the current edited body for this stopover immediately."""
        try:
            cm = get_config_manager()
            data = cm.get_value(OVERRIDES_KEY) or {}
            code = (self.stopover.code or "").upper()
            # Always persist current subject+body. Subject comes from computed self.subject (template or override)
            data[code] = {"subject": self.subject, "body": self.body_view.toPlainText()}
            cm.set_value(OVERRIDES_KEY, data)
        except Exception as e:
            # Non bloquant
            print(f"[StopoverEmailPreviewItem] auto persist failed: {e}")



class EmailPreviewTabWidget(QWidget):
    """
    Email preview tab (PySide6).

    Requirements implemented:
      - Scrollable page with top template (subject/body) used for all stopovers
      - Below, all stopovers one after another
      - For each stopover: left subject/body preview; right a page-size-adapted preview placeholder
      - Reuses centralized template logic via ConfigManager and StopoverEmailService
    """

    def __init__(self, controller=None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.controller = controller
        self._stopovers: List[Stopover] = []
        self._pdf_path: Optional[str] = None
        self._mappings: Dict[str, List[str]] = {}
        self._mapping_service = MappingService()
        self._email_service = EmailService()
        self._stopover_email_service = StopoverEmailService()
        self._build_ui()
        # track last applied template to detect global template changes
        self._last_template = None

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(6)

        # Header actions + filters
        header = QHBoxLayout()
        self.info_label = QLabel("Prévisualisation des emails par escale", self)
        header.addWidget(self.info_label)

        header.addSpacing(16)
        header.addWidget(QLabel("Filtre escale:"))
        # Remplace le champ texte par une liste déroulante d'escales
        self.filter_combo = QComboBox()
        self.filter_combo.addItem("Toutes les escales")
        self.filter_combo.currentIndexChanged.connect(self._rebuild_items_async)
        header.addWidget(self.filter_combo)

        header.addSpacing(8)
        header.addWidget(QLabel("Afficher:"))
        self.presence_combo = QComboBox()
        self.presence_combo.addItems(["Toutes", "Avec email", "Sans email"])
        self.presence_combo.currentIndexChanged.connect(self._rebuild_items_async)
        header.addWidget(self.presence_combo)
        # Défaut: s'assurer que "Toutes" est sélectionné au lancement
        self.presence_combo.setCurrentIndex(0)

        header.addStretch(1)

        # Global action: send to all
        self.send_all_button = QPushButton("Envoyer à toutes les escales")
        self.send_all_button.clicked.connect(self._send_all_stopovers)
        header.addWidget(self.send_all_button)

        # Option: ignorer les escales sans destinataires
        self.ignore_no_recipients = QComboBox()
        self.ignore_no_recipients.addItems(["Ignorer sans destinataire", "Inclure sans destinataire"])
        header.addWidget(self.ignore_no_recipients)

        root.addLayout(header)

        # Scroll area holding everything
        self.scroll = QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        self.scroll.setWidget(scroll_content)
        self.scroll_layout = QVBoxLayout(scroll_content)
        self.scroll_layout.setContentsMargins(8, 8, 8, 8)
        self.scroll_layout.setSpacing(12)
        root.addWidget(self.scroll, 1)

        # Top: global template group (reusable by escales)
        self.template_group = QGroupBox("Modèle global (réutilisé par les escales)")
        tgl = QVBoxLayout(self.template_group)
        tgl.setContentsMargins(8, 8, 8, 8)
        tgl.setSpacing(6)

        self.template_subject = QTextEdit()
        self.template_subject.setPlaceholderText("Sujet global avec paramètres dynamiques (ex: Stopover Report - {{stopover_code}})")
        self.template_subject.setFixedHeight(40)

        self.template_body = QTextEdit()
        self.template_body.setPlaceholderText("Corps de l'email avec paramètres dynamiques")
        self.template_body.setMinimumHeight(120)

        # Load current global subject/body from config manager
        try:
            t = get_config_manager().get_templates()
            subj_value = t.get("subject") or "Stopover Report - {{stopover_code}}"
            body_value = t.get("body") or self._email_service._get_default_template()
        except Exception:
            subj_value = "Stopover Report - {{stopover_code}}"
            body_value = self._email_service._get_default_template()

        self.template_subject.setPlainText(subj_value)
        self.template_body.setPlainText(body_value)

        # Save on edit with debounce to avoid rapid rebuilds
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.setInterval(400)
        self.template_subject.textChanged.connect(lambda: self._debounce_timer.start())
        self.template_body.textChanged.connect(lambda: self._debounce_timer.start())
        self._debounce_timer.timeout.connect(self._persist_templates_from_ui)

        tgl.addWidget(QLabel("Sujet"))
        tgl.addWidget(self.template_subject)
        tgl.addWidget(QLabel("Corps"))
        tgl.addWidget(self.template_body)
        self.scroll_layout.addWidget(self.template_group)

        # Container for stopover items
        self.items_container = QWidget()
        self.items_layout = QVBoxLayout(self.items_container)
        self.items_layout.setContentsMargins(0, 0, 0, 0)
        self.items_layout.setSpacing(10)
        self.scroll_layout.addWidget(self.items_container, 1)
        self.scroll_layout.addStretch(1)

    # ---------- Public API ----------

    def set_stopovers(self, stopovers: List[Stopover]):
        self._stopovers = stopovers or []
        # reconstruire le contenu de la combo d'escales
        self._rebuild_stopover_filter_combo()
        self._rebuild_items_async()

    def set_pdf_path(self, path: Optional[str]):
        self._pdf_path = path
        # Rebuild so each item can bind to the new pdf path for preview
        self._rebuild_items_async()

    def clear(self):
        self._stopovers = []
        self._pdf_path = None
        self._clear_items()

    def refresh_recipients_from_configs(self):
        # Called when mappings change
        try:
            self._mappings = self._mapping_service.get_all_mappings()
        except Exception:
            self._mappings = {}
        self._rebuild_items_async()

    # ---------- Internal ----------

    def _persist_templates_from_ui(self):
        # Debounced-ish immediate persist of templates into unified config
        subject_template = self.template_subject.toPlainText().strip() or "Stopover Report - {{stopover_code}}"
        body_template = self.template_body.toPlainText().strip() or self._email_service._get_default_template()
        try:
            # If the template content has changed versus last applied, clear overrides
            tpl_current = (subject_template, body_template)
            if self._last_template and self._last_template != tpl_current:
                try:
                    get_config_manager().set_value(OVERRIDES_KEY, {})  # destroy manual overrides
                except Exception:
                    pass
            get_config_manager().set_templates(subject_template, body_template)
            self._last_template = tpl_current
        except Exception as e:
            # Non-fatal; keep UI responsive
            print(f"[EmailPreviewTabWidget] Failed to persist templates: {e}")

        # After persisting, refresh the stopover previews to reflect changes
        self._rebuild_items_async()

    def _clear_items(self):
        while self.items_layout.count():
            item = self.items_layout.takeAt(0)
            w = item.widget()
            if w:
                w.setParent(None)
                w.deleteLater()

    def _rebuild_items_async(self):
        QTimer.singleShot(0, self._rebuild_items)

    def _rebuild_items(self):
        self._clear_items()
        if not self._stopovers:
            info = QLabel("Aucune escale détectée.")
            info.setAlignment(Qt.AlignCenter)
            self.items_layout.addWidget(info)
            return
        # s'assurer que la combo est remplie au premier affichage
        if self.filter_combo.count() <= 1:
            self._rebuild_stopover_filter_combo()

        # Ensure we have latest mappings
        try:
            if not self._mappings:
                self._mappings = self._mapping_service.get_all_mappings()
        except Exception:
            self._mappings = {}

        # Pull current global templates once
        try:
            t = get_config_manager().get_templates()
            subject_template = t.get("subject") or "Stopover Report - {{stopover_code}}"
            body_template = t.get("body") or self._email_service._get_default_template()
        except Exception:
            subject_template = "Stopover Report - {{stopover_code}}"
            body_template = self._email_service._get_default_template()

        # Load overrides map once
        try:
            overrides = get_config_manager().get_value(OVERRIDES_KEY) or {}
        except Exception:
            overrides = {}

        # Appliquer les filtres. Si la combo "Toutes" est sélectionnée et que les mappings ne sont pas encore chargés,
        # on veut quand même afficher toutes les escales.
        filtered = self._apply_filters(self._stopovers)
        if not filtered and (not hasattr(self, "presence_combo") or self.presence_combo.currentText() == "Toutes"):
            filtered = list(self._stopovers)

        for s in filtered:
            recipients = self._mappings.get(str(s.code).upper(), [])
            # Apply overrides per stopover if present, else use template
            code_uc = (s.code or "").upper()
            ov = overrides.get(code_uc) if isinstance(overrides, dict) else None
            if ov and isinstance(ov, dict):
                subject = (ov.get("subject") or subject_template).replace("{{stopover_code}}", s.code)
                body = (ov.get("body") or body_template).replace("{{stopover_code}}", s.code)
            else:
                subject = subject_template.replace("{{stopover_code}}", s.code)
                body = body_template.replace("{{stopover_code}}", s.code)
            # Mettre à jour self.subject/self.body pour l'item créé

            # Use page size if available on stopover; fallback to A4
            page_mm = self._extract_page_size_mm(s)

            item = StopoverEmailPreviewItem(
                stopover=s,
                recipients=recipients,
                subject=subject,
                body=body,
                page_size_mm=page_mm,
                pdf_path=self._pdf_path,
                on_send_one=self._send_one_stopover,
            )
            self.items_layout.addWidget(item)

    def _extract_page_size_mm(self, stopover: Stopover) -> Tuple[float, float]:
        """
        Try to infer page size from stopover object if attributes exist.
        Fallback to A4 (210 x 297 mm).
        """
        # Common attributes that might exist: page_width_mm, page_height_mm
        try:
            w = getattr(stopover, "page_width_mm", None)
            h = getattr(stopover, "page_height_mm", None)
            if isinstance(w, (int, float)) and isinstance(h, (int, float)) and w > 0 and h > 0:
                return float(w), float(h)
        except Exception:
            pass
        return 210.0, 297.0  # A4 portrait

    # ---------- Filters and sending ----------

    def _apply_filters(self, stopovers: List[Stopover]) -> List[Stopover]:
        # sélection d'escale via combo (1er item = Toutes les escales)
        selected = self.filter_combo.currentText() if hasattr(self, "filter_combo") else "Toutes les escales"
        mode = self.presence_combo.currentText() if hasattr(self, "presence_combo") else "Toutes"
        selected = (selected or "").strip().upper()

        def match_stopover(code: str) -> bool:
            uc = (code or "").upper()
            if not selected or selected == "TOUTES LES ESCALES":
                return True
            return uc == selected

        filtered = []
        for s in stopovers:
            if not match_stopover(s.code):
                continue
            emails = self._mappings.get(str(s.code).upper(), [])
            # Pour "Toutes", ne pas filtrer par présence de mail
            if mode == "Avec email" and not emails:
                continue
            if mode == "Sans email" and emails:
                continue
            filtered.append(s)
        # Si le filtre "Toutes" est sélectionné mais que rien n'apparaît (ex: mappings pas encore chargés),
        # renvoyer l'ensemble des escales en secours pour éviter une liste vide par défaut.
        if not filtered and mode == "Toutes":
            return list(stopovers)
        return filtered

    def _send_one_stopover(self, stopover: Stopover, subject: str, body: str, recipients: List[str]):
        # Use EmailService to send a single email with a single-page PDF attachment for the selected stopover.
        try:
            if not recipients:
                QMessageBox.warning(self, "Aucun destinataire", f"Aucun email n'est configuré pour {stopover.code}.")
                return
            service = self._email_service
            attachment = self._build_attachment_for_stopover(stopover)
            ok = service.send_email(
                to_emails=recipients,
                subject=subject,
                body=body,
                attachment_path=attachment
            )
            # Feedback message requested for 'envoyer individuellement'
            if ok:
                QMessageBox.information(self, "Envoyé", f"Email envoyé pour {stopover.code}.")
            else:
                QMessageBox.critical(self, "Erreur", f"Echec d'envoi pour {stopover.code}.")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Echec d'envoi pour {stopover.code}: {e}")

    def _send_all_stopovers(self):
        if not self._stopovers:
            QMessageBox.information(self, "Info", "Aucune escale à envoyer.")
            return
        # Confirm
        reply = QMessageBox.question(self, "Confirmation", "Envoyer pour toutes les escales filtrées ?")
        if reply != QMessageBox.Yes:
            return

        # Pull current templates
        try:
            t = get_config_manager().get_templates()
            subject_template = t.get("subject") or "Stopover Report - {{stopover_code}}"
            body_template = t.get("body") or self._email_service._get_default_template()
        except Exception:
            subject_template = "Stopover Report - {{stopover_code}}"
            body_template = self._email_service._get_default_template()

        # Load overrides map for sending
        try:
            overrides = get_config_manager().get_value(OVERRIDES_KEY) or {}
        except Exception:
            overrides = {}

        sent_count = 0
        skipped_no_rec = 0
        ignore_empty = (self.ignore_no_recipients.currentText() == "Ignorer sans destinataire") if hasattr(self, "ignore_no_recipients") else True

        for s in self._apply_filters(self._stopovers):
            try:
                recipients = self._mappings.get(str(s.code).upper(), [])
                if not recipients and ignore_empty:
                    skipped_no_rec += 1
                    continue
                code_uc = (s.code or "").upper()
                ov = overrides.get(code_uc) if isinstance(overrides, dict) else None
                if ov and isinstance(ov, dict):
                    subject = (ov.get("subject") or subject_template).replace("{{stopover_code}}", s.code)
                    body = (ov.get("body") or body_template).replace("{{stopover_code}}", s.code)
                else:
                    subject = subject_template.replace("{{stopover_code}}", s.code)
                    body = body_template.replace("{{stopover_code}}", s.code)
                attachment = self._build_attachment_for_stopover(s)
                self._email_service.send_email(
                    to_emails=recipients,
                    subject=subject,
                    body=body,
                    attachment_path=attachment
                )
                sent_count += 1
            except Exception as e:
                print(f"[EmailPreviewTabWidget] send all failed for {s.code}: {e}")

        # Simple summary message after bulk send (requested)
        msg = f"Envois effectués: {sent_count}"
        if skipped_no_rec:
            msg += f"\nIgnorés (sans destinataire): {skipped_no_rec}"
        QMessageBox.information(self, "Terminé", msg)

    def _build_attachment_for_stopover(self, stopover: Stopover) -> Optional[str]:
        """
        Build a one-page PDF attachment for the given stopover from the current PDF.
        Only the page corresponding to the stopover is included, as requested.
        """
        import os
        import tempfile
        try:
            if not self._pdf_path:
                return None
            page_num = getattr(stopover, "page_number", 1)
            if not isinstance(page_num, int) or page_num <= 0:
                page_num = 1
            # Create a temporary single-page PDF using PyMuPDF
            import fitz  # PyMuPDF
            tmp_dir = tempfile.gettempdir()
            base = os.path.splitext(os.path.basename(self._pdf_path))[0]
            out_path = os.path.join(tmp_dir, f"{base}_{stopover.code}_page{page_num}.pdf")
            # Build single-page doc
            with fitz.open(self._pdf_path) as src:
                if page_num < 1 or page_num > len(src):
                    return self._pdf_path  # fallback: whole PDF
                with fitz.open() as dst:
                    dst.insert_pdf(src, from_page=page_num - 1, to_page=page_num - 1)
                    dst.save(out_path)
            return out_path if os.path.exists(out_path) else self._pdf_path
        except Exception:
            # In case of any failure, fallback to sending the whole file
            return self._pdf_path

    def _rebuild_stopover_filter_combo(self):
        # Remplit la combo avec "Toutes les escales" + codes triés
        try:
            current = self.filter_combo.currentText() if hasattr(self, "filter_combo") else "Toutes les escales"
        except Exception:
            current = "Toutes les escales"
        if not hasattr(self, "filter_combo"):
            return
        self.filter_combo.blockSignals(True)
        self.filter_combo.clear()
        self.filter_combo.addItem("Toutes les escales")
        codes = sorted({(s.code or "").upper() for s in (self._stopovers or []) if getattr(s, "code", None)})
        for c in codes:
            self.filter_combo.addItem(c)
        # restaurer sélection si possible
        if current and current in [self.filter_combo.itemText(i) for i in range(self.filter_combo.count())]:
            self.filter_combo.setCurrentText(current)
        else:
            self.filter_combo.setCurrentIndex(0)
        self.filter_combo.blockSignals(False)
