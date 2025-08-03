"""PySide6 Stopover tab widget preserving behavior from Tkinter StopoverTabComponent."""

from typing import List, Optional, Callable
from PySide6.QtCore import Qt, QTimer, QSize, QEvent
from PySide6.QtGui import QPixmap, QImage
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QListWidget, QListWidgetItem,
    QGroupBox, QLabel, QMessageBox, QSplitter, QSizePolicy
)
from PIL import Image
from core.pdf_renderer import PDFRenderer
from models.stopover import Stopover
from utils.file_utils import validate_pdf_file


class StopoverTabWidget(QWidget):
    """
    UI component for the stopover pages tab using PySide6.

    Public methods kept compatible with original component:
      - set_pdf_path(path) -> bool
      - set_stopovers(stopovers: List[Stopover])
      - load_page_preview(stopover: Stopover, progress_callback: Optional[Callable[[str], None]])
      - clear()
    """

    def __init__(self, on_stopover_select: Callable[[Stopover], None] = None, controller=None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.on_stopover_select = on_stopover_select
        self.controller = controller
        self.stopovers: List[Stopover] = []
        self.pdf_renderer: Optional[PDFRenderer] = None
        self.current_pdf_path: Optional[str] = None

        self._last_rendered_image: Optional[Image.Image] = None

        # Track a last used external progress callback so helper setters can use it reliably
        self._progress_callback: Optional[Callable[[str], None]] = None

        self._build_ui()

    def _build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        splitter = QSplitter(Qt.Horizontal, self)

        # Left: Stopover list
        left_group = QGroupBox("", splitter)
        left_v = QVBoxLayout(left_group)
        left_v.setContentsMargins(8, 8, 8, 8)
        left_v.setSpacing(6)

        self.stopover_list = QListWidget(left_group)
        # Change 1: preview on single selection instead of double click
        self.stopover_list.currentItemChanged.connect(self._on_selection_changed)
        # Change 2: open email settings on double click (replaces right-click flow)
        self.stopover_list.itemDoubleClicked.connect(self._open_email_settings_for_item)
        # Remove obsolete context menu
        left_v.addWidget(self.stopover_list)

        # Right: Preview
        right_group = QGroupBox("", splitter)
        right_v = QVBoxLayout(right_group)
        right_v.setContentsMargins(8, 8, 8, 8)
        right_v.setSpacing(6)

        self.preview_label = QLabel("Sélectionner une escale pour afficher sa page", right_group)
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview_label.setObjectName("PreviewLabel")
        right_v.addWidget(self.preview_label)

        splitter.addWidget(left_group)
        splitter.addWidget(right_group)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)

        layout.addWidget(splitter)

        # Handle resize to refit preview
        right_group.installEventFilter(self)
        self._right_group = right_group

    # ===== Status helpers (ensure final updates in all paths) =====
    def _emit_status(self, message: str):
        try:
            if self._progress_callback:
                self._progress_callback(message)
        except Exception:
            pass

    def _set_status_loading(self):
        self._emit_status("Chargement de l’aperçu de la page…")

    def _set_status_error(self):
        self._emit_status("Aperçu indisponible")

    def _set_status_no_selection(self):
        # Prefer a quiet UI; explicit message when nothing is selected
        self._emit_status("Aucun aperçu")

    def _set_status_idle(self):
        # Use empty string to reduce noise as requested
        self._emit_status("")

    # -------- Public API (parity) --------

    def set_pdf_path(self, pdf_path: str) -> bool:
        if validate_pdf_file(pdf_path):
            self.current_pdf_path = pdf_path
            self.close_pdf_renderer()
            return True
        return False

    def set_stopovers(self, stopovers: List[Stopover]):
        self.stopovers = stopovers
        self._update_stopover_list()

    def clear(self):
        self.stopovers = []
        self.current_pdf_path = None
        self.stopover_list.clear()
        self.preview_label.clear()
        self.preview_label.setText("Sélectionner une escale pour afficher sa page")
        self.close_pdf_renderer()
        # When cleared, reflect idle/no selection status
        self._set_status_idle()

    # -------- Internal behavior --------

    def _update_stopover_list(self):
        self.stopover_list.clear()
        for s in self.stopovers:
            item = QListWidgetItem(s.code)
            self.stopover_list.addItem(item)

    # New: single-selection triggers preview callback
    def _on_selection_changed(self, current: Optional[QListWidgetItem], previous: Optional[QListWidgetItem]):
        try:
            if not current:
                # Selection cleared: show no preview message and idle status
                self.preview_label.setText("Sélectionner une escale pour afficher sa page")
                self.preview_label.setPixmap(QPixmap())
                self._set_status_no_selection()
                return
            code = current.text()
            selected = None
            for s in self.stopovers:
                if s.code == code:
                    selected = s
                    break
            if selected:
                # If external handler provided, keep compatibility
                if self.on_stopover_select:
                    self.on_stopover_select(selected)
                else:
                    # Fallback: render preview directly using local API if available
                    self.load_page_preview(selected)
        except Exception:
            # On unexpected errors during selection change, surface an error state
            self.preview_label.setText("Aperçu indisponible")
            self._set_status_error()

    # New: double click opens email settings dialog directly
    def _open_email_settings_for_item(self, item: QListWidgetItem):
        if not item:
            return
        try:
            code = item.text()
            selected = None
            for s in self.stopovers:
                if s.code == code:
                    selected = s
                    break
            if not selected:
                QMessageBox.critical(self, "Erreur", "Escale sélectionnée introuvable")
                return
            if not self.controller or not hasattr(self.controller, "stopover_email_service"):
                QMessageBox.critical(self, "Erreur", "Contrôleur indisponible")
                return
            from ui.pyside_stopover_email_dialog import StopoverEmailSettingsDialog
            dlg = StopoverEmailSettingsDialog(self, selected.code, self.controller.stopover_email_service)
            dlg.exec()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Échec d’ouverture de la boîte de dialogue de configuration email : {str(e)}")

    def load_page_preview(self, stopover: Stopover, progress_callback: Callable[[str], None] = None):
        # Persist the latest callback so helper setters can always emit status updates
        if progress_callback:
            self._progress_callback = progress_callback

        if not self.current_pdf_path:
            # No PDF selected: do not show loading; reflect no selection/idle
            self.preview_label.setText("Aucun aperçu")
            self._set_status_no_selection()
            return
        try:
            # Set loading immediately
            self._set_status_loading()

            if not self.pdf_renderer or self.pdf_renderer.pdf_path != self.current_pdf_path:
                self.close_pdf_renderer()
                self.pdf_renderer = PDFRenderer(self.current_pdf_path)

            img = self.pdf_renderer.get_page_image(stopover.page_number, max_width=1600, max_height=1600)
            self._last_rendered_image = img
            # Update preview asynchronously to ensure the widget has sizes
            QTimer.singleShot(0, self._fit_and_update_preview)
            # Ensure final idle status after successful scheduling of UI update
            QTimer.singleShot(0, self._set_status_idle)
        except Exception as e:
            # Error: update both preview text and status error
            self.preview_label.setText("Aperçu indisponible")
            self._set_status_error()
            # Keep error message visible; do not immediately reset to idle

    def eventFilter(self, watched, event):
        # Refit on container resize
        try:
            if watched is getattr(self, "_right_group", None) and event.type() == QEvent.Resize:
                QTimer.singleShot(50, self._fit_and_update_preview)
        except Exception:
            pass
        return super().eventFilter(watched, event)

    def _fit_and_update_preview(self):
        if self._last_rendered_image is None:
            return
        try:
            avail_w = max(1, self.preview_label.width() - 16)
            avail_h = max(1, self.preview_label.height() - 16)
            img = self._last_rendered_image
            iw, ih = img.size
            if iw <= 0 or ih <= 0:
                return
            scale = min(avail_w / iw, avail_h / ih)
            target_w = max(1, int(iw * scale))
            target_h = max(1, int(ih * scale))
            resized = img.resize((target_w, target_h), Image.LANCZOS)

            # Convert PIL Image to QPixmap
            qimg = QImage(resized.tobytes(), resized.width, resized.height, resized.width * 3, QImage.Format_RGB888) if resized.mode == "RGB" else None
            if qimg is None:
                resized = resized.convert("RGBA")
                qimg = QImage(resized.tobytes(), resized.width, resized.height, resized.width * 4, QImage.Format_RGBA8888)
            pix = QPixmap.fromImage(qimg)
            self.preview_label.setPixmap(pix)
            self.preview_label.setText("")
        except Exception:
            self.preview_label.setText("Aperçu indisponible")
            # Reflect error on status as well
            self._set_status_error()

    def close_pdf_renderer(self):
        if self.pdf_renderer:
            try:
                self.pdf_renderer.close()
            except Exception:
                pass
            self.pdf_renderer = None
