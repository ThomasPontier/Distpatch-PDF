"""PySide6 Stopover tab widget preserving behavior from Tkinter StopoverTabComponent."""

from typing import List, Optional, Callable
from PySide6.QtCore import Qt, QTimer, QSize, QEvent
from PySide6.QtGui import QPixmap, QImage
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QListWidget, QListWidgetItem,
    QGroupBox, QLabel, QMenu, QMessageBox, QSplitter, QSizePolicy, QFrame
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

        self._build_ui()

    def _build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        splitter = QSplitter(Qt.Horizontal, self)

        # Left: Stopover list
        left_group = QGroupBox("Stopover Pages", splitter)
        left_v = QVBoxLayout(left_group)
        left_v.setContentsMargins(8, 8, 8, 8)
        left_v.setSpacing(6)

        self.stopover_list = QListWidget(left_group)
        self.stopover_list.itemDoubleClicked.connect(self._on_double_click)
        self.stopover_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.stopover_list.customContextMenuRequested.connect(self._on_right_click)
        left_v.addWidget(self.stopover_list)

        # Right: Preview
        right_group = QGroupBox("Page Preview", splitter)
        right_v = QVBoxLayout(right_group)
        right_v.setContentsMargins(8, 8, 8, 8)
        right_v.setSpacing(6)

        self.preview_label = QLabel("Select a stopover to preview its page", right_group)
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
        self.preview_label.setText("Select a stopover to preview its page")
        self.close_pdf_renderer()

    # -------- Internal behavior --------

    def _update_stopover_list(self):
        self.stopover_list.clear()
        for s in self.stopovers:
            item = QListWidgetItem(s.code)
            self.stopover_list.addItem(item)

    def _on_double_click(self, item: QListWidgetItem):
        code = item.text()
        selected = None
        for s in self.stopovers:
            if s.code == code:
                selected = s
                break
        if selected and self.on_stopover_select:
            self.on_stopover_select(selected)

    def _on_right_click(self, pos):
        item = self.stopover_list.itemAt(pos)
        if not item:
            return
        self.stopover_list.setCurrentItem(item)
        menu = QMenu(self)
        action = menu.addAction("Configure Email Settings")
        action.triggered.connect(self._configure_email_settings)
        menu.exec(self.stopover_list.mapToGlobal(pos))

    def _configure_email_settings(self):
        if not self.controller:
            QMessageBox.critical(self, "Error", "Controller not available")
            return
        item = self.stopover_list.currentItem()
        if not item:
            QMessageBox.information(self, "Info", "Please select a stopover first")
            return
        code = item.text()
        selected = None
        for s in self.stopovers:
            if s.code == code:
                selected = s
                break
        if not selected:
            QMessageBox.critical(self, "Error", "Selected stopover not found")
            return
        try:
            # Dialog remains Tkinter-based in original; parity maintained by launching through controller if needed.
            from ui.stopover_email_dialog import StopoverEmailDialog
            # The Tk dialog expects a Tk parent; since we preserve behavior, call its show via controller service.
            dlg = StopoverEmailDialog(None, selected.code, self.controller.stopover_email_service)
            dlg.show()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open email configuration dialog: {str(e)}")

    def load_page_preview(self, stopover: Stopover, progress_callback: Callable[[str], None] = None):
        if not self.current_pdf_path:
            return
        try:
            if progress_callback:
                progress_callback("Loading page preview...")
            if not self.pdf_renderer or self.pdf_renderer.pdf_path != self.current_pdf_path:
                self.close_pdf_renderer()
                self.pdf_renderer = PDFRenderer(self.current_pdf_path)

            img = self.pdf_renderer.get_page_image(stopover.page_number, max_width=1600, max_height=1600)
            self._last_rendered_image = img
            # Update preview asynchronously to ensure the widget has sizes
            QTimer.singleShot(0, self._fit_and_update_preview)
        except Exception as e:
            err = f"Error loading page preview: {str(e)}"
            if progress_callback:
                progress_callback(err)
            self.preview_label.setText("Preview unavailable")

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
            self.preview_label.setText("Preview unavailable")

    def close_pdf_renderer(self):
        if self.pdf_renderer:
            try:
                self.pdf_renderer.close()
            except Exception:
                pass
            self.pdf_renderer = None
