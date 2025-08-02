from __future__ import annotations

from typing import Optional
from PySide6.QtCore import Qt, QSize, QPoint, Signal, QObject, QThread, QRect
from PySide6.QtGui import QPixmap, QImage, QAction, QWheelEvent, QMouseEvent, QPalette, QDesktopServices
from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QHBoxLayout, QPushButton, QFileDialog, QFrame, QToolBar, QStyle, QSizePolicy, QMessageBox, QScrollArea

from PIL import Image
import os

# We reuse core.pdf_renderer to render PIL Image from a PDF page.
# This widget converts PIL Image to QPixmap for display, provides zoom and scroll,
# and falls back to a placeholder if rendering fails or engine is unavailable.


def pil_to_qimage(img: Image.Image) -> QImage:
    """Convert PIL Image (RGB or RGBA or L) to QImage."""
    if img.mode == "RGB":
        r, g, b = img.split()
        img = Image.merge("RGB", (r, g, b))
        data = img.tobytes("raw", "RGB")
        qimg = QImage(data, img.width, img.height, QImage.Format.Format_RGB888)
        return qimg
    elif img.mode == "RGBA":
        data = img.tobytes("raw", "RGBA")
        qimg = QImage(data, img.width, img.height, QImage.Format.Format_RGBA8888)
        return qimg
    elif img.mode == "L":
        data = img.tobytes("raw", "L")
        qimg = QImage(data, img.width, img.height, QImage.Format.Format_Grayscale8)
        return qimg
    else:
        # Convert to RGB as a fallback
        rgb = img.convert("RGB")
        data = rgb.tobytes("raw", "RGB")
        qimg = QImage(data, rgb.width, rgb.height, QImage.Format.Format_RGB888)
        return qimg


class _RenderWorker(QObject):
    rendered = Signal(QPixmap)
    error = Signal(str)

    def __init__(self, pdf_path: str, page_number: int, max_w: int, max_h: int):
        super().__init__()
        self.pdf_path = pdf_path
        self.page_number = page_number
        self.max_w = max_w
        self.max_h = max_h
        # render hint: request higher DPI for better quality (used by renderer if applicable)
        self.target_dpi = 192

    def run(self):
        try:
            from core.pdf_renderer import PDFRenderer
            renderer = PDFRenderer(self.pdf_path)
            # Render at higher internal resolution for sharper preview, then we will downscale with SmoothTransformation
            img = renderer.get_page_image(self.page_number, max_width=self.max_w, max_height=self.max_h)
            renderer.close()
            qimg = pil_to_qimage(img)
            # Convert to QPixmap with no further scaling here (keep native high-res)
            self.rendered.emit(QPixmap.fromImage(qimg))
        except Exception as e:
            self.error.emit(str(e))


class PdfPreview(QWidget):
    """
    PDF page preview with zoom and scroll, backed by core.pdf_renderer.

    Signals:
      - openExternallyRequested(pdf_path: str, page_number: int)
    """
    openExternallyRequested = Signal(str, int)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._pdf_path: Optional[str] = None
        self._page_number: int = 1
        self._scale: float = 1.0
        self._base_pixmap: Optional[QPixmap] = None
        self._fit_to_view: bool = True  # always show full page initially (fit-to-container)

        # Thread members
        self.thread: Optional[QThread] = None
        self.worker: Optional[_RenderWorker] = None

        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(4)

        # Toolbar
        self.toolbar = QToolBar()
        self.zoomInAct = QAction("Zoom +", self)
        self.zoomOutAct = QAction("Zoom -", self)
        self.resetZoomAct = QAction("Reset", self)
        self.openExternAct = QAction("Open externally", self)
        # Keep icon/text as-is; behavior will open only the current page externally

        self.zoomInAct.triggered.connect(lambda: self._apply_zoom(1.1))
        self.zoomOutAct.triggered.connect(lambda: self._apply_zoom(1/1.1))
        # Reset returns to fit-to-view (full page)
        self.resetZoomAct.triggered.connect(self._reset_fit_to_view)
        self.openExternAct.triggered.connect(self._open_externally)

        self.toolbar.addAction(self.zoomInAct)
        self.toolbar.addAction(self.zoomOutAct)
        self.toolbar.addAction(self.resetZoomAct)
        self.toolbar.addSeparator()
        self.toolbar.addAction(self.openExternAct)
        root.addWidget(self.toolbar)

        # Scroll area with image label
        self.scrollArea = QScrollArea()
        self.scrollArea.setWidgetResizable(True)
        # Keep default styling; no special scaling tweaks required
        # self.scrollArea.setStyleSheet("QScrollArea { background: transparent; }")
        self.imageLabel = QLabel("No preview")
        self.imageLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.imageLabel.setBackgroundRole(QPalette.Base if hasattr(self.imageLabel, "setBackgroundRole") else None)  # safe guard
        # Keep default size policy (Expanding) to avoid unexpected scaling quirks
        # self.imageLabel.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.imageLabel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        container = QWidget()
        c_layout = QVBoxLayout(container)
        c_layout.setContentsMargins(0, 0, 0, 0)
        c_layout.addWidget(self.imageLabel)
        self.scrollArea.setWidget(container)
        root.addWidget(self.scrollArea)

        # Placeholder label for errors
        self.placeholder = QLabel("PDF preview unavailable")
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder.setVisible(False)
        root.addWidget(self.placeholder)

    def setDocument(self, pdf_path: str, page_number: int):
        """Set document path and target page."""
        self._pdf_path = pdf_path
        self._page_number = max(1, page_number)
        self._scale = 1.0
        self._base_pixmap = None
        self._fit_to_view = True  # every new doc starts fitted (full page visible)
        self._render_async()

    def _render_async(self):
        if not self._pdf_path or not os.path.exists(self._pdf_path):
            self._show_placeholder(f"File not found: {self._pdf_path or ''}")
            return

        # Ensure any previous thread is properly shut down
        self._cleanup_thread()

        # Start worker thread
        self.thread = QThread(self)
        # Restore previous stable render size to avoid regressions
        self.worker = _RenderWorker(self._pdf_path, self._page_number, 1200, 800)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.rendered.connect(self._on_rendered)
        self.worker.error.connect(self._on_render_error)
        # Clean up and ownership
        self.worker.rendered.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.thread.finished.connect(self._on_thread_finished)
        self.thread.start()

    def _on_rendered(self, pix: QPixmap):
        self.placeholder.setVisible(False)
        self._base_pixmap = pix
        # On first render or reset, show entire page fitted to scroll area
        self._fit_to_view = True
        self._apply_scaled_pixmap()

    def _on_render_error(self, msg: str):
        self._show_placeholder(f"Error rendering PDF: {msg}")

    def _apply_zoom(self, factor: float):
        # switching to manual zoom disables fit-to-view
        self._fit_to_view = False
        self._scale = max(0.1, min(5.0, self._scale * factor))
        self._apply_scaled_pixmap()

    def _set_zoom(self, scale: float):
        self._fit_to_view = False
        self._scale = max(0.1, min(5.0, scale))
        self._apply_scaled_pixmap()

    def _apply_scaled_pixmap(self):
        if not self._base_pixmap:
            return

        if self._fit_to_view:
            # Fit the entire page into the visible scroll area viewport
            viewport = self.scrollArea.viewport().size()
            if viewport.width() <= 0 or viewport.height() <= 0:
                scaled = self._base_pixmap
            else:
                scaled = self._base_pixmap.scaled(
                    viewport, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
                )
        else:
            size = self._base_pixmap.size()
            new_size = QSize(int(size.width() * self._scale), int(size.height() * self._scale))
            scaled = self._base_pixmap.scaled(
                new_size, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            )

        self.imageLabel.setPixmap(scaled)
        self.imageLabel.resize(scaled.size())

    def _show_placeholder(self, text: str):
        self.placeholder.setText(text)
        self.placeholder.setVisible(True)
        self.imageLabel.clear()
        self._fit_to_view = True

    def _open_externally(self):
        """
        Open only the currently previewed page externally.
        We export the single page as a temporary PNG (high quality) and open that,
        so the user only sees the relevant stopover page. Interface unchanged.
        """
        if self._base_pixmap is None:
            QMessageBox.information(self, "Open externally", "Aucun aperçu de page n'est disponible.")
            return

        try:
            # Save current scaled preview (or base pixmap) as a temporary image file
            from tempfile import NamedTemporaryFile
            tmp = NamedTemporaryFile(prefix="dispatch_page_", suffix=".png", delete=False)
            tmp_path = tmp.name
            tmp.close()

            # Use base pixmap for best quality
            pix = self._base_pixmap
            pix.save(tmp_path, "PNG")

            # Try to open image with OS default viewer
            opened = QDesktopServices.openUrl(f"file:///{os.path.abspath(tmp_path).replace('\\', '/')}")
            if not opened:
                try:
                    os.startfile(tmp_path)  # type: ignore[attr-defined]
                except Exception as e2:
                    QMessageBox.warning(self, "Open externally", f"Impossible d'ouvrir l'aperçu: {e2}")

            # Also emit signal for listeners with path and page number for any custom handling
            try:
                self.openExternallyRequested.emit(self._pdf_path or "", self._page_number)
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "Open externally", f"Erreur lors de l'ouverture externe: {e}")

    def _on_thread_finished(self):
        # Called when thread finishes; delete worker and thread safely
        if self.worker:
            self.worker.deleteLater()
            self.worker = None
        if self.thread:
            self.thread.deleteLater()
            self.thread = None

    def _cleanup_thread(self):
        # Ensure no running thread remains (prevents 'QThread: Destroyed while thread is still running')
        if self.thread:
            try:
                self.thread.quit()
                self.thread.wait(1500)
            except Exception:
                pass
            self._on_thread_finished()

    def closeEvent(self, event):
        # Ensure threads are cleaned up on widget close
        self._cleanup_thread()
        super().closeEvent(event)

    def _reset_fit_to_view(self):
        # Reset to full-page preview (no zoom cropping)
        self._fit_to_view = True
        self._scale = 1.0
        self._apply_scaled_pixmap()
