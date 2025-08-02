"""PDF rendering engine for page previews."""

import fitz  # PyMuPDF
from PIL import Image
import io

class PDFRenderer:
    """Handles PDF page rendering for preview display."""
    
    def __init__(self, pdf_path: str):
        """
        Initialize the PDF renderer.
        
        Args:
            pdf_path: Path to the PDF file
        """
        self.pdf_path = pdf_path
        self.doc = None
        self._open_document()
    
    def _open_document(self):
        """Open the PDF document."""
        try:
            self.doc = fitz.open(self.pdf_path)
        except Exception as e:
            raise ValueError(f"Error opening PDF: {str(e)}")
    
    def get_page_image(self, page_number: int, max_width: int = 1200, max_height: int = 800) -> Image.Image:
        """
        Get a PIL Image of the specified page.
        
        Args:
            page_number: 1-based page number
            max_width: Maximum width for the image
            max_height: Maximum height for the image
            
        Returns:
            PIL Image of the page
            
        Raises:
            ValueError: If page number is invalid
        """
        if not self.doc:
            raise ValueError("PDF document not opened")
        
        if page_number < 1 or page_number > len(self.doc):
            raise ValueError(f"Invalid page number: {page_number}")
        
        page = self.doc[page_number - 1]  # Convert to 0-based indexing
        
        # Get page dimensions
        rect = page.rect
        zoom = min(max_width / rect.width, max_height / rect.height)
        
        # Create transformation matrix for zoom
        mat = fitz.Matrix(zoom, zoom)
        
        # Render page to pixmap
        pix = page.get_pixmap(matrix=mat)
        
        # Convert to PIL Image
        img_data = pix.tobytes("ppm")
        return Image.open(io.BytesIO(img_data))
    
    def get_page_count(self) -> int:
        """Get the total number of pages in the PDF."""
        return len(self.doc) if self.doc else 0
    
    def close(self):
        """Close the PDF document."""
        if self.doc:
            self.doc.close()
            self.doc = None
    
    def __enter__(self):
        """Context manager entry."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close()
