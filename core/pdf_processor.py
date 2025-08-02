"""Simplified PDF processor using centralized detection engine."""

from typing import List
from core.detection_engine import DetectionEngine


class PDFProcessor:
    """Simplified PDF processor that delegates to DetectionEngine."""
    
    def __init__(self):
        """Initialize with centralized detection engine."""
        self.detector = DetectionEngine()
    
    def analyze_pdf(self, pdf_path: str) -> List:
        """Analyze PDF for stopover pages."""
        return self.detector.analyze_pdf(pdf_path)
    
    def test_detection(self, test_text: str) -> dict:
        """Test detection accuracy."""
        return self.detector.test_detection(test_text)
    
    def get_page_text(self, pdf_path: str, page_number: int) -> str:
        """Extract text from specific page."""
        try:
            import fitz
            with fitz.open(pdf_path) as doc:
                if 1 <= page_number <= len(doc):
                    return doc[page_number - 1].get_text() or ""
                return ""
        except Exception:
            return ""
