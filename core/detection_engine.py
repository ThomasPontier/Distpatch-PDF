"""Centralized detection engine for stopover pages."""

import re
from typing import Any, Dict, List, Optional

import fitz  # PyMuPDF

from models.stopover import Stopover


class DetectionEngine:
    """Single source of truth for stopover detection logic."""

    # Consolidated patterns for stopover detection
    STOPOVER_PATTERNS = [
        re.compile(r"\b([A-Z]{3})\s*-\s*Bilan\b", re.IGNORECASE),
        re.compile(r"\[([A-Z]{3})\]-Bilan\b", re.IGNORECASE),
        re.compile(r"([A-Z]{3})-Bilan\b", re.IGNORECASE),
    ]

    OBJECTIVES_PATTERN = re.compile(r"objectifs", re.IGNORECASE)

    @classmethod
    def extract_stopover_code(cls, text: str) -> Optional[str]:
        """Extract 3-letter airport code from text."""
        if not text:
            return None

        for pattern in cls.STOPOVER_PATTERNS:
            match = pattern.search(text)
            if match:
                code = match.group(1).upper()
                if len(code) == 3 and code.isalpha():
                    return code
        return None

    @classmethod
    def contains_objectives(cls, text: str) -> bool:
        """Check if text contains objectives keyword."""
        if not text:
            return False
        return bool(cls.OBJECTIVES_PATTERN.search(text))

    @classmethod
    def is_stopover_page(cls, text: str) -> bool:
        """Determine if page qualifies as stopover page."""
        return cls.extract_stopover_code(text) is not None and cls.contains_objectives(text)

    def analyze_pdf(self, pdf_path: str) -> List[Stopover]:
        """Analyze PDF and return stopover pages."""
        stopovers: List[Stopover] = []

        try:
            with fitz.open(pdf_path) as doc:
                for page_num in range(1, len(doc) + 1):
                    text = doc[page_num - 1].get_text()

                    code = self.extract_stopover_code(text)
                    if code and self.contains_objectives(text):
                        stopovers.append(Stopover(code=code, page_number=page_num))

        except Exception as e:
            # Preserve error propagation behavior
            raise ValueError(f"Error analyzing PDF: {str(e)}")

        return stopovers

    @classmethod
    def test_detection(cls, text: str) -> Dict[str, Any]:
        """Test detection accuracy with sample text."""
        code = cls.extract_stopover_code(text)
        has_objectives = cls.contains_objectives(text)

        return {
            "stopover_found": code,
            "has_objectives": has_objectives,
            "is_valid_stopover_page": bool(code and has_objectives),
        }
