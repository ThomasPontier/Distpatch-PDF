"""Data model for stopover information."""

from dataclasses import dataclass


@dataclass
class Stopover:
    """Represents a stopover page found in a PDF document."""
    
    code: str
    page_number: int
    
    def __str__(self):
        """String representation for display in GUI."""
        return f"{self.code} (Page {self.page_number})"
    
    def __repr__(self):
        """Detailed string representation for debugging."""
        return f"Stopover(code='{self.code}', page_number={self.page_number})"
