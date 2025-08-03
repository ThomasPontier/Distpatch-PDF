"""Essential file utilities."""

import os
import sys
from pathlib import Path


def validate_pdf_file(file_path: str) -> bool:
    """Validate if file exists and is a PDF."""
    if not file_path:
        return False

    path = Path(file_path)
    return path.exists() and path.is_file() and path.suffix.lower() == ".pdf"


def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS  # type: ignore[attr-defined]  # used when packaged
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
