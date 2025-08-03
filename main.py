#!/usr/bin/env python3
"""Main entry point for the PDF Stopover Analyzer application."""

import logging
import os
import sys

# Configure basic logging at module import time
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Add the current directory to sys.path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import PySide6 UI entry. Avoid importing Tkinter UI here to prevent importing win32com eagerly.
from ui.pyside_main_window import run_app as run_app_qt


def main() -> None:
    """Main application entry point."""
    try:
        # Launch the PySide6 UI
        run_app_qt()
    except Exception:
        logger.exception("Error starting application")
        sys.exit(1)


if __name__ == "__main__":
    main()
