#!/usr/bin/env python3
"""Main entry point for the PDF Stopover Analyzer application."""

import sys
import os

# Add the current directory to cls
#  path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import PySide6 UI entry. Avoid importing Tkinter UI here to prevent importing win32com eagerly.
from ui.pyside_main_window import run_app as run_app_qt


def main():
    """Main application entry point."""
    try:
        # Launch the new PySide6 UI (legacy Tkinter UI kept for reference)
        run_app_qt()
    except Exception as e:
        print(f"Error starting application: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
