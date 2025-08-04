#!/usr/bin/env python3
"""Build script for creating the Dispatch-SATISFACTION executable."""

import subprocess
import sys
import os
import shutil
import logging
from typing import NoReturn

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def remove_path(path: str) -> None:
    """Safely remove a file or directory if it exists, logging actions and errors."""
    if os.path.isdir(path):
        logger.info("Removing directory: %s", path)
        try:
            shutil.rmtree(path, ignore_errors=False)
        except Exception:
            logger.exception("Failed to remove directory: %s", path)
    elif os.path.isfile(path):
        logger.info("Removing file: %s", path)
        try:
            os.remove(path)
        except Exception:
            logger.exception("Failed to remove file: %s", path)

def main() -> None:
    """Build the application using PyInstaller."""
    logger.info("Building Dispatch-SATISFACTION...")

    # Avoid manually deleting build/ and dist/ to prevent Windows lock errors.
    # Delegate cleaning to PyInstaller with --clean and explicit paths.
    # Check if PyInstaller is installed (do not auto-install)
    try:
        import PyInstaller  # type: ignore  # noqa: F401
    except ImportError:
        logger.error(
            "PyInstaller is not installed. Install it first (e.g., via requirements-dev.txt) and re-run."
        )
        sys.exit(1)

    # Run PyInstaller with our spec file.
    # Use explicit work/dist paths so PyInstaller manages its own cleanup safely.
    try:
        subprocess.check_call([
            sys.executable, "-m", "PyInstaller",
            "--clean",
            "--workpath", "build",
            "--distpath", "dist",
            "distpatch_pdf.spec"
        ])
        logger.info("Build completed successfully!")
        logger.info("Executable location: dist/Dispatch-SATISFACTION.exe")
    except subprocess.CalledProcessError as e:
        logger.exception("Build failed with error: %s", e)
        sys.exit(1)

if __name__ == "__main__":
    main()
