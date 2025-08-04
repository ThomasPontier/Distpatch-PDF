#!/usr/bin/env python3
"""Build script for creating the Dispatch-SATISFACTION executable."""

import subprocess
import sys
import os
import shutil
import logging
import stat
import time
from typing import NoReturn

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def _on_rm_error(func, path, exc_info):
    """
    Error handler for shutil.rmtree.
    If the error is due to a read-only file, make it writable and retry.
    """
    try:
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
    except Exception:
        pass
    try:
        func(path)
    except Exception:
        logger.debug("Retry failed removing: %s", path, exc_info=True)


def _rmtree_robust(path: str, attempts: int = 5, delay: float = 0.4) -> None:
    """
    Robustly remove a directory tree on Windows by clearing attributes and retrying.
    """
    for attempt in range(1, attempts + 1):
        try:
            shutil.rmtree(path, onerror=_on_rm_error)
            return
        except FileNotFoundError:
            return
        except PermissionError as e:
            logger.warning("PermissionError removing %s (attempt %d/%d): %s", path, attempt, attempts, e)
        except OSError as e:
            logger.warning("OSError removing %s (attempt %d/%d): %s", path, attempt, attempts, e)
        time.sleep(delay)
    shutil.rmtree(path, onerror=_on_rm_error)


def remove_path(path: str) -> None:
    """Safely remove a file or directory if it exists, logging actions and errors with robust handling on Windows."""
    if os.path.isdir(path):
        logger.info("Removing directory: %s", path)
        try:
            _rmtree_robust(path)
        except Exception:
            logger.exception("Failed to remove directory: %s", path)
    elif os.path.isfile(path):
        logger.info("Removing file: %s", path)
        try:
            try:
                os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
            except Exception:
                pass
            os.remove(path)
        except Exception:
            logger.exception("Failed to remove file: %s", path)

def main() -> None:
    """Build the application using PyInstaller."""
    logger.info("Building Dispatch-SATISFACTION...")
    # Pre-clean build/ and dist/ to avoid leftover locks from previous runs
    remove_path("build")
    remove_path("dist")

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
