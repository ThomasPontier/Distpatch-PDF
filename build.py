#!/usr/bin/env python3
"""Build script for creating the Distpatch-PDF executable."""

import subprocess
import sys
import os
import shutil

def remove_path(path: str):
    """Safely remove a file or directory if it exists."""
    if os.path.isdir(path):
        print(f"Removing directory: {path}")
        shutil.rmtree(path, ignore_errors=True)
    elif os.path.isfile(path):
        print(f"Removing file: {path}")
        try:
            os.remove(path)
        except Exception:
            pass

def main():
    """Build the application using PyInstaller."""
    print("Building Distpatch-PDF...")

    # Clean previous outputs to avoid stale artifacts
    remove_path("build")
    remove_path("dist")

    # Check if PyInstaller is installed
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # Run PyInstaller with our spec file
    try:
        subprocess.check_call([
            sys.executable, "-m", "PyInstaller",
            "--clean",
            "distpatch_pdf.spec"
        ])
        print("Build completed successfully!")
        print("Executable location: dist/Distpatch-PDF.exe")
    except subprocess.CalledProcessError as e:
        print(f"Build failed with error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()