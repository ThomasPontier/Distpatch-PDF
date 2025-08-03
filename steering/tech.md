# Technology Stack

## Core Technologies
- **Python 3.7+**: Main programming language
- **PySide6**: Qt-based GUI framework for cross-platform desktop UI
- **PyMuPDF (fitz)**: PDF processing, text extraction, and page rendering
- **Pillow**: Image processing for PDF page previews
- **pywin32**: Windows Outlook integration via COM automation

## Architecture Pattern
- **MVC-style separation**: Controllers coordinate between UI and services
- **Service layer**: Dedicated services for email, configuration, mapping, and PDF processing
- **Centralized detection**: Single `DetectionEngine` class handles all stopover pattern matching
- **Configuration management**: Atomic JSON persistence with backup and crash safety

## Build System & Commands

### Development Setup
```bash
# Create virtual environment
python -m venv .venv

# Activate (Windows)
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run application
python main.py
```

### Building Executable
```bash
# Build with custom script (recommended)
python build.py

# Or build directly with PyInstaller
pyinstaller distpatch_pdf.spec
```

### Testing
```bash
# Run unit tests
python -m pytest tests/
```

## Key Dependencies
- **PyMuPDF==1.24.1**: PDF processing
- **PySide6==6.9.1**: GUI framework
- **Pillow==10.3.0**: Image handling
- **PyInstaller==6.8.0**: Executable creation
- **pywin32==306**: Windows COM integration

## Styling System
- **QSS stylesheets**: Centralized styling in `ui/style_pyside.qss`
- **Design tokens**: Consistent colors, spacing, and typography
- **Segoe UI font**: 12pt default with accessibility considerations