# PDF Stopover Analyzer

A standalone desktop application for analyzing PDF documents to find stopover pages based on specific text patterns.

## Features

- **PDF File Selection**: Browse and select PDF files from your computer
- **Stopover Detection**: Automatically finds pages containing both:
  - A stopover code in format `[XYZ]-Bilan` (where XYZ is any 3 uppercase letters)
  - The word "objectifs" (case-insensitive)
- **Visual Results**: Displays a list of all found stopover pages with their codes
- **Page Preview**: Double-click any stopover to see a preview of the corresponding PDF page
- **Cross-platform**: Works on Windows, macOS, and Linux

## Installation

### Requirements
- Python 3.7 or higher
- Required packages listed in `requirements.txt`

### Install Dependencies
```bash
pip install -r requirements.txt
```

### Run from Source
```bash
python main.py
```

### Create Executable (Windows)
1. Place your application icon in `assets/app.ico`
2. Run the build script:
```bash
python build.py
```

The executable will be created in the `dist` folder.

### Configuration Management
- On first run, the application creates a default configuration file at:
  - Windows (packaged): `%APPDATA%\Distpatch-PDF\config\app_config.json`
  - Development: `config/app_config.json`
- Configuration is automatically updated on each subsequent execution when settings change
- The configuration includes stopover mappings, email templates, and last sent timestamps

## Usage

1. **Launch the Application**: Run the executable or `python main.py`
2. **Select PDF**: Click "Select PDF" to choose a PDF file
3. **Analyze**: Click "Analyze" to process the PDF
4. **View Results**: Found stopovers will appear in the left panel
5. **Preview Pages**: Double-click any stopover to preview its page

## Project Structure

```
pdf_stopover_analyzer/
├── main.py                 # Application entry point
├── models/
│   └── stopover.py        # Data model for stopover information
├── ui/
│   └── main_window.py     # Main GUI window
├── core/
│   ├── pdf_analyzer.py    # PDF text analysis logic
│   └── pdf_renderer.py    # PDF page rendering
├── utils/
│   └── file_utils.py      # File handling utilities
├── services/
│   ├── config_manager.py  # Centralized configuration management
│   └── config_service.py  # Backward compatibility facade
├── tests/
│   └── test_pdf_analyzer.py  # Unit tests
├── config/                # Default configuration files
├── assets/                # Application assets (icons, images)
├── requirements.txt       # Python dependencies
├── distpatch_pdf.spec     # PyInstaller configuration
├── build.py              # Build script
└── README.md             # This file
```

## Technical Details

### Dependencies
- **PyMuPDF (fitz)**: PDF processing and rendering
- **Pillow**: Image handling for page previews
- **tkinter**: GUI framework (built-in with Python)
- **PyInstaller**: Executable creation
- **comtypes**: Windows COM interface for Outlook integration

### Pattern Matching
- Stopover codes must be exactly 3 uppercase letters followed by "-Bilan"
- The word "objectifs" must appear on the same page (case-insensitive)
- Examples: "CDG-Bilan", "LHR-Bilan", "JFK-Bilan"

## Testing

Run unit tests:
```bash
python -m pytest tests/
```

## Troubleshooting

### Common Issues

1. **"No module named 'fitz'"**
   - Install PyMuPDF: `pip install PyMuPDF`

2. **"Error opening PDF"**
   - Ensure the PDF file is not corrupted
   - Check file permissions

3. **"No stopover pages found"**
   - Verify the PDF contains pages with both required patterns
   - Check that stopover codes are in the correct format

## License

This project is provided as-is for educational and practical use.
