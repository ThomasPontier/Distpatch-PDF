# Project Structure

## Directory Organization

```
pdf_stopover_analyzer/
├── main.py                    # Application entry point
├── controllers/               # Application controllers (MVC pattern)
│   └── app_controller.py     # Main controller coordinating UI and services
├── core/                     # Core business logic
│   ├── detection_engine.py   # Centralized stopover detection patterns
│   ├── pdf_processor.py      # PDF analysis and text extraction
│   └── pdf_renderer.py       # PDF page rendering for previews
├── models/                   # Data models and structures
│   ├── stopover.py          # Stopover data class
│   └── template.py          # Email template and account models
├── services/                 # Service layer for external integrations
│   ├── accounts_service.py   # Account management
│   ├── config_manager.py     # Centralized configuration with atomic writes
│   ├── config_service.py     # Legacy config facade
│   ├── email_service.py      # Email sending functionality
│   ├── mapping_service.py    # Stopover mapping management
│   ├── pdf_attachment_service.py # PDF attachment handling
│   └── stopover_email_service.py # Stopover-specific email logic
├── ui/                       # User interface components
│   ├── components/           # Reusable UI components
│   ├── pyside_main_window.py # Main application window
│   ├── pyside_*_tab.py      # Tab-specific UI components
│   ├── pyside_*_dialog.py   # Dialog windows
│   ├── style_pyside.qss     # Centralized QSS stylesheet
│   └── pyside_tokens.py     # Design tokens and theming
├── utils/                    # Utility functions
│   └── file_utils.py        # File handling and validation
├── config/                   # Default configuration files
├── assets/                   # Application assets (icons, images)
├── build.py                  # Build script for executable creation
├── distpatch_pdf.spec       # PyInstaller configuration
└── requirements.txt         # Python dependencies
```

## Architectural Patterns

### Naming Conventions
- **Files**: Snake_case for Python files (`app_controller.py`)
- **Classes**: PascalCase (`AppController`, `DetectionEngine`)
- **Methods/Variables**: Snake_case (`analyze_pdf`, `current_pdf_path`)
- **Constants**: UPPER_SNAKE_CASE (`STOPOVER_PATTERNS`)

### Code Organization
- **Single responsibility**: Each service handles one domain (email, config, mapping)
- **Centralized detection**: All pattern matching logic in `DetectionEngine`
- **UI separation**: PySide6 components prefixed with `pyside_`
- **Configuration**: Atomic JSON writes with `.bak` backup files
- **Resource bundling**: Assets and stylesheets bundled via PyInstaller spec

### Import Patterns
- Relative imports within packages
- Absolute imports from project root
- Service dependencies injected through controllers
- UI components import from `ui.pyside_tokens` for consistent theming