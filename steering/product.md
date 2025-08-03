# Product Overview

PDF Stopover Analyzer is a standalone desktop application for analyzing PDF documents to identify stopover pages based on specific text patterns.

## Core Functionality
- **PDF Analysis**: Detects pages containing both a 3-letter airport code in format `[XYZ]-Bilan` and the word "objectifs" (case-insensitive)
- **Visual Interface**: PySide6-based GUI with tabbed interface for stopover management, mapping configuration, and email preview
- **Email Integration**: Outlook integration for sending stopover reports with customizable templates
- **Page Preview**: PDF page rendering for visual confirmation of detected stopovers

## Key Features
- Cross-platform desktop application (Windows, macOS, Linux)
- Real-time PDF processing with progress indication
- Configuration management with automatic persistence
- Account management for email services
- Template-based email generation with placeholder support

## Target Users
Business users who need to process PDF documents and extract specific stopover information for reporting and communication purposes.