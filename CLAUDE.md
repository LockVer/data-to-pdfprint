# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Data to PDF Print (data-to-pdfprint) is a Python CLI tool for reading Excel data and generating PDF labels. It supports custom templates and allows generating different styles of PDF labels from the same data.

## Development Commands

### Installation and Setup
```bash
# Install dependencies
pip install -r requirements.txt

# Install project in development mode
pip install -e .
```

### Testing
```bash
# Run all tests
python -m pytest tests/

# Run specific test file (when tests directory exists)
python -m pytest tests/test_excel_reader.py
```

### Code Quality
```bash
# Format code
black src/

# Lint code
flake8 src/
```

### Running the CLI
```bash
# Basic usage with default template
data-to-pdf --input data.xlsx --template basic

# Custom template usage
data-to-pdf --input data.xlsx --template custom_template

# Batch generation
data-to-pdf --input data.xlsx --template basic --output output_dir/
```

## Architecture

### Core Modules Structure

The project follows a modular architecture with clear separation of concerns:

- **CLI Layer** (`src/cli/`): Command-line interface and argument parsing
- **Data Layer** (`src/data/`): Excel reading and data processing
- **Template System** (`src/template/`): Template management and rendering (planned)
- **PDF Generation** (`src/pdf/`): ReportLab-based PDF creation
- **Configuration** (`src/config/`): Settings and configuration management
- **Utilities** (`src/utils/`): Common helper functions

### Key Dependencies

- `reportlab>=3.6.0` - PDF generation
- `pandas>=1.5.0` - Data manipulation
- `openpyxl>=3.1.0` - Excel file reading
- `click>=8.1.0` - CLI framework

### Template System Design

The template system is designed around these key concepts:
- **BaseTemplate**: Base class defining template interface
- **Template Manager**: Handles template creation, loading, and management
- **Builtin Templates**: Pre-defined common templates
- Templates are stored in `/templates/` directory (to be created)

### Data Flow

1. **Excel Reading**: `ExcelReader` class reads Excel files using pandas/openpyxl
2. **Data Processing**: `DataProcessor` extracts and validates specific fields
3. **Template Application**: Template system renders data according to selected template
4. **PDF Generation**: `PDFGenerator` creates final PDF using ReportLab

### Configuration Management

The `Settings` class manages:
- Project root directory paths
- Template directory location (`PROJECT_ROOT/templates`)
- Default output directory (`PROJECT_ROOT/output`)
- Configuration file loading/saving

## Development Notes

### Current Implementation Status
The codebase currently contains skeleton classes with method stubs. Most core functionality needs implementation.

### Missing Directories
- `/templates/` - Template files storage
- `/tests/` - Test files (planned structure exists in README)

### Entry Point
CLI entry point is configured in setup.py as `data-to-pdf=src.cli.main:main`

### File Naming Convention
Uses lowercase with underscores (snake_case) for Python files and modules.