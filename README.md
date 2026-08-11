# PPTX2PDF — Lightweight Windows PowerPoint to PDF Converter

A desktop application and CLI tool built with Python and CustomTkinter that converts PowerPoint `.pptx` presentations into `.pdf` files locally on Windows using LibreOffice in headless mode. Includes batch processing, background thread rendering, file collision options, unit tests, PyInstaller packaging, and automated GitHub Actions CI builds.

## Tech Stack
- **CustomTkinter** — Modern desktop GUI
- **LibreOffice** — Headless PowerPoint rendering engine (`soffice`)
- **PyInstaller** — Single-file Windows executable packaging
- **Pytest** — Unit testing suite with mocked dependencies
- **Pillow** — Application icon generation
- **GitHub Actions** — Automated Windows CI workflow

## Project Structure
```text
pptx2pdf/
├── app/
│   ├── __init__.py
│   ├── main.py                     # Entry point (GUI + CLI argument parser)
│   ├── gui.py                      # CustomTkinter GUI layout & background worker thread
│   ├── converter.py                # Core batch conversion engine via subprocess
│   ├── libreoffice.py              # System PATH, Program Files & Registry search
│   ├── utils.py                    # Path validation, safe naming & folder opener
│   └── config.py                   # App metadata, asset path resolution & logging
│
├── assets/
│   └── icon.ico                    # Generated application icon
│
├── dist/
│   └── PPTX2PDF.exe                # Built standalone Windows executable
│
├── tests/
│   ├── __init__.py
│   ├── test_converter.py           # Unit tests for conversion engine & error handling
│   └── test_libreoffice.py         # Unit tests for LibreOffice detection logic
│
├── .github/
│   └── workflows/
│       └── build.yml               # GitHub Actions CI workflow (Windows build + release)
│
├── .gitignore
├── LICENSE
├── README.md
├── requirements.txt                # Runtime dependencies (customtkinter, pillow)
├── requirements-dev.txt            # Development dependencies (pytest, pyinstaller, ruff)
├── build.bat                       # One-click PyInstaller build script
├── run.bat                         # Quick launcher for development mode
└── pptx2pdf.spec                   # PyInstaller bundle specification
```

## Architecture

```text
GUI / CLI Interface
      ↓
Conversion Engine (app/converter.py)
      ↓
Subprocess Staging (tempfile)
      ↓
LibreOffice Headless (soffice --headless --convert-to pdf)
      ↓
PDF Output Folder
```

## Features

- **Batch Conversion**: Select and convert multiple `.pptx` presentations simultaneously.
- **100% Local & Private**: All rendering happens locally via LibreOffice. No files are uploaded to any external server.
- **Non-Blocking GUI**: Conversions run on a background thread so the interface stays responsive with real-time progress updates.
- **Collision Strategies**: Option to overwrite existing PDFs or automatically generate unique output names (e.g. `presentation (1).pdf`).
- **LibreOffice Auto-Detection**: Automatically checks system PATH, Program Files, AppData, and the Windows Registry. Includes inline buttons to download or browse to custom `soffice.exe` locations.
- **Standalone Executable**: Packaged into a single `dist/PPTX2PDF.exe` file. End-users do not need Python installed.

## System Requirements
- **Operating System**: Windows 10 or Windows 11 (64-bit)
- **Engine Requirement**: [LibreOffice](https://www.libreoffice.org/download/) (Free & Open Source). Python is **NOT** required for end-users running the pre-compiled `.exe`.

## Installation & Download

### Method 1 — Download Executable (Recommended for End-Users)
No Python installation required.

1. Download **[PPTX2PDF.exe](https://github.com/Istiyak-Ishan/pptx2pdf-/blob/main/dist/PPTX2PDF.exe)**.
2. Install **[LibreOffice](https://www.libreoffice.org/download/)** if not already installed.
3. Double-click **`PPTX2PDF.exe`** to launch the application.

### Method 2 — Run from Source (Development Setup)
```bash
# 1. Clone repository
git clone https://github.com/Istiyak-Ishan/pptx2pdf-.git
cd pptx2pdf

# 2. Create virtual environment
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt
pip install -r requirements-dev.txt

# 4. Launch GUI
python -m app.main
```

## CLI Mode
Run conversions directly from the terminal without launching the GUI:

```bash
python -m app.main presentation.pptx --output C:\PDFs
python -m app.main pres1.pptx pres2.pptx -o C:\PDFs --no-overwrite
```

## Building the Executable
To compile the standalone Windows executable yourself using PyInstaller:

```bash
# Run build script
build.bat

# Or run PyInstaller directly
pyinstaller --clean pptx2pdf.spec
```

The compiled binary will be generated at:
```text
dist/PPTX2PDF.exe
```

## Running Unit Tests
Run the test suite via `pytest`:

```bash
pytest tests/ -v
```

## Notes on Rendering
Rendering quality depends on LibreOffice's PPTX import engine. While standard slides, text, shapes, and images render identically, complex Office 365 specific animations or custom proprietary Microsoft fonts may display slight visual differences.

## Future Improvements
- Add drag-and-drop file support directly onto the application window
- Add PPTX -> PNG/JPEG slide export mode
- Support portable bundled LibreOffice runner
- Add conversion history log viewer tab

## About
A lightweight Windows desktop application and CLI tool that converts PowerPoint presentations into PDFs locally using LibreOffice in headless mode. Built with Python, CustomTkinter, PyInstaller, and GitHub Actions.

**Executable Location**: [GitHub Releases Page](https://github.com/Istiyak-Ishan/pptx2pdf-/releases) or local build directory `dist/PPTX2PDF.exe`.

**Topics**: `python` `customtkinter` `pptx-to-pdf` `libreoffice` `desktop-app` `windows` `pyinstaller` `batch-converter`
