# Disassembly Wizard - Unified Desktop GUI Exporter

## Overview

A desktop application developed as part of the Futurdata Disassembly Wizard project. It integrates Rubin's `disassembly_loader` with modular exporters to convert product JSON models into comprehensive, formatted disassembly guides across multiple industry formats (PDF, Word, Markdown, HTML+JS, PowerPoint, and Text).

## System Requirements

- Python 3.10 or newer (Python 3.11/3.12 recommended)
- OS: Windows 10/11, macOS, or Linux
- Standard graphical environment supporting Tkinter

## Installation

### 1. Clone or Download the Repository

```bash
git clone <repository-url>
cd <repository-directory>
```

### 2. Create and Activate a Virtual Environment (Recommended)

**Windows (PowerShell):**

```powershell
python -m venv .venv
.venv\Scripts\activate
```

**macOS / Linux:**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

> **Note for Linux:** Tkinter may require system-level installation if it is not already present:
>
> ```bash
> sudo apt-get install python3-tk
> ```

## Running the Application

Launch the graphical interface from the repository root:

```bash
python app.py
```

(or `python3 app.py` on macOS/Linux).

## Application Features & Compliance

- **Cross-Platform Path Resolution:** Dynamic module importing via `importlib.util` and `sys.path` injection guarantees execution on any host operating system, seamlessly handling directory names containing hyphens (`-`).
- **Flexible Destinations:** Users can specify a dedicated output directory for exported files via the GUI, fulfilling thesis evaluation requirements.
- **Configurable Parameters:** Disassembly depth (`FULL` or `KEEP_MAIN`) and Bill of Materials (BoM) generation can be toggled per export run.

## Integrated Exporters

| Format | Extension | Description |
|--------|-----------|-------------|
| PDF | `.pdf` | Formatted visual report |
| Word | `.docx` | Structured document with BoM tables and active subroot tracking |
| Markdown | `.md` | Lightweight documentation ready for Git hosting |
| HTML + JS | `.html` | Interactive browser-based wizard with step navigation and component grading |
| PowerPoint | `.pptx` | Editable presentation slides |
| Text | `.txt` | Plain-text terminal-friendly summary |