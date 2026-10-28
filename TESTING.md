# Disassembly Wizard - System Verification & Testing Guide

This document defines the verification procedures, expected deliverables, and error-handling telemetry for the Desktop GUI Integration (`app.py`).

---

## 1. Test Matrix: Input Models vs. Expected Outputs

| Test Case | Input JSON Model | Selected Format | Disassembly Parameters | Expected Generated Output | Verification Criteria |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **TC-01** | `BialettiGioia.json` | **PDF (.pdf)** | Full Depth, BoM: Enabled | `BialettiGioia_guide.pdf` | Document renders product details, 7 disassembly actions, and complete BoM table. |
| **TC-02** | `BialettiGioia.json` | **Word (.docx)** | Full Depth, BoM: Enabled | `BialettiGioia_guide.docx` | Clean table formatting, safe extraction of tuple-based component entities, valid heading levels. |
| **TC-03** | `BialettiGioia.json` | **Markdown (.md)** | Full Depth, BoM: Disabled | `BialettiGioia_guide.md` | Markdown headings, bullet points, no BoM section present. |
| **TC-04** | `Nespresso.json` | **HTML + JS** | Full Depth, BoM: Enabled | `Nespresso_guide.html` | Self-contained interactive single-page wizard with navigation buttons and step counter. |
| **TC-05** | `Air_fryer_Philips_HD9252.json` | **PowerPoint (.pptx)**| KEEP_MAIN, BoM: Enabled | `Air_fryer_Philips_HD9252_guide.pptx` | Editable slides, title slide, paginated step cards, aspect-ratio preserved image boxes. |
| **TC-06** | `BialettiGioia.json` | **Text (.txt)** | Full Depth, BoM: Enabled | `BialettiGioia_guide.txt` | Formatted ASCII borders, line-wrapped instruction strings, clean summary metrics. |

---

## 2. Output Path Verification

| Scenario | Input Configuration | Execution | Result |
| :--- | :--- | :--- | :--- |
| **Default Destination** | Leave "Output Directory" empty. | Click `GENERATE EXPORT`. | File is saved in the source JSON directory with `<model_name>_guide.<ext>`. |
| **Custom Destination** | Select a custom folder via `Browse Output`. | Click `GENERATE EXPORT`. | File is generated inside the selected destination folder. |

---

## 3. Error Exploration & Boundary Testing

The system implements defensive exception handling to prevent application crashes and display clear user telemetry.

### Test Case E-01: Missing Input Model
- **Trigger:** Click `GENERATE EXPORT` without choosing a JSON file.
- **GUI Telemetry:** Warning dialog: *"Please select a JSON model file first."*
- **Application State:** Stable, export aborted cleanly.

### Test Case E-02: Malformed or Non-Model JSON
- **Trigger:** Load an invalid JSON file (e.g., empty file or corrupted syntax).
- **GUI Telemetry:** `[PARSING ERROR] ...` logged to the console; error message dialog displayed.
- **Application State:** Stable, system returns to ready state without exiting.

### Test Case E-03: Destination Permission Lock
- **Trigger:** Generate a `.docx` or `.pptx` file while the existing target file is open in Microsoft Office.
- **GUI Telemetry:** `[EXECUTION ERROR] Permission denied: <file_path>` displayed in console.
- **Application State:** Informative dialog prompting the user to close the open application.

### Test Case E-04: Schema Tuple Mismatch Recovery
- **Trigger:** Process steps where the loader returns `(Component, quantity)` tuples instead of plain objects.
- **Internal Handling:** Addressed via the `_extract_name()` helper across export routines, preventing `AttributeError: 'tuple' object has no attribute 'name'`.
- **Expected Result:** Export completes successfully (`[SUCCESS] File generated at...`).