# Futurdata JSON-to-PPTX Converter

## Purpose

This directory contains the PowerPoint export module developed for the
Futurdata Disassembly Wizard. The converter receives either:

1. a Builder JSON model that can be processed by the shared
   `disassembly_loader`, or
2. an already-normalized Loader Intermediate Representation (IR) JSON file.

It generates an editable `.pptx` disassembly guide. The source JSON is opened
read-only and is never modified.

## Architecture

```text
Builder JSON
    ↓
disassembly_loader.build_guide(...)
    ↓
Normalized guide / IR
    ↓
PPTX validation and rendering pipeline
    ↓
Editable PowerPoint presentation
```

The converter deliberately does not reconstruct the Builder graph. Graph
validation and linearization remain responsibilities of the shared Loader.

## Correct repository location

```text
src/main/pptx-converter/
```

The public entry point is:

```text
src/main/pptx-converter/pptx_exporter.py
```

The obsolete files/directories from the older implementation should be
removed after migration:

```text
src/main/pptx_exporter.py
src/main/futurdata_pptx/
src/main/requirements-pptx.txt
src/tests/test_pptx_exporter.py
```

## Main features

The generated presentation can include:

- product title and overview slides;
- Loader warnings and PPTX-side validation issues;
- aggregated tools and safety notices;
- a paginated Bill of Materials;
- parent disassembly operations;
- atomic actions grouped under each parent operation;
- removed components and remaining assemblies;
- nominal and measured weights;
- material, color, quality, and destination metadata;
- local, URL, and data-URI images when available;
- configurable first and last step;
- configurable maximum number of action groups;
- one to four action groups per slide;
- slide numbering and a final export summary.

All text, tables, shapes, and images remain editable in PowerPoint.

## Directory structure

```text
pptx-converter/
├── README.md
├── requirements.txt
├── requirements-dev.txt
├── main.py
├── pptx_exporter.py
├── gui_integration_example.py
├── data/
│   └── example_air_fryer.json
├── output/
│   └── example_air_fryer.pptx
└── pptx_core/
    ├── __init__.py
    ├── engine.py
    ├── loader.py
    ├── models.py
    ├── options.py
    ├── utils.py
    ├── validation.py
    └── exporters/
        ├── __init__.py
        ├── base.py
        ├── registry.py
        └── pptx_exporter.py
```

### File responsibilities

- `main.py` — easiest command-line launcher; runs the included example by
  default.
- `pptx_exporter.py` — public API used by the desktop GUI and scripts.
- `pptx_core/loader.py` — finds the shared Loader or reads normalized IR JSON.
- `pptx_core/models.py` — adapts dictionaries, dataclasses, and Loader objects
  to one stable internal model.
- `pptx_core/validation.py` — validates steps, components, branches, weights,
  images, and depth selections.
- `pptx_core/engine.py` — coordinates loading, validation, and rendering.
- `pptx_core/exporters/pptx_exporter.py` — creates the PowerPoint slides.
- `data/` — example input JSON.
- `output/` — official generated example.

## Requirements

- Python 3.10 or newer
- `python-pptx`
- `Pillow`
- `pytest` only when running the tests

Install the converter dependencies from the repository root.

### Windows PowerShell

Runtime only:

```powershell
python -m pip install -r "src/main/pptx-converter/requirements.txt"
```

Runtime plus tests:

```powershell
python -m pip install -r "src/main/pptx-converter/requirements-dev.txt"
```

### macOS/Linux shell

Runtime only:

```bash
python3 -m pip install -r src/main/pptx-converter/requirements.txt
```

Runtime plus tests:

```bash
python3 -m pip install -r src/main/pptx-converter/requirements-dev.txt
```

## Quick start

The simplest command runs the included example.

### Windows PowerShell

```powershell
python "src/main/pptx-converter/main.py"
```

### macOS/Linux shell

```bash
python3 src/main/pptx-converter/main.py
```

Expected output:

```text
src/main/pptx-converter/output/example_air_fryer.pptx
```

## Convert another JSON file

### Windows PowerShell — one line

```powershell
python "src/main/pptx-converter/pptx_exporter.py" "path/to/input.json" "path/to/output.pptx"
```

### Windows PowerShell — multiple lines

PowerShell uses the backtick character, not `\`, for line continuation.
There must be no spaces after the backtick.

```powershell
python "src/main/pptx-converter/pptx_exporter.py" `
  "path/to/input.json" `
  "path/to/output.pptx"
```

### macOS/Linux shell

```bash
python3 src/main/pptx-converter/pptx_exporter.py \
  path/to/input.json \
  path/to/output.pptx
```

## Useful command-line options

Export only Steps 3–10 and place two parent operations on each slide:

```powershell
python "src/main/pptx-converter/pptx_exporter.py" "src/main/pptx-converter/data/example_air_fryer.json" "partial.pptx" --start-step 3 --end-step 10 --groups-per-slide 2
```

Exclude the Bill of Materials and images:

```powershell
python "src/main/pptx-converter/pptx_exporter.py" "src/main/pptx-converter/data/example_air_fryer.json" "compact.pptx" --no-bom --no-images
```

## Public Python API

```python
from pptx_exporter import export_to_pptx

output = export_to_pptx(
    json_path="model.json",
    output_path="guide.pptx",
    depth=None,
    include_bom=True,
    start_step=None,
    end_step=None,
    max_action_groups=None,
    groups_per_slide=1,
    include_images=True,
)
```

The function returns the absolute path of the generated `.pptx` file.

## Integration with the team GUI

The folder name contains a hyphen, so it cannot be imported as a dotted Python
package. A GUI file located in `src/main` must add the converter directory to
`sys.path` before importing the public module:

```python
from pathlib import Path
import sys

PPTX_CONVERTER_DIR = Path(__file__).resolve().parent / "pptx-converter"
if str(PPTX_CONVERTER_DIR) not in sys.path:
    sys.path.insert(0, str(PPTX_CONVERTER_DIR))

from pptx_exporter import export_to_pptx
```

The current GUI call remains supported:

```python
generated_path = export_to_pptx(
    path,
    depth=spec,
    include_bom=bom_needed,
)
```

The preferred integration reuses the guide already created by the GUI. This
avoids parsing the same source twice:

```python
generated_path = export_to_pptx(
    json_path=path,
    guide=guide,
    depth=spec,
    include_bom=bom_needed,
    output_path=output_path,
)
```

## Branch-source handling

The converter never assumes that a new branch starts from the product merely
because the previous step has no continuation. The source is resolved in this
order:

1. explicit `step.source` supplied by the Loader;
2. product root for the first step;
3. previous `continues_as` for a linear continuation;
4. `Source not specified in JSON` for an unencoded new branch.

This prevents unrelated branches from being connected incorrectly. For exact
continuity, the Loader IR should provide `step.source` whenever a new branch
begins.

## Validation

Before saving the presentation, the converter checks:

- required product and step data;
- duplicate or invalid step indexes;
- invalid component weights;
- empty atomic actions;
- missing branch sources;
- product and operation mass balance;
- missing local image files;
- invalid keep-whole references.

Blocking errors stop strict exports. Non-blocking warnings are shown inside the
presentation.

## Run the tests

From the repository root:

```powershell
python -m pytest "src/tests/test_pptx_converter.py" -v
```

The tests cover:

- complete PPTX generation;
- BoM inclusion and exclusion;
- step filtering and grouped slides;
- reusing an already-built guide;
- safe branch handling;
- dictionary-based Loader image metadata;
- invalid JSON reporting;
- automatic `.pptx` extension handling.

## Troubleshooting

### `ModuleNotFoundError: No module named 'pptx'`

```powershell
python -m pip install -r "src/main/pptx-converter/requirements.txt"
```

### `ModuleNotFoundError: No module named 'disassembly_loader'`

The standalone converter automatically checks both known Loader locations:

```text
src/main/loader_se/disassembly_loader
loader_se/disassembly_loader
```

If neither exists, the input must already be normalized IR JSON.

### `PermissionError`

Close the generated `.pptx` in PowerPoint and run the export again.

### PowerShell reports `Unexpected token '\'`

Use a one-line command or PowerShell backticks as shown above. A backslash is
only a line-continuation character in Bash-like shells.

## Output and resource paths

The PPTX converter keeps each type of resource in one place:

```text
src/main/pptx-converter/
├── data/       # JSON input/example files
├── images/     # local image assets referenced by JSON
└── output/     # generated .pptx files only
```

Generated PowerPoint files must be written to `output/`. The converter does not require or duplicate PPTX/JSON reference files under `docs/ref/pptx-converter/`.

## Bundled image assets and Nespresso regression example

The converter now ships with the image assets used by the reference models:

```text
src/main/pptx-converter/
├── data/
│   ├── Nespresso.json
│   └── AirFryer_Philips_HD9252.json
├── images/
│   ├── machine_root.jpg
│   ├── action_remove_external.jpg
│   ├── water_pump.jpg
│   ├── thermoblock.jpg
│   └── ...
└── output/
    ├── Nespresso.pptx
    └── AirFryer_Philips_HD9252.pptx
```

Image references in the JSON are written as relative paths such as
`images/machine_root.jpg`. The converter resolves them independently of the
current shell directory by checking both the JSON directory and the converter
root. Local images are embedded directly in the generated PowerPoint, so the
`.pptx` remains self-contained after generation.

### Generate the Nespresso example with images

From the repository root on Windows PowerShell:

```powershell
python "src/main/pptx-converter/pptx_exporter.py" "src/main/pptx-converter/data/Nespresso.json" "src/main/pptx-converter/output/Nespresso.pptx"
```

Do **not** add `--no-images` when testing image export. The option deliberately
suppresses figures.

### Generate the Air Fryer example with images

```powershell
python "src/main/pptx-converter/pptx_exporter.py" "src/main/pptx-converter/data/AirFryer_Philips_HD9252.json" "src/main/pptx-converter/output/AirFryer_Philips_HD9252.pptx"
```

The automated tests also verify that the Nespresso presentation contains real
PowerPoint picture shapes, not just image placeholders.
