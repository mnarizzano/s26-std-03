"""Small executable wrapper for the Futurdata PPTX converter.

Running ``python main.py`` exports the included example from ``data/`` to
``output/``. Optional positional arguments let users select another JSON input
and output file without changing source code.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from pptx_exporter import export_to_pptx


CONVERTER_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = CONVERTER_DIR / "data" / "example_air_fryer.json"
DEFAULT_OUTPUT = CONVERTER_DIR / "output" / "example_air_fryer.pptx"


def build_parser() -> argparse.ArgumentParser:
    """Create the command-line parser used by the simple launcher."""

    parser = argparse.ArgumentParser(
        description="Generate an editable PPTX disassembly guide."
    )
    parser.add_argument(
        "json_path",
        nargs="?",
        default=str(DEFAULT_INPUT),
        help="Builder/IR JSON path. Defaults to the included Air Fryer example.",
    )
    parser.add_argument(
        "output_path",
        nargs="?",
        default=str(DEFAULT_OUTPUT),
        help="Destination .pptx path.",
    )
    parser.add_argument("--no-bom", action="store_true")
    parser.add_argument("--start-step", type=int)
    parser.add_argument("--end-step", type=int)
    parser.add_argument("--groups-per-slide", type=int, choices=(1, 2, 3, 4), default=1)
    parser.add_argument("--no-images", action="store_true")
    return parser


def main() -> int:
    """Execute one conversion and print the absolute output path."""

    args = build_parser().parse_args()
    generated = export_to_pptx(
        args.json_path,
        args.output_path,
        include_bom=not args.no_bom,
        start_step=args.start_step,
        end_step=args.end_step,
        groups_per_slide=args.groups_per_slide,
        include_images=not args.no_images,
    )
    print(f"PowerPoint generated successfully:\n{generated}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
