from __future__ import annotations

import argparse
from pathlib import Path

from .api import convert_xlsx_to_html, load_xlsx
from .model import ConvertOptions


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert .xlsx into Markdown or HTML")
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("-o", "--output", type=Path, required=True, help="Output path")
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "--sheetview",
        action="store_true",
        help="Render sheet-view style markdown with HTML table layout",
    )
    parser.add_argument(
        "--strict-unsupported",
        action="store_true",
        help="Fail if unsupported elements are detected",
    )
    mode_group.add_argument(
        "--full",
        action="store_true",
        help="Render full fidelity markdown (large output)",
    )
    parser.add_argument(
        "--html",
        action="store_true",
        help="Output standalone HTML (sheet-view reproduction)",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    options = ConvertOptions(
        strict_unsupported=args.strict_unsupported,
        output_mode="full" if args.full else ("sheetview" if args.sheetview else "work"),
    )
    if args.html:
        html = convert_xlsx_to_html(args.input, options=options)
        args.output.write_text(html, encoding="utf-8")
        return 0

    workbook = load_xlsx(args.input, options=options)
    args.output.write_text(workbook.markdown, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
