from __future__ import annotations

import argparse
from pathlib import Path

from .api import load_xlsx
from .model import ConvertOptions


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert .xlsx into a single Markdown file")
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("-o", "--output", type=Path, required=True, help="Output Markdown path")
    parser.add_argument(
        "--strict-unsupported",
        action="store_true",
        help="Fail if unsupported elements are detected",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    options = ConvertOptions(strict_unsupported=args.strict_unsupported)
    workbook = load_xlsx(args.input, options=options)
    args.output.write_text(workbook.markdown, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
