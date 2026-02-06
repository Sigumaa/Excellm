from __future__ import annotations

import argparse
from pathlib import Path

from .api import convert_xlsx_to_markdown


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert .xlsx into a single Markdown file")
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("-o", "--output", type=Path, required=True, help="Output Markdown path")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    markdown = convert_xlsx_to_markdown(args.input)
    args.output.write_text(markdown, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
