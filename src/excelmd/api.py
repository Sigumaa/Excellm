from __future__ import annotations

from pathlib import Path

from .model import ConvertOptions, WorkbookDoc
from .parser.ooxml import OOXMLWorkbookParser
from .render_html import render_workbook_html
from .render_markdown import render_workbook_markdown


def _parse_xlsx(path: str | Path, *, options: ConvertOptions | None = None) -> WorkbookDoc:
    opts = options or ConvertOptions()
    parser = OOXMLWorkbookParser(path, opts)
    return parser.parse()


def load_xlsx(path: str | Path, *, options: ConvertOptions | None = None) -> WorkbookDoc:
    workbook = _parse_xlsx(path, options=options)
    workbook.markdown = render_workbook_markdown(workbook)
    return workbook


def convert_xlsx_to_markdown(path: str | Path, *, options: ConvertOptions | None = None) -> str:
    workbook = _parse_xlsx(path, options=options)
    return render_workbook_markdown(workbook)


def convert_xlsx_to_html(path: str | Path, *, options: ConvertOptions | None = None) -> str:
    workbook = _parse_xlsx(path, options=options)
    return render_workbook_html(workbook)
