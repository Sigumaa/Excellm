from __future__ import annotations

from pathlib import Path

from .model import ConvertOptions, WorkbookDoc
from .parser.ooxml import OOXMLWorkbookParser
from .render_markdown import render_workbook_markdown


def load_xlsx(path: str | Path, *, options: ConvertOptions | None = None) -> WorkbookDoc:
    opts = options or ConvertOptions()
    parser = OOXMLWorkbookParser(path, opts)
    workbook = parser.parse()
    workbook.markdown = render_workbook_markdown(workbook)
    return workbook


def convert_xlsx_to_markdown(path: str | Path, *, options: ConvertOptions | None = None) -> str:
    doc = load_xlsx(path, options=options)
    return doc.markdown
