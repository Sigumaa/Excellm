from __future__ import annotations

from pathlib import Path

from .model import ConvertOptions, WorkbookDoc


def load_xlsx(path: str | Path, *, options: ConvertOptions | None = None) -> WorkbookDoc:
    raise NotImplementedError("load_xlsx is not implemented yet")


def convert_xlsx_to_markdown(path: str | Path, *, options: ConvertOptions | None = None) -> str:
    doc = load_xlsx(path, options=options)
    return doc.markdown
