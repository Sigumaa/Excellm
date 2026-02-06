from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

from excelmd.model import ConvertOptions
from excelmd.parser.ooxml import OOXMLWorkbookParser


def test_parse_shared_strings_with_rich_text(tmp_path: Path) -> None:
    workbook_path = tmp_path / "dummy.xlsx"
    shared_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\">
  <si><t>plain</t></si>
  <si>
    <r><t>rich</t></r>
    <r><t>Text</t></r>
  </si>
</sst>
"""

    with ZipFile(workbook_path, "w") as zf:
        zf.writestr("xl/sharedStrings.xml", shared_xml)

    parser = OOXMLWorkbookParser(workbook_path, ConvertOptions())
    with ZipFile(workbook_path) as zf:
        values = parser._parse_shared_strings(zf)  # noqa: SLF001

    assert values == ["plain", "richText"]
