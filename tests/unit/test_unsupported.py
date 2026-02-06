from __future__ import annotations

from zipfile import ZipFile

from excelmd.model import ConvertOptions
from excelmd.parser.drawing import parse_drawing_for_sheet


def test_unsupported_element_capture_in_drawing(tmp_path) -> None:
    xlsx = tmp_path / "mini.xlsx"
    drawing_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"
          xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:foo />
    <xdr:clientData />
  </xdr:twoCellAnchor>
</xdr:wsDr>
"""

    with ZipFile(xlsx, "w") as zf:
        zf.writestr("xl/drawings/drawing1.xml", drawing_xml)

    unsupported = []
    warnings = []
    with ZipFile(xlsx) as zf:
        parse_drawing_for_sheet(
            zip_file=zf,
            drawing_path="xl/drawings/drawing1.xml",
            content_types={},
            options=ConvertOptions(),
            unsupported=unsupported,
            warnings=warnings,
        )

    assert len(unsupported) == 1
    assert unsupported[0].tag == "foo"
