from __future__ import annotations

from excelmd.api import load_xlsx


def test_formula_and_merge_extraction(sample_files) -> None:
    doc = load_xlsx(sample_files["design"])

    assert doc.summary["formula_count"] > 0
    assert doc.summary["merge_count"] > 0
    assert any(cell.formula and cell.cached_value is not None for sheet in doc.sheets for cell in sheet.cells)


def test_style_xml_equivalent_exists(sample_files) -> None:
    doc = load_xlsx(sample_files["design"])
    styles = doc.styles_xml_equivalent

    assert styles["tag"] == "styleSheet"
    assert any(child["tag"] == "fonts" for child in styles.get("children", []))
