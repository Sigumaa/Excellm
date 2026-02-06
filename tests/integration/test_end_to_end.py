from __future__ import annotations

from excelmd.api import convert_xlsx_to_markdown, load_xlsx


def test_convert_all_samples_success(sample_files) -> None:
    for path in sample_files.values():
        markdown = convert_xlsx_to_markdown(path)
        assert markdown.startswith(f"# Workbook: {path.name}")
        assert "## Extraction Summary" in markdown


def test_hidden_sheet_is_included(sample_files) -> None:
    doc = load_xlsx(sample_files["design"])
    hidden_sheets = [sheet for sheet in doc.sheets if sheet.state != "visible"]

    assert any(sheet.name == "データ" for sheet in hidden_sheets)


def test_embedded_images_and_connectors(sample_files) -> None:
    design_doc = load_xlsx(sample_files["design"])
    flow_doc = load_xlsx(sample_files["flow"])

    assert design_doc.summary["embedded_image_count"] == 3
    assert flow_doc.summary["connector_count"] == 38


def test_markdown_contains_mermaid_and_unsupported_sections(sample_files) -> None:
    markdown = convert_xlsx_to_markdown(sample_files["flow"])

    assert "### Diagram Workspace" in markdown
    assert "```mermaid" in markdown
    assert "### Unsupported Elements" in markdown
