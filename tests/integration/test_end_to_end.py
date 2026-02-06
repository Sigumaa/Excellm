from __future__ import annotations

from excelmd.api import convert_xlsx_to_html, convert_xlsx_to_markdown, load_xlsx
from excelmd.model import ConvertOptions


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


def test_sheetview_mode_renders_html_grid(sample_files) -> None:
    markdown = convert_xlsx_to_markdown(
        sample_files["flow"],
        options=ConvertOptions(output_mode="sheetview"),
    )

    assert "## SheetView (Markdown + HTML)" in markdown
    assert '<table class="sv-grid">' in markdown
    assert '<div class="sv-overlay">' in markdown
    assert "sv-formula" not in markdown
    assert '=IF(INDIRECT("変更履歴!E2")<>"",INDIRECT("変更履歴!E2"),"")' not in markdown


def test_standalone_html_output(sample_files) -> None:
    html = convert_xlsx_to_html(sample_files["flow"])

    assert "<!doctype html>" in html.lower()
    assert "<html" in html.lower()
    assert '<table class="sv-grid">' in html
    assert 'class="sv-col-head"' in html
    assert 'class="sv-row-head"' in html
    assert "marker id=\"arrow-triangle\"" in html
    assert "Sheet: 1. ログイン(A101)" in html
    assert "sv-formula" not in html
    assert '=IF(INDIRECT("変更履歴!E2")<>"",INDIRECT("変更履歴!E2"),"")' not in html
    assert "sv-toolbar" not in html
    assert "<script>" not in html.lower()
    assert 'class="sv-page-break"' in html


def test_standalone_html_includes_footer_tokens(sample_files) -> None:
    html = convert_xlsx_to_html(sample_files["design"])

    assert 'class="sv-hf"' in html
    assert 'class="sv-hf-center"' in html
    assert "{page}" in html


def test_work_mode_prefers_displayed_values(sample_files) -> None:
    markdown = convert_xlsx_to_markdown(sample_files["design"])

    assert "### Calculated Cells (Displayed Results)" in markdown
    assert "### Formula Cells" not in markdown
    assert "サンプルシステム" in markdown
    assert '=IF(INDIRECT("変更履歴!E2")<>"",INDIRECT("変更履歴!E2"),"")' not in markdown
