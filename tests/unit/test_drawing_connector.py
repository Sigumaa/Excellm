from __future__ import annotations

from excelmd.api import load_xlsx


def test_group_shape_recursive_extraction(sample_files) -> None:
    doc = load_xlsx(sample_files["flow"])
    grouped_children = [obj for sheet in doc.sheets for obj in sheet.drawings if obj.parent_uid is not None]
    assert grouped_children, "Expected grouped child drawing objects"


def test_connector_direction_and_resolution(sample_files) -> None:
    doc = load_xlsx(sample_files["flow"])
    connectors = [conn for sheet in doc.sheets for conn in sheet.connectors]

    assert len(connectors) == 38
    assert any(conn.direction in {"forward", "reverse", "undirected"} for conn in connectors)
    assert sum(1 for conn in connectors if conn.resolved) >= 10
