from __future__ import annotations

import hashlib
from pathlib import Path

from excelmd.api import convert_xlsx_to_markdown, load_xlsx

from tests.helpers import ground_truth_counts


def test_summary_counts_match_ground_truth(sample_files) -> None:
    for path in sample_files.values():
        doc = load_xlsx(path)
        truth = ground_truth_counts(path)

        assert doc.summary["sheet_count"] == truth["sheet_count"]
        assert doc.summary["merge_count"] == truth["merge_count"]
        assert doc.summary["formula_count"] == truth["formula_count"]
        assert doc.summary["drawing_object_count"] == truth["drawing_object_count"]
        assert doc.summary["connector_count"] == truth["connector_count"]


def test_markdown_snapshot_hashes(sample_files) -> None:
    golden_dir = Path(__file__).resolve().parents[1] / "golden"
    mapping = {
        "design": "design.sha256",
        "screen_list": "screen_list.sha256",
        "flow": "flow.sha256",
    }

    for key, file_name in mapping.items():
        markdown = convert_xlsx_to_markdown(sample_files[key])
        digest = hashlib.sha256(markdown.encode("utf-8")).hexdigest()
        golden_path = golden_dir / file_name
        expected = golden_path.read_text(encoding="utf-8").strip()
        assert digest == expected
