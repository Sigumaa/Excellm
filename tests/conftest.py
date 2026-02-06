from __future__ import annotations

from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
WORKSPACE = ROOT

SAMPLE_FILES = {
    "design": WORKSPACE / "システム機能設計書(画面)_WA10202_プロジェクト照会.xlsx",
    "screen_list": WORKSPACE / "画面一覧_A1_プロジェクト管理システム.xlsx",
    "flow": WORKSPACE / "画面遷移図_プロジェクト管理システム.xlsx",
}


@pytest.fixture(scope="session")
def sample_files() -> dict[str, Path]:
    for key, path in SAMPLE_FILES.items():
        if not path.exists():
            raise FileNotFoundError(f"Missing sample file [{key}]: {path}")
    return SAMPLE_FILES
