from __future__ import annotations

from excelmd.parser.utils import (
    col_to_index,
    coord_to_rowcol,
    index_to_col,
    parse_range_ref,
    parse_sheet_scoped_range,
    rowcol_to_coord,
)


def test_col_index_roundtrip() -> None:
    assert col_to_index("A") == 1
    assert col_to_index("Z") == 26
    assert col_to_index("AA") == 27
    assert index_to_col(1) == "A"
    assert index_to_col(27) == "AA"
    assert index_to_col(52) == "AZ"


def test_coord_roundtrip() -> None:
    assert coord_to_rowcol("C12") == (12, 3)
    assert rowcol_to_coord(12, 3) == "C12"


def test_parse_range_ref() -> None:
    rng = parse_range_ref("$B$2:$D$4")
    assert (rng.start_row, rng.start_col, rng.end_row, rng.end_col) == (2, 2, 4, 4)


def test_parse_sheet_scoped_range() -> None:
    ranges = parse_sheet_scoped_range("'Sheet 1'!$A$1:$C$3,$D$5")
    assert len(ranges) == 2
    assert ranges[0].ref == "A1:C3"
    assert ranges[1].ref == "D5"
