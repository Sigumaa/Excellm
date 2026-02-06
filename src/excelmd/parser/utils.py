from __future__ import annotations

import posixpath
import re
from typing import Iterable

from ..model import RangeRef

CELL_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$")
RANGE_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$")
SHEET_RANGE_RE = re.compile(r"^(?:'([^']+)'|([^!]+))!(.+)$")


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def col_to_index(col: str) -> int:
    value = 0
    for char in col.upper():
        value = value * 26 + (ord(char) - 64)
    return value


def index_to_col(index: int) -> str:
    if index < 1:
        raise ValueError("Column index must be >= 1")
    result: list[str] = []
    value = index
    while value > 0:
        value, rem = divmod(value - 1, 26)
        result.append(chr(65 + rem))
    return "".join(reversed(result))


def coord_to_rowcol(coord: str) -> tuple[int, int]:
    match = CELL_RE.match(coord)
    if not match:
        raise ValueError(f"Invalid coordinate: {coord}")
    col = col_to_index(match.group(1))
    row = int(match.group(2))
    return row, col


def rowcol_to_coord(row: int, col: int) -> str:
    if row < 1 or col < 1:
        raise ValueError("row/col must be >= 1")
    return f"{index_to_col(col)}{row}"


def parse_range_ref(ref: str) -> RangeRef:
    normalized = ref.replace("$", "")
    range_match = RANGE_RE.match(normalized)
    if range_match:
        sc = col_to_index(range_match.group(1))
        sr = int(range_match.group(2))
        ec = col_to_index(range_match.group(3))
        er = int(range_match.group(4))
        return RangeRef(
            ref=normalized,
            start_row=min(sr, er),
            start_col=min(sc, ec),
            end_row=max(sr, er),
            end_col=max(sc, ec),
        )

    cell_match = CELL_RE.match(normalized)
    if not cell_match:
        raise ValueError(f"Invalid range reference: {ref}")

    col = col_to_index(cell_match.group(1))
    row = int(cell_match.group(2))
    return RangeRef(ref=normalized, start_row=row, start_col=col, end_row=row, end_col=col)


def parse_sheet_scoped_range(value: str) -> list[RangeRef]:
    raw = value.strip()
    if not raw:
        return []

    match = SHEET_RANGE_RE.match(raw)
    if match:
        payload = match.group(3)
    else:
        payload = raw

    refs: list[RangeRef] = []
    for part in payload.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            refs.append(parse_range_ref(part))
        except ValueError:
            continue
    return refs


def parse_sqref(sqref: str) -> list[RangeRef]:
    refs: list[RangeRef] = []
    for token in sqref.split():
        token = token.strip()
        if not token:
            continue
        try:
            refs.append(parse_range_ref(token))
        except ValueError:
            continue
    return refs


def iter_cells_in_range(rng: RangeRef) -> Iterable[tuple[int, int]]:
    for row in range(rng.start_row, rng.end_row + 1):
        for col in range(rng.start_col, rng.end_col + 1):
            yield row, col


def resolve_target(base_path: str, target: str) -> str:
    joined = posixpath.normpath(posixpath.join(posixpath.dirname(base_path), target))
    if joined.startswith("/"):
        joined = joined[1:]
    return joined


def xml_to_dict(element) -> dict:
    children = [xml_to_dict(child) for child in list(element)]
    text = (element.text or "").strip()
    payload: dict[str, object] = {
        "tag": local_name(element.tag),
        "attrs": dict(sorted(element.attrib.items())),
    }
    if text:
        payload["text"] = text
    if children:
        payload["children"] = children
    return payload
