from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal


@dataclass(slots=True)
class ConvertOptions:
    include_hidden_sheets: bool = True
    image_mode: Literal["data_uri"] = "data_uri"
    style_level: Literal["xml_equivalent"] = "xml_equivalent"
    strict_unsupported: bool = False
    output_mode: Literal["work", "full"] = "work"


@dataclass(slots=True)
class RangeRef:
    ref: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int


@dataclass(slots=True)
class DefinedName:
    name: str
    value: str
    local_sheet_id: int | None = None


@dataclass(slots=True)
class CellData:
    coord: str
    row: int
    col: int
    cell_type: str
    value: str
    formula: str | None
    cached_value: str | None
    style_id: str | None


@dataclass(slots=True)
class DataValidation:
    type: str | None
    sqref: str
    allow_blank: bool | None
    show_error_message: bool | None
    operator: str | None
    formula1: str | None
    formula2: str | None


@dataclass(slots=True)
class AnchorPoint:
    col: int
    row: int
    col_off: int
    row_off: int


@dataclass(slots=True)
class DrawingObject:
    object_uid: str
    object_id: str
    drawing_path: str
    kind: str
    name: str
    text: str
    anchor_type: str
    anchor_from: AnchorPoint | None
    anchor_to: AnchorPoint | None
    bbox: tuple[float, float, float, float]
    parent_uid: str | None
    image_target: str | None
    image_content_type: str | None
    image_data_uri: str | None
    raw_xml: str
    extra: dict[str, str] = field(default_factory=dict)


@dataclass(slots=True)
class ConnectorInfo:
    object_uid: str
    object_id: str
    drawing_path: str
    name: str
    text: str
    anchor_from: AnchorPoint | None
    anchor_to: AnchorPoint | None
    bbox: tuple[float, float, float, float]
    arrow_head: str | None
    arrow_tail: str | None
    direction: str
    source_uid: str | None
    target_uid: str | None
    resolved: bool
    distance_source: float | None
    distance_target: float | None
    raw_xml: str


@dataclass(slots=True)
class UnsupportedElement:
    scope: str
    location: str
    tag: str
    raw_xml: str


@dataclass(slots=True)
class RegionCellRow:
    coord: str
    value: str
    formula: str | None
    cached_value: str | None
    cell_type: str
    style_id: str | None
    merge_ref: str | None
    flags: list[str] = field(default_factory=list)


@dataclass(slots=True)
class CellRegion:
    region_id: int
    bounds: RangeRef
    rows: list[RegionCellRow]


@dataclass(slots=True)
class SheetDoc:
    index: int
    name: str
    state: str
    path: str
    dimension_ref: str
    cells: list[CellData] = field(default_factory=list)
    cell_map: dict[str, CellData] = field(default_factory=dict)
    merges: list[RangeRef] = field(default_factory=list)
    merge_map: dict[str, str] = field(default_factory=dict)
    data_validations: list[DataValidation] = field(default_factory=list)
    row_heights: dict[int, float] = field(default_factory=dict)
    col_widths: dict[int, float] = field(default_factory=dict)
    print_areas: list[RangeRef] = field(default_factory=list)
    print_titles: list[str] = field(default_factory=list)
    page_setup: dict[str, str] = field(default_factory=dict)
    page_margins: dict[str, str] = field(default_factory=dict)
    print_options: dict[str, str] = field(default_factory=dict)
    header_footer: dict[str, str] = field(default_factory=dict)
    page_breaks: dict[str, list[int]] = field(default_factory=dict)
    drawings: list[DrawingObject] = field(default_factory=list)
    connectors: list[ConnectorInfo] = field(default_factory=list)
    mermaid: str = ""
    regions: list[CellRegion] = field(default_factory=list)
    unsupported: list[UnsupportedElement] = field(default_factory=list)


@dataclass(slots=True)
class WorkbookDoc:
    source_path: Path
    options: ConvertOptions
    source_metadata: dict[str, Any] = field(default_factory=dict)
    styles_xml_equivalent: dict[str, Any] = field(default_factory=dict)
    defined_names: list[DefinedName] = field(default_factory=list)
    sheets: list[SheetDoc] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    summary: dict[str, Any] = field(default_factory=dict)
    markdown: str = ""
