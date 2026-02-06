from __future__ import annotations

import json
import re

from .model import SheetDoc, WorkbookDoc

_NUMBER_RE = re.compile(r"^[+-]?(?:\d+|\d+\.\d+)$")
_COORD_RE = re.compile(r"^([A-Z]+)(\d+)$")


def render_workbook_markdown(workbook: WorkbookDoc) -> str:
    if workbook.options.output_mode == "full":
        return _render_full_markdown(workbook)
    return _render_work_markdown(workbook)


def _render_work_markdown(workbook: WorkbookDoc) -> str:
    lines: list[str] = []

    lines.append(f"# Workbook: {workbook.source_path.name}")
    lines.append("")

    lines.append("## Source Metadata")
    lines.append("")
    _append_key_value_table(lines, workbook.source_metadata)

    lines.append("")
    lines.append("## Workbook Workboard")
    lines.append("")
    lines.append(f"- workbook_type: `{_infer_document_type(workbook)}`")
    lines.append(
        f"- global_metrics: `sheets={workbook.summary.get('sheet_count', 0)}, "
        f"cells={workbook.summary.get('cell_count', 0)}, merges={workbook.summary.get('merge_count', 0)}, "
        f"formulas={workbook.summary.get('formula_count', 0)}, drawings={workbook.summary.get('drawing_object_count', 0)}, "
        f"connectors={workbook.summary.get('connector_count', 0)}, images={workbook.summary.get('embedded_image_count', 0)}`"
    )

    lines.append("")
    lines.append("| index | sheet | state | inferred_role | print_area |")
    lines.append("|---:|---|---|---|---|")
    for sheet in workbook.sheets:
        lines.append(
            "| "
            + " | ".join(
                [
                    str(sheet.index),
                    _esc(sheet.name),
                    _esc(sheet.state),
                    _esc(_infer_sheet_role(sheet.name)),
                    _esc(", ".join(r.ref for r in sheet.print_areas) if sheet.print_areas else "(none)"),
                ]
            )
            + " |"
        )

    lines.append("")
    lines.append("## Defined Names")
    lines.append("")
    if not workbook.defined_names:
        lines.append("(none)")
    else:
        lines.append("| name | local_sheet_id | value |")
        lines.append("|---|---:|---|")
        for dn in workbook.defined_names:
            lines.append(
                "| "
                + " | ".join(
                    [
                        _esc(dn.name),
                        _esc("" if dn.local_sheet_id is None else str(dn.local_sheet_id)),
                        _esc(_short(dn.value, 160)),
                    ]
                )
                + " |"
            )

    for sheet in workbook.sheets:
        lines.append("")
        lines.append(f"## Sheet: {sheet.name} [{sheet.state}]")
        lines.append("")

        formula_count = sum(1 for c in sheet.cells if c.formula)
        image_count = sum(1 for obj in sheet.drawings if obj.image_target)

        lines.append("### Work Context")
        lines.append("")
        lines.append(f"- used_range: `{sheet.dimension_ref}`")
        lines.append(
            f"- metrics: `cells={len(sheet.cells)}, merges={len(sheet.merges)}, formulas={formula_count}, "
            f"regions={len(sheet.regions)}, validations={len(sheet.data_validations)}, drawings={len(sheet.drawings)}, "
            f"connectors={len(sheet.connectors)}, images={image_count}`"
        )
        lines.append(f"- print_areas: `{', '.join(r.ref for r in sheet.print_areas) if sheet.print_areas else '(none)'}`")
        lines.append(f"- print_titles: `{', '.join(sheet.print_titles) if sheet.print_titles else '(none)'}`")
        lines.append(f"- key_texts: `{ ' / '.join(_representative_texts(sheet, 14)) or '(none)' }`")

        lines.append("")
        lines.append("### Region Workspaces")
        lines.append("")
        if not sheet.regions:
            lines.append("(none)")
        else:
            for region in sheet.regions:
                _append_region_workspace(lines, region.region_id, region.bounds.ref, region.rows)

        lines.append("")
        lines.append("### Formula Cells")
        lines.append("")
        formula_cells = [c for c in sheet.cells if c.formula]
        if not formula_cells:
            lines.append("(none)")
        else:
            lines.append("| coord | value | formula | cached_value |")
            lines.append("|---|---|---|---|")
            for cell in formula_cells:
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            _esc(cell.coord),
                            _esc(_short(cell.value, 60)),
                            _esc(_short(cell.formula or "", 120)),
                            _esc(_short(cell.cached_value or "", 60)),
                        ]
                    )
                    + " |"
                )

        lines.append("")
        lines.append("### Input Rules")
        lines.append("")
        if not sheet.data_validations:
            lines.append("(none)")
        else:
            lines.append("| sqref | type | formula1 | formula2 |")
            lines.append("|---|---|---|---|")
            for dv in sheet.data_validations:
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            _esc(dv.sqref),
                            _esc(dv.type or ""),
                            _esc(_short(dv.formula1 or "", 80)),
                            _esc(_short(dv.formula2 or "", 80)),
                        ]
                    )
                    + " |"
                )

        lines.append("")
        lines.append("### Diagram Workspace")
        lines.append("")
        if not sheet.drawings and not sheet.connectors:
            lines.append("(none)")
        else:
            lines.append(
                f"- diagram_metrics: `nodes={sum(1 for o in sheet.drawings if o.kind != 'cxnSp')}, "
                f"raw_connectors={len(sheet.connectors)}, resolved_edges={sum(1 for c in sheet.connectors if c.resolved)}`"
            )

            examples = _connector_examples(sheet, 16)
            if examples:
                lines.append("")
                lines.append("| from | to | label | direction |")
                lines.append("|---|---|---|---|")
                for src, dst, label, direction in examples:
                    lines.append("| " + " | ".join([_esc(src), _esc(dst), _esc(label), _esc(direction)]) + " |")

            if sheet.mermaid:
                lines.append("")
                lines.append("```mermaid")
                lines.append(sheet.mermaid)
                lines.append("```")

        lines.append("")
        lines.append("### Image Assets")
        lines.append("")
        images = [obj for obj in sheet.drawings if obj.image_target]
        if not images:
            lines.append("(none)")
        else:
            lines.append("| object_uid | target | content_type | in_full_mode |")
            lines.append("|---|---|---|---|")
            for obj in images:
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            _esc(obj.object_uid),
                            _esc(obj.image_target or ""),
                            _esc(obj.image_content_type or ""),
                            "data_uri",
                        ]
                    )
                    + " |"
                )

        lines.append("")
        lines.append("### Unsupported Elements")
        lines.append("")
        if not sheet.unsupported:
            lines.append("(none)")
        else:
            lines.append("| scope | location | tag |")
            lines.append("|---|---|---|")
            for item in sheet.unsupported:
                lines.append("| " + " | ".join([_esc(item.scope), _esc(item.location), _esc(item.tag)]) + " |")

    lines.append("")
    lines.append("## Extraction Summary")
    lines.append("")
    _append_key_value_table(lines, workbook.summary)

    lines.append("")
    lines.append("## Warnings")
    lines.append("")
    if not workbook.warnings:
        lines.append("(none)")
    else:
        for warning in workbook.warnings:
            lines.append(f"- {warning}")

    lines.append("")
    lines.append("## Mode")
    lines.append("")
    lines.append("- current_output: `work` (operator-friendly)")
    lines.append("- switch_to_full_dump: `excel-md INPUT.xlsx -o OUTPUT.md --full`")

    return "\n".join(lines).rstrip() + "\n"


def _append_region_workspace(lines: list[str], region_id: int, bounds_ref: str, region_rows) -> None:
    lines.append(f"#### Region {region_id}: {bounds_ref}")
    lines.append("")

    row_map: dict[int, list[tuple[str, str]]] = {}
    for row in region_rows:
        if not _row_is_interesting(row):
            continue
        col, row_num = _coord_parts(row.coord)
        display = _cell_display(row)
        row_map.setdefault(row_num, []).append((col, display))

    if not row_map:
        lines.append("- (no visible work cells)")
        lines.append("")
        return

    lines.append("```text")
    for row_num in sorted(row_map):
        cells = sorted(row_map[row_num], key=lambda x: _col_index(x[0]))
        payload = " | ".join(f"{col}={text}" for col, text in cells)
        lines.append(f"R{row_num} | {payload}")
    lines.append("```")
    lines.append("")


def _row_is_interesting(row) -> bool:
    if (row.value or "").strip():
        return True
    if row.formula:
        return True
    if "data_validation" in row.flags:
        return True
    if "merged" in row.flags and row.merge_ref and row.coord == row.merge_ref.split(":", 1)[0]:
        return True
    return False


def _cell_display(row) -> str:
    value = (row.value or "").strip()
    if not value and row.formula:
        value = f"={_short(row.formula, 30)}"
    elif row.formula:
        value = f"{_short(value, 40)} {{={_short(row.formula, 24)}}}"
    else:
        value = _short(value, 48)

    markers: list[str] = []
    if row.merge_ref and row.coord == row.merge_ref.split(":", 1)[0]:
        markers.append("merge")
    if "data_validation" in row.flags:
        markers.append("dv")

    if markers:
        return f"{value} [{'|'.join(markers)}]" if value else f"[{'|'.join(markers)}]"
    return value or "."


def _connector_examples(sheet: SheetDoc, limit: int) -> list[tuple[str, str, str, str]]:
    node_label: dict[str, str] = {}
    for obj in sheet.drawings:
        label = (obj.text or "").strip() or (obj.name or "").strip() or obj.object_uid
        node_label[obj.object_uid] = _short(label, 50)

    rows: list[tuple[str, str, str, str]] = []
    for conn in sheet.connectors:
        if not conn.resolved or not conn.source_uid or not conn.target_uid:
            continue
        src = node_label.get(conn.source_uid, conn.source_uid)
        dst = node_label.get(conn.target_uid, conn.target_uid)
        label = _short((conn.text or "").strip() or "(no label)", 40)
        rows.append((src, dst, label, conn.direction))
        if len(rows) >= limit:
            break
    return rows


def _representative_texts(sheet: SheetDoc, limit: int) -> list[str]:
    candidates: list[tuple[int, int, str]] = []
    for cell in sheet.cells:
        value = (cell.value or "").strip()
        if _is_informative_text(value):
            col, row_num = _coord_parts(cell.coord)
            candidates.append((row_num, _col_index(col), _short(value, 60)))

    candidates.sort(key=lambda x: (x[0], x[1]))
    seen: set[str] = set()
    result: list[str] = []
    for _, _, text in candidates:
        if text in seen:
            continue
        seen.add(text)
        result.append(text)
        if len(result) >= limit:
            break
    return result


def _is_informative_text(value: str) -> bool:
    if not value:
        return False
    v = value.strip()
    if not v:
        return False
    if _NUMBER_RE.fullmatch(v):
        return False
    if v in {"-", "○", "×", "TRUE", "FALSE"}:
        return False
    return True


def _infer_document_type(workbook: WorkbookDoc) -> str:
    name = workbook.source_path.name
    if "遷移" in name:
        return "screen transition workbook"
    if "設計" in name:
        return "screen design workbook"
    if "一覧" in name:
        return "catalog workbook"
    if workbook.summary.get("connector_count", 0) > 0:
        return "diagram workbook"
    return "general workbook"


def _infer_sheet_role(name: str) -> str:
    normalized = name.strip()
    if normalized in {"表紙", "cover", "Cover"}:
        return "cover"
    if "変更履歴" in normalized:
        return "change log"
    if "目次" in normalized:
        return "table of contents"
    if "データ" in normalized:
        return "master/reference data"
    if "遷移" in normalized:
        return "transition diagram"
    if "画面" in normalized:
        return "screen spec"
    return "work sheet"


def _coord_parts(coord: str) -> tuple[str, int]:
    match = _COORD_RE.match(coord)
    if not match:
        return coord, 0
    return match.group(1), int(match.group(2))


def _col_index(col: str) -> int:
    value = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            return 0
        value = value * 26 + (ord(ch) - 64)
    return value


def _short(value: str, limit: int) -> str:
    if len(value) <= limit:
        return value
    return value[: limit - 3] + "..."


def _render_full_markdown(workbook: WorkbookDoc) -> str:
    lines: list[str] = []

    lines.append(f"# Workbook: {workbook.source_path.name}")
    lines.append("")

    lines.append("## Source Metadata")
    lines.append("")
    _append_key_value_table(lines, workbook.source_metadata)

    lines.append("")
    lines.append("## Styles (XML-equivalent)")
    lines.append("")
    lines.append("```json")
    lines.append(json.dumps(workbook.styles_xml_equivalent, ensure_ascii=False, indent=2))
    lines.append("```")

    lines.append("")
    lines.append("## Defined Names")
    lines.append("")
    if not workbook.defined_names:
        lines.append("(none)")
    else:
        lines.append("| name | local_sheet_id | value |")
        lines.append("|---|---:|---|")
        for dn in workbook.defined_names:
            lines.append(
                "| "
                + " | ".join(
                    [
                        _esc(dn.name),
                        _esc("" if dn.local_sheet_id is None else str(dn.local_sheet_id)),
                        _esc(dn.value),
                    ]
                )
                + " |"
            )

    for sheet in workbook.sheets:
        lines.append("")
        lines.append(f"## Sheet: {sheet.name} [{sheet.state}]")
        lines.append("")

        lines.append("### Sheet Metadata")
        lines.append("")
        _append_key_value_table(
            lines,
            {
                "sheet_index": sheet.index,
                "path": sheet.path,
                "dimension_ref": sheet.dimension_ref,
                "cell_count": len(sheet.cells),
                "merge_count": len(sheet.merges),
                "data_validation_count": len(sheet.data_validations),
                "drawing_object_count": len(sheet.drawings),
                "connector_count": len(sheet.connectors),
                "region_count": len(sheet.regions),
                "unsupported_count": len(sheet.unsupported),
            },
        )

        lines.append("")
        lines.append("### Print Metadata")
        lines.append("")
        lines.append(f"- print_areas: {', '.join(r.ref for r in sheet.print_areas) if sheet.print_areas else '(none)'}")
        lines.append(f"- print_titles: {', '.join(sheet.print_titles) if sheet.print_titles else '(none)'}")
        lines.append(f"- page_setup: `{json.dumps(sheet.page_setup, ensure_ascii=False)}`")
        lines.append(f"- page_margins: `{json.dumps(sheet.page_margins, ensure_ascii=False)}`")
        lines.append(f"- print_options: `{json.dumps(sheet.print_options, ensure_ascii=False)}`")
        lines.append(f"- header_footer: `{json.dumps(sheet.header_footer, ensure_ascii=False)}`")
        lines.append(f"- page_breaks: `{json.dumps(sheet.page_breaks, ensure_ascii=False)}`")

        lines.append("")
        lines.append("### Data Validations")
        lines.append("")
        if not sheet.data_validations:
            lines.append("(none)")
        else:
            lines.append("| type | sqref | formula1 | formula2 | allow_blank | show_error_message | operator |")
            lines.append("|---|---|---|---|---|---|---|")
            for dv in sheet.data_validations:
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            _esc(dv.type or ""),
                            _esc(dv.sqref),
                            _esc(dv.formula1 or ""),
                            _esc(dv.formula2 or ""),
                            _esc("" if dv.allow_blank is None else str(dv.allow_blank)),
                            _esc("" if dv.show_error_message is None else str(dv.show_error_message)),
                            _esc(dv.operator or ""),
                        ]
                    )
                    + " |"
                )

        lines.append("")
        lines.append("### Cell Regions")
        lines.append("")
        if not sheet.regions:
            lines.append("(none)")
        else:
            for region in sheet.regions:
                lines.append(f"#### Region {region.region_id}: {region.bounds.ref}")
                lines.append("")
                lines.append("| coord | value | formula | cached_value | type | style_id | merge_ref | flags |")
                lines.append("|---|---|---|---|---|---|---|---|")
                for row in region.rows:
                    lines.append(
                        "| "
                        + " | ".join(
                            [
                                _esc(row.coord),
                                _esc(row.value),
                                _esc(row.formula or ""),
                                _esc(row.cached_value or ""),
                                _esc(row.cell_type),
                                _esc(row.style_id or ""),
                                _esc(row.merge_ref or ""),
                                _esc(",".join(row.flags)),
                            ]
                        )
                        + " |"
                    )
                lines.append("")

        lines.append("### Drawings Raw Objects")
        lines.append("")
        if not sheet.drawings:
            lines.append("(none)")
        else:
            lines.append("| object_uid | kind | name | text | anchor_from | anchor_to | bbox | parent_uid | image_target |")
            lines.append("|---|---|---|---|---|---|---|---|---|")
            for obj in sheet.drawings:
                from_repr = _anchor_repr(obj.anchor_from)
                to_repr = _anchor_repr(obj.anchor_to)
                bbox_repr = f"{obj.bbox[0]:.2f},{obj.bbox[1]:.2f},{obj.bbox[2]:.2f},{obj.bbox[3]:.2f}"
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            _esc(obj.object_uid),
                            _esc(obj.kind),
                            _esc(obj.name),
                            _esc(obj.text),
                            _esc(from_repr),
                            _esc(to_repr),
                            _esc(bbox_repr),
                            _esc(obj.parent_uid or ""),
                            _esc(obj.image_target or ""),
                        ]
                    )
                    + " |"
                )

        lines.append("")
        lines.append("### Connectors (Raw + Inferred)")
        lines.append("")
        if not sheet.connectors:
            lines.append("(none)")
        else:
            lines.append(
                "| object_uid | name | direction | source_uid | target_uid | resolved | "
                "distance_source | distance_target | arrow_head | arrow_tail | text |"
            )
            lines.append("|---|---|---|---|---|---|---:|---:|---|---|---|")
            for conn in sheet.connectors:
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            _esc(conn.object_uid),
                            _esc(conn.name),
                            _esc(conn.direction),
                            _esc(conn.source_uid or ""),
                            _esc(conn.target_uid or ""),
                            _esc(str(conn.resolved)),
                            _esc("" if conn.distance_source is None else f"{conn.distance_source:.2f}"),
                            _esc("" if conn.distance_target is None else f"{conn.distance_target:.2f}"),
                            _esc(conn.arrow_head or ""),
                            _esc(conn.arrow_tail or ""),
                            _esc(conn.text),
                        ]
                    )
                    + " |"
                )

        lines.append("")
        lines.append("### Mermaid")
        lines.append("")
        if sheet.mermaid:
            lines.append("```mermaid")
            lines.append(sheet.mermaid)
            lines.append("```")
        else:
            lines.append("(no resolved edges)")

        lines.append("")
        lines.append("### Embedded Images")
        lines.append("")
        images = [obj for obj in sheet.drawings if obj.image_data_uri]
        if not images:
            lines.append("(none)")
        else:
            for idx, img_obj in enumerate(images, start=1):
                lines.append(f"#### Image {idx}: {img_obj.object_uid}")
                lines.append("")
                lines.append(f"- target: `{img_obj.image_target}`")
                lines.append(f"- content_type: `{img_obj.image_content_type}`")
                lines.append(f"![{_esc(img_obj.name or img_obj.object_uid)}]({img_obj.image_data_uri})")
                lines.append("")

        lines.append("### Unsupported Elements")
        lines.append("")
        if not sheet.unsupported:
            lines.append("(none)")
        else:
            lines.append("| scope | location | tag |")
            lines.append("|---|---|---|")
            for item in sheet.unsupported:
                lines.append("| " + " | ".join([_esc(item.scope), _esc(item.location), _esc(item.tag)]) + " |")
            lines.append("")
            for idx, item in enumerate(sheet.unsupported, start=1):
                lines.append(f"#### Unsupported {idx}: {item.tag}")
                lines.append("")
                lines.append(f"- scope: `{item.scope}`")
                lines.append(f"- location: `{item.location}`")
                lines.append("```xml")
                lines.append(item.raw_xml)
                lines.append("```")
                lines.append("")

    lines.append("## Extraction Summary")
    lines.append("")
    _append_key_value_table(lines, workbook.summary)

    lines.append("")
    lines.append("## Warnings")
    lines.append("")
    if not workbook.warnings:
        lines.append("(none)")
    else:
        for warning in workbook.warnings:
            lines.append(f"- {warning}")

    return "\n".join(lines).rstrip() + "\n"


def _append_key_value_table(lines: list[str], payload: dict) -> None:
    lines.append("| key | value |")
    lines.append("|---|---|")
    for key, value in payload.items():
        lines.append(f"| {_esc(str(key))} | {_esc(_compact(value))} |")


def _compact(value) -> str:
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def _esc(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", "<br>")


def _anchor_repr(anchor) -> str:
    if anchor is None:
        return ""
    return f"({anchor.col},{anchor.row},{anchor.col_off},{anchor.row_off})"
