from __future__ import annotations

import json

from .model import WorkbookDoc


def render_workbook_markdown(workbook: WorkbookDoc) -> str:
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
