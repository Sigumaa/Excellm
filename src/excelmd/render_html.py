from __future__ import annotations

import json
from html import escape as html_escape

from .model import AnchorPoint, SheetDoc, WorkbookDoc
from .parser.utils import index_to_col, parse_range_ref, rowcol_to_coord


def render_workbook_html(workbook: WorkbookDoc) -> str:
    parts: list[str] = []

    parts.append("<!doctype html>")
    parts.append('<html lang="ja">')
    parts.append("<head>")
    parts.append('<meta charset="utf-8">')
    parts.append('<meta name="viewport" content="width=device-width, initial-scale=1">')
    parts.append(f"<title>{html_escape(workbook.source_path.name)} - SheetView HTML</title>")
    parts.append(_html_css())
    parts.append("</head>")
    parts.append("<body>")

    parts.append('<main class="page">')
    parts.append(f"<h1>Workbook: {html_escape(workbook.source_path.name)}</h1>")

    parts.append("<section>")
    parts.append("<h2>Source Metadata</h2>")
    parts.append(_kv_table(workbook.source_metadata))
    parts.append("</section>")

    parts.append("<section>")
    parts.append("<h2>Extraction Summary</h2>")
    parts.append(_kv_table(workbook.summary))
    parts.append("</section>")

    for sheet in workbook.sheets:
        parts.append('<section class="sheet">')
        parts.append(f"<h2>Sheet: {html_escape(sheet.name)} [{html_escape(sheet.state)}]</h2>")
        print_areas = ", ".join(r.ref for r in sheet.print_areas) if sheet.print_areas else "(none)"
        parts.append(
            "<p class=\"meta\">"
            f"used_range=<code>{html_escape(sheet.dimension_ref)}</code> / "
            f"print_areas=<code>{html_escape(print_areas)}</code> / "
            f"hidden_rows=<code>{len(sheet.hidden_rows)}</code> / "
            f"hidden_cols=<code>{len(sheet.hidden_cols)}</code>"
            "</p>"
        )
        if sheet.pane:
            parts.append(f"<p class=\"meta\">pane=<code>{html_escape(json.dumps(sheet.pane, ensure_ascii=False))}</code></p>")

        ranges = _sheetview_ranges(sheet)
        for idx, rng in enumerate(ranges, start=1):
            parts.append(f"<h3>Range {idx}: {html_escape(rng.ref)}</h3>")
            parts.append(_render_sheet_range_html(sheet, rng.ref, workbook.style_css_map))

        if not ranges:
            parts.append('<p class="empty">No renderable range.</p>')

        if sheet.unsupported:
            parts.append("<details>")
            parts.append(f"<summary>Unsupported Elements ({len(sheet.unsupported)})</summary>")
            parts.append('<table class="simple">')
            parts.append("<thead><tr><th>scope</th><th>location</th><th>tag</th></tr></thead><tbody>")
            for item in sheet.unsupported:
                parts.append(
                    "<tr>"
                    f"<td>{html_escape(item.scope)}</td>"
                    f"<td>{html_escape(item.location)}</td>"
                    f"<td>{html_escape(item.tag)}</td>"
                    "</tr>"
                )
            parts.append("</tbody></table>")
            parts.append("</details>")

        parts.append("</section>")

    parts.append("<section>")
    parts.append("<h2>Warnings</h2>")
    if not workbook.warnings:
        parts.append("<p>(none)</p>")
    else:
        parts.append("<ul>")
        for warning in workbook.warnings:
            parts.append(f"<li>{html_escape(warning)}</li>")
        parts.append("</ul>")
    parts.append("</section>")

    parts.append("</main>")
    parts.append("</body>")
    parts.append("</html>")

    return "\n".join(parts) + "\n"


def _sheetview_ranges(sheet: SheetDoc):
    if sheet.print_areas:
        return sheet.print_areas
    try:
        return [parse_range_ref(sheet.dimension_ref)]
    except ValueError:
        return []


def _render_sheet_range_html(sheet: SheetDoc, range_ref: str, style_css_map: dict[str, str]) -> str:
    rng = parse_range_ref(range_ref)
    visible_rows = [r for r in range(rng.start_row, rng.end_row + 1) if r not in sheet.hidden_rows]
    visible_cols = [c for c in range(rng.start_col, rng.end_col + 1) if c not in sheet.hidden_cols]

    if not visible_rows or not visible_cols:
        return '<p class="empty">All rows/cols in this range are hidden.</p>'

    merge_anchor: dict[str, tuple[int, int, str]] = {}
    merge_covered: set[str] = set()
    for m in sheet.merges:
        if m.end_row < rng.start_row or m.start_row > rng.end_row:
            continue
        if m.end_col < rng.start_col or m.start_col > rng.end_col:
            continue

        m_rows = [r for r in visible_rows if m.start_row <= r <= m.end_row]
        m_cols = [c for c in visible_cols if m.start_col <= c <= m.end_col]
        if not m_rows or not m_cols:
            continue

        anchor_coord = rowcol_to_coord(m_rows[0], m_cols[0])
        merge_anchor[anchor_coord] = (len(m_rows), len(m_cols), m.ref)
        for rr in m_rows:
            for cc in m_cols:
                coord = rowcol_to_coord(rr, cc)
                if coord != anchor_coord:
                    merge_covered.add(coord)

    col_px: dict[int, float] = {col: _col_width_to_px(sheet.col_widths.get(col)) for col in visible_cols}
    row_px: dict[int, float] = {row: _row_height_to_px(sheet.row_heights.get(row)) for row in visible_rows}

    total_w = int(sum(col_px.values()))
    total_h = int(sum(row_px.values()))

    out: list[str] = []
    out.append('<div class="sv-wrap">')
    out.append('<div class="sv-canvas">')
    out.append('<table class="sv-grid">')
    out.append("<colgroup>")
    out.append('<col style="width:56px">')
    for col in visible_cols:
        out.append(f'<col style="width:{col_px[col]:.1f}px">')
    out.append("</colgroup>")
    out.append("<thead>")
    out.append('<tr class="sv-head-row">')
    out.append('<th class="sv-corner"></th>')
    for col in visible_cols:
        out.append(f'<th class="sv-col-head">{index_to_col(col)}</th>')
    out.append("</tr>")
    out.append("</thead>")
    out.append("<tbody>")

    for row in visible_rows:
        out.append(f'<tr style="height:{row_px[row]:.1f}px">')
        out.append(f'<th class="sv-row-head">{row}</th>')
        for col in visible_cols:
            coord = rowcol_to_coord(row, col)
            if coord in merge_covered:
                continue

            cell = sheet.cell_map.get(coord)
            style_id = cell.style_id if cell and cell.style_id is not None else "0"
            style_css = style_css_map.get(style_id, "")
            if style_css and not style_css.strip().endswith(";"):
                style_css += ";"

            attrs: list[str] = []
            if coord in merge_anchor:
                rowspan, colspan, merge_ref = merge_anchor[coord]
                attrs.append(f'rowspan="{rowspan}"')
                attrs.append(f'colspan="{colspan}"')
                attrs.append(f'data-merge="{html_escape(merge_ref)}"')

            text_html = _cell_html(
                cell.display_value if cell else "",
                cell.formula if cell else None,
            )
            classes = ["sv-cell"]
            if not text_html.strip():
                classes.append("sv-empty")
            attrs.append(f'class="{" ".join(classes)}"')
            attrs.append(f'data-coord="{coord}"')
            if style_css:
                attrs.append(f'style="{html_escape(style_css)}"')

            out.append(f"<td {' '.join(attrs)}>{text_html}</td>")
        out.append("</tr>")

    out.append("</tbody>")
    out.append("</table>")
    out.extend(_overlay_html(sheet, rng.start_col, rng.start_row, total_w, total_h))
    out.append("</div>")
    out.append("</div>")
    return "\n".join(out)


def _overlay_html(sheet: SheetDoc, start_col: int, start_row: int, total_w: int, total_h: int) -> list[str]:
    origin_x = (start_col - 1) * 64.0
    origin_y = (start_row - 1) * 20.0

    def intersects(bbox: tuple[float, float, float, float]) -> bool:
        x1, y1, x2, y2 = bbox
        rx1, ry1, rx2, ry2 = origin_x, origin_y, origin_x + total_w, origin_y + total_h
        return not (x2 < rx1 or x1 > rx2 or y2 < ry1 or y1 > ry2)

    drawing_map = {obj.object_uid: obj for obj in sheet.drawings}

    lines: list[str] = []
    lines.append('<div class="sv-overlay">')

    z = 10
    for obj in sheet.drawings:
        if obj.kind == "cxnSp" or not intersects(obj.bbox):
            continue

        x1, y1, x2, y2 = obj.bbox
        left = max(0.0, x1 - origin_x)
        top = max(0.0, y1 - origin_y)
        width = max(8.0, x2 - x1)
        height = max(8.0, y2 - y1)

        label = (obj.text or obj.name or obj.object_id).strip()
        safe_label = html_escape(label)

        classes = ["sv-shape"]
        if obj.kind == "pic":
            classes.append("pic")

        shape_style = _shape_style_css(obj.extra, z)
        z += 1

        if obj.kind == "pic" and obj.image_data_uri:
            body = (
                f'<img src="{obj.image_data_uri}" alt="{safe_label}" '
                'style="width:100%;height:100%;object-fit:contain;">'
            )
        else:
            body = safe_label or "&nbsp;"

        lines.append(
            f'<div class="{" ".join(classes)}" style="left:{left:.1f}px;top:{top:.1f}px;'
            f'width:{width:.1f}px;height:{height:.1f}px;{shape_style}">{body}</div>'
        )

    lines.append(f'<svg class="sv-lines" width="{total_w}" height="{total_h}" viewBox="0 0 {total_w} {total_h}">')
    lines.append("<defs>")
    lines.append('<marker id="arrow-triangle" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">')
    lines.append('<path d="M 0 0 L 10 5 L 0 10 z" fill="context-stroke" />')
    lines.append("</marker>")
    lines.append("</defs>")

    for conn in sheet.connectors:
        x1, y1 = _connector_point(conn.anchor_from, conn.bbox, True)
        x2, y2 = _connector_point(conn.anchor_to, conn.bbox, False)
        x1 -= origin_x
        x2 -= origin_x
        y1 -= origin_y
        y2 -= origin_y
        if max(x1, x2) < 0 or max(y1, y2) < 0 or min(x1, x2) > total_w or min(y1, y2) > total_h:
            continue

        source_obj = drawing_map.get(conn.object_uid)
        extra = source_obj.extra if source_obj else {}
        stroke = extra.get("line_color", "#ef4444")
        stroke_width = extra.get("line_width_px", "1.2")
        dash = extra.get("line_dash", "")
        dash_css = _dasharray_for(dash)
        marker_start = ' marker-start="url(#arrow-triangle)"' if _has_arrow(conn.arrow_head) else ""
        marker_end = ' marker-end="url(#arrow-triangle)"' if _has_arrow(conn.arrow_tail) else ""
        dash_attr = f' stroke-dasharray="{dash_css}"' if dash_css else ""

        lines.append(
            f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" '
            f'stroke="{stroke}" stroke-width="{stroke_width}" stroke-opacity="0.9"{dash_attr}{marker_start}{marker_end} />'
        )
    lines.append("</svg>")

    lines.append("</div>")
    return lines


def _shape_style_css(extra: dict[str, str], z_index: int) -> str:
    line = extra.get("line_color", "#fb7185")
    fill = extra.get("fill_color", "rgba(251,113,133,0.08)")
    width = extra.get("line_width_px", "1.0")
    dash = extra.get("line_dash")

    pieces = [
        f"z-index:{z_index}",
        f"border-color:{line}",
        f"border-width:{width}px",
        f"background:{_to_alpha(fill, 0.12)}",
    ]
    if dash:
        pieces.append(f"border-style:{_border_style_for_dash(dash)}")
    return ";".join(pieces) + ";"


def _to_alpha(color: str, alpha: float) -> str:
    c = color.strip()
    if c.startswith("#") and len(c) == 7:
        try:
            r = int(c[1:3], 16)
            g = int(c[3:5], 16)
            b = int(c[5:7], 16)
            return f"rgba({r},{g},{b},{alpha:.2f})"
        except ValueError:
            return c
    return c


def _border_style_for_dash(dash: str) -> str:
    mapping = {
        "dash": "dashed",
        "dot": "dotted",
        "dashDot": "dashed",
        "lgDash": "dashed",
    }
    return mapping.get(dash, "solid")


def _dasharray_for(dash: str) -> str:
    mapping = {
        "dash": "6 4",
        "dot": "2 3",
        "dashDot": "8 3 2 3",
        "lgDash": "10 4",
        "sysDot": "2 3",
        "sysDash": "6 4",
    }
    return mapping.get(dash, "")


def _has_arrow(value: str | None) -> bool:
    return bool(value and value.lower() != "none")


def _connector_point(anchor: AnchorPoint | None, bbox: tuple[float, float, float, float], is_start: bool) -> tuple[float, float]:
    if anchor is not None:
        return anchor.col * 64.0 + anchor.col_off / 9525.0, anchor.row * 20.0 + anchor.row_off / 9525.0
    return (bbox[0], bbox[1]) if is_start else (bbox[2], bbox[3])


def _cell_html(value: str, formula: str | None) -> str:
    safe_value = html_escape(value or "")
    if formula:
        safe_formula = html_escape(formula)
        if safe_value:
            return f'{safe_value}<span class="sv-formula">={safe_formula}</span>'
        return f'<span class="sv-formula">={safe_formula}</span>'
    return safe_value


def _col_width_to_px(width: float | None) -> float:
    if width is None:
        return 64.0
    return max(20.0, width * 7.0 + 5.0)


def _row_height_to_px(height: float | None) -> float:
    if height is None:
        return 20.0
    return max(12.0, height * (96.0 / 72.0))


def _kv_table(payload: dict) -> str:
    lines: list[str] = []
    lines.append('<table class="simple">')
    lines.append("<thead><tr><th>key</th><th>value</th></tr></thead><tbody>")
    for key, value in payload.items():
        if isinstance(value, (dict, list, tuple)):
            value_text = json.dumps(value, ensure_ascii=False)
        else:
            value_text = str(value)
        lines.append(f"<tr><td>{html_escape(str(key))}</td><td>{html_escape(value_text)}</td></tr>")
    lines.append("</tbody></table>")
    return "\n".join(lines)


def _html_css() -> str:
    return """<style>
:root {
  --line: #d0d7de;
  --line-head: #bcc6d4;
  --bg-soft: #f6f8fa;
  --bg-head: #edf2f7;
  --text: #111827;
  --shape: #fb7185;
  --shape-bg: rgba(251, 113, 133, 0.08);
  --pic: #3b82f6;
  --pic-bg: rgba(59, 130, 246, 0.06);
}
* { box-sizing: border-box; }
body { margin: 0; background: #fff; color: var(--text); font: 14px/1.5 -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
.page { max-width: 98vw; margin: 0 auto; padding: 18px 16px 80px; }
h1 { margin: 0 0 12px; font-size: 24px; }
h2 { margin: 28px 0 10px; font-size: 18px; }
h3 { margin: 22px 0 8px; font-size: 14px; }
.meta { margin: 0 0 10px; color: #4b5563; }
.empty { color: #6b7280; font-style: italic; }
.simple { border-collapse: collapse; width: 100%; max-width: 1200px; }
.simple th, .simple td { border: 1px solid var(--line); padding: 6px 8px; font-size: 12px; vertical-align: top; }
.simple th { background: var(--bg-soft); text-align: left; }
.sv-wrap { margin: 12px 0 28px; border: 1px solid var(--line); border-radius: 8px; overflow: auto; background: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
.sv-canvas { position: relative; display: inline-block; }
.sv-grid { border-collapse: collapse; font: 11px/1.25 'Yu Gothic UI', 'Meiryo', sans-serif; table-layout: fixed; background: #fff; }
.sv-grid th, .sv-grid td { border: 1px solid var(--line); }
.sv-grid .sv-head-row th,
.sv-grid .sv-row-head,
.sv-grid .sv-corner,
.sv-grid .sv-col-head { background: var(--bg-head); border-color: var(--line-head); color: #374151; font: 11px/1.2 'SF Mono', Menlo, Consolas, monospace; text-align: center; }
.sv-grid .sv-col-head { height: 24px; position: sticky; top: 0; z-index: 6; }
.sv-grid .sv-corner { width: 56px; min-width: 56px; position: sticky; top: 0; left: 0; z-index: 7; }
.sv-grid .sv-row-head { width: 56px; min-width: 56px; position: sticky; left: 0; z-index: 5; }
.sv-grid td { padding: 2px 4px; overflow: hidden; vertical-align: top; white-space: pre-wrap; }
.sv-grid .sv-empty { color: transparent; }
.sv-formula { display: block; margin-top: 2px; color: #6b7280; font-size: 10px; }
.sv-overlay { position: absolute; left: 56px; top: 24px; right: 0; bottom: 0; pointer-events: none; }
.sv-shape { position: absolute; border: 1px solid var(--shape); background: var(--shape-bg); color: var(--text); font: 10px/1.2 sans-serif; padding: 2px; overflow: hidden; }
.sv-shape.pic { border-color: var(--pic); background: var(--pic-bg); }
.sv-lines { position: absolute; inset: 0; overflow: visible; }
</style>"""
