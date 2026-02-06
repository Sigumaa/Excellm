from __future__ import annotations

import json
import re
from dataclasses import dataclass
from html import escape as html_escape

from .model import AnchorPoint, SheetDoc, WorkbookDoc
from .parser.utils import index_to_col, parse_range_ref, rowcol_to_coord

EMU_PER_PIXEL = 9525.0
ROW_HEADER_WIDTH = 56.0
COL_HEADER_HEIGHT = 24.0


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
            parts.append(_render_sheet_range_html(sheet, rng.ref, workbook.style_css_map, idx))

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


@dataclass(slots=True)
class _SheetGeometry:
    sheet: SheetDoc
    start_col: int
    start_row: int
    visible_cols: list[int]
    visible_rows: list[int]
    total_width: float
    total_height: float

    @classmethod
    def build(cls, sheet: SheetDoc, start_col: int, start_row: int, end_col: int, end_row: int) -> "_SheetGeometry":
        visible_cols = [c for c in range(start_col, end_col + 1) if c not in sheet.hidden_cols]
        visible_rows = [r for r in range(start_row, end_row + 1) if r not in sheet.hidden_rows]

        total_w = sum(_col_width_to_px(sheet.col_widths.get(c)) for c in visible_cols)
        total_h = sum(_row_height_to_px(sheet.row_heights.get(r)) for r in visible_rows)
        return cls(
            sheet=sheet,
            start_col=start_col,
            start_row=start_row,
            visible_cols=visible_cols,
            visible_rows=visible_rows,
            total_width=total_w,
            total_height=total_h,
        )

    def col_width(self, col: int) -> float:
        return _col_width_to_px(self.sheet.col_widths.get(col))

    def row_height(self, row: int) -> float:
        return _row_height_to_px(self.sheet.row_heights.get(row))

    def x_at_col(self, col: int, col_off: int = 0) -> float:
        x = 0.0
        if col >= self.start_col:
            for c in range(self.start_col, col):
                if c in self.sheet.hidden_cols:
                    continue
                x += self.col_width(c)
        else:
            for c in range(col, self.start_col):
                if c in self.sheet.hidden_cols:
                    continue
                x -= self.col_width(c)
        x += col_off / EMU_PER_PIXEL
        return x

    def y_at_row(self, row: int, row_off: int = 0) -> float:
        y = 0.0
        if row >= self.start_row:
            for r in range(self.start_row, row):
                if r in self.sheet.hidden_rows:
                    continue
                y += self.row_height(r)
        else:
            for r in range(row, self.start_row):
                if r in self.sheet.hidden_rows:
                    continue
                y -= self.row_height(r)
        y += row_off / EMU_PER_PIXEL
        return y

    def point(self, anchor: AnchorPoint) -> tuple[float, float]:
        return self.x_at_col(anchor.col, anchor.col_off), self.y_at_row(anchor.row, anchor.row_off)

    def freeze_lines(self) -> tuple[float | None, float | None]:
        if not self.sheet.pane:
            return None, None

        def _read_int(key: str) -> int:
            raw = self.sheet.pane.get(key)
            if raw is None:
                return 0
            try:
                return int(float(raw))
            except ValueError:
                return 0

        x_split = _read_int("xSplit")
        y_split = _read_int("ySplit")

        freeze_x = None
        freeze_y = None
        if x_split > 0:
            freeze_col = self.start_col + x_split
            freeze_x = self.x_at_col(freeze_col)
        if y_split > 0:
            freeze_row = self.start_row + y_split
            freeze_y = self.y_at_row(freeze_row)
        return freeze_x, freeze_y


def _render_sheet_range_html(sheet: SheetDoc, range_ref: str, style_css_map: dict[str, str], idx: int) -> str:
    rng = parse_range_ref(range_ref)
    geom = _SheetGeometry.build(sheet, rng.start_col, rng.start_row, rng.end_col, rng.end_row)

    if not geom.visible_rows or not geom.visible_cols:
        return '<p class="empty">All rows/cols in this range are hidden.</p>'

    merge_anchor: dict[str, tuple[int, int, str]] = {}
    merge_covered: set[str] = set()
    for m in sheet.merges:
        if m.end_row < rng.start_row or m.start_row > rng.end_row:
            continue
        if m.end_col < rng.start_col or m.start_col > rng.end_col:
            continue

        m_rows = [r for r in geom.visible_rows if m.start_row <= r <= m.end_row]
        m_cols = [c for c in geom.visible_cols if m.start_col <= c <= m.end_col]
        if not m_rows or not m_cols:
            continue

        anchor_coord = rowcol_to_coord(m_rows[0], m_cols[0])
        merge_anchor[anchor_coord] = (len(m_rows), len(m_cols), m.ref)
        for rr in m_rows:
            for cc in m_cols:
                coord = rowcol_to_coord(rr, cc)
                if coord != anchor_coord:
                    merge_covered.add(coord)

    out: list[str] = []
    out.append(f'<div class="sv-range" data-range-id="{idx}">')
    pad_top, pad_right, pad_bottom, pad_left = _sheet_padding_px(sheet)
    out.append('<div class="sv-wrap">')
    out.append(
        '<div class="sv-viewport" '
        f'style="padding:{pad_top:.1f}px {pad_right:.1f}px {pad_bottom:.1f}px {pad_left:.1f}px;">'
    )
    out.append('<div class="sv-canvas">')
    out.append('<table class="sv-grid">')
    out.append("<colgroup>")
    out.append(f'<col style="width:{ROW_HEADER_WIDTH}px">')
    for col in geom.visible_cols:
        out.append(f'<col style="width:{geom.col_width(col):.1f}px">')
    out.append("</colgroup>")
    out.append("<thead>")
    out.append('<tr class="sv-head-row">')
    out.append('<th class="sv-corner"></th>')
    for col in geom.visible_cols:
        out.append(f'<th class="sv-col-head">{index_to_col(col)}</th>')
    out.append("</tr>")
    out.append("</thead>")
    out.append("<tbody>")

    for row in geom.visible_rows:
        out.append(f'<tr style="height:{geom.row_height(row):.1f}px">')
        out.append(f'<th class="sv-row-head">{row}</th>')
        for col in geom.visible_cols:
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

            text_html = _cell_html(cell.display_value if cell else "")
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
    out.extend(_overlay_html(sheet, geom))
    out.append("</div>")
    out.append("</div>")
    hf_html = _header_footer_html(sheet)
    if hf_html:
        out.append(hf_html)
    out.append("</div>")
    out.append("</div>")
    return "\n".join(out)


def _overlay_html(sheet: SheetDoc, geom: _SheetGeometry) -> list[str]:
    drawing_map = {obj.object_uid: obj for obj in sheet.drawings}

    lines: list[str] = []
    lines.append('<div class="sv-overlay">')

    z = 10
    for obj in sheet.drawings:
        if obj.kind == "cxnSp":
            continue

        rect = _shape_rect(obj, geom)
        if rect is None:
            continue
        left, top, width, height = rect
        if left + width < 0 or top + height < 0:
            continue
        if left > geom.total_width or top > geom.total_height:
            continue

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

    lines.append(
        f'<svg class="sv-lines" width="{geom.total_width:.1f}" height="{geom.total_height:.1f}" '
        f'viewBox="0 0 {geom.total_width:.1f} {geom.total_height:.1f}">'
    )
    lines.append("<defs>")
    lines.append('<marker id="arrow-triangle" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">')
    lines.append('<path d="M 0 0 L 10 5 L 0 10 z" fill="context-stroke" />')
    lines.append("</marker>")
    lines.append("</defs>")

    for conn in sheet.connectors:
        p1, p2 = _connector_points(conn.anchor_from, conn.anchor_to, conn.bbox, geom)
        x1, y1 = p1
        x2, y2 = p2

        if max(x1, x2) < 0 or max(y1, y2) < 0 or min(x1, x2) > geom.total_width or min(y1, y2) > geom.total_height:
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
            f'stroke="{stroke}" stroke-width="{stroke_width}" stroke-opacity="0.92"{dash_attr}{marker_start}{marker_end} />'
        )

        if conn.text:
            mx = (x1 + x2) / 2
            my = (y1 + y2) / 2
            label = html_escape(conn.text.strip())
            lines.append(f'<text x="{mx:.1f}" y="{my - 2:.1f}" class="sv-line-label">{label}</text>')

    freeze_x, freeze_y = geom.freeze_lines()
    if freeze_x is not None:
        lines.append(f'<line class="sv-freeze" x1="{freeze_x:.1f}" y1="0" x2="{freeze_x:.1f}" y2="{geom.total_height:.1f}" />')
    if freeze_y is not None:
        lines.append(f'<line class="sv-freeze" x1="0" y1="{freeze_y:.1f}" x2="{geom.total_width:.1f}" y2="{freeze_y:.1f}" />')

    for row_break in sheet.page_breaks.get("row", []):
        y = geom.y_at_row(row_break + 1)
        if 0 <= y <= geom.total_height:
            lines.append(f'<line class="sv-page-break" x1="0" y1="{y:.1f}" x2="{geom.total_width:.1f}" y2="{y:.1f}" />')
    for col_break in sheet.page_breaks.get("col", []):
        x = geom.x_at_col(col_break + 1)
        if 0 <= x <= geom.total_width:
            lines.append(f'<line class="sv-page-break" x1="{x:.1f}" y1="0" x2="{x:.1f}" y2="{geom.total_height:.1f}" />')

    lines.append("</svg>")
    lines.append("</div>")
    return lines


def _shape_rect(obj, geom: _SheetGeometry) -> tuple[float, float, float, float] | None:
    if obj.anchor_from is not None and obj.anchor_to is not None:
        x1, y1 = geom.point(obj.anchor_from)
        x2, y2 = geom.point(obj.anchor_to)
        left = min(x1, x2)
        top = min(y1, y2)
        width = max(8.0, abs(x2 - x1))
        height = max(8.0, abs(y2 - y1))
        return left, top, width, height

    x1, y1, x2, y2 = obj.bbox
    width = max(8.0, x2 - x1)
    height = max(8.0, y2 - y1)
    return x1, y1, width, height


def _connector_points(
    from_anchor: AnchorPoint | None,
    to_anchor: AnchorPoint | None,
    bbox: tuple[float, float, float, float],
    geom: _SheetGeometry,
) -> tuple[tuple[float, float], tuple[float, float]]:
    if from_anchor is not None:
        p1 = geom.point(from_anchor)
    else:
        p1 = (bbox[0], bbox[1])

    if to_anchor is not None:
        p2 = geom.point(to_anchor)
    else:
        p2 = (bbox[2], bbox[3])

    return p1, p2


def _shape_style_css(extra: dict[str, str], z_index: int) -> str:
    line = (extra.get("line_color") or "").strip()
    fill = (extra.get("fill_color") or "").strip()
    width_raw = (extra.get("line_width_px") or "1.0").strip()
    dash = extra.get("line_dash")
    try:
        width = max(0.5, float(width_raw))
    except ValueError:
        width = 1.0

    pieces = [f"z-index:{z_index}"]
    if line:
        pieces.append(f"border-color:{line}")
        pieces.append(f"border-width:{width:.2f}px")
        pieces.append(f"border-style:{_border_style_for_dash(dash)}" if dash else "border-style:solid")
    else:
        pieces.append("border:none")

    if fill:
        pieces.append(f"background:{_to_alpha(fill, 0.12)}")
    else:
        pieces.append("background:transparent")

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


def _cell_html(value: str) -> str:
    return html_escape(value or "")


def _sheet_padding_px(sheet: SheetDoc) -> tuple[float, float, float, float]:
    margins = sheet.page_margins
    return (
        _margin_to_px(margins.get("top"), 8.0),
        _margin_to_px(margins.get("right"), 8.0),
        _margin_to_px(margins.get("bottom"), 8.0),
        _margin_to_px(margins.get("left"), 8.0),
    )


def _margin_to_px(value: str | None, default_px: float) -> float:
    if not value:
        return default_px
    try:
        return max(0.0, float(value) * 96.0)
    except ValueError:
        return default_px


def _header_footer_html(sheet: SheetDoc) -> str:
    header_raw = sheet.header_footer.get("oddHeader", "")
    footer_raw = sheet.header_footer.get("oddFooter", "")
    if not header_raw and not footer_raw:
        return ""

    rows: list[str] = ['<div class="sv-hf">']
    if header_raw:
        rows.append(_hf_row_html("header", _decode_header_footer(header_raw, sheet.name)))
    if footer_raw:
        rows.append(_hf_row_html("footer", _decode_header_footer(footer_raw, sheet.name)))
    rows.append("</div>")
    return "\n".join(rows)


def _hf_row_html(kind: str, sections: dict[str, str]) -> str:
    return (
        f'<div class="sv-hf-row sv-hf-{kind}">'
        f'<span class="sv-hf-left">{html_escape(sections.get("L", ""))}</span>'
        f'<span class="sv-hf-center">{html_escape(sections.get("C", ""))}</span>'
        f'<span class="sv-hf-right">{html_escape(sections.get("R", ""))}</span>'
        "</div>"
    )


def _decode_header_footer(raw: str, sheet_name: str) -> dict[str, str]:
    sections = _split_hf_sections(raw)
    return {
        side: _clean_hf_text(text, sheet_name).strip()
        for side, text in sections.items()
    }


def _split_hf_sections(raw: str) -> dict[str, str]:
    sections = {"L": "", "C": "", "R": ""}
    current = "C"
    i = 0
    while i < len(raw):
        if raw[i] == "&" and i + 1 < len(raw) and raw[i + 1] in {"L", "C", "R"}:
            current = raw[i + 1]
            i += 2
            continue
        sections[current] += raw[i]
        i += 1
    return sections


def _clean_hf_text(text: str, sheet_name: str) -> str:
    cleaned = text.replace("&&", "&")
    token_map = {
        "&P": "{page}",
        "&N": "{pages}",
        "&D": "{date}",
        "&T": "{time}",
        "&A": sheet_name,
        "&F": "{file}",
        "&Z": "{path}",
        "&G": "{image}",
    }
    for token, replacement in token_map.items():
        cleaned = cleaned.replace(token, replacement)

    cleaned = re.sub(r'&"[^"]*"', "", cleaned)
    cleaned = re.sub(r"&K[0-9A-Fa-f]{6}", "", cleaned)
    cleaned = re.sub(r"&[0-9]+", "", cleaned)
    cleaned = re.sub(r"&[BIESUXY]", "", cleaned)
    cleaned = cleaned.replace("\r", " ").replace("\n", " ")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned


def _col_width_to_px(width: float | None) -> float:
    if width is None:
        return 64.0
    # Match Excel-ish conversion more closely
    px = int(((256 * width + int(128 / 7)) / 256) * 7)
    return max(20.0, float(px))


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
.page { max-width: 99vw; margin: 0 auto; padding: 18px 14px 80px; }
h1 { margin: 0 0 12px; font-size: 24px; }
h2 { margin: 28px 0 10px; font-size: 18px; }
h3 { margin: 22px 0 8px; font-size: 14px; }
.meta { margin: 0 0 10px; color: #4b5563; }
.empty { color: #6b7280; font-style: italic; }
.simple { border-collapse: collapse; width: 100%; max-width: 1400px; }
.simple th, .simple td { border: 1px solid var(--line); padding: 6px 8px; font-size: 12px; vertical-align: top; }
.simple th { background: var(--bg-soft); text-align: left; }
.sv-range { margin: 8px 0 22px; }
.sv-wrap { border: 1px solid var(--line); border-radius: 8px; overflow: auto; background: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
.sv-viewport { padding: 8px; min-width: fit-content; }
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
.sv-grid td { padding: 2px 4px; overflow: hidden; vertical-align: top; white-space: pre-wrap; background: #fff; }
.sv-grid .sv-empty { color: transparent; }
.sv-overlay { position: absolute; left: 56px; top: 24px; right: 0; bottom: 0; pointer-events: none; }
.sv-shape { position: absolute; border: none; background: transparent; color: var(--text); font: 10px/1.2 sans-serif; padding: 0; overflow: hidden; }
.sv-shape.pic { border: none; background: transparent; }
.sv-lines { position: absolute; inset: 0; overflow: visible; }
.sv-line-label { font: 10px/1.1 sans-serif; fill: #1f2937; paint-order: stroke; stroke: #fff; stroke-width: 2px; }
.sv-freeze { stroke: #2563eb; stroke-width: 1.3; stroke-dasharray: 5 3; opacity: 0.9; }
.sv-page-break { stroke: #2563eb; stroke-width: 1; stroke-dasharray: 3 3; opacity: 0.75; }
.sv-hf { border-top: 1px dashed #cbd5e1; background: #f8fafc; padding: 6px 10px; }
.sv-hf-row { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 6px; font: 11px/1.3 'SF Mono', Menlo, Consolas, monospace; color: #334155; }
.sv-hf-header { margin-bottom: 4px; }
.sv-hf-left { text-align: left; }
.sv-hf-center { text-align: center; }
.sv-hf-right { text-align: right; }
</style>"""
