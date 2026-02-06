from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from ..model import (
    CellData,
    ConvertOptions,
    DataValidation,
    DefinedName,
    SheetDoc,
    UnsupportedElement,
    WorkbookDoc,
)
from .drawing import parse_drawing_for_sheet
from .namespaces import NS, PACKAGE_REL_NS, SPREADSHEET_NS
from .regions import build_sheet_regions
from .utils import (
    coord_to_rowcol,
    iter_cells_in_range,
    local_name,
    parse_range_ref,
    parse_sheet_scoped_range,
    resolve_target,
    xml_to_dict,
)


@dataclass(slots=True)
class _SheetRef:
    index: int
    rid: str
    name: str
    state: str
    path: str


class OOXMLWorkbookParser:
    def __init__(self, source_path: str | Path, options: ConvertOptions) -> None:
        self.source_path = Path(source_path)
        self.options = options
        self._theme_colors: dict[int, str] = {}
        self._style_numfmt_map: dict[str, str] = {}

        self._builtin_numfmts: dict[int, str] = {
            0: "General",
            1: "0",
            2: "0.00",
            3: "#,##0",
            4: "#,##0.00",
            9: "0%",
            10: "0.00%",
            11: "0.00E+00",
            12: "# ?/?",
            13: "# ??/??",
            14: "m/d/yyyy",
            15: "d-mmm-yy",
            16: "d-mmm",
            17: "mmm-yy",
            18: "h:mm AM/PM",
            19: "h:mm:ss AM/PM",
            20: "h:mm",
            21: "h:mm:ss",
            22: "m/d/yyyy h:mm",
            37: "#,##0 ;(#,##0)",
            38: "#,##0 ;[Red](#,##0)",
            39: "#,##0.00;(#,##0.00)",
            40: "#,##0.00;[Red](#,##0.00)",
            45: "mm:ss",
            46: "[h]:mm:ss",
            47: "mmss.0",
            48: "##0.0E+0",
            49: "@",
        }
        self._date_token_re = re.compile(r"(?:^|[^\\])(?:y+|m+|d+|h+|s+|AM/PM)", re.IGNORECASE)

    def parse(self) -> WorkbookDoc:
        if self.source_path.suffix.lower() != ".xlsx":
            raise ValueError("Only .xlsx is supported in this version")

        with ZipFile(self.source_path) as zip_file:
            workbook = WorkbookDoc(source_path=self.source_path, options=self.options)
            workbook.source_metadata = self._build_source_metadata(zip_file)
            content_types = self._parse_content_types(zip_file)
            shared_strings = self._parse_shared_strings(zip_file)
            self._theme_colors = self._parse_theme_colors(zip_file)
            styles_xml_equivalent, style_css_map, style_numfmt_map = self._parse_styles(zip_file)
            self._style_numfmt_map = style_numfmt_map
            workbook.styles_xml_equivalent = styles_xml_equivalent
            workbook.style_css_map = style_css_map

            wb_root = ET.fromstring(zip_file.read("xl/workbook.xml"))
            wb_rels = self._load_relationships(zip_file, "xl/_rels/workbook.xml.rels")
            sheet_refs = self._parse_sheet_refs(wb_root, wb_rels)

            defined_names, print_areas_by_sheet, print_titles_by_sheet = self._parse_defined_names(wb_root)
            workbook.defined_names = defined_names

            for sheet_ref in sheet_refs:
                if sheet_ref.state != "visible" and not self.options.include_hidden_sheets:
                    continue
                sheet_doc = self._parse_sheet(
                    zip_file=zip_file,
                    content_types=content_types,
                    shared_strings=shared_strings,
                    sheet_ref=sheet_ref,
                    print_areas=print_areas_by_sheet.get(sheet_ref.index, []),
                    print_titles=print_titles_by_sheet.get(sheet_ref.index, []),
                    warnings=workbook.warnings,
                )
                sheet_doc.regions = build_sheet_regions(sheet_doc)
                workbook.sheets.append(sheet_doc)

            unsupported_count = sum(len(sheet.unsupported) for sheet in workbook.sheets)
            if unsupported_count and self.options.strict_unsupported:
                raise RuntimeError(f"Unsupported elements detected: {unsupported_count}")

            workbook.summary = self._build_summary(workbook)
            return workbook

    def _build_source_metadata(self, zip_file: ZipFile) -> dict[str, str | int]:
        payload = self.source_path.read_bytes()
        return {
            "file_name": self.source_path.name,
            "file_size_bytes": len(payload),
            "sha256": hashlib.sha256(payload).hexdigest(),
            "zip_entries": len(zip_file.namelist()),
        }

    def _parse_content_types(self, zip_file: ZipFile) -> dict[str, str]:
        if "[Content_Types].xml" not in zip_file.namelist():
            return {}

        root = ET.fromstring(zip_file.read("[Content_Types].xml"))
        types: dict[str, str] = {}
        defaults: dict[str, str] = {}

        for child in list(root):
            tag = local_name(child.tag)
            if tag == "Default":
                ext = child.attrib.get("Extension", "").lower()
                ctype = child.attrib.get("ContentType", "")
                if ext and ctype:
                    defaults[ext] = ctype
            elif tag == "Override":
                part_name = child.attrib.get("PartName", "")
                ctype = child.attrib.get("ContentType", "")
                if part_name and ctype:
                    types[part_name] = ctype

        for path in zip_file.namelist():
            with_slash = "/" + path if not path.startswith("/") else path
            if with_slash in types:
                continue
            ext = Path(path).suffix.lower().lstrip(".")
            if ext in defaults:
                types[with_slash] = defaults[ext]

        return types

    def _parse_theme_colors(self, zip_file: ZipFile) -> dict[int, str]:
        theme_path = "xl/theme/theme1.xml"
        if theme_path not in zip_file.namelist():
            return {}

        root = ET.fromstring(zip_file.read(theme_path))
        clr_scheme = root.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme")
        if clr_scheme is None:
            return {}

        color_list: list[str] = []
        for child in list(clr_scheme):
            srgb = child.find("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
            if srgb is not None and srgb.attrib.get("val"):
                color_list.append("#" + srgb.attrib["val"].upper())
                continue
            sys_clr = child.find("{http://schemas.openxmlformats.org/drawingml/2006/main}sysClr")
            if sys_clr is not None and sys_clr.attrib.get("lastClr"):
                color_list.append("#" + sys_clr.attrib["lastClr"].upper())

        return {idx: color for idx, color in enumerate(color_list)}

    def _parse_styles(self, zip_file: ZipFile) -> tuple[dict, dict[str, str], dict[str, str]]:
        if "xl/styles.xml" not in zip_file.namelist():
            return {}, {}, {}
        root = ET.fromstring(zip_file.read("xl/styles.xml"))
        style_css_map, style_numfmt_map = self._build_style_maps(root)
        return xml_to_dict(root), style_css_map, style_numfmt_map

    def _build_style_maps(self, styles_root: ET.Element) -> tuple[dict[str, str], dict[str, str]]:
        fonts = [self._extract_font_css(font) for font in styles_root.findall("a:fonts/a:font", NS)]
        fills = [self._extract_fill_css(fill) for fill in styles_root.findall("a:fills/a:fill", NS)]
        borders = [self._extract_border_css(border) for border in styles_root.findall("a:borders/a:border", NS)]
        custom_numfmts = self._parse_custom_numfmts(styles_root)

        style_map: dict[str, str] = {}
        style_numfmt_map: dict[str, str] = {}
        xfs = styles_root.findall("a:cellXfs/a:xf", NS)
        for idx, xf in enumerate(xfs):
            parts: list[str] = []
            try:
                font_id = int(xf.attrib.get("fontId", "0"))
            except ValueError:
                font_id = 0
            try:
                fill_id = int(xf.attrib.get("fillId", "0"))
            except ValueError:
                fill_id = 0
            try:
                border_id = int(xf.attrib.get("borderId", "0"))
            except ValueError:
                border_id = 0

            if 0 <= font_id < len(fonts) and fonts[font_id]:
                parts.extend(fonts[font_id])
            if 0 <= fill_id < len(fills) and fills[fill_id]:
                parts.extend(fills[fill_id])
            if 0 <= border_id < len(borders) and borders[border_id]:
                parts.extend(borders[border_id])

            alignment = xf.find("a:alignment", NS)
            if alignment is not None:
                h = alignment.attrib.get("horizontal")
                v = alignment.attrib.get("vertical")
                if h:
                    parts.append(f"text-align: {h};")
                if v:
                    parts.append(f"vertical-align: {v};")
                if alignment.attrib.get("wrapText") == "1":
                    parts.append("white-space: pre-wrap;")
                else:
                    parts.append("white-space: pre;")

            style_map[str(idx)] = " ".join(parts).strip()
            try:
                num_fmt_id = int(xf.attrib.get("numFmtId", "0"))
            except ValueError:
                num_fmt_id = 0
            style_numfmt_map[str(idx)] = custom_numfmts.get(num_fmt_id) or self._builtin_numfmts.get(num_fmt_id, "")

        style_map.setdefault("0", "")
        style_numfmt_map.setdefault("0", "")
        return style_map, style_numfmt_map

    def _parse_custom_numfmts(self, styles_root: ET.Element) -> dict[int, str]:
        result: dict[int, str] = {}
        for num_fmt in styles_root.findall("a:numFmts/a:numFmt", NS):
            raw_id = num_fmt.attrib.get("numFmtId")
            code = num_fmt.attrib.get("formatCode")
            if raw_id is None or code is None:
                continue
            try:
                fmt_id = int(raw_id)
            except ValueError:
                continue
            result[fmt_id] = code
        return result

    def _extract_font_css(self, font: ET.Element) -> list[str]:
        parts: list[str] = []
        name = font.find("a:name", NS)
        if name is not None and name.attrib.get("val"):
            parts.append(f"font-family: '{name.attrib['val']}';")
        size = font.find("a:sz", NS)
        if size is not None and size.attrib.get("val"):
            parts.append(f"font-size: {size.attrib['val']}pt;")
        if font.find("a:b", NS) is not None:
            parts.append("font-weight: 700;")
        if font.find("a:i", NS) is not None:
            parts.append("font-style: italic;")
        if font.find("a:u", NS) is not None:
            parts.append("text-decoration: underline;")

        color = font.find("a:color", NS)
        color_hex = self._extract_color(color)
        if color_hex:
            parts.append(f"color: {color_hex};")
        return parts

    def _extract_fill_css(self, fill: ET.Element) -> list[str]:
        pattern = fill.find("a:patternFill", NS)
        if pattern is None:
            return []
        pattern_type = pattern.attrib.get("patternType")
        if pattern_type in {None, "none"}:
            return []
        fg = self._extract_color(pattern.find("a:fgColor", NS))
        bg = self._extract_color(pattern.find("a:bgColor", NS))
        color = fg or bg
        if not color:
            return []
        return [f"background-color: {color};"]

    def _extract_border_css(self, border: ET.Element) -> list[str]:
        parts: list[str] = []
        for side in ("left", "right", "top", "bottom"):
            elem = border.find(f"a:{side}", NS)
            if elem is None:
                continue
            style = elem.attrib.get("style")
            if not style:
                continue
            color = self._extract_color(elem.find("a:color", NS)) or "#6b7280"
            width, pattern = self._border_style(style)
            parts.append(f"border-{side}: {width}px {pattern} {color};")
        return parts

    def _border_style(self, style: str) -> tuple[int, str]:
        mapping = {
            "thin": (1, "solid"),
            "medium": (2, "solid"),
            "thick": (3, "solid"),
            "dotted": (1, "dotted"),
            "dashed": (1, "dashed"),
            "double": (3, "double"),
            "hair": (1, "solid"),
            "mediumDashed": (2, "dashed"),
            "dashDot": (1, "dashed"),
            "mediumDashDot": (2, "dashed"),
            "dashDotDot": (1, "dashed"),
            "mediumDashDotDot": (2, "dashed"),
            "slantDashDot": (1, "dashed"),
        }
        return mapping.get(style, (1, "solid"))

    def _extract_color(self, color_elem: ET.Element | None) -> str | None:
        if color_elem is None:
            return None
        rgb = color_elem.attrib.get("rgb")
        if rgb:
            hex_rgb = rgb[-6:] if len(rgb) >= 6 else rgb
            return f"#{hex_rgb.upper()}"
        theme = color_elem.attrib.get("theme")
        if theme is not None:
            try:
                theme_idx = int(theme)
            except ValueError:
                theme_idx = -1
            base = self._theme_colors.get(theme_idx)
            if base:
                tint_raw = color_elem.attrib.get("tint")
                if tint_raw is not None:
                    try:
                        tint = float(tint_raw)
                        return self._apply_tint(base, tint)
                    except ValueError:
                        return base
                return base
        if color_elem.attrib.get("auto") == "1":
            return "#000000"
        indexed = color_elem.attrib.get("indexed")
        if indexed is not None:
            try:
                idx = int(indexed)
            except ValueError:
                return None
            indexed_map = {
                0: "#000000",
                1: "#FFFFFF",
                2: "#FF0000",
                3: "#00FF00",
                4: "#0000FF",
                5: "#FFFF00",
                6: "#FF00FF",
                7: "#00FFFF",
                8: "#000000",
                9: "#FFFFFF",
            }
            return indexed_map.get(idx)
        return None

    def _apply_tint(self, hex_color: str, tint: float) -> str:
        hex_value = hex_color.lstrip("#")
        if len(hex_value) != 6:
            return hex_color
        r = int(hex_value[0:2], 16)
        g = int(hex_value[2:4], 16)
        b = int(hex_value[4:6], 16)

        if tint < 0:
            factor = 1.0 + tint
            r = int(r * factor)
            g = int(g * factor)
            b = int(b * factor)
        else:
            r = int(r * (1.0 - tint) + 255 * tint)
            g = int(g * (1.0 - tint) + 255 * tint)
            b = int(b * (1.0 - tint) + 255 * tint)

        r = min(255, max(0, r))
        g = min(255, max(0, g))
        b = min(255, max(0, b))
        return f"#{r:02X}{g:02X}{b:02X}"

    def _parse_shared_strings(self, zip_file: ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in zip_file.namelist():
            return []

        root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
        values: list[str] = []
        for si in root.findall(f"{{{SPREADSHEET_NS}}}si"):
            direct = si.find(f"{{{SPREADSHEET_NS}}}t")
            if direct is not None:
                values.append(direct.text or "")
                continue
            texts: list[str] = []
            for txt in si.findall(f".//{{{SPREADSHEET_NS}}}t"):
                texts.append(txt.text or "")
            values.append("".join(texts))
        return values

    def _load_relationships(self, zip_file: ZipFile, path: str) -> dict[str, str]:
        if path not in zip_file.namelist():
            return {}
        root = ET.fromstring(zip_file.read(path))
        rels: dict[str, str] = {}
        for rel in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
            rel_id = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            if rel_id and target:
                rels[rel_id] = target
        return rels

    def _parse_sheet_refs(self, wb_root: ET.Element, wb_rels: dict[str, str]) -> list[_SheetRef]:
        sheet_refs: list[_SheetRef] = []
        for idx, sheet in enumerate(wb_root.findall("a:sheets/a:sheet", NS)):
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
            target = wb_rels.get(rid)
            if not target:
                continue
            path = resolve_target("xl/workbook.xml", target)
            sheet_refs.append(
                _SheetRef(
                    index=idx,
                    rid=rid,
                    name=sheet.attrib.get("name", f"Sheet{idx+1}"),
                    state=sheet.attrib.get("state", "visible"),
                    path=path,
                )
            )
        return sheet_refs

    def _parse_defined_names(
        self,
        wb_root: ET.Element,
    ) -> tuple[list[DefinedName], dict[int, list], dict[int, list[str]]]:
        defined_names: list[DefinedName] = []
        print_areas: dict[int, list] = {}
        print_titles: dict[int, list[str]] = {}

        for dn in wb_root.findall("a:definedNames/a:definedName", NS):
            name = dn.attrib.get("name", "")
            value = (dn.text or "").strip()
            local_sheet_id_raw = dn.attrib.get("localSheetId")
            local_sheet_id = int(local_sheet_id_raw) if local_sheet_id_raw is not None else None
            defined_names.append(DefinedName(name=name, value=value, local_sheet_id=local_sheet_id))

            if local_sheet_id is None:
                continue
            if name == "_xlnm.Print_Area":
                print_areas.setdefault(local_sheet_id, []).extend(parse_sheet_scoped_range(value))
            elif name == "_xlnm.Print_Titles":
                print_titles.setdefault(local_sheet_id, []).append(value)

        return defined_names, print_areas, print_titles

    def _parse_sheet(
        self,
        zip_file: ZipFile,
        content_types: dict[str, str],
        shared_strings: list[str],
        sheet_ref: _SheetRef,
        print_areas,
        print_titles,
        warnings: list[str],
    ) -> SheetDoc:
        root = ET.fromstring(zip_file.read(sheet_ref.path))

        dim_elem = root.find("a:dimension", NS)
        dimension_ref = dim_elem.attrib.get("ref", "A1") if dim_elem is not None else "A1"

        sheet = SheetDoc(
            index=sheet_ref.index,
            name=sheet_ref.name,
            state=sheet_ref.state,
            path=sheet_ref.path,
            dimension_ref=dimension_ref,
            print_areas=list(print_areas),
            print_titles=list(print_titles),
        )

        self._parse_rows_cols_cells(root, shared_strings, sheet)
        self._parse_merges(root, sheet)
        self._parse_data_validations(root, sheet)
        self._parse_sheet_view(root, sheet)
        self._parse_sheet_print_metadata(root, sheet)
        self._parse_sheet_unsupported(root, sheet)
        self._parse_sheet_drawings(zip_file, content_types, sheet, warnings)

        return sheet

    def _parse_rows_cols_cells(self, root: ET.Element, shared_strings: list[str], sheet: SheetDoc) -> None:
        for row_elem in root.findall(".//a:sheetData/a:row", NS):
            row_idx = int(row_elem.attrib.get("r", "0"))
            ht = row_elem.attrib.get("ht")
            if row_elem.attrib.get("hidden") == "1":
                sheet.hidden_rows.add(row_idx)
            if ht:
                try:
                    sheet.row_heights[row_idx] = float(ht)
                except ValueError:
                    pass

            for cell_elem in row_elem.findall("a:c", NS):
                coord = cell_elem.attrib.get("r")
                if not coord:
                    continue

                row, col = coord_to_rowcol(coord)
                cell_type = cell_elem.attrib.get("t", "n")
                style_id = cell_elem.attrib.get("s")

                formula_elem = cell_elem.find("a:f", NS)
                formula = None
                if formula_elem is not None:
                    formula = (formula_elem.text or "").strip() or None
                    if formula is None and formula_elem.attrib:
                        formula = f"<formula:{dict(formula_elem.attrib)}>"

                value_elem = cell_elem.find("a:v", NS)
                cached_value = value_elem.text if value_elem is not None else None
                display_value = self._decode_cell_value(
                    cell_elem,
                    cell_type,
                    cached_value,
                    shared_strings,
                    style_id,
                )

                cell = CellData(
                    coord=coord,
                    row=row,
                    col=col,
                    cell_type=cell_type,
                    value=display_value,
                    display_value=display_value,
                    formula=formula,
                    cached_value=cached_value,
                    style_id=style_id,
                )
                sheet.cells.append(cell)
                sheet.cell_map[coord] = cell

        for col_elem in root.findall(".//a:cols/a:col", NS):
            start = int(col_elem.attrib.get("min", "0"))
            end = int(col_elem.attrib.get("max", "0"))
            if col_elem.attrib.get("hidden") == "1":
                for idx in range(start, end + 1):
                    sheet.hidden_cols.add(idx)

            width = col_elem.attrib.get("width")
            if width is None:
                continue
            try:
                width_value = float(width)
            except ValueError:
                continue
            for idx in range(start, end + 1):
                sheet.col_widths[idx] = width_value

    def _decode_cell_value(
        self,
        cell_elem: ET.Element,
        cell_type: str,
        cached_value: str | None,
        shared_strings: list[str],
        style_id: str | None,
    ) -> str:
        if cell_type == "s":
            if cached_value is None:
                return ""
            try:
                idx = int(cached_value)
            except ValueError:
                return ""
            return shared_strings[idx] if 0 <= idx < len(shared_strings) else ""

        if cell_type == "inlineStr":
            inline = cell_elem.find("a:is", NS)
            if inline is None:
                return ""
            direct = inline.find("a:t", NS)
            if direct is not None:
                return direct.text or ""
            return "".join((node.text or "") for node in inline.findall(".//a:t", NS))

        if cell_type == "str":
            return cached_value or ""

        if cell_type == "b":
            return "TRUE" if cached_value == "1" else "FALSE"

        if cell_type in {"e"}:
            return cached_value or ""

        return self._format_number(cached_value, style_id)

    def _format_number(self, value: str | None, style_id: str | None) -> str:
        if value is None:
            return ""
        raw = value.strip()
        if raw == "":
            return ""
        try:
            number = float(raw)
        except ValueError:
            return raw

        fmt = self._style_numfmt_map.get(style_id or "0", "")
        if not fmt or fmt.lower() == "general":
            return self._normalize_general_number(number, raw)

        primary = fmt.split(";")[0]
        if self._is_date_format(primary):
            return self._format_excel_date(number, primary)
        if "%" in primary:
            return self._format_percent(number, primary)
        if any(token in primary for token in ("0", "#")):
            return self._format_decimal(number, primary)
        return self._normalize_general_number(number, raw)

    def _normalize_general_number(self, number: float, raw: str) -> str:
        if abs(number - round(number)) < 1e-11:
            return str(int(round(number)))
        return raw

    def _is_date_format(self, fmt: str) -> bool:
        cleaned = self._strip_quoted(fmt)
        return bool(self._date_token_re.search(cleaned))

    def _strip_quoted(self, fmt: str) -> str:
        out: list[str] = []
        in_quote = False
        for ch in fmt:
            if ch == '"':
                in_quote = not in_quote
                continue
            if not in_quote:
                out.append(ch)
        return "".join(out)

    def _format_excel_date(self, number: float, fmt: str) -> str:
        base = datetime(1899, 12, 30)
        dt = base + timedelta(days=number)
        cleaned = fmt.lower()
        has_date = any(t in cleaned for t in ("y", "d", "m"))
        has_time = any(t in cleaned for t in ("h", "s")) or "am/pm" in cleaned
        if has_date and has_time:
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        if has_time:
            return dt.strftime("%H:%M:%S")
        return dt.strftime("%Y-%m-%d")

    def _format_percent(self, number: float, fmt: str) -> str:
        decimals = 0
        if "." in fmt:
            after = fmt.split(".", 1)[1]
            decimals = sum(1 for ch in after if ch in {"0", "#"})
        value = number * 100
        return f"{value:.{decimals}f}%"

    def _format_decimal(self, number: float, fmt: str) -> str:
        use_grouping = "," in fmt.split(".", 1)[0]
        decimals = 0
        if "." in fmt:
            after = fmt.split(".", 1)[1]
            decimals = sum(1 for ch in after if ch in {"0", "#"})

        if use_grouping:
            return f"{number:,.{decimals}f}"
        return f"{number:.{decimals}f}"

    def _parse_sheet_view(self, root: ET.Element, sheet: SheetDoc) -> None:
        pane = root.find("a:sheetViews/a:sheetView/a:pane", NS)
        if pane is not None:
            sheet.pane = dict(pane.attrib)

    def _parse_merges(self, root: ET.Element, sheet: SheetDoc) -> None:
        for merge in root.findall(".//a:mergeCells/a:mergeCell", NS):
            ref = merge.attrib.get("ref")
            if not ref:
                continue
            try:
                rng = parse_range_ref(ref)
            except ValueError:
                continue
            sheet.merges.append(rng)
            for row, col in iter_cells_in_range(rng):
                coord = self._row_col_to_coord(row, col)
                sheet.merge_map[coord] = rng.ref

    def _parse_data_validations(self, root: ET.Element, sheet: SheetDoc) -> None:
        for dv in root.findall(".//a:dataValidations/a:dataValidation", NS):
            sqref = dv.attrib.get("sqref", "")
            if not sqref:
                continue
            sheet.data_validations.append(
                DataValidation(
                    type=dv.attrib.get("type"),
                    sqref=sqref,
                    allow_blank=self._to_bool(dv.attrib.get("allowBlank")),
                    show_error_message=self._to_bool(dv.attrib.get("showErrorMessage")),
                    operator=dv.attrib.get("operator"),
                    formula1=(dv.findtext("a:formula1", default="", namespaces=NS) or None),
                    formula2=(dv.findtext("a:formula2", default="", namespaces=NS) or None),
                )
            )

    def _parse_sheet_print_metadata(self, root: ET.Element, sheet: SheetDoc) -> None:
        print_options = root.find("a:printOptions", NS)
        if print_options is not None:
            sheet.print_options = dict(print_options.attrib)

        page_margins = root.find("a:pageMargins", NS)
        if page_margins is not None:
            sheet.page_margins = dict(page_margins.attrib)

        page_setup = root.find("a:pageSetup", NS)
        if page_setup is not None:
            sheet.page_setup = dict(page_setup.attrib)

        header_footer = root.find("a:headerFooter", NS)
        if header_footer is not None:
            hf: dict[str, str] = dict(header_footer.attrib)
            for child in list(header_footer):
                hf[local_name(child.tag)] = child.text or ""
            sheet.header_footer = hf

        row_breaks = root.findall("a:rowBreaks/a:brk", NS)
        col_breaks = root.findall("a:colBreaks/a:brk", NS)
        if row_breaks:
            sheet.page_breaks["row"] = [int(b.attrib.get("id", "0")) for b in row_breaks]
        if col_breaks:
            sheet.page_breaks["col"] = [int(b.attrib.get("id", "0")) for b in col_breaks]

    def _parse_sheet_unsupported(self, root: ET.Element, sheet: SheetDoc) -> None:
        allowed_tags = {
            "dimension",
            "sheetViews",
            "sheetFormatPr",
            "cols",
            "sheetData",
            "sheetCalcPr",
            "sheetProtection",
            "protectedRanges",
            "scenarios",
            "autoFilter",
            "sortState",
            "dataConsolidate",
            "customSheetViews",
            "mergeCells",
            "phoneticPr",
            "conditionalFormatting",
            "dataValidations",
            "hyperlinks",
            "printOptions",
            "pageMargins",
            "pageSetup",
            "headerFooter",
            "rowBreaks",
            "colBreaks",
            "customProperties",
            "cellWatches",
            "ignoredErrors",
            "smartTags",
            "drawing",
            "legacyDrawing",
            "legacyDrawingHF",
            "picture",
            "oleObjects",
            "controls",
            "webPublishItems",
            "tableParts",
            "extLst",
            "sheetPr",
        }

        for child in list(root):
            tag = local_name(child.tag)
            if tag in allowed_tags:
                continue
            sheet.unsupported.append(
                UnsupportedElement(
                    scope="worksheet",
                    location=sheet.path,
                    tag=tag,
                    raw_xml=ET.tostring(child, encoding="unicode"),
                )
            )

    def _parse_sheet_drawings(
        self,
        zip_file: ZipFile,
        content_types: dict[str, str],
        sheet: SheetDoc,
        warnings: list[str],
    ) -> None:
        rels_path = self._worksheet_rels_path(sheet.path)
        rels = self._load_relationships(zip_file, rels_path)
        if not rels:
            return

        drawing_targets: list[str] = []
        root = ET.fromstring(zip_file.read(rels_path))
        for rel in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
            rel_type = rel.attrib.get("Type", "")
            target = rel.attrib.get("Target", "")
            if rel_type.endswith("/drawing") and target:
                drawing_targets.append(resolve_target(sheet.path, target))

        for drawing_path in drawing_targets:
            if drawing_path not in zip_file.namelist():
                warnings.append(f"Missing drawing part: {drawing_path}")
                continue
            drawing_objs, connectors, mermaid = parse_drawing_for_sheet(
                zip_file=zip_file,
                drawing_path=drawing_path,
                content_types=content_types,
                options=self.options,
                unsupported=sheet.unsupported,
                warnings=warnings,
            )
            sheet.drawings.extend(drawing_objs)
            sheet.connectors.extend(connectors)
            if mermaid:
                if sheet.mermaid:
                    sheet.mermaid += "\n\n" + mermaid
                else:
                    sheet.mermaid = mermaid

    def _worksheet_rels_path(self, sheet_path: str) -> str:
        parent, file_name = sheet_path.rsplit("/", 1)
        return f"{parent}/_rels/{file_name}.rels"

    def _build_summary(self, workbook: WorkbookDoc) -> dict[str, int]:
        total_cells = sum(len(sheet.cells) for sheet in workbook.sheets)
        total_merges = sum(len(sheet.merges) for sheet in workbook.sheets)
        total_formulas = sum(1 for sheet in workbook.sheets for cell in sheet.cells if cell.formula)
        total_drawings = sum(len(sheet.drawings) for sheet in workbook.sheets)
        total_connectors = sum(len(sheet.connectors) for sheet in workbook.sheets)
        total_images = sum(1 for sheet in workbook.sheets for obj in sheet.drawings if obj.image_data_uri)
        total_regions = sum(len(sheet.regions) for sheet in workbook.sheets)
        total_unsupported = sum(len(sheet.unsupported) for sheet in workbook.sheets)

        return {
            "sheet_count": len(workbook.sheets),
            "defined_name_count": len(workbook.defined_names),
            "cell_count": total_cells,
            "merge_count": total_merges,
            "formula_count": total_formulas,
            "drawing_object_count": total_drawings,
            "connector_count": total_connectors,
            "embedded_image_count": total_images,
            "region_count": total_regions,
            "unsupported_count": total_unsupported,
            "warning_count": len(workbook.warnings),
        }

    def _to_bool(self, value: str | None) -> bool | None:
        if value is None:
            return None
        return value in {"1", "true", "TRUE"}

    def _row_col_to_coord(self, row: int, col: int) -> str:
        letters: list[str] = []
        n = col
        while n > 0:
            n, rem = divmod(n - 1, 26)
            letters.append(chr(65 + rem))
        return "".join(reversed(letters)) + str(row)
