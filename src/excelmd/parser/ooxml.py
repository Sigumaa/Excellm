from __future__ import annotations

import hashlib
from dataclasses import dataclass
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

    def parse(self) -> WorkbookDoc:
        if self.source_path.suffix.lower() != ".xlsx":
            raise ValueError("Only .xlsx is supported in this version")

        with ZipFile(self.source_path) as zip_file:
            workbook = WorkbookDoc(source_path=self.source_path, options=self.options)
            workbook.source_metadata = self._build_source_metadata(zip_file)
            content_types = self._parse_content_types(zip_file)
            shared_strings = self._parse_shared_strings(zip_file)
            workbook.styles_xml_equivalent = self._parse_styles_xml_equivalent(zip_file)

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

    def _parse_styles_xml_equivalent(self, zip_file: ZipFile) -> dict:
        if "xl/styles.xml" not in zip_file.namelist():
            return {}
        root = ET.fromstring(zip_file.read("xl/styles.xml"))
        return xml_to_dict(root)

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
        self._parse_sheet_print_metadata(root, sheet)
        self._parse_sheet_unsupported(root, sheet)
        self._parse_sheet_drawings(zip_file, content_types, sheet, warnings)

        return sheet

    def _parse_rows_cols_cells(self, root: ET.Element, shared_strings: list[str], sheet: SheetDoc) -> None:
        for row_elem in root.findall(".//a:sheetData/a:row", NS):
            row_idx = int(row_elem.attrib.get("r", "0"))
            ht = row_elem.attrib.get("ht")
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
                value = self._decode_cell_value(cell_elem, cell_type, cached_value, shared_strings)

                cell = CellData(
                    coord=coord,
                    row=row,
                    col=col,
                    cell_type=cell_type,
                    value=value,
                    formula=formula,
                    cached_value=cached_value,
                    style_id=style_id,
                )
                sheet.cells.append(cell)
                sheet.cell_map[coord] = cell

        for col_elem in root.findall(".//a:cols/a:col", NS):
            width = col_elem.attrib.get("width")
            if width is None:
                continue
            try:
                width_value = float(width)
            except ValueError:
                continue
            start = int(col_elem.attrib.get("min", "0"))
            end = int(col_elem.attrib.get("max", "0"))
            for idx in range(start, end + 1):
                sheet.col_widths[idx] = width_value

    def _decode_cell_value(
        self,
        cell_elem: ET.Element,
        cell_type: str,
        cached_value: str | None,
        shared_strings: list[str],
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

        return cached_value or ""

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
