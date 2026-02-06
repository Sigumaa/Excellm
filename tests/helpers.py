from __future__ import annotations

import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


def ground_truth_counts(path) -> dict[str, int]:
    with ZipFile(path) as zf:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        sheet_count = len(wb.findall(f"{{{SPREADSHEET_NS}}}sheets/{{{SPREADSHEET_NS}}}sheet"))

        merge_count = 0
        formula_count = 0
        for name in zf.namelist():
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                continue
            root = ET.fromstring(zf.read(name))
            merge_count += len(root.findall(f".//{{{SPREADSHEET_NS}}}mergeCell"))
            formula_count += len(root.findall(f".//{{{SPREADSHEET_NS}}}f"))

        drawing_count = 0
        connector_count = 0
        for name in zf.namelist():
            if not re.match(r"xl/drawings/drawing\d+\.xml$", name):
                continue
            root = ET.fromstring(zf.read(name))
            d_count, c_count = _count_drawing_objects(root)
            drawing_count += d_count
            connector_count += c_count

        return {
            "sheet_count": sheet_count,
            "merge_count": merge_count,
            "formula_count": formula_count,
            "drawing_object_count": drawing_count,
            "connector_count": connector_count,
        }


def _count_drawing_objects(root: ET.Element) -> tuple[int, int]:
    total = 0
    connectors = 0

    for anchor in list(root):
        for child in list(anchor):
            tag = _local_name(child.tag)
            if tag in {"from", "to", "clientData", "pos", "ext", "nvGrpSpPr", "grpSpPr"}:
                continue
            t, c = _count_object_recursive(child)
            total += t
            connectors += c

    return total, connectors


def _count_object_recursive(elem: ET.Element) -> tuple[int, int]:
    tag = _local_name(elem.tag)
    recognized = {"sp", "cxnSp", "pic", "grpSp", "graphicFrame"}
    if tag not in recognized:
        return 0, 0

    total = 1
    connectors = 1 if tag == "cxnSp" else 0

    if tag == "grpSp":
        for child in list(elem):
            child_tag = _local_name(child.tag)
            if child_tag in {"nvGrpSpPr", "grpSpPr"}:
                continue
            t, c = _count_object_recursive(child)
            total += t
            connectors += c

    return total, connectors


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
