from __future__ import annotations

import base64
import math
import mimetypes
from collections import defaultdict
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from ..model import AnchorPoint, ConnectorInfo, ConvertOptions, DrawingObject, UnsupportedElement
from .namespaces import DOCUMENT_REL_NS, DRAWING_MAIN_NS, PACKAGE_REL_NS, SHEET_DRAWING_NS
from .utils import local_name, resolve_target

EMU_PER_PIXEL = 9525.0
DEFAULT_COL_PX = 64.0
DEFAULT_ROW_PX = 20.0


def parse_drawing_for_sheet(
    zip_file: ZipFile,
    drawing_path: str,
    content_types: dict[str, str],
    options: ConvertOptions,
    unsupported: list[UnsupportedElement],
    warnings: list[str],
) -> tuple[list[DrawingObject], list[ConnectorInfo], str]:
    root = ET.fromstring(zip_file.read(drawing_path))
    rels_path = _rels_path_for(drawing_path)
    rel_map = _load_relationship_map(zip_file, rels_path)

    drawing_objects: list[DrawingObject] = []
    connectors: list[ConnectorInfo] = []
    uid_counter: defaultdict[str, int] = defaultdict(int)

    for anchor in list(root):
        anchor_tag = local_name(anchor.tag)
        if anchor_tag not in {"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"}:
            unsupported.append(
                UnsupportedElement(
                    scope="drawing",
                    location=drawing_path,
                    tag=anchor_tag,
                    raw_xml=ET.tostring(anchor, encoding="unicode"),
                )
            )
            continue

        anchor_from, anchor_to, bbox = _parse_anchor(anchor, anchor_tag)

        for child in list(anchor):
            child_tag = local_name(child.tag)
            if child_tag in {"from", "to", "clientData", "pos", "ext"}:
                continue

            if child_tag not in {"sp", "cxnSp", "pic", "grpSp", "graphicFrame"}:
                unsupported.append(
                    UnsupportedElement(
                        scope="drawing",
                        location=drawing_path,
                        tag=child_tag,
                        raw_xml=ET.tostring(child, encoding="unicode"),
                    )
                )
                continue

            _walk_drawing_object(
                zip_file=zip_file,
                drawing_path=drawing_path,
                element=child,
                kind=child_tag,
                content_types=content_types,
                rel_map=rel_map,
                options=options,
                anchor_type=anchor_tag,
                anchor_from=anchor_from,
                anchor_to=anchor_to,
                bbox=bbox,
                parent_uid=None,
                drawing_objects=drawing_objects,
                connectors=connectors,
                uid_counter=uid_counter,
            )

    infer_connectors(drawing_objects, connectors, warnings)
    mermaid = build_mermaid(drawing_objects, connectors)
    return drawing_objects, connectors, mermaid


def _walk_drawing_object(
    zip_file: ZipFile,
    drawing_path: str,
    element: ET.Element,
    kind: str,
    content_types: dict[str, str],
    rel_map: dict[str, str],
    options: ConvertOptions,
    anchor_type: str,
    anchor_from: AnchorPoint | None,
    anchor_to: AnchorPoint | None,
    bbox: tuple[float, float, float, float],
    parent_uid: str | None,
    drawing_objects: list[DrawingObject],
    connectors: list[ConnectorInfo],
    uid_counter: defaultdict[str, int],
) -> None:
    object_id, name = _extract_identity(element, kind)
    if not object_id:
        object_id = f"auto-{len(drawing_objects)+1}"

    raw_uid = f"{drawing_path}:{object_id}"
    uid_counter[raw_uid] += 1
    object_uid = raw_uid if uid_counter[raw_uid] == 1 else f"{raw_uid}#{uid_counter[raw_uid]}"

    text = _extract_text(element)
    image_target = None
    image_content_type = None
    image_data_uri = None
    extra: dict[str, str] = _extract_shape_style(element)

    if kind == "pic":
        image_target, image_content_type, image_data_uri = _extract_picture(
            zip_file=zip_file,
            drawing_path=drawing_path,
            pic_element=element,
            rel_map=rel_map,
            content_types=content_types,
            options=options,
        )

    if kind == "cxnSp":
        arrow_head, arrow_tail = _extract_connector_arrows(element)
        extra["arrow_head"] = arrow_head or "none"
        extra["arrow_tail"] = arrow_tail or "none"

    obj = DrawingObject(
        object_uid=object_uid,
        object_id=object_id,
        drawing_path=drawing_path,
        kind=kind,
        name=name,
        text=text,
        anchor_type=anchor_type,
        anchor_from=anchor_from,
        anchor_to=anchor_to,
        bbox=bbox,
        parent_uid=parent_uid,
        image_target=image_target,
        image_content_type=image_content_type,
        image_data_uri=image_data_uri,
        raw_xml=ET.tostring(element, encoding="unicode"),
        extra=extra,
    )
    drawing_objects.append(obj)

    if kind == "cxnSp":
        connectors.append(
            ConnectorInfo(
                object_uid=object_uid,
                object_id=object_id,
                drawing_path=drawing_path,
                name=name,
                text=text,
                anchor_from=anchor_from,
                anchor_to=anchor_to,
                bbox=bbox,
                arrow_head=extra.get("arrow_head"),
                arrow_tail=extra.get("arrow_tail"),
                direction="undirected",
                source_uid=None,
                target_uid=None,
                resolved=False,
                distance_source=None,
                distance_target=None,
                raw_xml=ET.tostring(element, encoding="unicode"),
            )
        )

    if kind == "grpSp":
        for child in list(element):
            child_tag = local_name(child.tag)
            if child_tag in {"nvGrpSpPr", "grpSpPr"}:
                continue
            if child_tag not in {"sp", "cxnSp", "pic", "grpSp", "graphicFrame"}:
                continue
            _walk_drawing_object(
                zip_file=zip_file,
                drawing_path=drawing_path,
                element=child,
                kind=child_tag,
                content_types=content_types,
                rel_map=rel_map,
                options=options,
                anchor_type=anchor_type,
                anchor_from=anchor_from,
                anchor_to=anchor_to,
                bbox=bbox,
                parent_uid=object_uid,
                drawing_objects=drawing_objects,
                connectors=connectors,
                uid_counter=uid_counter,
            )


def _rels_path_for(drawing_path: str) -> str:
    prefix, file_name = drawing_path.rsplit("/", 1)
    return f"{prefix}/_rels/{file_name}.rels"


def _load_relationship_map(zip_file: ZipFile, rels_path: str) -> dict[str, str]:
    if rels_path not in zip_file.namelist():
        return {}
    root = ET.fromstring(zip_file.read(rels_path))
    rel_map: dict[str, str] = {}
    for rel in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if not rel_id or not target:
            continue
        rel_map[rel_id] = target
    return rel_map


def _parse_anchor(
    anchor: ET.Element,
    anchor_tag: str,
) -> tuple[AnchorPoint | None, AnchorPoint | None, tuple[float, float, float, float]]:
    if anchor_tag == "twoCellAnchor":
        from_elem = anchor.find(f"{{{SHEET_DRAWING_NS}}}from")
        to_elem = anchor.find(f"{{{SHEET_DRAWING_NS}}}to")
        anchor_from = _parse_anchor_point(from_elem)
        anchor_to = _parse_anchor_point(to_elem)
        bbox = _bbox_from_anchor(anchor_from, anchor_to)
        return anchor_from, anchor_to, bbox

    if anchor_tag == "oneCellAnchor":
        from_elem = anchor.find(f"{{{SHEET_DRAWING_NS}}}from")
        ext_elem = anchor.find(f"{{{SHEET_DRAWING_NS}}}ext")
        anchor_from = _parse_anchor_point(from_elem)
        if anchor_from is None:
            return None, None, (0.0, 0.0, 0.0, 0.0)

        ext_cx = int(ext_elem.attrib.get("cx", "0")) if ext_elem is not None else 0
        ext_cy = int(ext_elem.attrib.get("cy", "0")) if ext_elem is not None else 0
        add_cols = max(1, int(round((ext_cx / EMU_PER_PIXEL) / DEFAULT_COL_PX)))
        add_rows = max(1, int(round((ext_cy / EMU_PER_PIXEL) / DEFAULT_ROW_PX)))
        anchor_to = AnchorPoint(
            col=anchor_from.col + add_cols,
            row=anchor_from.row + add_rows,
            col_off=anchor_from.col_off,
            row_off=anchor_from.row_off,
        )
        bbox = _bbox_from_anchor(anchor_from, anchor_to)
        return anchor_from, anchor_to, bbox

    pos = anchor.find(f"{{{SHEET_DRAWING_NS}}}pos")
    ext = anchor.find(f"{{{SHEET_DRAWING_NS}}}ext")
    x = int(pos.attrib.get("x", "0")) / EMU_PER_PIXEL if pos is not None else 0.0
    y = int(pos.attrib.get("y", "0")) / EMU_PER_PIXEL if pos is not None else 0.0
    w = int(ext.attrib.get("cx", "0")) / EMU_PER_PIXEL if ext is not None else 0.0
    h = int(ext.attrib.get("cy", "0")) / EMU_PER_PIXEL if ext is not None else 0.0
    return None, None, (x, y, x + w, y + h)


def _parse_anchor_point(elem: ET.Element | None) -> AnchorPoint | None:
    if elem is None:
        return None
    return AnchorPoint(
        col=int(elem.findtext(f"{{{SHEET_DRAWING_NS}}}col", default="0")),
        row=int(elem.findtext(f"{{{SHEET_DRAWING_NS}}}row", default="0")),
        col_off=int(elem.findtext(f"{{{SHEET_DRAWING_NS}}}colOff", default="0")),
        row_off=int(elem.findtext(f"{{{SHEET_DRAWING_NS}}}rowOff", default="0")),
    )


def _bbox_from_anchor(
    anchor_from: AnchorPoint | None,
    anchor_to: AnchorPoint | None,
) -> tuple[float, float, float, float]:
    if anchor_from is None or anchor_to is None:
        return (0.0, 0.0, 0.0, 0.0)

    x1, y1 = _point_to_xy(anchor_from)
    x2, y2 = _point_to_xy(anchor_to)
    return (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))


def _point_to_xy(point: AnchorPoint) -> tuple[float, float]:
    x = point.col * DEFAULT_COL_PX + (point.col_off / EMU_PER_PIXEL)
    y = point.row * DEFAULT_ROW_PX + (point.row_off / EMU_PER_PIXEL)
    return x, y


def _extract_identity(element: ET.Element, kind: str) -> tuple[str, str]:
    path_by_kind = {
        "sp": f"{{{SHEET_DRAWING_NS}}}nvSpPr/{{{SHEET_DRAWING_NS}}}cNvPr",
        "cxnSp": f"{{{SHEET_DRAWING_NS}}}nvCxnSpPr/{{{SHEET_DRAWING_NS}}}cNvPr",
        "pic": f"{{{SHEET_DRAWING_NS}}}nvPicPr/{{{SHEET_DRAWING_NS}}}cNvPr",
        "grpSp": f"{{{SHEET_DRAWING_NS}}}nvGrpSpPr/{{{SHEET_DRAWING_NS}}}cNvPr",
        "graphicFrame": f"{{{SHEET_DRAWING_NS}}}nvGraphicFramePr/{{{SHEET_DRAWING_NS}}}cNvPr",
    }
    c_nv_pr = element.find(path_by_kind.get(kind, ""))
    if c_nv_pr is None:
        return "", ""
    return c_nv_pr.attrib.get("id", ""), c_nv_pr.attrib.get("name", "")


def _extract_text(element: ET.Element) -> str:
    fragments: list[str] = []
    for txt in element.findall(f".//{{{DRAWING_MAIN_NS}}}t"):
        if txt.text:
            fragments.append(txt.text)
    return "".join(fragments).strip()


def _extract_shape_style(element: ET.Element) -> dict[str, str]:
    extra: dict[str, str] = {}
    sp_pr = element.find(f"{{{SHEET_DRAWING_NS}}}spPr")
    if sp_pr is None:
        return extra

    line = sp_pr.find(f"{{{DRAWING_MAIN_NS}}}ln")
    if line is not None:
        width = line.attrib.get("w")
        if width:
            try:
                extra["line_width_px"] = f"{max(1.0, int(width) / EMU_PER_PIXEL):.2f}"
            except ValueError:
                pass
        line_color = _extract_drawing_color(line)
        if line_color:
            extra["line_color"] = line_color
        dash = line.find(f"{{{DRAWING_MAIN_NS}}}prstDash")
        if dash is not None and dash.attrib.get("val"):
            extra["line_dash"] = dash.attrib["val"]

    fill_color = _extract_fill_color(sp_pr)
    if fill_color:
        extra["fill_color"] = fill_color

    return extra


def _extract_fill_color(sp_pr: ET.Element) -> str | None:
    solid = sp_pr.find(f"{{{DRAWING_MAIN_NS}}}solidFill")
    if solid is not None:
        color = _extract_drawing_color(solid)
        if color:
            return color
    gradient = sp_pr.find(f"{{{DRAWING_MAIN_NS}}}gradFill")
    if gradient is not None:
        first_stop = gradient.find(f".//{{{DRAWING_MAIN_NS}}}gs")
        if first_stop is not None:
            color = _extract_drawing_color(first_stop)
            if color:
                return color
    return None


def _extract_drawing_color(node: ET.Element) -> str | None:
    for tag in ("srgbClr", "sysClr", "schemeClr", "prstClr"):
        c = node.find(f".//{{{DRAWING_MAIN_NS}}}{tag}")
        if c is None:
            continue
        if tag == "srgbClr":
            val = c.attrib.get("val")
            if val:
                return f"#{val.upper()}"
        if tag == "sysClr":
            last = c.attrib.get("lastClr")
            if last:
                return f"#{last.upper()}"
        if tag in {"schemeClr", "prstClr"}:
            val = c.attrib.get("val")
            if val:
                fallback = {
                    "dk1": "#000000",
                    "lt1": "#FFFFFF",
                    "dk2": "#1F2937",
                    "lt2": "#F3F4F6",
                    "accent1": "#4F46E5",
                    "accent2": "#16A34A",
                    "accent3": "#F59E0B",
                    "accent4": "#0EA5E9",
                    "accent5": "#EC4899",
                    "accent6": "#A855F7",
                }
                return fallback.get(val, "#6B7280")
    return None


def _extract_picture(
    zip_file: ZipFile,
    drawing_path: str,
    pic_element: ET.Element,
    rel_map: dict[str, str],
    content_types: dict[str, str],
    options: ConvertOptions,
) -> tuple[str | None, str | None, str | None]:
    blip = pic_element.find(f".//{{{DRAWING_MAIN_NS}}}blip")
    if blip is None:
        return None, None, None

    rel_id = blip.attrib.get(f"{{{DOCUMENT_REL_NS}}}embed")
    if not rel_id:
        return None, None, None

    target = rel_map.get(rel_id)
    if not target:
        return None, None, None

    media_path = resolve_target(drawing_path, target)
    if media_path not in zip_file.namelist():
        return media_path, None, None

    content_type = _guess_content_type(media_path, content_types)
    if options.image_mode != "data_uri":
        return media_path, content_type, None

    payload = zip_file.read(media_path)
    encoded = base64.b64encode(payload).decode("ascii")
    data_uri = f"data:{content_type};base64,{encoded}"
    return media_path, content_type, data_uri


def _guess_content_type(path: str, content_types: dict[str, str]) -> str:
    normalized = "/" + path if not path.startswith("/") else path
    if normalized in content_types:
        return content_types[normalized]
    guessed, _ = mimetypes.guess_type(path)
    return guessed or "application/octet-stream"


def _extract_connector_arrows(cxn_element: ET.Element) -> tuple[str | None, str | None]:
    line = cxn_element.find(f".//{{{DRAWING_MAIN_NS}}}ln")
    if line is None:
        return None, None
    head = line.find(f"{{{DRAWING_MAIN_NS}}}headEnd")
    tail = line.find(f"{{{DRAWING_MAIN_NS}}}tailEnd")
    head_type = head.attrib.get("type") if head is not None else None
    tail_type = tail.attrib.get("type") if tail is not None else None
    return head_type, tail_type


def infer_connectors(
    drawing_objects: list[DrawingObject],
    connectors: list[ConnectorInfo],
    warnings: list[str],
    threshold: float = 220.0,
) -> None:
    node_objects = [obj for obj in drawing_objects if obj.kind != "cxnSp"]
    node_map = {obj.object_uid: obj for obj in node_objects}

    for connector in connectors:
        from_pt = _connector_endpoint(connector.anchor_from, connector.bbox, start=True)
        to_pt = _connector_endpoint(connector.anchor_to, connector.bbox, start=False)

        has_head = _has_arrow(connector.arrow_head)
        has_tail = _has_arrow(connector.arrow_tail)

        if has_head and has_tail:
            connector.direction = "bidirectional"
            source_point = from_pt
            target_point = to_pt
        elif has_tail:
            connector.direction = "forward"
            source_point = from_pt
            target_point = to_pt
        elif has_head:
            connector.direction = "reverse"
            source_point = to_pt
            target_point = from_pt
        else:
            connector.direction = "undirected"
            source_point = from_pt
            target_point = to_pt

        source_uid, d_src = _nearest_node(source_point, node_objects)
        target_uid, d_tgt = _nearest_node(target_point, node_objects)

        if d_src is not None and d_src > threshold:
            source_uid = None
        if d_tgt is not None and d_tgt > threshold:
            target_uid = None

        connector.source_uid = source_uid
        connector.target_uid = target_uid
        connector.distance_source = d_src
        connector.distance_target = d_tgt
        connector.resolved = bool(
            source_uid and target_uid and source_uid != target_uid and source_uid in node_map and target_uid in node_map
        )

        if not connector.resolved:
            warnings.append(f"Unresolved connector: {connector.object_uid}")


def build_mermaid(drawing_objects: list[DrawingObject], connectors: list[ConnectorInfo]) -> str:
    resolved = [c for c in connectors if c.resolved and c.source_uid and c.target_uid]
    if not resolved:
        return ""

    used_node_uids: set[str] = set()
    for conn in resolved:
        used_node_uids.add(conn.source_uid or "")
        used_node_uids.add(conn.target_uid or "")

    nodes = [obj for obj in drawing_objects if obj.object_uid in used_node_uids]
    node_id_map: dict[str, str] = {}
    lines: list[str] = ["flowchart TD"]

    for idx, node in enumerate(nodes, start=1):
        node_id = f"N{idx}"
        node_id_map[node.object_uid] = node_id
        label = _mermaid_escape(node.text.strip() or node.name or node.object_id)
        lines.append(f"    {node_id}[\"{label}\"]")

    for conn in resolved:
        source_id = node_id_map.get(conn.source_uid or "")
        target_id = node_id_map.get(conn.target_uid or "")
        if not source_id or not target_id:
            continue

        edge_label = _mermaid_escape(conn.text.strip()) if conn.text.strip() else ""
        if conn.direction == "bidirectional":
            if edge_label:
                lines.append(f"    {source_id} <-->|\"{edge_label}\"| {target_id}")
            else:
                lines.append(f"    {source_id} <--> {target_id}")
        elif conn.direction == "undirected":
            if edge_label:
                lines.append(f"    {source_id} ---|\"{edge_label}\"| {target_id}")
            else:
                lines.append(f"    {source_id} --- {target_id}")
        else:
            if edge_label:
                lines.append(f"    {source_id} -- \"{edge_label}\" --> {target_id}")
            else:
                lines.append(f"    {source_id} --> {target_id}")

    return "\n".join(lines)


def _connector_endpoint(
    anchor: AnchorPoint | None,
    bbox: tuple[float, float, float, float],
    *,
    start: bool,
) -> tuple[float, float]:
    if anchor is not None:
        return _point_to_xy(anchor)
    if start:
        return bbox[0], bbox[1]
    return bbox[2], bbox[3]


def _nearest_node(
    point: tuple[float, float],
    nodes: list[DrawingObject],
) -> tuple[str | None, float | None]:
    best_uid: str | None = None
    best_dist: float | None = None
    for node in nodes:
        dist = _point_to_bbox_distance(point, node.bbox)
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best_uid = node.object_uid
    return best_uid, best_dist


def _point_to_bbox_distance(
    point: tuple[float, float],
    bbox: tuple[float, float, float, float],
) -> float:
    px, py = point
    x1, y1, x2, y2 = bbox

    if x1 <= px <= x2 and y1 <= py <= y2:
        return 0.0

    dx = max(x1 - px, 0.0, px - x2)
    dy = max(y1 - py, 0.0, py - y2)
    return math.hypot(dx, dy)


def _has_arrow(marker: str | None) -> bool:
    return bool(marker and marker.lower() != "none")


def _mermaid_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace('"', "\\\"").replace("\n", " ").strip()
