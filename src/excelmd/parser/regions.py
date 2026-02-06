from __future__ import annotations

from collections import deque

from ..model import CellRegion, RangeRef, RegionCellRow, SheetDoc
from .utils import iter_cells_in_range, parse_range_ref, parse_sqref, rowcol_to_coord


def build_sheet_regions(sheet: SheetDoc) -> list[CellRegion]:
    base_ranges = sheet.print_areas[:] if sheet.print_areas else _fallback_base_ranges(sheet)
    if not base_ranges:
        return []

    def in_base(row: int, col: int) -> bool:
        for rng in base_ranges:
            if rng.start_row <= row <= rng.end_row and rng.start_col <= col <= rng.end_col:
                return True
        return False

    occupied: set[tuple[int, int]] = set()

    for cell in sheet.cells:
        if not in_base(cell.row, cell.col):
            continue
        if cell.value or cell.formula or (cell.style_id not in (None, "0")):
            occupied.add((cell.row, cell.col))

    for merge_rng in sheet.merges:
        for row, col in iter_cells_in_range(merge_rng):
            if in_base(row, col):
                occupied.add((row, col))

    dv_coords: set[tuple[int, int]] = set()
    for dv in sheet.data_validations:
        for rng in parse_sqref(dv.sqref):
            for row, col in iter_cells_in_range(rng):
                if in_base(row, col):
                    dv_coords.add((row, col))
                    occupied.add((row, col))

    if not occupied:
        return []

    components = _connected_components(occupied)
    components.sort(key=lambda comp: (min(r for r, _ in comp), min(c for _, c in comp)))

    regions: list[CellRegion] = []
    for region_idx, component in enumerate(components, start=1):
        min_row = min(row for row, _ in component)
        max_row = max(row for row, _ in component)
        min_col = min(col for _, col in component)
        max_col = max(col for _, col in component)

        bounds_ref = f"{rowcol_to_coord(min_row, min_col)}:{rowcol_to_coord(max_row, max_col)}"
        bounds = RangeRef(
            ref=bounds_ref,
            start_row=min_row,
            start_col=min_col,
            end_row=max_row,
            end_col=max_col,
        )

        rows: list[RegionCellRow] = []
        for row, col in sorted(component):
            coord = rowcol_to_coord(row, col)
            cell = sheet.cell_map.get(coord)
            merge_ref = sheet.merge_map.get(coord)
            flags: list[str] = []
            if merge_ref is not None:
                flags.append("merged")
            if (row, col) in dv_coords:
                flags.append("data_validation")

            if cell is None:
                rows.append(
                    RegionCellRow(
                        coord=coord,
                        value="",
                        formula=None,
                        cached_value=None,
                        cell_type="virtual",
                        style_id=None,
                        merge_ref=merge_ref,
                        flags=flags + ["virtual"],
                    )
                )
                continue

            if cell.value:
                flags.append("has_value")
            if cell.formula:
                flags.append("has_formula")
            if cell.cached_value not in (None, ""):
                flags.append("has_cached")
            if cell.style_id not in (None, "0"):
                flags.append("non_default_style")

            rows.append(
                RegionCellRow(
                    coord=coord,
                    value=cell.value,
                    formula=cell.formula,
                    cached_value=cell.cached_value,
                    cell_type=cell.cell_type,
                    style_id=cell.style_id,
                    merge_ref=merge_ref,
                    flags=flags,
                )
            )

        regions.append(CellRegion(region_id=region_idx, bounds=bounds, rows=rows))

    return regions


def _connected_components(points: set[tuple[int, int]]) -> list[set[tuple[int, int]]]:
    remaining = set(points)
    components: list[set[tuple[int, int]]] = []

    while remaining:
        start = remaining.pop()
        queue = deque([start])
        comp = {start}

        while queue:
            row, col = queue.popleft()
            for neighbor in ((row - 1, col), (row + 1, col), (row, col - 1), (row, col + 1)):
                if neighbor in remaining:
                    remaining.remove(neighbor)
                    comp.add(neighbor)
                    queue.append(neighbor)

        components.append(comp)

    return components


def _fallback_base_ranges(sheet: SheetDoc) -> list[RangeRef]:
    if sheet.dimension_ref:
        try:
            return [parse_range_ref(sheet.dimension_ref)]
        except ValueError:
            pass

    if not sheet.cells:
        return []

    min_row = min(cell.row for cell in sheet.cells)
    max_row = max(cell.row for cell in sheet.cells)
    min_col = min(cell.col for cell in sheet.cells)
    max_col = max(cell.col for cell in sheet.cells)
    ref = f"{rowcol_to_coord(min_row, min_col)}:{rowcol_to_coord(max_row, max_col)}"
    return [
        RangeRef(
            ref=ref,
            start_row=min_row,
            start_col=min_col,
            end_row=max_row,
            end_col=max_col,
        )
    ]
