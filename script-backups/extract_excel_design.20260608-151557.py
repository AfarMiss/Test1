#!/usr/bin/env python3
"""Extract structured content from one Excel design sheet."""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a_dml": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
SS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
OD_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
DOC_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def configure_stdio_utf8() -> None:
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")


def col_to_num(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + ord(char) - 64
    return value


def num_to_col(value: int) -> str:
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def cell_position(ref: str | None) -> tuple[int, int]:
    match = re.match(r"([A-Z]+)(\d+)", ref or "")
    if not match:
        return (999999, 999999)
    return (int(match.group(2)), col_to_num(match.group(1)))


def range_bounds(ref: str) -> dict:
    start, _, end = ref.partition(":")
    if not end:
        end = start
    start_row, start_col = cell_position(start)
    end_row, end_col = cell_position(end)
    return {
        "ref": ref,
        "start_row": min(start_row, end_row),
        "start_col": min(start_col, end_col),
        "end_row": max(start_row, end_row),
        "end_col": max(start_col, end_col),
    }


def range_ref(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    start = f"{num_to_col(start_col)}{start_row}"
    end = f"{num_to_col(end_col)}{end_row}"
    return start if start == end else f"{start}:{end}"


def text_from_si(node: ET.Element | None) -> str:
    if node is None:
        return ""
    return "".join(t.text or "" for t in node.iter(f"{SS}t"))


def sheet_path_from_target(target: str) -> str:
    if target.startswith("/"):
        target = target[1:]
    if target.startswith("xl/"):
        return target
    return "xl/" + target


def resolve_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    base = Path(base_part).parent
    normalized = (base / target).as_posix()
    parts: list[str] = []
    for part in normalized.split("/"):
        if part == "..":
            if parts:
                parts.pop()
        elif part and part != ".":
            parts.append(part)
    return "/".join(parts)


def rels_path_for(part: str) -> str:
    path = Path(part)
    return (path.parent / "_rels" / f"{path.name}.rels").as_posix()


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return [text_from_si(si) for si in root.findall("a:si", NS)]


def read_rels(zf: zipfile.ZipFile, rels_path: str) -> dict[str, str]:
    if rels_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rels_path))
    return {rel.attrib["Id"]: rel.attrib["Target"] for rel in root.findall(f"{REL_NS}Relationship")}


def read_style_features(zf: zipfile.ZipFile) -> dict[str, dict]:
    if "xl/styles.xml" not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read("xl/styles.xml"))
    fonts_node = root.find("a:fonts", NS)
    fills_node = root.find("a:fills", NS)
    cell_xfs_node = root.find("a:cellXfs", NS)
    fonts = list(fonts_node) if fonts_node is not None else []
    fills = list(fills_node) if fills_node is not None else []
    cell_xfs = list(cell_xfs_node) if cell_xfs_node is not None else []
    features = {}
    for index, xf in enumerate(cell_xfs):
        font_id = int(xf.attrib.get("fontId", "0"))
        fill_id = int(xf.attrib.get("fillId", "0"))
        font = fonts[font_id] if font_id < len(fonts) else None
        fill = fills[fill_id] if fill_id < len(fills) else None
        size_node = font.find("a:sz", NS) if font is not None else None
        bold_node = font.find("a:b", NS) if font is not None else None
        pattern_node = fill.find("a:patternFill", NS) if fill is not None else None
        pattern_type = pattern_node.attrib.get("patternType") if pattern_node is not None else None
        features[str(index)] = {
            "font_size": float(size_node.attrib["val"]) if size_node is not None and "val" in size_node.attrib else None,
            "bold": bold_node is not None,
            "filled": pattern_type not in (None, "none"),
        }
    return features


def is_dispimg_formula(value: str) -> bool:
    normalized = value.strip()
    normalized = normalized[1:].strip() if normalized.startswith("=") else normalized
    normalized = normalized.removeprefix("_xlfn.").strip()
    return bool(re.fullmatch(r"DISPIMG\s*\(.*\)", normalized, flags=re.IGNORECASE))


def cell_text(cell: ET.Element, shared: list[str]) -> str:
    formula_node = cell.find("a:f", NS)
    formula = formula_node.text if formula_node is not None and formula_node.text else ""

    text = ""
    if cell.attrib.get("t") == "inlineStr":
        text = text_from_si(cell.find("a:is", NS))
    else:
        value_node = cell.find("a:v", NS)
        if value_node is not None and value_node.text is not None:
            text = value_node.text
            if cell.attrib.get("t") == "s":
                try:
                    text = shared[int(text)]
                except (ValueError, IndexError):
                    pass

    text = html.unescape(str(text)).replace("\r", " ").replace("\n", " ").strip()
    formula = html.unescape(str(formula)).replace("\r", " ").replace("\n", " ").strip()
    if is_dispimg_formula(text) or is_dispimg_formula(formula):
        return ""
    if text:
        return text
    if formula:
        return f"[formula] {formula}"
    return ""


def normalize_sheet_name(name: str) -> str:
    return name.strip()


def parse_patterns(value: str) -> list[str]:
    patterns = [item.strip() for item in re.split(r"[,，]", value) if item.strip()]
    if not patterns:
        raise ValueError("Sheet pattern list cannot be empty.")
    return patterns


def select_sheet(sheets: list[dict], patterns: list[str]) -> dict | None:
    for pattern in patterns:
        for sheet in sheets:
            normalized = normalize_sheet_name(sheet["name"])
            if pattern in normalized:
                selected = sheet.copy()
                selected["matched_pattern"] = pattern
                return selected
    return None


def ensure_clean_output_dir(output_dir: Path, source_path: Path) -> None:
    resolved = output_dir.resolve()
    source_resolved = source_path.resolve()
    if str(resolved) in ("", str(resolved.anchor)):
        raise ValueError(f"Refusing to use unsafe output directory: {output_dir}")
    if resolved == source_resolved or resolved == source_resolved.parent:
        raise ValueError("Output directory cannot be the source file or source file directory.")
    if resolved.exists():
        if not resolved.is_dir():
            raise ValueError(f"Output path exists and is not a directory: {resolved}")
        shutil.rmtree(resolved)
    resolved.mkdir(parents=True, exist_ok=True)


def read_merged_ranges(root: ET.Element) -> list[dict]:
    ranges = []
    for node in root.findall(".//a:mergeCell", NS):
        ref = node.attrib.get("ref")
        if ref:
            ranges.append(range_bounds(ref))
    ranges.sort(key=lambda item: (item["start_row"], item["start_col"], item["end_row"], item["end_col"]))
    return ranges


def merge_for_cell(row: int, col: int, merged_ranges: list[dict]) -> dict | None:
    for merged in merged_ranges:
        if (
            merged["start_row"] <= row <= merged["end_row"]
            and merged["start_col"] <= col <= merged["end_col"]
        ):
            return merged
    return None


def cell_bounds(cell: dict) -> dict:
    merged = cell.get("merge")
    if merged:
        return {
            "start_row": merged["start_row"],
            "start_col": merged["start_col"],
            "end_row": merged["end_row"],
            "end_col": merged["end_col"],
        }
    return {
        "start_row": cell["row"],
        "start_col": cell["col"],
        "end_row": cell["row"],
        "end_col": cell["col"],
    }


def parse_cells(root: ET.Element, shared: list[str], merged_ranges: list[dict]) -> list[dict]:
    cells = []
    for cell in root.findall(".//a:c", NS):
        text = cell_text(cell, shared)
        if not text:
            continue
        ref = cell.attrib.get("r", "")
        row, col = cell_position(ref)
        merged = merge_for_cell(row, col, merged_ranges)
        bounds = cell_bounds({"row": row, "col": col, "merge": merged})
        cells.append(
            {
                "ref": ref,
                "row": row,
                "col": col,
                "text": text,
                "style": cell.attrib.get("s"),
                "merge_ref": merged["ref"] if merged else None,
                "bounds": bounds,
            }
        )
    cells.sort(key=lambda item: (item["row"], item["col"]))
    return cells


def rows_from_cells(cells: list[dict]) -> list[dict]:
    grouped: dict[int, list[dict]] = {}
    for cell in cells:
        grouped.setdefault(cell["row"], []).append(cell)

    rows = []
    for row in sorted(grouped):
        row_cells = sorted(grouped[row], key=lambda item: item["col"])
        rows.append(
            {
                "row": row,
                "cells": [
                    {
                        "ref": cell["ref"],
                        "row": cell["row"],
                        "col": cell["col"],
                        "text": cell["text"],
                        "style": cell.get("style"),
                        "merge_ref": cell.get("merge_ref"),
                        "bounds": cell.get("bounds"),
                    }
                    for cell in row_cells
                ],
            }
        )
    return rows


def style_counts(cells: list[dict]) -> dict[str | None, int]:
    counts: dict[str | None, int] = {}
    for cell in cells:
        style = cell.get("style")
        counts[style] = counts.get(style, 0) + 1
    return counts


def title_style_set(cells: list[dict], style_features: dict[str, dict]) -> set[str | None]:
    counts = style_counts(cells)
    if not counts:
        return set()
    body_style = max(counts, key=counts.get)
    body_size = (style_features.get(body_style or "") or {}).get("font_size")
    if body_size is None:
        sized_styles = [
            style
            for style in counts
            if (style_features.get(style or "") or {}).get("font_size") is not None
        ]
        if not sized_styles:
            return set()
        body_size = min(
            (style_features.get(style or "") or {}).get("font_size")
            for style in sized_styles
        )
    return {
        style
        for style in counts
        if (style_features.get(style or "") or {}).get("font_size") is not None
        and (style_features.get(style or "") or {}).get("font_size") > body_size
        and (
            (style_features.get(style or "") or {}).get("bold")
            or (style_features.get(style or "") or {}).get("filled")
        )
    }


def uses_fallback_title_styles(cells: list[dict], style_features: dict[str, dict]) -> bool:
    counts = style_counts(cells)
    if not counts:
        return False
    body_style = max(counts, key=counts.get)
    body_size = (style_features.get(body_style or "") or {}).get("font_size")
    return body_size is None


def is_title_row(row: dict, title_styles: set[str | None]) -> bool:
    cells = row.get("cells", [])
    if not cells:
        return False
    return all(cell.get("style") in title_styles for cell in cells if cell.get("text"))


def row_title_lines(row_block: dict, title_styles: set[str | None]) -> list[str]:
    if row_block.get("type") != "row":
        return []
    content_cells = [cell for cell in row_block.get("cells", []) if cell.get("text")]
    if not content_cells:
        return []
    if len(content_cells) != 1:
        return []
    if not all(cell.get("style") in title_styles for cell in content_cells):
        return []
    return [cell["text"] for cell in content_cells]


def row_block_from_cells(row: int, cells: list[dict]) -> dict | None:
    if not cells:
        return None
    return {
        "type": "row",
        "cells": cells,
        "lines": [cell["text"] for cell in cells if cell.get("text")],
        "source_refs": [cell["ref"] for cell in cells],
        "bounds": {
            "start_row": row,
            "end_row": row,
            "start_col": min(cell["col"] for cell in cells),
            "end_col": max(cell["col"] for cell in cells),
        },
    }


def title_cells(cells: list[dict], title_styles: set[str | None]) -> list[dict]:
    first_row = min((cell["row"] for cell in cells), default=None)
    return [
        cell
        for cell in cells
        if cell["row"] != first_row
        and cell.get("style") in title_styles
    ]


def cells_grouped_by_row(cells: list[dict]) -> dict[int, list[dict]]:
    grouped: dict[int, list[dict]] = {}
    for cell in cells:
        grouped.setdefault(cell["row"], []).append(cell)
    for row_cells in grouped.values():
        row_cells.sort(key=lambda item: item["col"])
    return grouped


def single_cell_rows(cells: list[dict]) -> list[dict]:
    return [
        row_cells[0]
        for row_cells in cells_grouped_by_row(cells).values()
        if len(row_cells) == 1
    ]


def cell_center_col(cell: dict) -> float:
    bounds = cell.get("bounds") or {}
    start_col = bounds.get("start_col", cell["col"])
    end_col = bounds.get("end_col", cell["col"])
    return (start_col + end_col) / 2


def cell_in_column_ranges(cell: dict, ranges: list[dict]) -> bool:
    center = cell_center_col(cell)
    return any(item["start_col"] <= center <= item["end_col"] for item in ranges)


def text_column_ranges_from_cells(cells: list[dict]) -> list[dict]:
    intervals = []
    for cell in cells:
        bounds = cell.get("bounds") or {}
        intervals.append(
            (
                bounds.get("start_col", cell["col"]),
                bounds.get("end_col", cell["col"]),
            )
        )
    return [
        {"col": start, "start_col": start, "end_col": end}
        for start, end in merge_intervals(intervals)
    ]


def recurring_isolated_title_cells(cells: list[dict], title_styles: set[str | None]) -> list[dict]:
    isolated = single_cell_rows(cells)
    if not isolated:
        return []
    isolated_refs = {cell["ref"] for cell in isolated}
    body_cells = [
        cell
        for cell in cells
        if cell["ref"] not in isolated_refs and cell.get("style") not in title_styles
    ]
    body_ranges = text_column_ranges_from_cells(body_cells)
    if not body_ranges:
        return []

    grouped: dict[tuple[int, str | None], list[dict]] = {}
    for cell in isolated:
        if cell.get("style") in title_styles or cell_in_column_ranges(cell, body_ranges):
            continue
        grouped.setdefault((cell["col"], cell.get("style")), []).append(cell)

    titles = []
    for group in grouped.values():
        if len(group) > 1:
            titles.extend(group)
    return sorted(titles, key=lambda cell: (cell["row"], cell["col"]))


def section_title_cells(cells: list[dict], title_styles: set[str | None]) -> list[dict]:
    by_ref: dict[str, dict] = {}
    styled_titles = title_cells(cells, title_styles)
    if styled_titles:
        return sorted(styled_titles, key=lambda cell: (cell["row"], cell["col"]))
    for cell in recurring_isolated_title_cells(cells, title_styles):
        by_ref.setdefault(cell["ref"], cell)
    return sorted(by_ref.values(), key=lambda cell: (cell["row"], cell["col"]))


def has_recurring_isolated_titles(cells: list[dict], title_styles: set[str | None]) -> bool:
    return not title_cells(cells, title_styles) and bool(recurring_isolated_title_cells(cells, title_styles))


def column_range_keys_for_cells(cells: list[dict], ranges: list[dict]) -> list[tuple[int, int]]:
    return sorted(
        {
            (
                column_range_for_col(cell["col"], ranges)["start_col"],
                column_range_for_col(cell["col"], ranges)["end_col"],
            )
            for cell in cells
        }
    )


def merge_intervals(intervals: list[tuple[int, int]]) -> list[tuple[int, int]]:
    if not intervals:
        return []
    ordered = sorted(intervals)
    merged = [ordered[0]]
    for start, end in ordered[1:]:
        prev_start, prev_end = merged[-1]
        if start <= prev_end:
            merged[-1] = (prev_start, max(prev_end, end))
        else:
            merged.append((start, end))
    return merged


def layout_column_ranges(cells: list[dict], images: list[dict], title_styles: set[str | None]) -> list[dict]:
    intervals: list[tuple[int, int]] = []
    for cell in cells:
        if cell.get("style") in title_styles:
            continue
        bounds = cell.get("bounds") or {}
        intervals.append((bounds.get("start_col", cell["col"]), bounds.get("end_col", cell["col"])))
    for image in images:
        anchor = image.get("anchor") or {}
        start_col = anchor.get("from_col")
        end_col = anchor.get("to_col") or start_col
        if start_col is not None and end_col is not None:
            intervals.append((min(start_col, end_col), max(start_col, end_col)))

    merged = merge_intervals(intervals)
    if len(merged) <= 1:
        return [{"col": merged[0][0] if merged else 1, "start_col": 1, "end_col": 999999}]

    gaps = [
        (merged[index][1] + 1, merged[index + 1][0] - 1)
        for index in range(len(merged) - 1)
        if merged[index][1] + 1 <= merged[index + 1][0] - 1
    ]
    if not gaps:
        return [{"col": merged[0][0], "start_col": merged[0][0], "end_col": merged[-1][1]}]

    max_gap_width = max(end - start + 1 for start, end in gaps)
    separator_gaps = {(start, end) for start, end in gaps if end - start + 1 == max_gap_width}

    ranges = []
    current_start, current_end = merged[0]
    for index in range(len(merged) - 1):
        gap = (merged[index][1] + 1, merged[index + 1][0] - 1)
        if gap in separator_gaps:
            ranges.append({"col": current_start, "start_col": current_start, "end_col": current_end})
            current_start, current_end = merged[index + 1]
        else:
            current_end = merged[index + 1][1]
    ranges.append({"col": current_start, "start_col": current_start, "end_col": current_end})
    return ranges


def text_flow_ranges_from_rows(rows: dict[int, list[dict]]) -> list[dict]:
    intervals: list[tuple[int, int]] = []
    row_flow_candidates: list[list[tuple[int, int]]] = []
    for row_cells in rows.values():
        row_intervals = []
        for cell in row_cells:
            bounds = cell.get("bounds") or {}
            row_intervals.append(
                (
                    bounds.get("start_col", cell["col"]),
                    bounds.get("end_col", cell["col"]),
                )
            )
        numbered_cells = [
            cell
            for cell in row_cells
            if leading_number_token(cell.get("text", "")) is not None
            and re.fullmatch(r"\s*\d+\s*", cell.get("text", ""))
        ]
        if len(numbered_cells) > 1:
            sorted_numbered = sorted(numbered_cells, key=lambda cell: cell_center_col(cell))
            row_ranges = []
            for index, cell in enumerate(sorted_numbered):
                start_col = cell["col"]
                end_col = (
                    sorted_numbered[index + 1]["col"] - 1
                    if index + 1 < len(sorted_numbered)
                    else max(end for _, end in row_intervals)
                )
                row_ranges.append((start_col, end_col))
            row_flow_candidates.append(row_ranges)

        row_merged = merge_intervals(row_intervals)
        intervals.extend(row_merged)
    if row_flow_candidates:
        flattened = [item for row_ranges in row_flow_candidates for item in row_ranges]
        candidate_ranges = merge_intervals(flattened)
        if len(candidate_ranges) > 1:
            max_col = max(end for _, end in intervals) if intervals else candidate_ranges[-1][1]
            expanded_ranges = []
            for index, (start, end) in enumerate(candidate_ranges):
                next_start = candidate_ranges[index + 1][0] if index + 1 < len(candidate_ranges) else max_col + 1
                expanded_ranges.append((start, max(end, next_start - 1)))
            candidate_ranges = expanded_ranges
            return [
                {"col": start, "start_col": start, "end_col": end}
                for start, end in candidate_ranges
            ]
    merged = merge_intervals(intervals)
    if len(merged) <= 1:
        return []
    gaps = [
        (merged[index][1] + 1, merged[index + 1][0] - 1)
        for index in range(len(merged) - 1)
        if merged[index][1] + 1 <= merged[index + 1][0] - 1
    ]
    if not gaps:
        return []
    separator_gaps = set(gaps)
    if not separator_gaps:
        return []
    ranges = []
    current_start, current_end = merged[0]
    for index in range(len(merged) - 1):
        gap = (merged[index][1] + 1, merged[index + 1][0] - 1)
        if gap in separator_gaps:
            ranges.append({"col": current_start, "start_col": current_start, "end_col": current_end})
            current_start, current_end = merged[index + 1]
        else:
            current_end = merged[index + 1][1]
    ranges.append({"col": current_start, "start_col": current_start, "end_col": current_end})
    if len(ranges) <= 1:
        return []

    synchronized_number_rows = 0
    for row_cells in rows.values():
        numbered_ranges = set()
        for flow_range in ranges:
            cells_in_range = [
                cell
                for cell in row_cells
                if flow_range["start_col"] <= cell_center_col(cell) <= flow_range["end_col"]
            ]
            if any(leading_number_token(cell.get("text", "")) is not None for cell in cells_in_range):
                numbered_ranges.add((flow_range["start_col"], flow_range["end_col"]))
        if len(numbered_ranges) > 1:
            synchronized_number_rows += 1
    return ranges if synchronized_number_rows > 1 else []


def range_has_numbered_text(rows: dict[int, list[dict]], flow_range: dict) -> bool:
    for row_cells in rows.values():
        cells_in_range = [
            cell
            for cell in row_cells
            if flow_range["start_col"] <= cell_center_col(cell) <= flow_range["end_col"]
        ]
        if any(leading_number_token(cell.get("text", "")) is not None for cell in cells_in_range):
            return True
    return False


def expand_flow_ranges_to_cover_cells(flow_ranges: list[dict], cells: list[dict]) -> list[dict]:
    if not flow_ranges or not cells:
        return flow_ranges
    sorted_ranges = sorted(flow_ranges, key=lambda item: (item["start_col"], item["end_col"]))
    min_col = min((cell.get("bounds") or {}).get("start_col", cell["col"]) for cell in cells)
    max_col = max((cell.get("bounds") or {}).get("end_col", cell["col"]) for cell in cells)
    expanded = []
    for index, flow_range in enumerate(sorted_ranges):
        start_col = min_col if index == 0 else flow_range["start_col"]
        end_col = (
            sorted_ranges[index + 1]["start_col"] - 1
            if index + 1 < len(sorted_ranges)
            else max_col
        )
        expanded.append(
            {
                "col": flow_range["col"],
                "start_col": start_col,
                "end_col": max(end_col, flow_range["end_col"]),
            }
        )
    return expanded


def column_range_for_col(col: int, ranges: list[dict]) -> dict:
    for item in ranges:
        if item["start_col"] <= col <= item["end_col"]:
            return item
    return min(
        ranges,
        key=lambda item: min(abs(col - item["start_col"]), abs(col - item["end_col"])),
    )


def sanitize_cell_for_output(cell: dict) -> dict:
    return {
        "ref": cell["ref"],
        "row": cell["row"],
        "col": cell["col"],
        "text": cell["text"],
        "style": cell.get("style"),
        "merge_ref": cell.get("merge_ref"),
        "bounds": cell.get("bounds"),
    }


def clone_cell_for_row(cell: dict) -> dict:
    return sanitize_cell_for_output(cell)


def logical_lines_for_cells(cells: list[dict]) -> list[str]:
    line = " ".join(cell["text"].strip() for cell in cells if cell.get("text", "").strip())
    return [line] if line else []


def list_item_number(text: str) -> str | None:
    match = re.match(r"^\s*(\d+)[.．、]", text)
    return match.group(1) if match else None


def leading_number_token(text: str) -> str | None:
    match = re.match(r"^\s*(\d+)(?:[.．、]|\s+|$)", text)
    return match.group(1) if match else None


def block_from_cells(row: int, cells: list[dict]) -> dict | None:
    public_cells = [clone_cell_for_row(cell) for cell in cells]
    if not public_cells:
        return None
    return {
        "type": "row",
        "cells": public_cells,
        "lines": logical_lines_for_cells(public_cells),
        "source_refs": [cell["ref"] for cell in public_cells],
        "bounds": {
            "start_row": row,
            "end_row": row,
            "start_col": min(cell["col"] for cell in public_cells),
            "end_col": max(cell["col"] for cell in public_cells),
        },
    }


def table_value_for_cell(row: int, col: int, cells_by_position: dict[tuple[int, int], dict]) -> str:
    cell = cells_by_position.get((row, col))
    if cell is not None:
        return cell.get("text", "")
    for candidate in cells_by_position.values():
        bounds = candidate.get("bounds") or {}
        if (
            bounds.get("start_row", candidate.get("row")) <= row <= bounds.get("end_row", candidate.get("row"))
            and bounds.get("start_col", candidate.get("col")) <= col <= bounds.get("end_col", candidate.get("col"))
        ):
            return candidate.get("text", "")
    return ""


def table_block_from_grid(
    rows: list[int],
    cols: list[int],
    cells_by_position: dict[tuple[int, int], dict],
    flow_order: int | None = None,
    title: str | None = None,
) -> dict:
    header = [
        table_value_for_cell(rows[0], col, cells_by_position)
        for col in cols
    ]
    body = [
        [
            table_value_for_cell(row, col, cells_by_position)
            for col in cols
        ]
        for row in rows[1:]
    ]
    source_refs = []
    for row in rows:
        for col in cols:
            cell = cells_by_position.get((row, col))
            if cell is not None:
                source_refs.append(cell["ref"])
    block = {
        "type": "table",
        "header": header,
        "rows": body,
        "source_refs": source_refs,
        "bounds": {
            "start_row": min(rows),
            "end_row": max(rows),
            "start_col": min(cols),
            "end_col": max(cols),
        },
    }
    if title:
        block["title"] = title
    if flow_order is not None:
        block["_flow_order"] = flow_order
    return block


def row_cell_count_in_cols(row_cells: list[dict], cols: list[int]) -> int:
    return sum(1 for cell in row_cells if cell["col"] in cols)


def table_caption_for_cols(row_number: int, cols: list[int], rows: dict[int, list[dict]]) -> dict | None:
    previous_cells = rows.get(row_number - 1, [])
    candidates = [
        cell
        for cell in previous_cells
        if min(cols) <= cell["col"] <= max(cols)
    ]
    if len(candidates) == 1:
        return candidates[0]
    return None


def detect_table_at_row(row_index: int, row_numbers: list[int], rows: dict[int, list[dict]]) -> tuple[list[int], list[int], dict | None] | None:
    row_number = row_numbers[row_index]
    row_cells = rows.get(row_number, [])
    if len(row_cells) < 2:
        return None

    cells_by_col = {cell["col"]: cell for cell in row_cells}
    sorted_cols = sorted(cells_by_col)
    best_cols: list[int] = []
    best_rows: list[int] = []
    best_caption: dict | None = None

    for start_index in range(len(sorted_cols)):
        cols = [sorted_cols[start_index]]
        for next_index in range(start_index + 1, len(sorted_cols)):
            if sorted_cols[next_index] != cols[-1] + 1:
                break
            cols.append(sorted_cols[next_index])
            if len(cols) < 3:
                continue
            table_rows = [row_number]
            scan_index = row_index + 1
            while scan_index < len(row_numbers):
                next_row = row_numbers[scan_index]
                next_cells = rows.get(next_row, [])
                count = row_cell_count_in_cols(next_cells, cols)
                if count == 0:
                    break
                table_rows.append(next_row)
                scan_index += 1
            if len(table_rows) > 2 and (len(cols), len(table_rows)) > (len(best_cols), len(best_rows)):
                best_cols = list(cols)
                best_rows = list(table_rows)
                best_caption = table_caption_for_cols(row_number, cols, rows)

    if best_cols and best_rows:
        return best_rows, best_cols, best_caption
    return None


def append_blocks_from_rows(section: dict, rows: dict[int, list[dict]], flow_order: int | None = None) -> None:
    row_numbers = sorted(rows)
    cells_by_position = {
        (cell["row"], cell["col"]): cell
        for row_cells in rows.values()
        for cell in row_cells
    }
    index = 0
    while index < len(row_numbers):
        if index + 1 < len(row_numbers):
            next_detected = detect_table_at_row(index + 1, row_numbers, rows)
            if next_detected is not None:
                _, _, next_caption = next_detected
                if next_caption and next_caption["row"] == row_numbers[index]:
                    index += 1
                    continue
        detected = detect_table_at_row(index, row_numbers, rows)
        if detected is None:
            row_number = row_numbers[index]
            block = block_from_cells(row_number, rows[row_number])
            if block is not None:
                if flow_order is not None:
                    block["_flow_order"] = flow_order
                append_to_section(section, block)
            index += 1
            continue

        table_rows, table_cols, table_caption = detected
        table_col_set = set(table_cols)
        pending_text_blocks: list[dict] = []
        caption_ref = table_caption.get("ref") if table_caption else None
        if table_caption and table_caption["row"] in rows:
            text_cells = [
                cell
                for cell in rows[table_caption["row"]]
                if cell.get("ref") != caption_ref
            ]
            block = block_from_cells(table_caption["row"], text_cells)
            if block is not None:
                if flow_order is not None:
                    block["_flow_order"] = flow_order
                pending_text_blocks.append(block)
        for row_number in table_rows:
            text_cells = [
                cell
                for cell in rows[row_number]
                if cell["col"] not in table_col_set
            ]
            block = block_from_cells(row_number, text_cells)
            if block is not None:
                if flow_order is not None:
                    block["_flow_order"] = flow_order
                pending_text_blocks.append(block)
        for block in pending_text_blocks:
            append_to_section(section, block)
        append_to_section(
            section,
            table_block_from_grid(
                table_rows,
                table_cols,
                cells_by_position,
                flow_order,
                table_caption.get("text") if table_caption else None,
            ),
        )
        index += len(table_rows)


def image_block(image: dict) -> dict:
    anchor = image.get("anchor") or {}
    return {
        "type": "image",
        "image_indexes": [image["index"]],
        "bounds": {
            "start_row": anchor.get("from_row") or 0,
            "start_col": anchor.get("from_col") or 0,
            "end_row": anchor.get("to_row") or anchor.get("from_row") or 0,
            "end_col": anchor.get("to_col") or anchor.get("from_col") or 0,
        },
    }


def make_section(title: str | None, title_cell: dict | None = None) -> dict:
    bounds = None
    refs = []
    if title_cell is not None:
        refs = [title_cell["ref"]]
        bounds = {
            "start_row": title_cell["row"],
            "end_row": title_cell["row"],
            "start_col": title_cell["col"],
            "end_col": title_cell["col"],
        }
    return {
        "type": "section",
        "title": title,
        "blocks": [],
        "source_refs": refs,
        "bounds": bounds,
    }


def append_to_section(section: dict, block: dict) -> None:
    section["blocks"].append(block)
    section["source_refs"].extend(block_source_refs(block))
    bounds = block_bounds(block)
    if section["bounds"] is None:
        section["bounds"] = bounds.copy()
    else:
        section["bounds"] = {
            "start_row": min(section["bounds"]["start_row"], bounds["start_row"]),
            "start_col": min(section["bounds"]["start_col"], bounds["start_col"]),
            "end_row": max(section["bounds"]["end_row"], bounds["end_row"]),
            "end_col": max(section["bounds"]["end_col"], bounds["end_col"]),
        }


def remove_title_cells(cells: list[dict], title_styles: set[str | None]) -> list[dict]:
    return [cell for cell in cells if cell.get("style") not in title_styles]


def remove_cells_by_ref(cells: list[dict], refs: set[str]) -> list[dict]:
    return [cell for cell in cells if cell["ref"] not in refs]


def group_cells_for_sections(
    cells: list[dict],
    title_styles: set[str | None],
    images: list[dict],
    fallback_title_styles: bool = False,
) -> list[dict]:
    title_candidates = section_title_cells(cells, title_styles)
    has_isolated_titles = has_recurring_isolated_titles(cells, title_styles)
    ranges = layout_column_ranges(cells, images, title_styles)
    rows_by_number = cells_grouped_by_row(cells)

    def first_row_layout_title_cells() -> list[dict]:
        first_row = min((cell["row"] for cell in cells), default=None)
        if first_row is None:
            return []
        first_row_cells = [
            cell
            for cell in cells
            if cell["row"] == first_row and cell.get("text")
        ]
        if len(first_row_cells) <= 1:
            return []
        title_refs = {cell["ref"] for cell in title_candidates}
        candidates = []
        for cell in first_row_cells:
            if cell["ref"] in title_refs:
                continue
            col_range = column_range_for_col(cell["col"], ranges)
            range_key = (col_range["start_col"], col_range["end_col"])
            has_body = any(
                body_cell["row"] > cell["row"]
                and range_key[0] <= cell_center_col(body_cell) <= range_key[1]
                for body_cell in cells
            )
            if has_body:
                candidates.append(cell)
        return candidates

    def image_anchor_title_cells() -> list[dict]:
        candidates_by_ref: dict[str, dict] = {}
        for image in images:
            anchor = image.get("anchor") or {}
            image_row = anchor.get("from_row")
            image_col = anchor.get("from_col")
            if image_row is None or image_col is None:
                continue
            col_range = column_range_for_col(image_col, ranges)
            range_key = (col_range["start_col"], col_range["end_col"])
            prior_rows = [
                row_number
                for row_number, row_cells in rows_by_number.items()
                if row_number < image_row
                and any(range_key[0] <= cell_center_col(cell) <= range_key[1] for cell in row_cells)
            ]
            for row_number in sorted(prior_rows, reverse=True):
                row_cells = [
                    cell
                    for cell in rows_by_number[row_number]
                    if range_key[0] <= cell_center_col(cell) <= range_key[1]
                    and cell.get("text")
                ]
                if len(row_cells) == 1:
                    candidates_by_ref.setdefault(row_cells[0]["ref"], row_cells[0])
                    break
        return sorted(candidates_by_ref.values(), key=lambda cell: (cell["row"], cell["col"]))

    layout_title_refs: set[str] = set()
    if not title_candidates:
        layout_title_candidates = first_row_layout_title_cells()
        layout_title_refs = {cell["ref"] for cell in layout_title_candidates}
        title_candidates = layout_title_candidates + image_anchor_title_cells()

    if not title_candidates:
        section = make_section(None)
        for row in rows_from_cells(cells):
            block = block_from_cells(row["row"], row["cells"])
            if block is not None:
                append_to_section(section, block)
        for image in images:
            append_to_section(section, image_block(image))
        return [section]

    def first_row_column_title_cells() -> list[dict]:
        first_row = min((cell["row"] for cell in cells), default=None)
        if first_row is None:
            return []
        first_row_titles = [
            cell
            for cell in cells
            if cell["row"] == first_row and cell.get("style") in title_styles and cell.get("text")
        ]
        if len(first_row_titles) <= 1:
            return []
        first_row_titles.sort(key=lambda cell: cell["col"])
        document_title = first_row_titles[0]
        document_title_ref = document_title["ref"]
        document_title_text = document_title.get("text", "")
        existing_refs = {cell["ref"] for cell in title_candidates}
        candidates = []
        for cell in first_row_titles:
            if cell["ref"] == document_title_ref or cell["ref"] in existing_refs:
                continue
            shared_prefix = ""
            for left_char, right_char in zip(document_title_text, cell.get("text", "")):
                if left_char != right_char:
                    break
                shared_prefix += left_char
            if not shared_prefix or len(shared_prefix) * 2 < len(document_title_text):
                continue
            title_range = column_range_for_col(cell["col"], ranges)
            range_key = (title_range["start_col"], title_range["end_col"])
            later_same_range_title_rows = [
                title["row"]
                for title in title_candidates
                if title["row"] > cell["row"]
                and range_key
                == (
                    column_range_for_col(title["col"], ranges)["start_col"],
                    column_range_for_col(title["col"], ranges)["end_col"],
                )
            ]
            end_row = min(later_same_range_title_rows) - 1 if later_same_range_title_rows else 999999
            has_body = False
            for row_cells in cells_grouped_by_row(cells).values():
                for body_cell in row_cells:
                    if body_cell["row"] <= cell["row"] or body_cell["row"] > end_row:
                        continue
                    if body_cell.get("style") in title_styles:
                        continue
                    if range_key[0] <= cell_center_col(body_cell) <= range_key[1]:
                        has_body = True
                        break
                if has_body:
                    break
            if has_body:
                candidates.append(cell)
        return candidates

    title_candidates = title_candidates + first_row_column_title_cells()
    sorted_title_candidates = sorted(title_candidates, key=lambda cell: (cell["row"], cell["col"]))
    candidate_title_refs = {title["ref"] for title in sorted_title_candidates}
    cells_by_row_without_candidate_titles = cells_grouped_by_row(
        remove_cells_by_ref(cells, candidate_title_refs)
    )

    def title_range_key(title: dict) -> tuple[int, int]:
        title_range = column_range_for_col(title["col"], ranges)
        return (title_range["start_col"], title_range["end_col"])

    def title_has_following_content(title: dict, next_title: dict | None) -> bool:
        range_key = title_range_key(title)
        end_row = next_title["row"] - 1 if next_title else 999999
        for row_number, row_cells in cells_by_row_without_candidate_titles.items():
            if not (title["row"] < row_number <= end_row):
                continue
            if any(range_key[0] <= cell_center_col(cell) <= range_key[1] for cell in row_cells):
                return True
        return False

    def title_has_following_text_in_own_range(title: dict, next_title: dict | None) -> bool:
        title_bounds = title.get("bounds") or {}
        start_col = title_bounds.get("start_col", title["col"])
        end_col = title_bounds.get("end_col", title["col"])
        end_row = next_title["row"] - 1 if next_title else 999999
        for row_number, row_cells in cells_by_row_without_candidate_titles.items():
            if not (title["row"] < row_number <= end_row):
                continue
            if any(start_col <= cell_center_col(cell) <= end_col for cell in row_cells):
                return True
        return False

    def text_extends_beyond_title_image_band(title: dict, next_title: dict | None) -> bool:
        title_bounds = title.get("bounds") or {}
        title_start_col = title_bounds.get("start_col", title["col"])
        title_end_col = title_bounds.get("end_col", title["col"])
        all_range_keys = {(range_item["start_col"], range_item["end_col"]) for range_item in ranges}
        other_ranges = [
            item
            for item in all_range_keys
            if not (item[0] <= title_start_col and title_end_col <= item[1])
        ]
        if not other_ranges:
            other_ranges = list(all_range_keys)
        image_bottom_rows = []
        for image in images:
            anchor = image.get("anchor") or {}
            image_row = anchor.get("from_row") or 0
            if image_row <= title["row"]:
                continue
            if next_title and image_row >= next_title["row"]:
                continue
            image_bottom_rows.append(anchor.get("to_row") or image_row)
        if not image_bottom_rows:
            return False
        image_bottom = max(image_bottom_rows)
        end_row = next_title["row"] - 1 if next_title else 999999
        for row_number, row_cells in cells_by_row_without_candidate_titles.items():
            if not (image_bottom < row_number <= end_row):
                continue
            if any(
                range_key[0] <= cell_center_col(cell) <= range_key[1]
                for cell in row_cells
                for range_key in other_ranges
            ):
                return True
        return False

    def title_has_following_text_in_image_range(title: dict, next_title: dict | None) -> bool:
        image_ranges = []
        for image in images:
            anchor = image.get("anchor") or {}
            image_row = anchor.get("from_row") or 0
            if image_row <= title["row"]:
                continue
            if next_title and image_row >= next_title["row"]:
                continue
            start_col = anchor.get("from_col") or 0
            end_col = anchor.get("to_col") or start_col
            if start_col and end_col:
                image_ranges.append((min(start_col, end_col), max(start_col, end_col)))
        if not image_ranges:
            return False
        end_row = next_title["row"] - 1 if next_title else 999999
        for row_number, row_cells in cells_by_row_without_candidate_titles.items():
            if not (title["row"] < row_number <= end_row):
                continue
            if any(
                image_range[0] <= cell_center_col(cell) <= image_range[1]
                for cell in row_cells
                for image_range in image_ranges
            ):
                return True
        return False

    def title_has_following_number_start(title: dict, next_title: dict | None) -> bool:
        end_row = next_title["row"] - 1 if next_title else 999999
        for row_number, row_cells in cells_by_row_without_candidate_titles.items():
            if not (title["row"] < row_number <= end_row):
                continue
            text = " ".join(cell["text"].strip() for cell in row_cells if cell.get("text", "").strip())
            if leading_number_token(text) == "1":
                return True
        return False

    def numbered_values_between(
        range_key: tuple[int, int],
        start_row: int,
        end_row: int,
    ) -> list[int]:
        numbers = []
        for row_number, row_cells in cells_by_row_without_candidate_titles.items():
            if not (start_row < row_number <= end_row):
                continue
            row_range_cells = [
                cell
                for cell in row_cells
                if range_key[0] <= cell_center_col(cell) <= range_key[1]
            ]
            text = " ".join(cell["text"].strip() for cell in row_range_cells if cell.get("text", "").strip())
            number = leading_number_token(text)
            if number is not None:
                numbers.append(int(number))
        return numbers

    def is_numbered_continuation_title(title: dict, previous_title: dict | None, next_title: dict | None) -> bool:
        if previous_title is None:
            return False
        all_range_keys = {(range_item["start_col"], range_item["end_col"]) for range_item in ranges}
        candidate_ranges = [
            item
            for item in all_range_keys
            if item != title_range_key(title)
        ]
        if not candidate_ranges:
            candidate_ranges = list(all_range_keys)
        end_row = next_title["row"] - 1 if next_title else 999999
        for range_key in candidate_ranges:
            before_numbers = numbered_values_between(range_key, previous_title["row"], title["row"] - 1)
            after_numbers = numbered_values_between(range_key, title["row"], end_row)
            if before_numbers and after_numbers and after_numbers[0] == max(before_numbers) + 1:
                return True
        return False

    demoted_title_refs: set[str] = set()
    continuation_title_refs: set[str] = set()
    for index, title in enumerate(sorted_title_candidates):
        if index == 0:
            continue
        next_title = (
            sorted_title_candidates[index + 1]
            if index + 1 < len(sorted_title_candidates)
            else None
        )
        previous_titles = [
            candidate
            for candidate in sorted_title_candidates[:index]
            if title_range_key(candidate) != title_range_key(title)
            and title_has_following_content(candidate, title)
        ]
        previous_title = sorted_title_candidates[index - 1]
        numbered_continuation = (
            is_numbered_continuation_title(title, previous_title, next_title)
            and not title_has_following_text_in_own_range(title, next_title)
            and not title_has_following_text_in_image_range(title, next_title)
            and not title_has_following_number_start(title, next_title)
            and not text_extends_beyond_title_image_band(title, next_title)
        )
        no_own_text_continuation = (
            previous_titles
            and not title_has_following_content(title, next_title)
            and not title_has_following_text_in_own_range(title, next_title)
        )
        if fallback_title_styles and no_own_text_continuation:
            demoted_title_refs.add(title["ref"])
        if numbered_continuation and previous_title["ref"] not in continuation_title_refs:
            continuation_title_refs.add(title["ref"])

    sorted_titles = [
        title
        for title in sorted_title_candidates
        if title["ref"] not in demoted_title_refs
    ]
    demoted_titles_by_row: dict[int, list[dict]] = {}
    for title in sorted_title_candidates:
        if title["ref"] in demoted_title_refs:
            demoted_titles_by_row.setdefault(title["row"], []).append(title)
    continuation_targets_by_ref: dict[str, dict] = {}
    continuation_ranges_by_ref: dict[str, set[tuple[int, int]]] = {}
    continuation_end_row_by_ref: dict[str, int] = {}
    for index, title in enumerate(sorted_title_candidates):
        if title["ref"] not in continuation_title_refs or index == 0:
            continue
        previous_title = sorted_title_candidates[index - 1]
        next_title = (
            sorted_title_candidates[index + 1]
            if index + 1 < len(sorted_title_candidates)
            else None
        )
        end_row = next_title["row"] - 1 if next_title else 999999
        candidate_ranges = [
            item
            for item in {(range_item["start_col"], range_item["end_col"]) for range_item in ranges}
            if item != title_range_key(title)
        ]
        if not candidate_ranges:
            candidate_ranges = [
                (range_item["start_col"], range_item["end_col"])
                for range_item in ranges
            ]
        matched_ranges = set()
        for range_key in candidate_ranges:
            before_numbers = numbered_values_between(range_key, previous_title["row"], title["row"] - 1)
            after_numbers = numbered_values_between(range_key, title["row"], end_row)
            if before_numbers and after_numbers and after_numbers[0] == max(before_numbers) + 1:
                matched_ranges.add(range_key)
        if matched_ranges:
            continuation_targets_by_ref[title["ref"]] = previous_title
            continuation_ranges_by_ref[title["ref"]] = matched_ranges
            continuation_end_row_by_ref[title["ref"]] = end_row

    sections = [make_section(None)]
    section_by_title: dict[str, dict] = {}
    title_rows: dict[int, list[dict]] = {}
    for title in sorted_titles:
        title_rows.setdefault(title["row"], []).append(title)
    title_refs = {title["ref"] for title in sorted_titles}

    all_range_keys = {(item["start_col"], item["end_col"]) for item in ranges}
    titled_range_keys = set()
    for title in sorted_titles:
        title_range = column_range_for_col(title["col"], ranges)
        titled_range_keys.add((title_range["start_col"], title_range["end_col"]))
    untitled_range_keys = all_range_keys - titled_range_keys

    title_range_keys_by_ref: dict[str, set[tuple[int, int]]] = {}
    for title in sorted_titles:
        section = make_section(title["text"], title)
        sections.append(section)
        section_by_title[title["ref"]] = section
        title_range = column_range_for_col(title["col"], ranges)
        title_range_keys_by_ref[title["ref"]] = {
            (title_range["start_col"], title_range["end_col"])
        }
        if len(title_rows[title["row"]]) == 1:
            title_range_keys_by_ref[title["ref"]].update(untitled_range_keys)

    title_band_end_by_ref: dict[str, int] = {}
    for index, title in enumerate(sorted_titles):
        later_titles = [
            candidate
            for candidate in sorted_titles[index + 1 :]
            if candidate["row"] > title["row"]
        ]
        if later_titles:
            title_band_end_by_ref[title["ref"]] = later_titles[0]["row"] - 1
        else:
            title_band_end_by_ref[title["ref"]] = 999999

    image_rows_by_title_ref: dict[str, list[tuple[int, int]]] = {}
    visual_rows_by_title_ref: dict[str, list[tuple[int, int]]] = {}
    image_bounds_by_title_ref: dict[str, list[dict]] = {}
    for image in images:
        anchor = image.get("anchor") or {}
        row = anchor.get("from_row")
        col = anchor.get("from_col")
        if row is None or col is None:
            continue
        col_range = column_range_for_col(col, ranges)
        range_key = (col_range["start_col"], col_range["end_col"])
        candidates = [
            title
            for title in sorted_titles
            if range_key in title_range_keys_by_ref[title["ref"]] and title["row"] <= row
        ]
        if not candidates:
            continue
        title = max(candidates, key=lambda cell: cell["row"])
        end_row = anchor.get("to_row") or row
        end_col = anchor.get("to_col") or col
        image_rows_by_title_ref.setdefault(title["ref"], []).append(
            (min(row, end_row), max(row, end_row))
        )
        visual_rows_by_title_ref.setdefault(title["ref"], []).append(
            (title["row"] + 1, max(row, end_row))
        )
        image_bounds_by_title_ref.setdefault(title["ref"], []).append(
            {
                "start_row": min(row, end_row),
                "end_row": max(row, end_row),
                "start_col": min(col, end_col),
                "end_col": max(col, end_col),
            }
        )

    row_number_by_title_ref: dict[str, dict[int, set[str]]] = {}
    row_number_by_range_key: dict[tuple[int, int], dict[int, set[str]]] = {}
    content_rows_by_title_ref: dict[str, set[int]] = {}
    content_cells_by_row = cells_grouped_by_row(remove_cells_by_ref(cells, title_refs))
    content_rows_by_range_key: dict[tuple[int, int], set[int]] = {}
    for range_item in ranges:
        range_key = (range_item["start_col"], range_item["end_col"])
        for row_number, row_cells in content_cells_by_row.items():
            row_range_cells = [
                cell
                for cell in row_cells
                if range_key[0] <= cell_center_col(cell) <= range_key[1]
            ]
            if any(cell.get("text", "").strip() for cell in row_range_cells):
                content_rows_by_range_key.setdefault(range_key, set()).add(row_number)
            number = list_item_number(" ".join(cell["text"].strip() for cell in row_range_cells if cell.get("text", "").strip()))
            if number:
                row_number_by_range_key.setdefault(range_key, {}).setdefault(row_number, set()).add(number)
    for title in sorted_titles:
        title_range = column_range_for_col(title["col"], ranges)
        range_key = (title_range["start_col"], title_range["end_col"])
        for row_number, row_cells in content_cells_by_row.items():
            if not (title["row"] < row_number <= title_band_end_by_ref[title["ref"]]):
                continue
            row_range_cells = [
                cell
                for cell in row_cells
                if range_key[0] <= cell_center_col(cell) <= range_key[1]
            ]
            if any(cell.get("text", "").strip() for cell in row_range_cells):
                content_rows_by_title_ref.setdefault(title["ref"], set()).add(row_number)
            number = list_item_number(" ".join(cell["text"].strip() for cell in row_range_cells if cell.get("text", "").strip()))
            if number:
                row_number_by_title_ref.setdefault(title["ref"], {}).setdefault(row_number, set()).add(number)

    def active_numbers_for_title(title: dict, row: int) -> set[str]:
        numbered_rows = [
            row_number
            for row_number in row_number_by_title_ref.get(title["ref"], {})
            if title["row"] < row_number <= row
        ]
        if not numbered_rows:
            return set()
        return row_number_by_title_ref[title["ref"]][max(numbered_rows)]

    def active_numbers_for_range(range_key: tuple[int, int], row: int) -> set[str]:
        numbered_rows = [
            row_number
            for row_number in row_number_by_range_key.get(range_key, {})
            if row_number <= row
        ]
        if not numbered_rows:
            return set()
        return row_number_by_range_key[range_key][max(numbered_rows)]

    def has_synchronized_numbered_flow(
        title: dict,
        baseline_range_key: tuple[int, int],
        row: int,
        allow_seen_numbers: bool = True,
    ) -> bool:
        title_seen_numbers: set[str] = set()
        if allow_seen_numbers:
            for row_number, numbers in row_number_by_title_ref.get(title["ref"], {}).items():
                if title["row"] < row_number <= row:
                    title_seen_numbers.update(numbers)
        for row_number in sorted(row_number_by_range_key.get(baseline_range_key, {})):
            if row_number > row:
                break
            if row_number <= title["row"]:
                continue
            baseline_numbers = active_numbers_for_range(baseline_range_key, row_number)
            title_numbers = active_numbers_for_title(title, row_number)
            if baseline_numbers and title_numbers and baseline_numbers.intersection(title_numbers):
                return True
            if (
                allow_seen_numbers
                and baseline_numbers
                and title_seen_numbers
                and baseline_numbers.intersection(title_seen_numbers)
            ):
                return True
        return False

    def row_band_title_after(
        row: int,
        baseline_title: dict,
        baseline_range_key: tuple[int, int],
        protect_baseline_image_band: bool = True,
        allow_seen_numbers: bool = True,
        content_col: int | None = None,
    ) -> dict | None:
        baseline_visual_protected = protect_baseline_image_band and any(
            start_row <= row <= end_row
            for start_row, end_row in visual_rows_by_title_ref.get(baseline_title["ref"], [])
        )
        row_band_candidates = []
        for title in sorted_titles:
            if not (title["row"] <= row <= title_band_end_by_ref[title["ref"]]):
                continue
            if title["row"] <= baseline_title["row"]:
                continue
            baseline_has_prior_content = any(
                baseline_title["row"] < content_row < title["row"]
                for content_row in content_rows_by_range_key.get(baseline_range_key, set())
            )
            title_row_has_number = bool(row_number_by_title_ref.get(title["ref"], {}).get(row))
            if not baseline_has_prior_content:
                continue
            image_bands = image_rows_by_title_ref.get(title["ref"], [])
            if image_bands:
                first_image_row = min(start_row for start_row, _ in image_bands)
                in_title_visual_band = any(start_row <= row <= end_row for start_row, end_row in image_bands)
                before_title_visual_band = title["row"] < row < first_image_row
                title_row_has_content = row in content_rows_by_title_ref.get(title["ref"], set())
            else:
                in_title_visual_band = False
                before_title_visual_band = False
                title_row_has_content = False
            if not (in_title_visual_band or before_title_visual_band or title_row_has_content):
                continue
            if baseline_visual_protected and not in_title_visual_band:
                continue
            if baseline_visual_protected and in_title_visual_band and content_col is not None:
                containing_image_bounds = [
                    bounds
                    for bounds in image_bounds_by_title_ref.get(title["ref"], [])
                    if bounds["start_row"] <= row <= bounds["end_row"]
                ]
                if containing_image_bounds and content_col < min(bounds["start_col"] for bounds in containing_image_bounds):
                    continue
            if has_synchronized_numbered_flow(title, baseline_range_key, row, allow_seen_numbers):
                row_band_candidates.append(title)
        if row_band_candidates:
            return max(row_band_candidates, key=lambda cell: cell["row"])
        return None

    def section_for_row(row: int) -> dict:
        candidates = [
            title
            for title in sorted_titles
            if title["row"] <= row
        ]
        if not candidates:
            return sections[0]
        same_range_title = max(candidates, key=lambda cell: cell["row"])
        same_range = column_range_for_col(same_range_title["col"], ranges)
        same_range_key = (same_range["start_col"], same_range["end_col"])
        row_band_title = row_band_title_after(row, same_range_title, same_range_key)
        if row_band_title:
            return section_by_title[row_band_title["ref"]]
        return section_by_title[same_range_title["ref"]]

    def section_for_position(
        row: int,
        col: int,
        protect_baseline_image_band: bool = True,
        allow_seen_numbers: bool = True,
    ) -> dict:
        col_range = column_range_for_col(col, ranges)
        range_key = (col_range["start_col"], col_range["end_col"])
        candidates = [
            title
            for title in sorted_titles
            if range_key in title_range_keys_by_ref[title["ref"]] and title["row"] <= row
        ]
        if not candidates:
            return sections[0]
        same_range_title = max(candidates, key=lambda cell: cell["row"])
        row_band_title = row_band_title_after(
            row,
            same_range_title,
            range_key,
            protect_baseline_image_band,
            allow_seen_numbers=allow_seen_numbers,
            content_col=col,
        )
        if row_band_title:
            return section_by_title[row_band_title["ref"]]
        return section_by_title[same_range_title["ref"]]

    def continuation_section_for_position(row: int, col: int) -> dict | None:
        col_range = column_range_for_col(col, ranges)
        range_key = (col_range["start_col"], col_range["end_col"])
        candidates = []
        for title in sorted_title_candidates:
            if title["ref"] not in continuation_targets_by_ref:
                continue
            if row <= title["row"] or row > continuation_end_row_by_ref.get(title["ref"], 999999):
                continue
            if range_key not in continuation_ranges_by_ref.get(title["ref"], set()):
                continue
            candidates.append(title)
        if not candidates:
            return None
        title = max(candidates, key=lambda cell: cell["row"])
        target_title = continuation_targets_by_ref.get(title["ref"])
        if target_title is None:
            return None
        return section_by_title.get(target_title["ref"])

    grouped_by_section_row: dict[int, dict[int, list[dict]]] = {}
    for row in rows_from_cells(remove_cells_by_ref(cells, title_refs)):
        buckets: dict[int, list[dict]] = {}
        for cell in row["cells"]:
            section = continuation_section_for_position(row["row"], cell["col"])
            if section is None:
                section = section_for_position(row["row"], cell["col"])
            section_index = sections.index(section)
            buckets.setdefault(section_index, []).append(cell)
        for section_index, bucket_cells in buckets.items():
            grouped_by_section_row.setdefault(section_index, {})[row["row"]] = bucket_cells

    for section_index, rows in grouped_by_section_row.items():
        section = sections[section_index]
        row_segments: list[dict[int, list[dict]]] = []
        current_segment: dict[int, list[dict]] = {}
        for row_number in sorted(rows):
            if row_number in demoted_titles_by_row and current_segment:
                row_segments.append(current_segment)
                current_segment = {}
            current_segment[row_number] = rows[row_number]
        if current_segment:
            row_segments.append(current_segment)

        flow_order = 0
        for segment_rows in row_segments:
            starts_with_demoted_title = min(segment_rows) in demoted_titles_by_row
            if starts_with_demoted_title:
                append_blocks_from_rows(section, segment_rows, flow_order)
                flow_order += 1
                continue

            section_cells = [cell for row_cells in segment_rows.values() for cell in row_cells]
            flow_ranges = text_flow_ranges_from_rows(segment_rows)
            if flow_ranges and not fallback_title_styles:
                flow_ranges = [
                    flow_range
                    for flow_range in flow_ranges
                    if range_has_numbered_text(segment_rows, flow_range)
                ]
                flow_ranges = expand_flow_ranges_to_cover_cells(flow_ranges, section_cells)
            range_keys = (
                [(item["start_col"], item["end_col"]) for item in flow_ranges]
                if flow_ranges
                else column_range_keys_for_cells(section_cells, ranges)
            )
            use_column_flows = len(range_keys) > 1
            if use_column_flows:
                for range_key in range_keys:
                    start_col, end_col = range_key
                    flow_rows: dict[int, list[dict]] = {}
                    for row_number in sorted(segment_rows):
                        row_cells = [
                            cell
                            for cell in segment_rows[row_number]
                            if start_col <= cell_center_col(cell) <= end_col
                        ]
                        if row_cells:
                            flow_rows[row_number] = row_cells
                    append_blocks_from_rows(section, flow_rows, flow_order)
                    flow_order += 1
            else:
                append_blocks_from_rows(section, segment_rows, flow_order)
                flow_order += 1

    for image in images:
        anchor = image.get("anchor") or {}
        row = anchor.get("from_row")
        col = anchor.get("from_col")
        if row is None or col is None:
            append_to_section(sections[0], image_block(image))
        else:
            append_to_section(
                section_for_position(row, col, False, allow_seen_numbers=False),
                image_block(image),
            )

    first_title_row = min(title["row"] for title in sorted_titles)
    leading_section = sections[0]
    if has_isolated_titles and leading_section["title"] is None and leading_section["blocks"]:
        leading_blocks = leading_section["blocks"]
        leading_title_like = True
        for block in leading_blocks:
            if block.get("type") != "row":
                leading_title_like = False
                break
            bounds = block_bounds(block)
            if bounds["start_row"] >= first_title_row:
                leading_title_like = False
                break
            content_cells = [cell for cell in block.get("cells", []) if cell.get("text")]
            if not content_cells or not all(cell.get("style") in title_styles for cell in content_cells):
                leading_title_like = False
                break
        if leading_title_like:
            leading_section["blocks"] = []
            leading_section["source_refs"] = []
            leading_section["bounds"] = None

    return [
        section
        for section in sections
        if section["title"] is not None or section["blocks"]
    ]


def row_without_title_cells(row_block: dict, title_styles: set[str | None]) -> dict | None:
    if row_block.get("type") != "row":
        return row_block
    cells = [
        cell
        for cell in row_block.get("cells", [])
        if cell.get("style") not in title_styles
    ]
    if not cells:
        return None
    clone = row_block.copy()
    clone["cells"] = cells
    clone["lines"] = logical_lines_for_cells(cells)
    clone["source_refs"] = [cell["ref"] for cell in cells]
    clone["bounds"] = {
        "start_row": row_block["bounds"]["start_row"],
        "end_row": row_block["bounds"]["end_row"],
        "start_col": min(cell["col"] for cell in cells),
        "end_col": max(cell["col"] for cell in cells),
    }
    return clone


def public_cell(cell: dict) -> dict:
    return {
        "ref": cell["ref"],
        "row": cell["row"],
        "col": cell["col"],
        "text": cell["text"],
        "style": cell.get("style"),
        "merge_ref": cell.get("merge_ref"),
        "bounds": cell["bounds"],
    }


def block_from_row(row: dict) -> dict:
    cells = row["cells"]
    refs = [cell["ref"] for cell in cells]
    return {
        "type": "row",
        "cells": cells,
        "lines": logical_lines_for_cells(cells),
        "source_refs": refs,
        "bounds": {
            "start_row": row["row"],
            "end_row": row["row"],
            "start_col": min(cell["col"] for cell in cells),
            "end_col": max(cell["col"] for cell in cells),
        },
    }


def same_table_shape(previous: dict, current: dict) -> bool:
    previous_cols = [cell["col"] for cell in previous["cells"]]
    current_cols = [cell["col"] for cell in current["cells"]]
    return len(previous_cols) > 1 and previous_cols == current_cols


def table_block_from_rows(table_rows: list[dict]) -> dict:
    header = [cell["text"] for cell in table_rows[0]["cells"]]
    body = []
    source_refs = []
    for row in table_rows:
        source_refs.extend(cell["ref"] for cell in row["cells"])
    for row in table_rows[1:]:
        body.append([cell["text"] for cell in row["cells"]])
    return {
        "type": "table",
        "header": header,
        "rows": body,
        "source_refs": source_refs,
        "bounds": {
            "start_row": table_rows[0]["row"],
            "end_row": table_rows[-1]["row"],
            "start_col": min(cell["col"] for row in table_rows for cell in row["cells"]),
            "end_col": max(cell["col"] for row in table_rows for cell in row["cells"]),
        },
    }


def build_blocks(rows: list[dict]) -> list[dict]:
    blocks: list[dict] = []
    index = 0
    while index < len(rows):
        table_rows = [rows[index]]
        next_index = index + 1
        while next_index < len(rows) and same_table_shape(table_rows[-1], rows[next_index]):
            table_rows.append(rows[next_index])
            next_index += 1

        if len(table_rows) > 1:
            blocks.append(table_block_from_rows(table_rows))
            index = next_index
            continue

        blocks.append(block_from_row(rows[index]))
        index += 1
    return blocks


def block_bounds(block: dict) -> dict:
    return block.get("bounds") or {
        "start_row": 0,
        "start_col": 0,
        "end_row": 0,
        "end_col": 0,
    }


def block_source_refs(block: dict) -> list[str]:
    refs = list(block.get("source_refs", []))
    for child in block.get("blocks", []):
        refs.extend(block_source_refs(child))
    return refs


def build_sections(blocks: list[dict], images: list[dict], title_styles: set[str | None]) -> list[dict]:
    sections: list[dict] = []
    current = {"type": "section", "title": None, "blocks": [], "source_refs": [], "bounds": None}
    image_index = 0
    sorted_images = sorted(
        images,
        key=lambda image: (
            (image.get("anchor") or {}).get("from_row") or 999999,
            (image.get("anchor") or {}).get("from_col") or 999999,
        ),
    )

    def flush_current() -> None:
        nonlocal current
        if current["blocks"] or current["title"] is not None:
            sections.append(current)
        current = {"type": "section", "title": None, "blocks": [], "source_refs": [], "bounds": None}

    def append_block(block: dict) -> None:
        current["blocks"].append(block)
        current["source_refs"].extend(block_source_refs(block))
        bounds = block_bounds(block)
        if current["bounds"] is None:
            current["bounds"] = bounds.copy()
        else:
            current["bounds"] = {
                "start_row": min(current["bounds"]["start_row"], bounds["start_row"]),
                "start_col": min(current["bounds"]["start_col"], bounds["start_col"]),
                "end_row": max(current["bounds"]["end_row"], bounds["end_row"]),
                "end_col": max(current["bounds"]["end_col"], bounds["end_col"]),
            }

    def append_due_images(before_row: int) -> None:
        nonlocal image_index
        while image_index < len(sorted_images):
            image = sorted_images[image_index]
            image_row = (image.get("anchor") or {}).get("from_row")
            if image_row is None or image_row > before_row:
                break
            anchor = image.get("anchor") or {}
            append_block(
                {
                    "type": "image",
                    "image_indexes": [image["index"]],
                    "bounds": {
                        "start_row": anchor.get("from_row") or 0,
                        "start_col": anchor.get("from_col") or 0,
                        "end_row": anchor.get("to_row") or anchor.get("from_row") or 0,
                        "end_col": anchor.get("to_col") or anchor.get("from_col") or 0,
                    },
                }
            )
            image_index += 1

    for block in blocks:
        bounds = block_bounds(block)
        append_due_images(bounds["start_row"])
        titles = row_title_lines(block, title_styles)
        if titles:
            flush_current()
            current["title"] = " ".join(titles)
            remaining = row_without_title_cells(block, title_styles)
            if remaining is not None:
                append_block(remaining)
            continue
        append_block(block)

    append_due_images(999999)
    flush_current()
    return sections


def append_public_text(public_blocks: list[dict], lines: list[str]) -> None:
    lines = [line for line in lines if line]
    if not lines:
        return
    if public_blocks and public_blocks[-1].get("type") == "text":
        public_blocks[-1]["lines"].extend(lines)
    else:
        public_blocks.append({"type": "text", "lines": lines})


def append_public_image(public_blocks: list[dict], image_indexes: list[int]) -> None:
    if not image_indexes:
        return
    if public_blocks and public_blocks[-1].get("type") == "image":
        public_blocks[-1]["image_indexes"].extend(image_indexes)
    else:
        public_blocks.append({"type": "image", "image_indexes": list(image_indexes)})


def row_overlap(a: dict, b: dict) -> int:
    start = max(a["start_row"], b["start_row"])
    end = min(a["end_row"], b["end_row"])
    return max(0, end - start + 1)


def col_gap(a: dict, b: dict) -> int:
    if a["end_col"] < b["start_col"]:
        return b["start_col"] - a["end_col"]
    if b["end_col"] < a["start_col"]:
        return a["start_col"] - b["end_col"]
    return 0


def row_gap(a: dict, b: dict) -> int:
    if a["end_row"] < b["start_row"]:
        return b["start_row"] - a["end_row"]
    if b["end_row"] < a["start_row"]:
        return a["start_row"] - b["end_row"]
    return 0


def public_table_block(block: dict) -> dict:
    table = {"header": block.get("header", []), "rows": block.get("rows", [])}
    if block.get("title"):
        table["title"] = block["title"]
    return table


def public_section_blocks(blocks: list[dict]) -> list[dict]:
    row_like_blocks = [
        block
        for block in blocks
        if (
            block.get("type") in ("row", "text")
            or ("lines" in block and "type" not in block)
        )
        and any(line for line in block.get("lines", []))
    ]
    table_like_blocks = [
        block
        for block in blocks
        if block.get("type") == "table" or ("header" in block and "rows" in block)
    ]
    image_like_blocks = [
        block
        for block in blocks
        if block.get("type") == "image" or ("image_indexes" in block and "type" not in block)
    ]
    row_flow_orders = {block.get("_flow_order", 0) for block in row_like_blocks}
    if row_like_blocks and not table_like_blocks and len(row_flow_orders) == 1:
        lines: list[str] = []
        image_indexes: list[int] = []
        for block in sorted(row_like_blocks, key=lambda item: (block_bounds(item)["start_row"], block_bounds(item)["start_col"])):
            lines.extend(line for line in block.get("lines", []) if line)
        for block in sorted(image_like_blocks, key=lambda item: (block_bounds(item)["start_row"], block_bounds(item)["start_col"])):
            image_indexes.extend(block.get("image_indexes", []))
        result = {"lines": lines}
        if image_indexes:
            result["image_indexes"] = image_indexes
        return [result]

    if row_like_blocks and table_like_blocks and not image_like_blocks and len(row_flow_orders) == 1:
        lines: list[str] = []
        result_blocks: list[dict] = []
        for block in sorted(row_like_blocks, key=lambda item: (block_bounds(item)["start_row"], block_bounds(item)["start_col"])):
            lines.extend(line for line in block.get("lines", []) if line)
        if lines:
            result_blocks.append({"lines": lines})
        for block in sorted(table_like_blocks, key=lambda item: (block_bounds(item)["start_row"], block_bounds(item)["start_col"])):
            result_blocks.append(public_table_block(block))
        return result_blocks

    if row_like_blocks and image_like_blocks and not table_like_blocks:
        flow_items: list[dict] = []
        for flow_order in sorted(row_flow_orders):
            flow_rows = [
                block
                for block in row_like_blocks
                if block.get("_flow_order", 0) == flow_order
            ]
            if not flow_rows:
                continue
            sorted_rows = sorted(flow_rows, key=lambda item: (block_bounds(item)["start_row"], block_bounds(item)["start_col"]))
            lines = [
                line
                for block in sorted_rows
                for line in block.get("lines", [])
                if line
            ]
            bounds_list = [block_bounds(block) for block in sorted_rows]
            flow_items.append(
                {
                    "flow_order": flow_order,
                    "block": {"lines": lines},
                    "bounds": {
                        "start_row": min(bounds["start_row"] for bounds in bounds_list),
                        "end_row": max(bounds["end_row"] for bounds in bounds_list),
                        "start_col": min(bounds["start_col"] for bounds in bounds_list),
                        "end_col": max(bounds["end_col"] for bounds in bounds_list),
                    },
                }
            )

        standalone_images: list[dict] = []
        for image in sorted(image_like_blocks, key=lambda item: (block_bounds(item)["start_row"], block_bounds(item)["start_col"])):
            image_bounds = block_bounds(image)
            if not flow_items:
                standalone_images.append(image)
                continue
            later_flow_items = [
                flow
                for flow in flow_items
                if flow["bounds"]["start_row"] <= image_bounds["start_row"]
            ]
            target_items = later_flow_items if later_flow_items else flow_items
            target = min(
                target_items,
                key=lambda flow: (
                    col_gap(image_bounds, flow["bounds"]),
                    row_gap(image_bounds, flow["bounds"]),
                    flow["bounds"]["start_col"],
                    -row_overlap(image_bounds, flow["bounds"]),
                ),
            )
            target["block"].setdefault("image_indexes", []).extend(image.get("image_indexes", []))

        if len(flow_items) == 2 and sum(1 for flow in flow_items if flow["block"].get("image_indexes")) == 1:
            sorted_flow_items = sorted(
                flow_items,
                key=lambda flow: (
                    0 if flow["block"].get("image_indexes") else 1,
                    flow["flow_order"],
                    flow["bounds"]["start_row"],
                    flow["bounds"]["start_col"],
                ),
            )
        else:
            sorted_flow_items = sorted(
                flow_items,
                key=lambda flow: (
                    flow["bounds"]["start_row"],
                    flow["flow_order"],
                    flow["bounds"]["start_col"],
                ),
            )
        result_blocks = [flow["block"] for flow in sorted_flow_items]
        for image in standalone_images:
            image_indexes = image.get("image_indexes", [])
            if image_indexes:
                result_blocks.append({"image_indexes": image_indexes})
        return result_blocks

    public_blocks: list[dict] = []
    pending_images: list[int] = []
    last_text_block: dict | None = None
    last_flow_order: int | None = None

    def flush_images() -> None:
        nonlocal pending_images
        if not pending_images:
            return
        if last_text_block is not None and "image_indexes" not in last_text_block:
            last_text_block["image_indexes"] = list(pending_images)
        else:
            public_blocks.append({"image_indexes": list(pending_images)})
        pending_images = []

    ordered_blocks = sorted(
        blocks,
        key=lambda item: (
            item.get("_flow_order", 0),
            block_bounds(item)["start_row"],
            block_bounds(item)["start_col"],
        ),
    )
    for block in ordered_blocks:
        block_type = block.get("type")
        if block_type == "image":
            pending_images.extend(block.get("image_indexes", []))
            continue
        if block_type == "table":
            flush_images()
            public_blocks.append(public_table_block(block))
            last_text_block = None
            last_flow_order = None
            continue
        if block_type in ("row", "text"):
            lines = [line for line in block.get("lines", []) if line]
            if not lines:
                continue
            flow_order = block.get("_flow_order", 0)
            if last_text_block is not None and last_flow_order == flow_order:
                last_text_block.setdefault("lines", []).extend(lines)
            else:
                last_text_block = {"lines": lines}
                if pending_images:
                    last_text_block["image_indexes"] = list(pending_images)
                    pending_images = []
                public_blocks.append(last_text_block)
                last_flow_order = flow_order
            continue
        public = public_block(block)
        if public is not None:
            if "image_indexes" in public and "lines" not in public and "header" not in public:
                pending_images.extend(public.get("image_indexes", []))
            elif "lines" in public:
                flush_images()
                public_blocks.append(public)
                last_text_block = public
                last_flow_order = block.get("_flow_order", 0)
            else:
                flush_images()
                public_blocks.append(public)
                last_text_block = None
                last_flow_order = None

    flush_images()
    return public_blocks


def public_block(block: dict) -> dict | None:
    block_type = block.get("type")
    if block_type == "section":
        return {
            "title": block.get("title"),
            "blocks": public_section_blocks(block.get("blocks", [])),
        }
    if block_type == "table":
        return public_table_block(block)
    if block_type == "image":
        image_indexes = block.get("image_indexes", [])
        return {"image_indexes": image_indexes} if image_indexes else None
    if block_type in ("row", "text"):
        lines = [line for line in block.get("lines", []) if line]
        if not lines:
            return None
        return {
            "lines": lines,
        }
    return None


def public_sections(blocks: list[dict]) -> list[dict]:
    sections = []
    loose_blocks = []
    for block in blocks:
        public = public_block(block)
        if public is None:
            continue
        if "blocks" in public:
            sections.append(public)
        else:
            loose_blocks.append(public)
    if loose_blocks and not sections:
        sections.insert(0, {"title": None, "blocks": loose_blocks})
    while len(sections) > 1 and sections[0].get("title") is None:
        sections.pop(0)
    return sections


def public_images(images: list[dict]) -> list[dict]:
    result = []
    for image in images:
        item = {
            "index": image["index"],
            "path": image.get("path"),
        }
        if image.get("error"):
            item["error"] = image["error"]
        result.append(item)
    return result


def convert_to_webp(source: Path, target: Path) -> None:
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError(
            "Pillow is required to convert Excel images to WebP. "
            "Install Pillow or do not use images for visual analysis."
        ) from exc

    target.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(source) as image:
        if image.mode not in ("RGB", "RGBA"):
            image = image.convert("RGBA" if "A" in image.getbands() else "RGB")
        image.save(target, "WEBP", quality=92, method=6)


def export_media_to_webp(
    zf: zipfile.ZipFile,
    media_items: list[dict],
    output_dir: Path | None,
    sheet_name: str,
) -> list[dict]:
    exported = []
    safe_sheet = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff_-]+", "_", sheet_name).strip("_")

    for item in media_items:
        index = item["index"]
        media_part = item["source"]
        suffix = Path(media_part).suffix.lower()
        webp_path = output_dir / f"{safe_sheet}_{index}.webp" if output_dir else None

        try:
            if webp_path:
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as temp_file:
                    raw_path = Path(temp_file.name)
                    temp_file.write(zf.read(media_part))
                try:
                    convert_to_webp(raw_path, webp_path)
                finally:
                    raw_path.unlink(missing_ok=True)
            exported.append(
                {
                    "index": index,
                    "source": media_part,
                    "path": str(webp_path) if webp_path else None,
                    "anchor": item.get("anchor"),
                }
            )
        except Exception as exc:
            exported.append(
                {
                    "index": index,
                    "source": media_part,
                    "path": None,
                    "anchor": item.get("anchor"),
                    "error": str(exc),
                }
            )

    return exported


def drawing_media_for_sheet(zf: zipfile.ZipFile, sheet_path: str) -> list[dict]:
    sheet_rels = read_rels(zf, rels_path_for(sheet_path))
    if not sheet_rels:
        return []

    root = ET.fromstring(zf.read(sheet_path))
    drawing_parts = []
    for drawing in root.findall("a:drawing", NS):
        rid = drawing.attrib.get(f"{DOC_REL_NS}id")
        target = sheet_rels.get(rid or "")
        if target:
            drawing_parts.append(resolve_part(sheet_path, target))

    media_items: list[dict] = []
    seen: set[tuple[str, int, int]] = set()
    for drawing_part in drawing_parts:
        drawing_rels = read_rels(zf, rels_path_for(drawing_part))
        if not drawing_rels:
            continue
        drawing_root = ET.fromstring(zf.read(drawing_part))
        for anchor in list(drawing_root):
            from_node = anchor.find("xdr:from", NS)
            to_node = anchor.find("xdr:to", NS)
            blip = anchor.find(".//a_dml:blip", NS)
            if blip is None:
                continue
            rid = blip.attrib.get(f"{DOC_REL_NS}embed")
            target = drawing_rels.get(rid or "")
            if not target:
                continue
            media_part = resolve_part(drawing_part, target)
            if not media_part.startswith("xl/media/"):
                continue

            from_row = from_col = to_row = to_col = None
            if from_node is not None:
                row_node = from_node.find("xdr:row", NS)
                col_node = from_node.find("xdr:col", NS)
                if row_node is not None and row_node.text is not None:
                    from_row = int(row_node.text) + 1
                if col_node is not None and col_node.text is not None:
                    from_col = int(col_node.text) + 1
            if to_node is not None:
                row_node = to_node.find("xdr:row", NS)
                col_node = to_node.find("xdr:col", NS)
                if row_node is not None and row_node.text is not None:
                    to_row = int(row_node.text) + 1
                if col_node is not None and col_node.text is not None:
                    to_col = int(col_node.text) + 1

            key = (media_part, from_row or -1, from_col or -1)
            if key in seen:
                continue
            seen.add(key)
            media_items.append(
                {
                    "source": media_part,
                    "anchor": {
                        "from_row": from_row,
                        "from_col": from_col,
                        "to_row": to_row,
                        "to_col": to_col,
                    },
                }
            )

    media_items.sort(
        key=lambda item: (
            (item.get("anchor") or {}).get("from_row") or 999999,
            (item.get("anchor") or {}).get("from_col") or 999999,
            item["source"],
        )
    )
    for index, item in enumerate(media_items, start=1):
        item["index"] = index
    return media_items


def parse_main_sheet(
    zf: zipfile.ZipFile,
    sheet_path: str,
    shared: list[str],
    style_features: dict[str, dict],
    images: list[dict],
) -> dict:
    root = ET.fromstring(zf.read(sheet_path))
    merged_ranges = read_merged_ranges(root)
    cells = parse_cells(root, shared, merged_ranges)
    title_styles = title_style_set(cells, style_features)
    blocks = group_cells_for_sections(
        cells,
        title_styles,
        images,
        uses_fallback_title_styles(cells, style_features),
    )
    return {
        "sections": public_sections(blocks),
    }


def parse_xlsx(path: Path, output_dir: Path, main_patterns: list[str]) -> dict:
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Only .xlsx files are supported by this script.")

    with zipfile.ZipFile(path) as zf:
        shared = read_shared_strings(zf)
        style_features = read_style_features(zf)
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        workbook_rels = read_rels(zf, "xl/_rels/workbook.xml.rels")

        sheets = []
        for sheet in workbook.find("a:sheets", NS):
            sheet_name = sheet.attrib["name"]
            rid = sheet.attrib.get(OD_REL_ID)
            target = workbook_rels.get(rid or "")
            if not target:
                continue
            sheets.append(
                {
                    "name": sheet_name,
                    "normalized_name": normalize_sheet_name(sheet_name),
                    "path": sheet_path_from_target(target),
                }
            )

        selected_main = select_sheet(sheets, main_patterns)
        if not selected_main:
            raise ValueError(
                "No main sheet matched. "
                f"main sheet patterns: {', '.join(main_patterns)}"
            )

        warnings = []
        errors = []
        media_parts = drawing_media_for_sheet(zf, selected_main["path"])
        exported_images = export_media_to_webp(zf, media_parts, output_dir, selected_main["name"])
        main_sheet = parse_main_sheet(zf, selected_main["path"], shared, style_features, exported_images)
        main_sheet.update(
            {
                "name": selected_main["name"],
                "images": public_images(exported_images),
            }
        )

        for image in main_sheet.get("images", []):
            if image.get("error"):
                warnings.append(f"Image conversion failed: index {image['index']} - {image['error']}")

        return {
            "file": str(path),
            "main_sheet": main_sheet,
            "warnings": warnings,
            "errors": errors,
        }


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("xlsx_path", type=Path)
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Directory for extract.json and exported WebP images. It is deleted before each run.",
    )
    parser.add_argument(
        "--main-sheets",
        required=True,
        help="Comma-separated fuzzy main sheet patterns in priority order.",
    )
    args = parser.parse_args(argv)

    main_patterns = parse_patterns(args.main_sheets)
    ensure_clean_output_dir(args.output, args.xlsx_path)
    data = parse_xlsx(args.xlsx_path, args.output, main_patterns)
    output_path = args.output / "extract.json"
    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(data, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    configure_stdio_utf8()
    try:
        raise SystemExit(main(sys.argv[1:]))
    except Exception as exc:
        print(f"extract_excel_design.py: error: {exc}", file=sys.stderr)
        raise SystemExit(1)
