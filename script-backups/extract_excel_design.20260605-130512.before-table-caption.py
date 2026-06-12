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
        return set()
    return {
        style
        for style in counts
        if (style_features.get(style or "") or {}).get("font_size") is not None
        and (style_features.get(style or "") or {}).get("font_size") > body_size
    }


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
    if flow_order is not None:
        block["_flow_order"] = flow_order
    return block


def row_cell_count_in_cols(row_cells: list[dict], cols: list[int]) -> int:
    return sum(1 for cell in row_cells if cell["col"] in cols)


def detect_table_at_row(row_index: int, row_numbers: list[int], rows: dict[int, list[dict]]) -> tuple[list[int], list[int]] | None:
    row_number = row_numbers[row_index]
    row_cells = rows.get(row_number, [])
    if len(row_cells) < 2:
        return None

    cells_by_col = {cell["col"]: cell for cell in row_cells}
    sorted_cols = sorted(cells_by_col)
    best_cols: list[int] = []
    best_rows: list[int] = []

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

    if best_cols and best_rows:
        return best_rows, best_cols
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

        table_rows, table_cols = detected
        table_col_set = set(table_cols)
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
                append_to_section(section, block)
        append_to_section(section, table_block_from_grid(table_rows, table_cols, cells_by_position, flow_order))
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


def group_cells_for_sections(cells: list[dict], title_styles: set[str | None], images: list[dict]) -> list[dict]:
    titles = section_title_cells(cells, title_styles)
    has_isolated_titles = has_recurring_isolated_titles(cells, title_styles)
    if not titles:
        section = make_section(None)
        for row in rows_from_cells(cells):
            block = block_from_cells(row["row"], row["cells"])
            if block is not None:
                append_to_section(section, block)
        for image in images:
            append_to_section(section, image_block(image))
        return [section]

    ranges = layout_column_ranges(cells, images, title_styles)
    sections = [make_section(None)]
    section_by_title: dict[str, dict] = {}
    sorted_titles = sorted(titles, key=lambda cell: (cell["row"], cell["col"]))
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

    def section_for_row(row: int) -> dict:
        candidates = [
            title
            for title in sorted_titles
            if title["row"] <= row
        ]
        if not candidates:
            return sections[0]
        return section_by_title[max(candidates, key=lambda cell: cell["row"])["ref"]]

    def section_for_position(row: int, col: int) -> dict:
        col_range = column_range_for_col(col, ranges)
        range_key = (col_range["start_col"], col_range["end_col"])
        candidates = [
            title
            for title in sorted_titles
            if range_key in title_range_keys_by_ref[title["ref"]] and title["row"] <= row
        ]
        if not candidates:
            return sections[0]
        return section_by_title[max(candidates, key=lambda cell: cell["row"])["ref"]]

    grouped_by_section_row: dict[int, dict[int, list[dict]]] = {}
    for row in rows_from_cells(remove_cells_by_ref(cells, title_refs)):
        buckets: dict[int, list[dict]] = {}
        for cell in row["cells"]:
            section = section_for_position(row["row"], cell["col"])
            section_index = sections.index(section)
            buckets.setdefault(section_index, []).append(cell)
        for section_index, bucket_cells in buckets.items():
            grouped_by_section_row.setdefault(section_index, {})[row["row"]] = bucket_cells

    for section_index, rows in grouped_by_section_row.items():
        section = sections[section_index]
        section_cells = [cell for row_cells in rows.values() for cell in row_cells]
        range_keys = column_range_keys_for_cells(section_cells, ranges)
        use_column_flows = len(range_keys) > 1
        if use_column_flows:
            for flow_order, range_key in enumerate(range_keys):
                start_col, end_col = range_key
                flow_rows: dict[int, list[dict]] = {}
                for row_number in sorted(rows):
                    row_cells = [
                        cell
                        for cell in rows[row_number]
                        if start_col <= cell_center_col(cell) <= end_col
                    ]
                    if row_cells:
                        flow_rows[row_number] = row_cells
                append_blocks_from_rows(section, flow_rows, flow_order)
        else:
            append_blocks_from_rows(section, rows)

    for image in images:
        anchor = image.get("anchor") or {}
        row = anchor.get("from_row")
        col = anchor.get("from_col")
        if row is None or col is None:
            append_to_section(sections[0], image_block(image))
        else:
            append_to_section(section_for_position(row, col), image_block(image))

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
            public_blocks.append({"header": block.get("header", []), "rows": block.get("rows", [])})
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
        return {
            "header": block.get("header", []),
            "rows": block.get("rows", []),
        }
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
    blocks = group_cells_for_sections(cells, title_style_set(cells, style_features), images)
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
