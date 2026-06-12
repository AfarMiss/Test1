#!/usr/bin/env python3
"""Generate Markdown reading views directly from Excel design docs.

This script reads only the source xlsx package: worksheet XML, styles, merged
cells, drawing anchors, and media parts.
"""

from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass, replace
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image


ROOT = Path(__file__).resolve().parents[2]
DOC_DIR = ROOT / ".codex" / "design-doc-analysis"
SOURCE_ROOT = Path(r"D:\DaoGuangZhanShen\design\策划案")
OUTPUT_ROOT = DOC_DIR / "all-design-md-v3"
MD_LIST = DOC_DIR / "md-list.md"
MAIN_SHEET_PATTERNS = ("界面&功能", "界面", "功能")

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a_dml": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
DOC_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
OD_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


@dataclass
class Cell:
    ref: str
    row: int
    col: int
    text: str
    style: str | None
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    is_title: bool = False


@dataclass
class ImageItem:
    index: int
    source: str
    path: str
    from_row: int
    from_col: int
    to_row: int
    to_col: int


def configure_stdio() -> None:
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


def cell_position(ref: str) -> tuple[int, int]:
    match = re.match(r"([A-Z]+)(\d+)", ref or "")
    if not match:
        return (999999, 999999)
    return int(match.group(2)), col_to_num(match.group(1))


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


def text_from_si(node: ET.Element | None) -> str:
    if node is None:
        return ""
    pieces = []
    for text_node in node.findall(".//a:t", NS):
        pieces.append(text_node.text or "")
    return "".join(pieces)


def normalize_text(value: str) -> str:
    return re.sub(r"[ \t]+", " ", value.replace("\r", "\n")).strip()


def is_dispimg_formula(value: str) -> bool:
    normalized = value.strip()
    if normalized.startswith("="):
        normalized = normalized[1:].strip()
    normalized = normalized.removeprefix("_xlfn.").strip()
    return bool(re.fullmatch(r"DISPIMG\s*\(.*\)", normalized, flags=re.IGNORECASE))


def dispimg_id(value: str) -> str | None:
    normalized = value.strip()
    if normalized.startswith("="):
        normalized = normalized[1:].strip()
    normalized = normalized.removeprefix("_xlfn.").strip()
    match = re.fullmatch(r'DISPIMG\s*\(\s*"([^"]+)"\s*,.*\)', normalized, flags=re.IGNORECASE)
    return match.group(1) if match else None


def resolve_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = Path(base_part).parent
    parts: list[str] = []
    for part in (base / target).as_posix().split("/"):
        if part == "..":
            if parts:
                parts.pop()
        elif part and part != ".":
            parts.append(part)
    return "/".join(parts)


def rels_path_for(part: str) -> str:
    path = Path(part)
    return (path.parent / "_rels" / f"{path.name}.rels").as_posix()


def sheet_path_from_target(target: str) -> str:
    target = target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return f"xl/{target}"


def read_rels(zf: zipfile.ZipFile, rels_path: str) -> dict[str, str]:
    if rels_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rels_path))
    return {rel.attrib["Id"]: rel.attrib["Target"] for rel in root.findall(f"{REL_NS}Relationship")}


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return [text_from_si(si) for si in root.findall("a:si", NS)]


def read_styles(zf: zipfile.ZipFile) -> dict[str, dict]:
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
    if is_dispimg_formula(text) or is_dispimg_formula(formula):
        return ""
    if text:
        return normalize_text(text.replace("\n", " "))
    if formula:
        return normalize_text(f"[formula] {formula}")
    return ""


def read_merged_ranges(root: ET.Element) -> list[dict]:
    ranges = []
    for node in root.findall(".//a:mergeCell", NS):
        ref = node.attrib.get("ref")
        if ref:
            ranges.append(range_bounds(ref))
    return ranges


def merge_for_cell(row: int, col: int, merged_ranges: list[dict]) -> dict | None:
    for merged in merged_ranges:
        if merged["start_row"] <= row <= merged["end_row"] and merged["start_col"] <= col <= merged["end_col"]:
            return merged
    return None


def parse_cells(root: ET.Element, shared: list[str], merged_ranges: list[dict]) -> list[Cell]:
    cells = []
    for node in root.findall(".//a:c", NS):
        text = cell_text(node, shared)
        if not text:
            continue
        ref = node.attrib.get("r", "")
        row, col = cell_position(ref)
        merged = merge_for_cell(row, col, merged_ranges)
        bounds = merged or {"start_row": row, "start_col": col, "end_row": row, "end_col": col}
        cells.append(
            Cell(
                ref=ref,
                row=row,
                col=col,
                text=text,
                style=node.attrib.get("s"),
                start_row=bounds["start_row"],
                start_col=bounds["start_col"],
                end_row=bounds["end_row"],
                end_col=bounds["end_col"],
            )
        )
    return sorted(cells, key=lambda item: (item.row, item.col))


def select_sheet(zf: zipfile.ZipFile) -> tuple[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    workbook_rels = read_rels(zf, "xl/_rels/workbook.xml.rels")
    sheets = []
    for sheet in workbook.find("a:sheets", NS):
        name = sheet.attrib["name"]
        rid = sheet.attrib.get(OD_REL_ID)
        target = workbook_rels.get(rid or "")
        if target:
            sheets.append((name, sheet_path_from_target(target)))
    for pattern in MAIN_SHEET_PATTERNS:
        for name, path in sheets:
            if pattern in name.strip():
                return name, path
    raise ValueError(f"no main sheet matched: {', '.join(MAIN_SHEET_PATTERNS)}")


def convert_to_webp(zf: zipfile.ZipFile, source: str, target: Path) -> None:
    suffix = Path(source).suffix.lower()
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        raw = Path(tmp.name)
        tmp.write(zf.read(source))
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        with Image.open(raw) as image:
            if image.mode not in ("RGB", "RGBA"):
                image = image.convert("RGBA" if "A" in image.getbands() else "RGB")
            image.save(target, "WEBP", quality=92, method=6)
    finally:
        raw.unlink(missing_ok=True)


def image_area(item: tuple[int, int, int, int, str]) -> int:
    from_row, from_col, to_row, to_col, _source = item
    return max(0, to_row - from_row + 1) * max(0, to_col - from_col + 1)


def image_contains(outer: tuple[int, int, int, int, str], inner: tuple[int, int, int, int, str]) -> bool:
    outer_from_row, outer_from_col, outer_to_row, outer_to_col, _outer_source = outer
    inner_from_row, inner_from_col, inner_to_row, inner_to_col, _inner_source = inner
    return (
        outer_from_row <= inner_from_row
        and outer_from_col <= inner_from_col
        and outer_to_row >= inner_to_row
        and outer_to_col >= inner_to_col
        and image_area(outer) > image_area(inner)
    )


def is_cell_image(image: ImageItem) -> bool:
    return image.from_row == image.to_row and image.from_col == image.to_col


def filter_contained_images(
    raw_items: list[tuple[int, int, int, int, str]],
) -> list[tuple[int, int, int, int, str]]:
    return [
        item
        for item in raw_items
        if not any(image_contains(other, item) for other in raw_items if other is not item)
    ]


def read_cell_image_sources(zf: zipfile.ZipFile) -> dict[str, str]:
    if "xl/cellimages.xml" not in zf.namelist():
        return {}
    rels = read_rels(zf, "xl/_rels/cellimages.xml.rels")
    root = ET.fromstring(zf.read("xl/cellimages.xml"))
    result: dict[str, str] = {}
    for cell_image in root:
        c_nv_pr = cell_image.find(".//xdr:cNvPr", NS)
        blip = cell_image.find(".//a_dml:blip", NS)
        if c_nv_pr is None or blip is None:
            continue
        image_id = c_nv_pr.attrib.get("name")
        rid = blip.attrib.get(f"{DOC_REL_NS}embed")
        target = rels.get(rid or "")
        if image_id and target:
            result[image_id] = resolve_part("xl/cellimages.xml", target)
    return result


def worksheet_dispimg_items(zf: zipfile.ZipFile, sheet_path: str) -> list[tuple[int, int, int, int, str]]:
    id_to_source = read_cell_image_sources(zf)
    if not id_to_source:
        return []
    root = ET.fromstring(zf.read(sheet_path))
    items: list[tuple[int, int, int, int, str]] = []
    for cell in root.findall(".//a:c", NS):
        ref = cell.attrib.get("r", "")
        row, col = cell_position(ref)
        formula_node = cell.find("a:f", NS)
        value_node = cell.find("a:v", NS)
        formula = formula_node.text if formula_node is not None and formula_node.text else ""
        value = value_node.text if value_node is not None and value_node.text else ""
        image_id = dispimg_id(formula) or dispimg_id(value)
        source = id_to_source.get(image_id or "")
        if source:
            items.append((row, col, row, col, source))
    return items


def drawing_images(zf: zipfile.ZipFile, sheet_path: str, output_dir: Path) -> list[ImageItem]:
    root = ET.fromstring(zf.read(sheet_path))
    sheet_rels = read_rels(zf, rels_path_for(sheet_path))
    drawing_parts = []
    for drawing in root.findall("a:drawing", NS):
        rid = drawing.attrib.get(f"{DOC_REL_NS}id")
        target = sheet_rels.get(rid or "")
        if target:
            drawing_parts.append(resolve_part(sheet_path, target))

    raw_items = []
    seen = set()
    for drawing_part in drawing_parts:
        drawing_rels = read_rels(zf, rels_path_for(drawing_part))
        drawing_root = ET.fromstring(zf.read(drawing_part))
        for anchor in list(drawing_root):
            from_node = anchor.find("xdr:from", NS)
            to_node = anchor.find("xdr:to", NS)
            blip = anchor.find(".//a_dml:blip", NS)
            if from_node is None or blip is None:
                continue
            rid = blip.attrib.get(f"{DOC_REL_NS}embed")
            target = drawing_rels.get(rid or "")
            if not target:
                continue
            source = resolve_part(drawing_part, target)
            if not source.startswith("xl/media/"):
                continue

            def marker(node: ET.Element | None, name: str) -> int:
                if node is None:
                    return 0
                child = node.find(f"xdr:{name}", NS)
                return int(child.text) + 1 if child is not None and child.text is not None else 0

            from_row = marker(from_node, "row")
            from_col = marker(from_node, "col")
            to_row = marker(to_node, "row") if to_node is not None else from_row
            to_col = marker(to_node, "col") if to_node is not None else from_col
            key = (source, from_row, from_col)
            if key in seen:
                continue
            seen.add(key)
            raw_items.append((from_row, from_col, to_row or from_row, to_col or from_col, source))

    for item in worksheet_dispimg_items(zf, sheet_path):
        key = (item[4], item[0], item[1])
        if key in seen:
            continue
        seen.add(key)
        raw_items.append(item)

    raw_items = filter_contained_images(raw_items)
    raw_items.sort(key=lambda item: (item[0], item[1], item[4]))
    images = []
    for index, (from_row, from_col, to_row, to_col, source) in enumerate(raw_items, start=1):
        rel_path = f"image_{index}.webp"
        convert_to_webp(zf, source, output_dir / rel_path)
        images.append(ImageItem(index, source, rel_path, from_row, from_col, to_row, to_col))
    return images


def detect_titles(cells: list[Cell], styles: dict[str, dict]) -> list[Cell]:
    counts: dict[str | None, int] = {}
    for cell in cells:
        counts[cell.style] = counts.get(cell.style, 0) + 1
    sized = [
        styles.get(style or "", {}).get("font_size")
        for style in counts
        if styles.get(style or "", {}).get("font_size") is not None
    ]
    body_size = min(sized) if sized else None
    titles = []
    for cell in cells:
        feature = styles.get(cell.style or "", {})
        size = feature.get("font_size")
        if cell.row <= 1:
            continue
        if size is not None and body_size is not None and size > body_size and (feature.get("bold") or feature.get("filled")):
            cell.is_title = True
            titles.append(cell)
    if not titles:
        titles = fallback_titles(cells, styles, counts)
    return sorted(titles, key=lambda item: (item.row, item.col))


def fallback_titles(cells: list[Cell], styles: dict[str, dict], counts: dict[str | None, int]) -> list[Cell]:
    titles: list[Cell] = []
    for cell in cells:
        if cell.row <= 1:
            continue
        text = cell.text.strip()
        feature = styles.get(cell.style or "", {})
        distinctive = feature.get("filled") or feature.get("bold") or counts.get(cell.style, 0) <= 8
        short_left_cell = len(text) <= 32 and cell.start_col == cell.end_col and cell.start_col <= 5
        following_rows = [
            other
            for other in cells
            if cell.row < other.row <= cell.row + 5
        ]
        has_following_content = bool(following_rows)
        if distinctive and short_left_cell and has_following_content:
            cell.is_title = True
            titles.append(cell)
    return titles


def title_lanes(titles: list[Cell]) -> list[int]:
    bands: list[tuple[int, int]] = []
    for title in sorted((item for item in titles if item.row > 1), key=lambda item: (item.start_col, item.end_col)):
        placed = False
        for index, (left, right) in enumerate(bands):
            if horizontal_overlap(title.start_col, title.end_col, left, right) > 0:
                bands[index] = (min(left, title.start_col), max(right, title.end_col))
                placed = True
                break
        if not placed:
            bands.append((title.start_col, title.end_col))
    cols = sorted(left for left, _right in bands)
    if not cols:
        return [1]
    return cols


def lane_bounds_for_title(title: Cell, lanes: list[int]) -> tuple[int, int]:
    lane = max((col for col in lanes if col <= title.start_col), default=lanes[0])
    index = lanes.index(lane)
    left = 1 if index == 0 else lane
    right = lanes[index + 1] - 1 if index + 1 < len(lanes) else 999999
    return left, right


def horizontal_overlap(a_left: int, a_right: int, b_left: int, b_right: int) -> int:
    return max(0, min(a_right, b_right) - max(a_left, b_left) + 1)


def assign_to_title(
    start_row: int,
    start_col: int,
    end_col: int,
    titles: list[Cell],
    lanes: list[int],
) -> Cell | None:
    candidates = []
    for title in titles:
        if title.row > start_row:
            continue
        left, right = lane_bounds_for_title(title, lanes)
        overlap = horizontal_overlap(start_col, end_col, left, right)
        if overlap <= 0:
            continue
        if title.row == 1 and title.start_col > 1:
            has_later_lane_title = any(
                other.ref != title.ref
                and other.row > title.row
                and other.row <= start_row
                and horizontal_overlap(start_col, end_col, *lane_bounds_for_title(other, lanes)) > 0
                and horizontal_overlap(title.start_col, title.end_col, *lane_bounds_for_title(other, lanes)) > 0
                for other in titles
            )
            if has_later_lane_title:
                continue
        distance = 0 if start_col >= title.start_col else title.start_col - end_col
        candidates.append((title.row, overlap, -abs(distance), title))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    return candidates[-1][3]


def should_attach_to_right_title_prefix(cell: Cell, current_title: Cell, right_title: Cell) -> bool:
    text = cell.text.strip()
    if len(text) > 24:
        return False
    if not (re.fullmatch(r"\d+[.、]?.*", text) or is_standalone_prefix(text)):
        return False
    same_row_right_cell = right_title.row == cell.row and right_title.start_col == cell.end_col + 1
    before_right_title = cell.row > right_title.row and cell.end_col + 1 == right_title.start_col
    below_same_row_titles = current_title.row == right_title.row and cell.row > right_title.row
    return (
        (before_right_title or (below_same_row_titles and cell.end_col + 1 == right_title.start_col))
        and current_title.start_col < cell.start_col < right_title.start_col
        and not same_row_right_cell
    )


def adjacent_right_title_for_cell(cell: Cell, current_title: Cell, titles: list[Cell]) -> Cell | None:
    candidates = [
        title
        for title in titles
        if title.row <= cell.row
        and not (title.row == 1 and current_title.row > 1)
        and title.start_col == cell.end_col + 1
        and title.start_col > current_title.start_col
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item.row, item.start_col))
    right_title = candidates[-1]
    if should_attach_to_right_title_prefix(cell, current_title, right_title):
        return right_title
    return None


def is_standalone_prefix(text: str) -> bool:
    stripped = text.strip()
    return bool(stripped) and len(stripped) <= 3 and not re.search(r"[\w\u4e00-\u9fff]", stripped)


def filter_prefixed_titles(cells: list[Cell], titles: list[Cell], lanes: list[int]) -> list[Cell]:
    filtered: list[Cell] = []
    for title in titles:
        left_prefixes = [
            cell
            for cell in cells
            if cell.row == title.row
            and cell.end_col + 1 == title.start_col
            and is_standalone_prefix(cell.text)
        ]
        parent_title = None
        if left_prefixes:
            prefix = left_prefixes[-1]
            parent_title = assign_to_title(prefix.start_row, prefix.start_col, prefix.end_col, titles, lanes)
        if parent_title is not None and parent_title.ref != title.ref and parent_title.row < title.row:
            title.is_title = False
            continue
        filtered.append(title)
    return filtered


def add_top_preamble_titles(
    cells: list[Cell],
    titles: list[Cell],
    styles: dict[str, dict],
    lanes: list[int],
    images: list[ImageItem],
) -> list[Cell]:
    added: list[Cell] = []
    for cell in cells:
        if cell.row != 1 or cell.start_col <= 1 or not cell.text.strip():
            continue
        feature = styles.get(cell.style or "", {})
        if not feature.get("bold"):
            continue
        later_right_titles = [
            title
            for title in titles
            if title.row > cell.row
            and title.start_col >= cell.start_col - 1
        ]
        if feature.get("filled") and later_right_titles:
            continue
        left_titles = [
            title
            for title in titles
            if title.row > cell.row and title.start_col < cell.start_col
        ]
        if feature.get("filled") and len(left_titles) < 2:
            continue
        first_title_row = min((title.row for title in later_right_titles), default=999999)
        if feature.get("filled"):
            first_left_title = min(left_titles, key=lambda item: item.row)
            later_left_titles = [title for title in left_titles if title.row > first_left_title.row]
            first_title_row = min((title.row for title in later_left_titles), default=999999)
            first_left_has_visual_body = any(
                image.from_row > first_left_title.row
                and image.from_row < first_title_row
                and image.from_col < cell.start_col
                for image in images
            )
            if first_left_has_visual_body:
                continue
        has_preamble_body = any(
            other.row > cell.row
            and other.row < first_title_row
            and other.start_col >= cell.start_col - 1
            for other in cells
        )
        if feature.get("filled"):
            has_numbered_rule = any(
                other.row > cell.row
                and other.row < first_title_row
                and other.start_col >= cell.start_col - 1
                and first_number(other.text) is not None
                for other in cells
            )
            if not has_numbered_rule:
                continue
        if has_preamble_body:
            cell.is_title = True
            added.append(cell)
    if not added:
        return titles
    return sorted([*titles, *added], key=lambda item: (item.row, item.col))


def adjacent_left_prefix_title_for_cell(cell: Cell, titles: list[Cell]) -> Cell | None:
    text = cell.text.strip()
    if len(text) > 24:
        return None
    if not (re.fullmatch(r"\d+[.、]?.*", text) or is_standalone_prefix(text)):
        return None
    candidates = [
        title
        for title in titles
        if title.row < cell.row and title.start_col == cell.end_col + 1
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda item: item.row)
    return candidates[-1]


def same_row_left_prefix_title_for_cell(cell: Cell, cells: list[Cell], titles: list[Cell], lanes: list[int]) -> Cell | None:
    left_prefixes = [
        other
        for other in cells
        if other.row == cell.row
        and other.end_col + 1 == cell.start_col
        and is_standalone_prefix(other.text)
    ]
    if not left_prefixes:
        return None
    prefix = left_prefixes[-1]
    return assign_to_title(prefix.start_row, prefix.start_col, prefix.end_col, titles, lanes)


def immediate_right_content_title_for_prefix(
    cell: Cell,
    cells: list[Cell],
    titles: list[Cell],
    lanes: list[int],
) -> Cell | None:
    stripped = cell.text.strip()
    if not (is_standalone_prefix(stripped) or re.fullmatch(r"\d+[.、]", stripped)):
        return None
    right_cells = [
        other
        for other in cells
        if other.row == cell.row
        and other.start_col == cell.end_col + 1
        and not other.is_title
        and other.text.strip()
    ]
    if not right_cells:
        return None
    right_cell = sorted(right_cells, key=lambda item: item.start_col)[0]
    return assign_to_title(right_cell.start_row, right_cell.start_col, right_cell.end_col, titles, lanes)


def top_preamble_title_for_cell(cell: Cell, titles: list[Cell]) -> Cell | None:
    top_titles = [
        title
        for title in titles
        if title.row == 1
        and title.start_col > 1
        and cell.start_col >= title.start_col - 1
    ]
    if not top_titles:
        return None
    left_titles = [
        title
        for title in titles
        if title.row > 1 and title.start_col < cell.start_col
    ]
    if not left_titles:
        return None
    first_left_title = min(left_titles, key=lambda item: item.row)
    later_left_titles = [
        title
        for title in left_titles
        if title.row > first_left_title.row
    ]
    section_end_row = min((title.row for title in later_left_titles), default=999999)
    if not (first_left_title.row < cell.row < section_end_row):
        return None
    same_band_later_titles = [
        title
        for title in titles
        if title.row > 1
        and title.start_col >= min(item.start_col for item in top_titles) - 1
        and title.row <= cell.row
    ]
    if same_band_later_titles:
        return None
    return max(top_titles, key=lambda item: item.start_col)


def first_number(text: str) -> int | None:
    match = re.match(r"\s*(\d+)[.、]", text)
    return int(match.group(1)) if match else None


def section_number_span(cells: list[Cell]) -> tuple[int | None, int | None]:
    numbers = [number for cell in cells if (number := first_number(cell.text)) is not None]
    if not numbers:
        return None, None
    return min(numbers), max(numbers)


def should_continue_previous_flow(
    cell: Cell,
    current_title: Cell,
    previous_title: Cell,
    previous_cells: list[Cell],
    current_images: list[ImageItem],
) -> bool:
    number = first_number(cell.text)
    if number is None:
        return False
    if not current_images:
        return False
    first_image_row = min(image.from_row for image in current_images)
    if cell.row >= first_image_row:
        return False
    _previous_min, previous_max = section_number_span(previous_cells)
    if previous_max is None or number != previous_max + 1:
        return False
    previous_flow_cols = [item.start_col for item in previous_cells if item.start_col > previous_title.start_col]
    if not previous_flow_cols:
        return False
    previous_flow_col = min(previous_flow_cols)
    same_flow_col = abs(cell.start_col - previous_flow_col) <= 1
    title_in_different_lane = current_title.start_col < previous_flow_col
    return same_flow_col and title_in_different_lane


def visual_section_for_cell(
    cell: Cell,
    assigned_title: Cell,
    titles: list[Cell],
    images_by_title: dict[str, list[ImageItem]],
) -> Cell | None:
    paired_left_visual_titles = [
        title
        for title in titles
        if title.start_col < assigned_title.start_col
        and abs(title.row - assigned_title.row) <= 2
        and images_by_title.get(title.ref)
    ]
    if paired_left_visual_titles:
        return None
    earlier_left_visual_titles = [
        title
        for title in titles
        if title.row < assigned_title.row
        and title.start_col < assigned_title.start_col
        and images_by_title.get(title.ref)
    ]
    if earlier_left_visual_titles:
        return None

    candidates = []
    for title in titles:
        if title.row > cell.start_row or title.row < assigned_title.row:
            continue
        if title.start_col >= cell.start_col:
            continue
        visual_images = [
            image
            for image in images_by_title.get(title.ref, [])
            if image.from_col < cell.start_col and image.from_row <= cell.start_row <= image.to_row
        ]
        if not visual_images:
            continue
        candidates.append((title.row, max(image.to_row for image in visual_images), title.start_col, title))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    return candidates[-1][3]


def visual_section_for_auxiliary_image(
    image: ImageItem,
    assigned_title: Cell,
    titles: list[Cell],
    images_by_title: dict[str, list[ImageItem]],
) -> Cell | None:
    same_row_left_titles = [
        title
        for title in titles
        if title.row == assigned_title.row and title.start_col < assigned_title.start_col
    ]
    if not same_row_left_titles:
        return None
    candidates = []
    for title in titles:
        if title.row <= assigned_title.row or title.row > image.from_row:
            continue
        if title.start_col >= assigned_title.start_col:
            continue
        visual_images = [
            item
            for item in images_by_title.get(title.ref, [])
            if item.from_col < image.from_col and item.from_row <= image.from_row <= item.to_row
        ]
        if not visual_images:
            continue
        candidates.append((title.row, max(item.to_row for item in visual_images), title.start_col, title))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    return candidates[-1][3]


def same_row_visual_table_title_for_cell(
    cell: Cell,
    cells: list[Cell],
    titles: list[Cell],
    lanes: list[int],
    images_by_title: dict[str, list[ImageItem]],
) -> Cell | None:
    candidates = []
    for title in titles:
        if title.row >= cell.row or title.start_col >= cell.start_col:
            continue
        title_images = images_by_title.get(title.ref, [])
        if len(title_images) < 2:
            continue
        if any(len(group) > 1 for group in image_groups(title_images)):
            continue
        first_image_row = min(image.from_row for image in title_images)
        if not (first_image_row - 2 <= cell.row < first_image_row):
            continue
        same_row_left_cells = [
            other
            for other in cells
            if other.row == cell.row
            and not other.is_title
            and other.end_col < cell.start_col
            and assign_to_title(other.start_row, other.start_col, other.end_col, titles, lanes) == title
        ]
        if same_row_left_cells:
            candidates.append((title.row, title.start_col, title))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[-1][2]


def cell_lines_for_row(row_cells: list[Cell]) -> list[tuple[int, str]]:
    if not row_cells:
        return []
    cells = sorted(row_cells, key=lambda item: item.col)
    groups: list[list[Cell]] = []
    current: list[Cell] = []
    for cell in cells:
        if current and cell.start_col > current[-1].end_col + 1:
            groups.append(current)
            current = []
        current.append(cell)
    if current:
        groups.append(current)
    result = []
    for group in groups:
        non_prefix = [c for c in group if not is_standalone_prefix(c.text)]
        if len(non_prefix) >= 3:
            # 多列结构：每个 cell 独立为 lane，保留列结构
            for cell in group:
                if cell.text.strip():
                    result.append((cell.start_col, cell.text.strip()))
        else:
            text = " ".join(item.text for item in group if item.text).strip()
            if text:
                result.append((group[0].start_col, text))
    return result


def table_matrix(lines_by_row: list[tuple[int, list[tuple[int, str]]]]) -> tuple[list[int], list[list[str]]]:
    lane_starts = sorted({start for _, items in lines_by_row for start, _ in items})
    lane_map: dict[int, int] = {}
    canonical: list[int] = []
    for lane in lane_starts:
        raw_previous = lane - 1
        if raw_previous in lane_map:
            appears_with_raw_previous = any(
                {start for start, _ in items}.issuperset({raw_previous, lane})
                for _, items in lines_by_row
            )
            if not appears_with_raw_previous:
                lane_map[lane] = lane_map[raw_previous]
                continue
        if canonical:
            prev = canonical[-1]
            appears_together = any(
                {start for start, _ in items}.issuperset({prev, lane})
                for _, items in lines_by_row
            )
            if lane == prev + 1 and not appears_together:
                lane_map[lane] = prev
                continue
        lane_map[lane] = lane
        canonical.append(lane)
    anchor_items = next((items for _, items in lines_by_row if len(items) > 1), lines_by_row[0][1])
    header_raw_starts = {start for start, _ in anchor_items}
    header_lanes = {lane_map[start] for start in header_raw_starts}
    continuation_map: dict[int, int] = {}
    for lane in canonical:
        if lane in header_lanes:
            continue
        previous_header_lanes = [item for item in canonical if item < lane and item in header_lanes]
        paired_with_previous = sum(
            1
            for _, items in lines_by_row
            if lane in {lane_map[start] for start, _ in items}
            and any(
                previous in {lane_map[start] for start, _ in items}
                for previous in previous_header_lanes
            )
        )
        if previous_header_lanes and paired_with_previous <= 1:
            continuation_map[lane] = previous_header_lanes[-1]
    lane_starts = [lane for lane in canonical if lane not in continuation_map]
    rows = []
    for _, items in lines_by_row:
        by_lane: dict[int, str] = {}
        for start, text in items:
            lane = continuation_map.get(lane_map[start], lane_map[start])
            by_lane[lane] = f"{by_lane[lane]} {text}".strip() if lane in by_lane else text
        rows.append([by_lane.get(start, "") for start in lane_starts])
    if len(lane_starts) >= 3:
        first_values = [row[0] for row in rows if row[0]]
        leading_empty_count = 0
        for row in rows:
            if row[0]:
                break
            leading_empty_count += 1
        if first_values and len(first_values) <= 2 and leading_empty_count >= 3:
            adjusted_rows = []
            for row in rows:
                adjusted = list(row)
                if adjusted[0] and not adjusted[1]:
                    adjusted[1] = adjusted[0]
                adjusted_rows.append(adjusted[1:])
            rows = adjusted_rows
            lane_starts = lane_starts[1:]
    rows = align_leading_single_side_row(rows)
    if len(lane_starts) >= 3:
        keep_indexes = []
        sparse_indexes = []
        for index, _lane in enumerate(lane_starts):
            values = [row[index] for row in rows if row[index]]
            sparse_caption = (
                len(values) == 1
                and len(values[0].strip()) <= 20
                and not re.match(r"\s*(?:\d+(?:[.、\s]|$)|[▶>\-])", values[0].strip())
            )
            if sparse_caption:
                sparse_indexes.append(index)
            else:
                keep_indexes.append(index)
        if keep_indexes and len(keep_indexes) < len(lane_starts):
            adjusted_rows = [list(row) for row in rows]
            for sparse_index in sparse_indexes:
                if sparse_index <= 0 or sparse_index >= len(lane_starts) - 1:
                    continue
                target_index = sparse_index - 1
                if target_index == sparse_index or target_index >= len(lane_starts):
                    continue
                for row in adjusted_rows:
                    if row[sparse_index]:
                        row[target_index] = f"{row[target_index]} {row[sparse_index]}".strip() if row[target_index] else row[sparse_index]
                        row[sparse_index] = ""
            rows = [[row[index] for index in keep_indexes] for row in adjusted_rows]
            lane_starts = [lane_starts[index] for index in keep_indexes]
    return lane_starts, rows


def render_table(lines_by_row: list[tuple[int, list[tuple[int, str]]]]) -> list[str]:
    # Markdown tables are semantic reading blocks: Excel blank rows are only
    # grouping evidence and are not emitted as empty table rows.
    lane_starts, rows = table_matrix(lines_by_row)
    width = max(1, len(lane_starts))
    output = [
        "- | " + " | ".join("" for _ in range(width)) + " |",
        "  | " + " | ".join("---" for _ in range(width)) + " |",
    ]
    for row in rows:
        output.append("  | " + " | ".join(cell.replace("|", "\\|") for cell in row) + " |")
    return output


def align_leading_single_side_row(rows: list[list[str]]) -> list[list[str]]:
    if len(rows) < 2 or len(rows[0]) < 2:
        return rows
    first = rows[0]
    second = rows[1]
    first_left_empty = not first[0]
    first_right_only = any(first[1:])
    second_has_left = bool(second[0])
    second_has_right = any(second[1:])
    if first_left_empty and first_right_only and second_has_left and second_has_right:
        adjusted = [list(row) for row in rows]
        adjusted[0][0] = adjusted[1][0]
        adjusted[1][0] = ""
        return adjusted
    return rows


def has_header_based_table(lines_by_row: list[tuple[int, list[tuple[int, str]]]]) -> bool:
    if len(lines_by_row) < 2:
        return False
    _header_row, header_items = lines_by_row[0]
    if len(header_items) < 2:
        return False
    if any(re.match(r"\s*(?:\d+(?:[.、\s]|$)|[▶>\-])", text.strip()) for _, text in header_items):
        return False
    header_starts = [start for start, _ in header_items]
    content_rows = lines_by_row[1:]
    if not any(items for _, items in content_rows):
        return False

    used_headers: set[int] = set()
    for _row, items in content_rows:
        for start, _text in items:
            candidates = [header for header in header_starts if header <= start]
            if not candidates:
                return False
            used_headers.add(candidates[-1])
    return bool(used_headers)


def table_shape_detected(lines_by_row: list[tuple[int, list[tuple[int, str]]]]) -> bool:
    if not lines_by_row:
        return False
    lane_count = len({start for _, items in lines_by_row for start, _ in items})
    multi_lane_rows = sum(1 for _, items in lines_by_row if len(items) > 1)
    return (lane_count > 1 and multi_lane_rows >= 2) or has_header_based_table(lines_by_row)


def row_clusters(lines_by_row: list[tuple[int, list[tuple[int, str]]]]) -> list[list[tuple[int, list[tuple[int, str]]]]]:
    clusters: list[list[tuple[int, list[tuple[int, str]]]]] = []
    for item in lines_by_row:
        if clusters and item[0] == clusters[-1][-1][0] + 1:
            clusters[-1].append(item)
        else:
            clusters.append([item])
    return clusters


def image_groups(images: list[ImageItem]) -> list[list[ImageItem]]:
    groups: list[list[ImageItem]] = []
    current: list[ImageItem] = []
    current_end = -1
    for image in sorted(images, key=lambda item: (item.from_row, item.from_col)):
        if current and image.from_row > current_end:
            groups.append(sorted(current, key=lambda item: item.from_col))
            current = []
        current.append(image)
        current_end = max(current_end, image.to_row)
    if current:
        groups.append(sorted(current, key=lambda item: item.from_col))
    return groups


def image_column_groups(images: list[ImageItem]) -> list[list[ImageItem]]:
    groups: list[list[ImageItem]] = []
    for image in sorted(images, key=lambda item: (item.from_col, item.from_row)):
        placed = False
        for group in groups:
            group_left = min(item.from_col for item in group)
            group_right = max(item.to_col for item in group)
            if horizontal_overlap(image.from_col, image.to_col, group_left, group_right) > 0:
                group.append(image)
                placed = True
                break
        if not placed:
            groups.append([image])
    return [sorted(group, key=lambda item: (item.from_row, item.from_col)) for group in groups]


def leading_visual_image_group(
    rows: list[tuple[int, list[tuple[int, str]]]],
    images: list[ImageItem],
) -> list[ImageItem]:
    if len(images) < 2 or not rows:
        return []
    text_cols = [start for _, items in rows for start, _ in items]
    if not text_cols:
        return []
    min_text_col = min(text_cols)
    candidates = []
    for group in image_column_groups(images):
        min_text_col = min(text_cols)
        if len(group) < 2 or max(image.to_col for image in group) >= min_text_col:
            continue
        restart = False
        for image in group[1:]:
            nearby_rows = [
                item
                for item in rows
                if image.from_row - 1 <= item[0] <= image.from_row + 1
                and any(start > image.to_col for start, _ in item[1])
            ]
            if any(
                first_number(text) == 1
                for _, items in nearby_rows
                for _, text in items
            ):
                restart = True
                break
        if not restart:
            candidates.append(group)
    if not candidates:
        return []
    return max(
        candidates,
        key=lambda group: (
            len(group),
            max(image.to_row for image in group) - min(image.from_row for image in group),
            -min(image.from_col for image in group),
        ),
    )


def image_local_heading_items(
    image: ImageItem,
    rows: list[tuple[int, list[tuple[int, str]]]],
) -> set[tuple[int, int, str]]:
    candidates: set[tuple[int, int, str]] = set()
    for row, items in rows:
        if not (image.from_row - 2 <= row <= image.from_row):
            continue
        in_band = [(start, text) for start, text in items if image.from_col <= start <= image.to_col]
        left_items = [(start, text) for start, text in items if start < image.from_col]
        heading_items = in_band
        if heading_items and left_items:
            continue
        if not heading_items and len(left_items) == 1:
            heading_items = left_items
        if len(heading_items) != 1:
            continue
        start, text = heading_items[0]
        stripped = text.strip()
        if len(stripped) > 16:
            continue
        if is_label_text(stripped):
            continue
        if re.match(r"\s*(?:\d+(?:[.、\s]|$)|[▶>\-])", stripped) or is_standalone_prefix(stripped):
            continue
        candidates.add((row, start, text))
    return candidates


def has_repeated_short_label(cells: list[Cell], row_start: int, row_end: int) -> bool:
    by_row: dict[int, dict[str, int]] = {}
    for cell in cells:
        if not (row_start <= cell.row <= row_end):
            continue
        text = cell.text.strip()
        if len(text) > 20:
            continue
        by_row.setdefault(cell.row, {})[text] = by_row.setdefault(cell.row, {}).get(text, 0) + 1
    return any(count >= 2 for counts in by_row.values() for count in counts.values())


def cell_center_col(cell: Cell) -> float:
    return (cell.start_col + cell.end_col) / 2


def image_center_col(image: ImageItem) -> float:
    return (image.from_col + image.to_col) / 2


def horizontally_contains_cell(image: ImageItem, cell: Cell) -> bool:
    center = cell_center_col(cell)
    return image.from_col <= center <= image.to_col


def cell_is_covered_by_image(cell: Cell, image: ImageItem) -> bool:
    return (
        image.from_row <= cell.row <= image.to_row
        and horizontal_overlap(cell.start_col, cell.end_col, image.from_col, image.to_col) > 0
    )


def cell_is_image_start_heading(cell: Cell, image: ImageItem) -> bool:
    text = cell.text.strip()
    return (
        cell.row == image.from_row
        and cell.start_col <= image.to_col
        and len(text) <= 20
        and not re.match(r"\s*\d+(?:[.、\s]|$)", text)
        and not is_standalone_prefix(text)
    )


def filter_cells_hidden_by_images(cells: list[Cell], images: list[ImageItem]) -> list[Cell]:
    visible: list[Cell] = []
    for cell in cells:
        covering_images = [image for image in images if cell_is_covered_by_image(cell, image)]
        if not covering_images:
            # Check if cell is in an image's row range but not column range (noise cell)
            in_image_row = any(image.from_row <= cell.row <= image.to_row for image in images)
            if in_image_row and len(cell.text.strip()) <= 2 and re.match(r'^[.\-·]+$', cell.text.strip()):
                continue
            visible.append(cell)
            continue
        if any(cell_is_image_start_heading(cell, image) for image in covering_images):
            visible.append(cell)
            continue
        if re.match(r"\s*\d+(?:[.、\s]|$)", cell.text.strip()) or is_standalone_prefix(cell.text):
            visible.append(cell)
            continue
        has_same_row_external_flow = any(
            other.row == cell.row
            and other.ref != cell.ref
            and any(other.start_col > image.to_col for image in covering_images)
            for other in cells
        )
        has_same_row_external_duplicate = any(
            other.row == cell.row
            and other.ref != cell.ref
            and other.text.strip() == cell.text.strip()
            and any(other.start_col > image.to_col for image in covering_images)
            for other in cells
        )
        if has_same_row_external_duplicate:
            visible.append(cell)
            continue
        if not has_same_row_external_flow:
            visible.append(cell)
    return visible


def has_aligned_caption_row(cells: list[Cell], group: list[ImageItem]) -> bool:
    group_start = min(image.from_row for image in group)
    before_by_row: dict[int, list[Cell]] = {}
    for cell in cells:
        if cell.row >= group_start:
            continue
        before_by_row.setdefault(cell.row, []).append(cell)

    for row_cells in before_by_row.values():
        image_indexes: set[int] = set()
        for cell in row_cells:
            aligned_images = [
                image
                for image in group
                if horizontally_contains_cell(image, cell)
            ]
            if not aligned_images:
                continue
            image = min(
                aligned_images,
                key=lambda item: abs(cell_center_col(cell) - image_center_col(item)),
            )
            image_indexes.add(image.index)
        if len(image_indexes) >= 2:
            return True
    return False


def is_label_text(text: str) -> bool:
    return text.strip().endswith((":", "："))


def use_card_rendering(cells: list[Cell], images: list[ImageItem]) -> bool:
    for group in image_groups(images):
        if len(group) < 2:
            continue
        row_start = min(image.from_row for image in group)
        row_end = max(image.to_row for image in group)
        if has_repeated_short_label(cells, row_start, row_end):
            return True
        if has_aligned_caption_row(cells, group):
            return True
    return False


def nearest_image(cell: Cell, images: list[ImageItem]) -> ImageItem:
    center = cell_center_col(cell)
    return min(
        images,
        key=lambda image: (
            abs(center - image_center_col(image)),
            abs(cell.row - image.from_row),
            image.from_col,
        ),
    )


def cell_can_be_image_caption(cell: Cell, image: ImageItem) -> bool:
    if len(cell.text.strip()) > 20:
        return False
    if cell.start_col > image.to_col:
        return False
    return horizontal_overlap(cell.start_col, cell.end_col, image.from_col, image.to_col) > 0 or cell.end_col < image.from_col


def emphasized_group_heading_row(
    before_cells: list[Cell],
    title_start: int | None,
    section_left: int,
    styles: dict[str, dict],
) -> int | None:
    if title_start is None:
        return None
    rows = sorted({cell.row for cell in before_cells if cell.row < title_start})
    for row in reversed(rows):
        row_cells = [cell for cell in before_cells if cell.row == row]
        if len(row_cells) != 1:
            continue
        cell = row_cells[0]
        style = styles.get(cell.style or "", {})
        emphasized = bool(style.get("bold")) or bool(style.get("filled"))
        if emphasized and cell.start_col == section_left and not is_label_text(cell.text):
            return row
    return None


def render_card_section(cells: list[Cell], images: list[ImageItem], styles: dict[str, dict]) -> list[str]:
    output: list[str] = []
    consumed_rows: set[int] = set()
    groups = image_groups(images)
    previous_group_end = -1
    section_left = min((cell.start_col for cell in cells), default=1)

    for group in groups:
        group_start = min(image.from_row for image in group)
        group_end = max(image.to_row for image in group)
        before_cells = [
            cell
            for cell in cells
            if previous_group_end < cell.row < group_start and cell.row not in consumed_rows
        ]
        candidate_rows = sorted({cell.row for cell in before_cells})
        title_start: int | None = None
        for row in candidate_rows:
            if sum(1 for cell in before_cells if cell.row == row) >= 2:
                title_start = row
        if title_start is None and candidate_rows:
            title_start = candidate_rows[-1]

        group_heading_row = emphasized_group_heading_row(before_cells, title_start, section_left, styles)
        if title_start is not None:
            pre_rows: dict[int, list[Cell]] = {}
            for cell in before_cells:
                if cell.row < title_start and cell.row != group_heading_row:
                    pre_rows.setdefault(cell.row, []).append(cell)
            pre_lines = [(row, cell_lines_for_row(row_cells)) for row, row_cells in sorted(pre_rows.items())]
            pre_lines = [(row, items) for row, items in pre_lines if items]
            output.extend(render_rows(pre_lines))
            if pre_lines:
                output.append("")
            if group_heading_row is not None:
                heading_cells = [cell for cell in before_cells if cell.row == group_heading_row]
                output.append(f"## {heading_cells[0].text}")
                output.append("")
                consumed_rows.add(group_heading_row)

        title_cells: dict[int, list[Cell]] = {image.index: [] for image in group}
        if title_start is not None:
            for cell in before_cells:
                if title_start <= cell.row < group_start:
                    assigned_image = nearest_image(cell, group)
                    if cell_can_be_image_caption(cell, assigned_image):
                        title_cells[assigned_image.index].append(cell)
                        consumed_rows.add(cell.row)

        prop_cells: dict[int, list[Cell]] = {image.index: [] for image in group}
        for cell in cells:
            if not (group_start <= cell.row <= group_end):
                continue
            assigned_image = nearest_image(cell, group)
            if cell.row < assigned_image.from_row and not is_label_text(cell.text):
                if cell_can_be_image_caption(cell, assigned_image):
                    title_cells[assigned_image.index].append(cell)
                else:
                    prop_cells[assigned_image.index].append(cell)
            else:
                prop_cells[assigned_image.index].append(cell)
            consumed_rows.add(cell.row)

        for image in group:
            heading_cells = sorted(title_cells.get(image.index, []), key=lambda item: (item.row, item.col))
            heading = " / ".join(dict.fromkeys(cell.text for cell in heading_cells))
            if heading:
                output.append(f"### {heading}")
                output.append("")
            output.append(f"- ![image {image.index}](./{image.path})")
            props_by_row: dict[int, list[Cell]] = {}
            for cell in sorted(prop_cells.get(image.index, []), key=lambda item: (item.row, item.col)):
                props_by_row.setdefault(cell.row, []).append(cell)
            prop_lines = [(row, cell_lines_for_row(row_cells)) for row, row_cells in sorted(props_by_row.items())]
            prop_lines = [(row, items) for row, items in prop_lines if items]
            if prop_lines:
                output.append("")
                output.extend(render_rows(prop_lines))
            output.append("")

        previous_group_end = group_end

    after_cells = [cell for cell in cells if cell.row > previous_group_end and cell.row not in consumed_rows]
    after_rows: dict[int, list[Cell]] = {}
    for cell in after_cells:
        after_rows.setdefault(cell.row, []).append(cell)
    after_lines = [(row, cell_lines_for_row(row_cells)) for row, row_cells in sorted(after_rows.items())]
    after_lines = [(row, items) for row, items in after_lines if items]
    output.extend(render_rows(after_lines))

    compacted: list[str] = []
    for line in output:
        if line == "" and compacted and compacted[-1] == "":
            continue
        compacted.append(line)
    while compacted and compacted[-1] == "":
        compacted.pop()
    return compacted


def split_list_table_cells(cells: list[Cell]) -> tuple[list[Cell], list[Cell]] | None:
    if not cells:
        return None
    starts = sorted({cell.start_col for cell in cells})
    gaps = [(right - left, left, right) for left, right in zip(starts, starts[1:])]
    large_gaps = [item for item in gaps if item[0] >= 4]
    if not large_gaps:
        return None
    _gap, left_col, right_col = max(large_gaps)
    left_cells = [cell for cell in cells if cell.start_col <= left_col]
    right_cells = [cell for cell in cells if cell.start_col >= right_col]
    if not left_cells or not right_cells:
        return None
    right_rows: dict[int, list[Cell]] = {}
    for cell in right_cells:
        right_rows.setdefault(cell.row, []).append(cell)
    right_lines = [
        (row, [(cell.start_col, cell.text) for cell in sorted(row_cells, key=lambda item: item.start_col)])
        for row, row_cells in sorted(right_rows.items())
    ]
    right_lines = [(row, items) for row, items in right_lines if items]
    if not table_shape_detected(right_lines):
        return None
    left_cols = sorted({cell.start_col for cell in left_cells})
    left_is_prefix_list = len(left_cols) <= 2 and (
        len(left_cols) == 1 or left_cols[1] == left_cols[0] + 1
    )
    if not left_is_prefix_list:
        return None
    return left_cells, right_cells


def render_mixed_list_table_section(cells: list[Cell]) -> list[str] | None:
    split = split_list_table_cells(cells)
    if split is None:
        return None
    left_cells, right_cells = split

    left_rows: dict[int, list[Cell]] = {}
    for cell in left_cells:
        left_rows.setdefault(cell.row, []).append(cell)
    left_lines = [(row, cell_lines_for_row(row_cells)) for row, row_cells in sorted(left_rows.items())]
    left_lines = [(row, items) for row, items in left_lines if items]

    right_rows: dict[int, list[Cell]] = {}
    for cell in right_cells:
        right_rows.setdefault(cell.row, []).append(cell)
    right_lines = [
        (row, [(cell.start_col, cell.text) for cell in sorted(row_cells, key=lambda item: item.start_col)])
        for row, row_cells in sorted(right_rows.items())
    ]
    right_lines = [(row, items) for row, items in right_lines if items]

    output = render_rows(left_lines)
    if output and right_lines:
        output.append("")
    if right_lines and len(right_lines[0][1]) == 1 and any(len(items) >= 3 for _, items in right_lines[1:]):
        output.append(f"- {right_lines[0][1][0][1]}")
        output.append("")
        output.extend(render_table(right_lines[1:]))
    else:
        output.extend(render_table(right_lines))
    return output


def visual_group_start(image: ImageItem, rows: list[tuple[int, list[tuple[int, str]]]]) -> int:
    for cluster in row_clusters(rows):
        cluster_start = cluster[0][0]
        cluster_end = cluster[-1][0]
        min_text_col = min(start for _, items in cluster for start, _ in items)
        max_text_col = max(start for _, items in cluster for start, _ in items)
        touches_left_edge = image.to_col == min_text_col and max_text_col > image.to_col
        visually_left = image.to_col < min_text_col or touches_left_edge
        vertically_related = cluster_start <= image.to_row and cluster_end >= image.from_row
        if visually_left and vertically_related:
            previous_rows = [item for item in rows if item[0] < cluster_start]
            if previous_rows and rows_render_as_table(cluster):
                previous_cluster = row_clusters(previous_rows)[-1]
                combined_cluster = previous_cluster + cluster
                previous_cols = [start for _, items in previous_cluster for start, _ in items]
                cluster_cols = [start for _, items in cluster for start, _ in items]
                previous_cluster_end = previous_cluster[-1][0]
                previous_near_image = previous_cluster_end >= image.from_row - 5
                same_table_band = (
                    previous_cols
                    and rows_render_as_table(combined_cluster)
                    and cluster_cols
                    and min(previous_cols) == min(cluster_cols)
                    and max(previous_cols) <= max(cluster_cols)
                    and previous_near_image
                )
                if same_table_band:
                    return previous_cluster[0][0]
            if previous_rows and rows_render_as_table(cluster):
                header_row, header_items = previous_rows[-1]
                no_intervening_text = all(
                    row == header_row or row >= cluster_start
                    for row, _ in rows
                    if row < cluster_start
                )
                header_in_table_band = (
                    header_row < image.from_row
                    and no_intervening_text
                    and len(header_items) > 1
                    and all(min_text_col <= start <= max_text_col for start, _ in header_items)
                )
                if header_in_table_band:
                    return header_row
            return cluster_start
    return image.from_row


def side_note_cluster(
    image: ImageItem,
    rows: list[tuple[int, list[tuple[int, str]]]],
) -> list[tuple[int, list[tuple[int, str]]]] | None:
    for cluster in row_clusters(rows):
        cluster_start = cluster[0][0]
        cluster_end = cluster[-1][0]
        max_text_col = max(start for _, items in cluster for start, _ in items)
        visually_right = image.from_col > max_text_col
        vertically_related = cluster_start <= image.to_row and cluster_end >= image.from_row
        if visually_right and vertically_related:
            return cluster
    return None


def append_suffix(text: str, suffix: str | None) -> str:
    return f"{text} {suffix}".strip() if suffix else text


def render_rows(
    lines_by_row: list[tuple[int, list[tuple[int, str]]]],
    row_suffixes: dict[int, str] | None = None,
) -> list[str]:
    if not lines_by_row:
        return []
    row_suffixes = row_suffixes or {}
    if table_shape_detected(lines_by_row):
        if row_suffixes:
            adjusted: list[tuple[int, list[tuple[int, str]]]] = []
            for row, items in lines_by_row:
                suffix = row_suffixes.get(row)
                if suffix and items:
                    adjusted_items = list(items)
                    start, text = adjusted_items[-1]
                    adjusted_items[-1] = (start, append_suffix(text, suffix))
                    adjusted.append((row, adjusted_items))
                else:
                    adjusted.append((row, items))
            lines_by_row = adjusted
        return render_table(lines_by_row)
    output = []
    for row, items in lines_by_row:
        for _, text in items:
            output.append(f"- {append_suffix(text, row_suffixes.get(row))}")
    return output


def append_row_suffixes(
    lines_by_row: list[tuple[int, list[tuple[int, str]]]],
    row_suffixes: dict[int, str],
) -> list[tuple[int, list[tuple[int, str]]]]:
    if not row_suffixes:
        return lines_by_row
    adjusted = []
    for row, items in lines_by_row:
        suffix = row_suffixes.get(row)
        if suffix and items:
            adjusted_items = list(items)
            start, text = adjusted_items[-1]
            adjusted_items[-1] = (start, append_suffix(text, suffix))
            adjusted.append((row, adjusted_items))
        else:
            adjusted.append((row, items))
    return adjusted


def rows_render_as_table(lines_by_row: list[tuple[int, list[tuple[int, str]]]]) -> bool:
    return table_shape_detected(lines_by_row)


def below_auxiliary_target_row(
    image_group: list[ImageItem],
    before_rows: list[tuple[int, list[tuple[int, str]]]],
) -> int | None:
    if not before_rows:
        return None
    group_left = min(image.from_col for image in image_group)
    group_right = max(image.to_col for image in image_group)
    candidates = []
    for row, items in before_rows:
        item_cols = [start for start, _ in items]
        if not item_cols:
            continue
        same_or_near_flow = any(start <= group_right + 1 for start in item_cols)
        if same_or_near_flow:
            candidates.append(row)
    return candidates[-1] if candidates else None


def detached_image_heading(
    image_group: list[ImageItem],
    before_rows: list[tuple[int, list[tuple[int, str]]]],
) -> tuple[
    str | None,
    list[tuple[int, list[tuple[int, str]]]],
    tuple[int, list[tuple[int, str]]] | None,
]:
    if not before_rows:
        return None, before_rows, None
    image_start = min(image.from_row for image in image_group)
    row, items = before_rows[-1]
    if row != image_start - 1 or len(items) < 2:
        return None, before_rows, None
    group_left = min(image.from_col for image in image_group)
    group_right = max(image.to_col for image in image_group)
    sorted_items = sorted(items, key=lambda item: item[0])
    left_start, left_text = sorted_items[0]
    if left_start >= group_left:
        return None, before_rows, None
    other_starts = [start for start, _ in sorted_items[1:]]
    if not other_starts or min(other_starts) <= group_right:
        return None, before_rows, None
    if re.match(r"\s*(?:\d+(?:[.、\s]|$)|[▶>\-])", left_text.strip()):
        return None, before_rows, None
    remaining = [(start, text) for start, text in items if (start, text) != (left_start, left_text)]
    adjusted_before = list(before_rows[:-1])
    carry_row = (row, remaining) if remaining else None
    return left_text.strip(), adjusted_before, carry_row


def detached_image_heading_row(
    image_group: list[ImageItem],
    rows: list[tuple[int, list[tuple[int, str]]]],
) -> int | None:
    heading, _adjusted, _carry = detached_image_heading(
        image_group,
        [item for item in rows if item[0] < min(image.from_row for image in image_group)],
    )
    if not heading:
        return None
    return min(image.from_row for image in image_group) - 1


def nearby_image_heading(
    image_group: list[ImageItem],
    rows: list[tuple[int, list[tuple[int, str]]]],
    heading_band: tuple[int, int] | None = None,
    title_band: tuple[int, int] | None = None,
) -> tuple[
    str | None,
    list[tuple[int, list[tuple[int, str]]]],
]:
    image_start = min(image.from_row for image in image_group)
    group_left = min(image.from_col for image in image_group)
    group_right = max(image.to_col for image in image_group)
    if heading_band is not None and horizontal_overlap(group_left, group_right, heading_band[0], heading_band[1]) <= 0:
        return None, rows
    if title_band is not None and horizontal_overlap(group_left, group_right, title_band[0], title_band[1]) <= 0:
        return None, rows
    competing_left_text = [
        text
        for row, items in rows
        if image_start - 2 <= row <= image_start
        for start, text in items
        if start < group_left and text.strip()
    ]
    if competing_left_text:
        return None, rows
    heading_rows = {
        row
        for row, items in rows
        if row < image_start
        and any(group_left <= start <= group_right for start, _ in items)
        and not any(start < group_left for start, _ in items)
    }
    closest_heading_row = max(heading_rows) if heading_rows else None
    candidates = []
    for row, items in rows:
        if row not in (image_start - 1, image_start, closest_heading_row):
            continue
        heading_items = [(start, text) for start, text in items if group_left <= start <= group_right]
        if len(heading_items) != 1:
            continue
        start, text = heading_items[0]
        if re.match(r"\s*\d+(?:[.、\s]|$)", text.strip()) or is_standalone_prefix(text):
            continue
        if len(text.strip()) > 16:
            continue
        if row < image_start - 1:
            intervening = [item for item in rows if row < item[0] < image_start]
            if any(any(group_left <= start <= group_right for start, _ in items) for _, items in intervening):
                continue
            previous_same_band_rows = [
                item_row
                for item_row, items in rows
                if item_row < row
                and row - item_row <= 3
                and any(group_left <= start <= group_right for start, _ in items)
            ]
            if previous_same_band_rows:
                continue
        candidates.append((row, start, text, text.strip()))
    if not candidates:
        return None, rows
    row, start, raw_heading, heading = sorted(candidates, key=lambda item: (abs(item[0] - image_start), -item[1]))[0]
    adjusted = []
    for current_row, items in rows:
        if current_row == row:
            remaining = [(item_start, text) for item_start, text in items if (item_start, text) != (start, raw_heading)]
            if remaining:
                adjusted.append((current_row, remaining))
        else:
            adjusted.append((current_row, items))
    return heading, adjusted


def previous_output_is_list_item(output: list[str]) -> bool:
    for line in reversed(output):
        if not line:
            continue
        return line.startswith("- ")
    return False


def append_list_image_item(output: list[str], image_line: str) -> None:
    if output and output[-1] == "":
        previous = next((line for line in reversed(output[:-1]) if line), "")
        if previous.startswith("- "):
            output.pop()
    output.append(f"- {image_line}")


def append_nested_image_item(output: list[str], heading: str, image_line: str) -> None:
    if output and output[-1] == "":
        previous = next((line for line in reversed(output[:-1]) if line), "")
        if previous.startswith("- "):
            output.pop()
    output.append(f"- **{heading}**")
    output.append(f"  - {image_line}")


def render_nested_rows(lines_by_row: list[tuple[int, list[tuple[int, str]]]]) -> list[str]:
    nested = []
    for line in render_rows(lines_by_row):
        if line.startswith("- "):
            nested.append(f"  - {line[2:]}")
        else:
            nested.append(f"  {line}")
    return nested


def canonical_lane_map(lines_by_row: list[tuple[int, list[tuple[int, str]]]]) -> dict[int, int]:
    lane_starts = sorted({start for _, items in lines_by_row for start, _ in items})
    lane_map: dict[int, int] = {}
    canonical: list[int] = []
    for lane in lane_starts:
        raw_previous = lane - 1
        if raw_previous in lane_map:
            appears_with_raw_previous = any(
                {start for start, _ in items}.issuperset({raw_previous, lane})
                for _, items in lines_by_row
            )
            if not appears_with_raw_previous:
                lane_map[lane] = lane_map[raw_previous]
                continue
        if canonical:
            prev = canonical[-1]
            appears_together = any(
                {start for start, _ in items}.issuperset({prev, lane})
                for _, items in lines_by_row
            )
            if lane == prev + 1 and not appears_together:
                lane_map[lane] = prev
                continue
        lane_map[lane] = lane
        canonical.append(lane)
    return lane_map


def split_nested_heading_local_rows(
    rows: list[tuple[int, list[tuple[int, str]]]],
    image_group: list[ImageItem],
) -> tuple[
    list[tuple[int, list[tuple[int, str]]]],
    list[tuple[int, list[tuple[int, str]]]],
]:
    if not rows:
        return [], rows
    group_right = max(image.to_col for image in image_group)
    right_rows = [
        (row, [(start, text) for start, text in items if start > group_right])
        for row, items in rows
        if any(start > group_right for start, _text in items)
    ]
    if not right_rows:
        return [], rows
    lane_map = canonical_lane_map(right_rows)
    right_lanes = sorted({lane_map[start] for _row, items in right_rows for start, _text in items})
    if not right_lanes:
        return [], rows
    local_lane = right_lanes[0]
    has_parent_lane = len(right_lanes) > 1
    if not has_parent_lane and table_shape_detected(rows):
        return [], rows

    local_rows: list[tuple[int, list[tuple[int, str]]]] = []
    remaining_rows: list[tuple[int, list[tuple[int, str]]]] = []
    for row, items in rows:
        local_items = [(start, text) for start, text in items if start > group_right and lane_map.get(start) == local_lane]
        remaining_items = [(start, text) for start, text in items if (start, text) not in local_items]
        if local_items:
            local_rows.append((row, local_items))
        if remaining_items:
            remaining_rows.append((row, remaining_items))
    if has_parent_lane and local_rows:
        local_row_numbers = {row for row, _items in local_rows}
        paired_rows = {
            row
            for row, items in remaining_rows
            if row in local_row_numbers
            and any(start > group_right for start, _text in items)
        }
        if len(paired_rows) >= 2 and len(paired_rows) * 2 >= len(local_rows):
            return [], rows
    return local_rows, remaining_rows


def lane_continues_outside_image_span(
    rows: list[tuple[int, list[tuple[int, str]]]],
    image: ImageItem,
    local_rows: list[tuple[int, list[tuple[int, str]]]],
) -> bool:
    if not local_rows:
        return False
    right_rows = [
        (row, [(start, text) for start, text in items if start > image.to_col])
        for row, items in rows
        if any(start > image.to_col for start, _text in items)
    ]
    if not right_rows:
        return False
    lane_map = canonical_lane_map(right_rows)
    local_lanes = {
        lane_map[start]
        for _row, items in local_rows
        for start, _text in items
        if start in lane_map
    }
    if not local_lanes:
        return False
    if min(local_lanes) == min(lane_map.values()):
        return False
    outside_same_lane_rows = {
        row
        for row, items in right_rows
        if not (image.from_row <= row <= image.to_row)
        and any(lane_map.get(start) in local_lanes for start, _text in items)
    }
    return len(outside_same_lane_rows) >= 2


def heading_for_image(
    image: ImageItem,
    rows: list[tuple[int, list[tuple[int, str]]]],
) -> tuple[str | None, list[tuple[int, list[tuple[int, str]]]]]:
    candidates = []
    for row, items in rows:
        if not (image.from_row - 2 <= row <= image.from_row):
            continue
        in_band = [(start, text) for start, text in items if image.from_col <= start <= image.to_col]
        left_items = [(start, text) for start, text in items if start < image.from_col]
        heading_items = in_band
        if not heading_items and len(left_items) == 1:
            heading_items = left_items
        if len(heading_items) != 1:
            continue
        start, text = heading_items[0]
        stripped = text.strip()
        if len(stripped) > 16:
            continue
        if re.match(r"\s*\d+(?:[.、\s]|$)", stripped) or is_standalone_prefix(stripped):
            continue
        candidates.append((abs(row - image.from_row), row, start, text, stripped))
    if not candidates:
        return None, rows
    _distance, heading_row, heading_start, raw_heading, heading = sorted(candidates)[0]
    adjusted = []
    for row, items in rows:
        if row == heading_row:
            remaining = [(start, text) for start, text in items if (start, text) != (heading_start, raw_heading)]
            if remaining:
                adjusted.append((row, remaining))
        else:
            adjusted.append((row, items))
    return heading, adjusted


def append_visual_image_group(
    output: list[str],
    image_group: list[ImageItem],
    rows: list[tuple[int, list[tuple[int, str]]]],
) -> list[tuple[int, list[tuple[int, str]]]]:
    adjusted_rows = rows
    for row_group in image_groups(image_group):
        if len(row_group) > 1:
            append_list_image_item(output, " ".join(f"![image {item.index}](./{item.path})" for item in row_group))
            output.append("")
            continue
        image = row_group[0]
        heading, adjusted_rows = heading_for_image(image, adjusted_rows)
        if heading:
            append_nested_image_item(output, heading, f"![image {image.index}](./{image.path})")
        else:
            append_list_image_item(output, f"![image {image.index}](./{image.path})")
        output.append("")
    return adjusted_rows


def visual_stack_table_section(
    rows: list[tuple[int, list[tuple[int, str]]]],
    images: list[ImageItem],
) -> list[str] | None:
    if len(images) < 2 or not rows:
        return None
    candidates = []
    for group in image_column_groups(images):
        if len(group) < 2:
            continue
        first_image = group[0]
        local_heading_items = {
            item
            for image in group
            for item in image_local_heading_items(image, rows)
        }
        if not any(image_local_heading_items(image, rows) for image in group[1:]):
            first_width = first_image.to_col - first_image.from_col + 1
            if not any((image.to_col - image.from_col + 1) * 2 <= first_width for image in group[1:]):
                continue
            _remaining_rows, local_rows_by_image = split_image_local_list_rows(rows, group, {})
            if not any(local_rows_by_image.get(image.index) for image in group[1:]):
                continue
        table_rows = []
        for row, items in rows:
            adjusted_items = [
                (start, text)
                for start, text in items
                if (row, start, text) not in local_heading_items
            ]
            if adjusted_items:
                table_rows.append((row, adjusted_items))
        if not table_shape_detected(table_rows):
            continue
        group_right = max(image.to_col for image in group)
        right_item_count = sum(
            1
            for _row, items in table_rows
            for start, _text in items
            if start > group_right
        )
        left_item_count = sum(
            1
            for _row, items in table_rows
            for start, _text in items
            if start <= group_right
        )
        if right_item_count <= left_item_count:
            continue
        candidates.append(group)
    if not candidates:
        return None

    image_group = max(
        candidates,
        key=lambda group: (
            len(group),
            max(image.to_row for image in group) - min(image.from_row for image in group),
            -min(image.from_col for image in group),
        ),
    )
    if {image.index for image in image_group} != {image.index for image in images}:
        return None

    adjusted_rows = rows
    image_headings: dict[int, str | None] = {}
    for row_group in image_groups(image_group):
        if len(row_group) > 1:
            continue
        image = row_group[0]
        heading, adjusted_rows = heading_for_image(image, adjusted_rows)
        image_headings[image.index] = heading
    adjusted_rows, local_rows_by_image = split_image_local_list_rows(adjusted_rows, image_group, image_headings)

    output: list[str] = []
    for row_group in image_groups(image_group):
        if len(row_group) > 1:
            append_list_image_item(output, " ".join(f"![image {item.index}](./{item.path})" for item in row_group))
            output.append("")
            continue
        image = row_group[0]
        heading = image_headings.get(image.index)
        if heading:
            append_nested_image_item(output, heading, f"![image {image.index}](./{image.path})")
        else:
            append_list_image_item(output, f"![image {image.index}](./{image.path})")
        image_local_rows = local_rows_by_image.get(image.index, [])
        if image_local_rows:
            output.extend(render_nested_rows(image_local_rows))
        output.append("")
    group_right = max(image.to_col for image in image_group)
    first_image_row = min(image.from_row for image in image_group)
    table_min_col = min((start for _row, items in adjusted_rows for start, _text in items), default=999999)
    side_note_rows = {
        row
        for row, items in adjusted_rows
        if row < first_image_row
        and len(items) == 1
        and items[0][0] > group_right
        and items[0][0] != table_min_col
    }
    side_notes = [item for item in adjusted_rows if item[0] in side_note_rows]
    table_rows = [item for item in adjusted_rows if item[0] not in side_note_rows]
    if side_notes:
        if output and output[-1] != "":
            output.append("")
        output.extend(render_rows(side_notes))
        output.append("")
    if output and output[-1] != "":
        output.append("")
    output.extend(render_rows(table_rows))
    while output and output[-1] == "":
        output.pop()
    return output


def primary_left_visual_images(
    rows: list[tuple[int, list[tuple[int, str]]]],
    images: list[ImageItem],
) -> list[ImageItem]:
    if not rows or not images:
        return []
    text_min_col = min(start for _row, items in rows for start, _text in items)
    text_max_col = max(start for _row, items in rows for start, _text in items)
    candidates = [
        group
        for group in image_column_groups(images)
        if (
            max(image.to_col for image in group) < text_min_col
            or (
                max(image.to_col for image in group) == text_min_col
                and text_max_col > text_min_col
            )
        )
    ]
    if not candidates:
        return []
    return max(
        candidates,
        key=lambda group: (
            max(image.to_row for image in group) - min(image.from_row for image in group),
            len(group),
            -min(image.from_col for image in group),
        ),
    )


def weak_single_image_visual_table(rows: list[tuple[int, list[tuple[int, str]]]]) -> bool:
    if len(rows) < 5:
        return False
    lane_starts, matrix_rows = table_matrix(rows)
    if len(lane_starts) != 2:
        return False
    multi_lane_rows = sum(1 for _row, items in rows if len(items) > 1)
    if multi_lane_rows != 1:
        return False
    left_nonempty = sum(1 for left, _right in matrix_rows if left)
    right_nonempty = sum(1 for _left, right in matrix_rows if right)
    return left_nonempty >= 1 and right_nonempty >= 5


def single_image_caption_table_section(
    rows: list[tuple[int, list[tuple[int, str]]]],
    images: list[ImageItem],
    raw_rows: list[tuple[int, list[Cell]]] | None = None,
) -> list[str] | None:
    if len(images) != 1 or not rows:
        return None
    image = images[0]
    lead_in_rows = [
        item
        for item in rows
        if image.from_row - 1 <= item[0] < image.from_row
        and len(item[1]) == 1
    ]
    pre_rows = [item for item in rows if item[0] < image.from_row - 1]
    if any(len(items) > 1 for _row, items in pre_rows):
        return None
    body_rows = lead_in_rows + [item for item in rows if item[0] >= image.from_row]
    if len(body_rows) < 4:
        return None

    lane_counts: dict[int, int] = {}
    for _row, items in body_rows:
        for start, _text in items:
            lane_counts[start] = lane_counts.get(start, 0) + 1
    if len(lane_counts) < 2:
        return None

    left_lane = min(lane_counts)
    if left_lane > image.to_col + 2 or lane_counts[left_lane] != 1:
        return None

    right_item_count = sum(count for start, count in lane_counts.items() if start != left_lane)
    if right_item_count < 4:
        return None

    left_values = [
        " ".join(text for start, text in sorted(items, key=lambda item: item[0]) if start == left_lane).strip()
        for _row, items in body_rows
        if any(start == left_lane for start, _text in items)
    ]
    left_value = " ".join(value for value in left_values if value).strip()
    if not left_value or len(left_value) > 40:
        return None

    right_values_by_row = []
    for row, items in body_rows:
        value = " ".join(
            text
            for start, text in sorted(items, key=lambda item: item[0])
            if start != left_lane
        ).strip()
        if value:
            right_values_by_row.append((row, value))
    if len(right_values_by_row) < 4:
        return None

    field_table: list[tuple[int, list[tuple[int, str]]]] = []
    field_rows = set()
    if raw_rows:
        raw_by_row = {
            row: sorted(row_cells, key=lambda item: item.start_col)
            for row, row_cells in raw_rows
        }
        candidates = []
        for row, row_cells in raw_by_row.items():
            if row < image.from_row:
                continue
            right_cells = [cell for cell in row_cells if cell.start_col != left_lane and cell.start_col > image.to_col]
            if len(right_cells) >= 3:
                candidates.append((row, right_cells))
        for first, second in zip(candidates, candidates[1:]):
            if second[0] != first[0] + 1:
                continue
            first_cols = [cell.start_col for cell in first[1]]
            second_cols = [cell.start_col for cell in second[1]]
            if len(first_cols) == len(second_cols) and first_cols == second_cols:
                field_rows.update({first[0], second[0]})
                field_cols = first_cols
                field_start_by_col = {cell.start_col: cell.start_col for cell in first[1]}
                next_row = second[0] + 1
                while next_row in raw_by_row:
                    right_cells = [
                        cell
                        for cell in raw_by_row[next_row]
                        if cell.start_col != left_lane and cell.start_col > image.to_col
                    ]
                    if not right_cells:
                        break
                    table_cells = [
                        cell
                        for cell in right_cells
                        if cell.start_col in field_start_by_col
                    ]
                    if len(table_cells) != len(right_cells):
                        break
                    used_cols = {cell.start_col for cell in table_cells}
                    sparse_continuation = len(table_cells) >= 2 or any(col != field_cols[0] for col in used_cols)
                    if not sparse_continuation:
                        break
                    field_rows.add(next_row)
                    next_row += 1
                break
        if field_rows:
            for row in sorted(field_rows):
                field_table.append((row, [(cell.start_col, cell.text) for cell in raw_by_row[row] if cell.start_col != left_lane and cell.start_col > image.to_col]))
            right_values_by_row = [
                (row, value)
                for row, value in right_values_by_row
                if row not in field_rows
            ]
            if len(right_values_by_row) < 2:
                return None

    right_values = [value for _row, value in right_values_by_row]
    output: list[str] = []
    if pre_rows:
        output.extend(render_rows(pre_rows))
        output.append("")
    append_list_image_item(output, f"![image {image.index}](./{image.path})")
    output.append("")
    output.extend(["- |  |  |", "  | --- | --- |"])
    output.append(f"  | {left_value.replace('|', '\\|')} | {right_values[0].replace('|', '\\|')} |")
    for value in right_values[1:]:
        output.append(f"  |  | {value.replace('|', '\\|')} |")
    if field_table:
        output.append("")
        output.extend(render_table(field_table))
    while output and output[-1] == "":
        output.pop()
    return output


def image_markdown(image: ImageItem, alt_text: str | None = None) -> str:
    alt = alt_text.strip() if alt_text and alt_text.strip() else f"image {image.index}"
    alt = alt.replace("[", "\\[").replace("]", "\\]")
    return f"![{alt}](./{image.path})"


def caption_text_candidate(text: str) -> bool:
    stripped = text.strip()
    if not stripped or len(stripped) > 20:
        return False
    if is_label_text(stripped):
        return False
    if re.match(r"\s*(?:\d+(?:[.、\s]|$)|[▶>\-])", stripped):
        return False
    return True


def normalize_caption_text(text: str) -> str:
    stripped = text.strip()
    if len(stripped) >= 2 and (
        (stripped[0] == "（" and stripped[-1] == "）")
        or (stripped[0] == "(" and stripped[-1] == ")")
    ):
        return stripped[1:-1].strip()
    return stripped


def image_caption_candidates(
    images: list[ImageItem],
    rows: list[tuple[int, list[tuple[int, str]]]],
) -> tuple[dict[int, str], set[tuple[int, int, str]]]:
    captions: dict[int, str] = {}
    consumed: set[tuple[int, int, str]] = set()
    for row, items in sorted(rows, key=lambda item: item[0]):
        candidates = [
            (start, text)
            for start, text in items
            if caption_text_candidate(text)
        ]
        if not candidates:
            continue
        matches: list[tuple[ImageItem, int, str, int]] = []
        used_images: set[int] = set()
        valid = True
        for start, text in sorted(candidates, key=lambda item: item[0]):
            image_candidates = [
                (
                    min(abs(row - image.from_row), abs(row - image.to_row)),
                    image,
                )
                for image in images
                if image.index not in captions
                and image.index not in used_images
                and image.from_col <= start <= image.to_col
                and min(abs(row - image.from_row), abs(row - image.to_row)) <= 2
            ]
            if not image_candidates:
                valid = False
                break
            image_candidates.sort(key=lambda item: (item[0], item[1].from_col, item[1].index))
            if len(image_candidates) > 1 and image_candidates[0][0] == image_candidates[1][0]:
                valid = False
                break
            image = image_candidates[0][1]
            used_images.add(image.index)
            matches.append((image, start, text, row))
        if not valid:
            continue
        for image, start, text, match_row in matches:
            captions[image.index] = normalize_caption_text(text)
            consumed.add((match_row, start, text))
    return captions, consumed


def primary_visual_table_section(
    rows: list[tuple[int, list[tuple[int, str]]]],
    images: list[ImageItem],
    caption_rows: list[tuple[int, list[tuple[int, str]]]] | None = None,
) -> list[str] | None:
    if not images:
        return None
    if len(images) < 2 and not weak_single_image_visual_table(rows):
        return None
    if len(images) >= 2 and not table_shape_detected(rows):
        return None
    primary_images = primary_left_visual_images(rows, images)
    if not primary_images:
        return None
    if len(primary_images) != 1:
        return None
    primary_indexes = {image.index for image in primary_images}
    auxiliary_images = [image for image in images if image.index not in primary_indexes]
    image_captions, caption_items = image_caption_candidates(auxiliary_images, caption_rows or rows)
    auxiliary_heading_items = {
        item
        for image in auxiliary_images
        for item in image_local_heading_items(image, rows)
    } | caption_items
    adjusted_table_rows = []
    for row, items in rows:
        adjusted_items = [
            (start, text)
            for start, text in items
            if (row, start, text) not in auxiliary_heading_items
        ]
        if adjusted_items:
            adjusted_table_rows.append((row, adjusted_items))

    output: list[str] = []
    adjusted_rows = append_visual_image_group(output, primary_images, adjusted_table_rows)
    lane_starts, matrix_rows = table_matrix(adjusted_rows)
    if len(lane_starts) != 2:
        output.extend(render_rows(adjusted_rows))
        if auxiliary_images:
            output.append("")
            for row_group in image_groups(auxiliary_images):
                append_list_image_item(output, " ".join(f"![image {item.index}](./{item.path})" for item in row_group))
        while output and output[-1] == "":
            output.pop()
        return output

    row_numbers = [row for row, _items in adjusted_rows]
    lane_items: list[list[tuple[int, str]]] = [[], []]
    for row_number, row_values in zip(row_numbers, matrix_rows):
        for index, value in enumerate(row_values):
            if value:
                lane_items[index].append((row_number, value))

    suffixes: list[dict[int, list[tuple[ImageItem, str]]]] = [{}, {}]
    for image in sorted(auxiliary_images, key=lambda item: (item.from_row, item.from_col)):
        image_line = image_markdown(image, image_captions.get(image.index))
        preceding_lanes = [index for index, lane in enumerate(lane_starts) if lane <= image.from_col]
        lane_index = preceding_lanes[-1] if preceding_lanes else 0
        lane_rows = [row for row, _text in lane_items[lane_index]]
        preceding_rows = [row for row in lane_rows if row <= image.from_row]
        target_row = max(preceding_rows) if preceding_rows else (lane_rows[0] if lane_rows else image.from_row)
        suffixes[lane_index].setdefault(target_row, []).append((image, image_line))

    compacted_lanes: list[list[str]] = []
    for lane_index, items in enumerate(lane_items):
        compacted = []
        for row_number, value in items:
            suffix = " ".join(
                image_line
                for _image, image_line in sorted(
                    suffixes[lane_index].get(row_number, []),
                    key=lambda item: (item[0].from_col, item[0].from_row, item[0].index),
                )
            )
            compacted.append(append_suffix(value, suffix))
        compacted_lanes.append(compacted)

    width = max(len(compacted_lanes[0]), len(compacted_lanes[1]))
    output.extend(["- |  |  |", "  | --- | --- |"])
    for index in range(width):
        left = compacted_lanes[0][index] if index < len(compacted_lanes[0]) else ""
        right = compacted_lanes[1][index] if index < len(compacted_lanes[1]) else ""
        output.append(f"  | {left.replace('|', '\\|')} | {right.replace('|', '\\|')} |")
    while output and output[-1] == "":
        output.pop()
    return output


def split_image_local_list_rows(
    rows: list[tuple[int, list[tuple[int, str]]]],
    images: list[ImageItem],
    image_headings: dict[int, str | None] | None = None,
) -> tuple[
    list[tuple[int, list[tuple[int, str]]]],
    dict[int, list[tuple[int, list[tuple[int, str]]]]],
]:
    by_image: dict[int, list[tuple[int, list[tuple[int, str]]]]] = {}
    local_items: set[tuple[int, int, str]] = set()
    image_headings = image_headings or {}
    for image in images[1:]:
        if image_headings.get(image.index):
            span_rows = [
                (row, items)
                for row, items in rows
                if image.from_row <= row <= image.to_row
            ]
            image_local_rows, _remaining = split_nested_heading_local_rows(span_rows, [image])
            if image_local_rows:
                if lane_continues_outside_image_span(rows, image, image_local_rows):
                    continue
                by_image[image.index] = image_local_rows
                local_items.update((row, start, text) for row, items in image_local_rows for start, text in items)
                continue
        else:
            image_cluster = [
                (row, items)
                for row, items in rows
                if image.from_row - 2 <= row
                and row <= image.to_row + 20
                and any(start > image.to_col for start, _text in items)
            ]
            image_cluster_cols = {start for _row, items in image_cluster for start, _text in items}
            if (
                image_cluster
                and any(row <= image.from_row for row, _items in image_cluster)
                and any(row >= image.from_row for row, _items in image_cluster)
                and len(image_cluster_cols) >= 2
                and table_shape_detected(image_cluster)
            ):
                by_image[image.index] = image_cluster
                local_items.update((row, start, text) for row, items in image_cluster for start, text in items)
                continue
        candidates = [
            (row, items)
            for row, items in rows
            if image.from_row <= row <= image.to_row
            and len(items) == 1
            and image.to_col < items[0][0]
        ]
        if len(candidates) >= 2 and not table_shape_detected(candidates):
            by_image[image.index] = candidates
            local_items.update((row, start, text) for row, items in candidates for start, text in items)
    adjusted_rows = []
    for row, items in rows:
        remaining_items = [(start, text) for start, text in items if (row, start, text) not in local_items]
        if remaining_items:
            adjusted_rows.append((row, remaining_items))
    return adjusted_rows, by_image


def visual_stack_text_section(
    rows: list[tuple[int, list[tuple[int, str]]]],
    images: list[ImageItem],
) -> list[str] | None:
    if len(images) < 2 or not rows:
        return None
    candidates = []
    for group in image_column_groups(images):
        if len(group) < 2:
            continue
        first_image = group[0]
        same_visual_span = all(
            image.from_col == first_image.from_col and image.to_col == first_image.to_col
            for image in group[1:]
        )
        if not same_visual_span:
            continue
        later_heading_items = {
            image.index: image_local_heading_items(image, rows)
            for image in group[1:]
        }
        if not any(later_heading_items.values()):
            _remaining_rows, local_rows_by_image = split_image_local_list_rows(rows, group, {})
            if not any(local_rows_by_image.get(image.index) for image in group[1:]):
                continue
        if any(
            heading_items and not any(row == image.from_row for row, _start, _text in heading_items)
            for image in group[1:]
            for heading_items in [later_heading_items[image.index]]
        ):
            continue
        adjusted_rows = rows
        for image in group:
            _heading, adjusted_rows = heading_for_image(image, adjusted_rows)
        if table_shape_detected(adjusted_rows):
            continue
        if any(row > first_image.to_row for row, _items in adjusted_rows):
            continue
        later_images = group[1:]
        overlaps_later_image_body = any(
            image.from_row <= row <= image.to_row
            for image in later_images
            for row, _items in adjusted_rows
        )
        if overlaps_later_image_body:
            continue
        candidates.append(group)
    if not candidates:
        return None

    image_group = max(
        candidates,
        key=lambda group: (
            len(group),
            max(image.to_row for image in group) - min(image.from_row for image in group),
            -min(image.from_col for image in group),
        ),
    )
    if {image.index for image in image_group} != {image.index for image in images}:
        return None

    output: list[str] = []
    adjusted_rows = append_visual_image_group(output, image_group, rows)
    output.extend(render_rows(adjusted_rows))
    while output and output[-1] == "":
        output.pop()
    return output


def visual_stack_individual_table_section(
    rows: list[tuple[int, list[tuple[int, str]]]],
    images: list[ImageItem],
) -> list[str] | None:
    if len(images) < 2 or not rows:
        return None
    candidates = []
    for group in image_column_groups(images):
        if len(group) < 2:
            continue
        first_image = group[0]
        same_visual_span = all(
            image.from_col == first_image.from_col and image.to_col == first_image.to_col
            for image in group[1:]
        )
        if not same_visual_span:
            continue
        if {image.index for image in group} != {image.index for image in images}:
            continue
        candidates.append(sorted(group, key=lambda item: (item.from_row, item.from_col)))
    if not candidates:
        return None

    image_group = max(
        candidates,
        key=lambda group: (
            len(group),
            max(image.to_row for image in group) - min(image.from_row for image in group),
            -min(image.from_col for image in group),
        ),
    )

    adjusted_rows = rows
    image_headings: dict[int, str | None] = {}
    for image in image_group:
        heading, adjusted_rows = heading_for_image(image, adjusted_rows)
        image_headings[image.index] = heading

    local_rows_by_image: dict[int, list[tuple[int, list[tuple[int, str]]]]] = {
        image.index: [] for image in image_group
    }
    outside_rows: list[tuple[int, list[tuple[int, str]]]] = []
    for row, items in adjusted_rows:
        target_image: ImageItem | None = None
        for index, image in enumerate(image_group):
            next_start = image_group[index + 1].from_row if index + 1 < len(image_group) else 999999
            if image.from_row <= row < next_start and any(start >= image.to_col for start, _text in items):
                target_image = image
                break
        if target_image is None:
            outside_rows.append((row, items))
        else:
            local_rows_by_image[target_image.index].append((row, items))

    if outside_rows:
        first_image_row = min(image.from_row for image in image_group)
        last_image_next = 999999
        if any(row >= first_image_row and row < last_image_next for row, _items in outside_rows):
            return None

    local_table_count = sum(
        1
        for image in image_group
        if table_shape_detected(local_rows_by_image.get(image.index, []))
    )
    if local_table_count < 2:
        return None
    if any(
        local_rows_by_image.get(image.index)
        and not table_shape_detected(local_rows_by_image[image.index])
        for image in image_group
    ):
        return None

    output: list[str] = []
    leading_rows = [item for item in outside_rows if item[0] < image_group[0].from_row]
    trailing_rows = [item for item in outside_rows if item[0] >= image_group[-1].from_row]
    if leading_rows:
        output.extend(render_rows(leading_rows))
        output.append("")
    for image in image_group:
        heading = image_headings.get(image.index)
        if heading:
            append_nested_image_item(output, heading, f"![image {image.index}](./{image.path})")
        else:
            append_list_image_item(output, f"![image {image.index}](./{image.path})")
        local_rows = local_rows_by_image.get(image.index, [])
        if local_rows:
            output.append("")
            output.extend(render_table(local_rows))
        output.append("")
    if trailing_rows:
        output.extend(render_rows(trailing_rows))
    while output and output[-1] == "":
        output.pop()
    return output


def split_rows_for_images(
    lines_by_row: list[tuple[int, list[tuple[int, str]]]],
    images: list[ImageItem],
    title_band: tuple[int, int] | None = None,
    raw_rows: list[tuple[int, list[Cell]]] | None = None,
) -> list[str]:
    output: list[str] = []
    rows = list(lines_by_row)
    single_image_table = single_image_caption_table_section(rows, images, raw_rows=raw_rows)
    if single_image_table is not None:
        return single_image_table
    stack_individual_table_body = visual_stack_individual_table_section(rows, images)
    if stack_individual_table_body is not None:
        return stack_individual_table_body
    stack_table_body = visual_stack_table_section(rows, images)
    if stack_table_body is not None:
        return stack_table_body
    stack_text_body = visual_stack_text_section(rows, images)
    if stack_text_body is not None:
        return stack_text_body
    column_groups = image_column_groups(images)
    heading_band = None
    if column_groups:
        left_group = min(column_groups, key=lambda group: min(image.from_col for image in group))
        heading_band = (min(image.from_col for image in left_group), max(image.to_col for image in left_group))
    leading_group = leading_visual_image_group(rows, images)
    if leading_group:
        consumed = {image.index for image in leading_group}
        rows = append_visual_image_group(output, leading_group, rows)
        images = [image for image in images if image.index not in consumed]
    image_queue = sorted(images, key=lambda item: (item.from_row, item.from_col))

    while rows or image_queue:
        if not image_queue:
            output.extend(render_rows(rows))
            break
        image = image_queue.pop(0)
        image_group = [image]
        while image_queue and image_queue[0].from_row == image.from_row:
            candidate = image_queue.pop(0)
            image_group.append(candidate)
        group_start = visual_group_start(image, rows)
        nearby_heading, rows = nearby_image_heading(image_group, rows, heading_band, title_band)
        if nearby_heading:
            group_start = min(group_start, min(image.from_row for image in image_group))
        before = [item for item in rows if item[0] < group_start]
        after_or_same = [item for item in rows if item[0] >= group_start]
        image_heading, before, carry_row = detached_image_heading(image_group, before)
        if nearby_heading and not image_heading:
            image_heading = nearby_heading
        if before:
            output.extend(render_rows(before))
            rows = after_or_same
        elif carry_row is not None:
            rows = after_or_same
        side_cluster = side_note_cluster(image, rows)
        if side_cluster:
            side_rows = {row for row, _ in side_cluster}
            row_suffixes = {
                image.from_row: " ".join(f"![image {item.index}](./{item.path})" for item in image_group)
            }
            output.extend(render_rows(side_cluster, row_suffixes))
            rows = [item for item in rows if item[0] not in side_rows]
            output.append("")
        else:
            target_row = below_auxiliary_target_row(image_group, before)
            image_line = " ".join(f"![image {item.index}](./{item.path})" for item in image_group)
            if target_row is not None and before:
                append_list_image_item(output, image_line)
                rows = after_or_same
            else:
                if image_heading:
                    append_nested_image_item(output, image_heading, image_line)
                    rows = after_or_same
                    next_group_start = visual_group_start(image_queue[0], rows) if image_queue else 999999
                    group = [item for item in rows if item[0] < next_group_start]
                    if carry_row is not None:
                        group = [carry_row] + group
                    rows = [item for item in rows if item[0] >= next_group_start]
                    local_group, group = split_nested_heading_local_rows(group, image_group)
                    if local_group:
                        output.extend(render_nested_rows(local_group))
                    if group:
                        output.append("")
                    output.extend(render_rows(group))
                    if rows:
                        output.append("")
                    continue
                if before and not rows_render_as_table(before):
                    append_list_image_item(output, image_line)
                elif (
                    not before
                    and len(image_group) >= 2
                    and rows
                    and previous_output_is_list_item(output)
                ):
                    append_list_image_item(output, image_line)
                else:
                    if before and output and output[-1] != "":
                        output.append("")
                    append_list_image_item(output, image_line)
                    output.append("")
                    if image_queue:
                        next_image = image_queue[0]
                        next_group_start = visual_group_start(next_image, rows)
                        next_before = [item for item in rows if item[0] < next_group_start]
                        next_heading, adjusted_next_before, next_carry = detached_image_heading([next_image], next_before)
                        if next_heading:
                            next_image_group = [image_queue.pop(0)]
                            while image_queue and image_queue[0].from_row == next_image.from_row:
                                next_image_group.append(image_queue.pop(0))
                            next_next_group_start = visual_group_start(image_queue[0], rows) if image_queue else 999999
                            group = list(adjusted_next_before)
                            if next_carry is not None:
                                group.append(next_carry)
                            group.extend(item for item in rows if next_group_start <= item[0] < next_next_group_start)
                            rows = [item for item in rows if item[0] >= next_next_group_start]
                            append_nested_image_item(
                                output,
                                next_heading,
                                " ".join(f"![image {item.index}](./{item.path})" for item in next_image_group),
                            )
                            if group:
                                output.append("")
                            output.extend(render_rows(group))
                            if rows:
                                output.append("")
                            continue
                next_group_start = visual_group_start(image_queue[0], rows) if image_queue else 999999
                if image_queue:
                    next_heading_row = detached_image_heading_row([image_queue[0]], rows)
                    if next_heading_row is None:
                        next_nearby_heading, _next_rows = nearby_image_heading([image_queue[0]], rows, heading_band, title_band)
                        if next_nearby_heading:
                            next_heading_row = image_queue[0].from_row - 1
                    if next_heading_row is not None:
                        next_group_start = min(next_group_start, next_heading_row)
                group = [item for item in rows if item[0] < next_group_start]
                if carry_row is not None:
                    group = [carry_row] + group
                rows = [item for item in rows if item[0] >= next_group_start]
                output.extend(render_rows(group))
                if rows:
                    output.append("")

    compacted: list[str] = []
    for line in output:
        if line == "" and compacted and compacted[-1] == "":
            continue
        compacted.append(line)
    output = compacted
    while output and output[-1] == "":
        output.pop()
    return output


def safe_dir_name(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r'[<>:"/\\|?*【】]+', "_", stem).strip(" _")


def parse_md_list(section_names: tuple[str, ...] = ("已验收需保护", "未纳入保护")) -> list[str]:
    text = MD_LIST.read_text(encoding="utf-8")
    in_input_list = False
    files = []
    for line in text.splitlines():
        if line.startswith("## "):
            in_input_list = line.strip() in {f"## {name}" for name in section_names}
            continue
        if not in_input_list:
            continue
        match = re.match(r"- `(.+?\.xlsx)`", line.strip())
        if match:
            file_name = match.group(1)
            if file_name not in files:
                files.append(file_name)
    return files


def normalize_match_text(value: str) -> str:
    normalized = value.strip().strip('"').strip("'")
    if normalized.lower().endswith(".xlsx"):
        normalized = normalized[:-5]
    return safe_dir_name(normalized).lower()


def file_matches_query(file_name: str, query: str) -> bool:
    raw_query = query.strip().strip('"').strip("'")
    if not raw_query:
        return False
    stem = Path(file_name).stem.strip()
    output_name = safe_dir_name(file_name)
    if raw_query.isdigit():
        return stem.startswith(f"{raw_query}-") or stem == raw_query
    normalized_query = normalize_match_text(raw_query)
    candidates = {
        normalize_match_text(file_name),
        normalize_match_text(stem),
        normalize_match_text(output_name),
    }
    return normalized_query in candidates


def filter_files(files: list[str], query: str | None) -> list[str]:
    if not query:
        return files
    return [file_name for file_name in files if file_matches_query(file_name, query)]


def section_output_sort_key(title: Cell, titles: list[Cell], lanes: list[int]) -> tuple[float, int, int]:
    if title.row == 1 and title.start_col > 1:
        left_titles = [
            other
            for other in titles
            if other.row > title.row and other.start_col < title.start_col
        ]
        later_same_lane_titles = [
            other
            for other in titles
            if other.ref != title.ref
            and other.row > title.row
            and horizontal_overlap(title.start_col, title.end_col, *lane_bounds_for_title(other, lanes)) > 0
        ]
        if later_same_lane_titles:
            first_later_row = min(other.row for other in later_same_lane_titles)
            preceding_left_titles = [
                other
                for other in titles
                if other.row > title.row
                and other.row < first_later_row
                and other.start_col < title.start_col
            ]
            if preceding_left_titles:
                return (max(other.row for other in preceding_left_titles) + 0.5, title.start_col, title.end_col)
        if left_titles:
            first_left_title = min(left_titles, key=lambda item: (item.row, item.start_col))
            return (first_left_title.row + 0.5, title.start_col, title.end_col)
    return (float(title.row), title.start_col, title.end_col)


def paired_rule_cells_by_visual_title(
    titles: list[Cell],
    sections: dict[str, dict],
    images_by_title: dict[str, list[ImageItem]],
) -> tuple[dict[str, list[Cell]], set[str]]:
    by_visual: dict[str, list[Cell]] = {}
    consumed_rule_titles: set[str] = set()
    titles_by_row: dict[int, list[Cell]] = {}
    for title in titles:
        titles_by_row.setdefault(title.row, []).append(title)

    for row_titles in titles_by_row.values():
        ordered = sorted(row_titles, key=lambda item: item.start_col)
        if len(ordered) < 2:
            continue
        for rule_title in ordered[1:]:
            if sections[rule_title.ref]["images"]:
                continue
            rule_cells = list(sections[rule_title.ref]["cells"])
            if not rule_cells:
                continue
            visual_titles = sorted(
                (
                    title
                    for title in titles
                    if title.row > rule_title.row
                    and title.start_col < rule_title.start_col
                    and images_by_title.get(title.ref)
                ),
                key=lambda item: (item.row, item.start_col),
            )
            if not visual_titles:
                continue
            has_leading_rule_cells = min(cell.row for cell in rule_cells) < visual_titles[0].row
            attached_any = False
            for index, visual_title in enumerate(visual_titles):
                next_row = visual_titles[index + 1].row if index + 1 < len(visual_titles) else 999999
                cells_for_visual = [
                    cell
                    for cell in rule_cells
                    if visual_title.row <= cell.row < next_row
                ]
                if cells_for_visual:
                    by_visual.setdefault(visual_title.ref, []).extend(cells_for_visual)
                    attached_any = True
            if attached_any and not has_leading_rule_cells:
                consumed_rule_titles.add(rule_title.ref)
    return by_visual, consumed_rule_titles


def trim_cells_before_later_visual_title(
    title: Cell,
    titles: list[Cell],
    sections: dict[str, dict],
    images_by_title: dict[str, list[ImageItem]],
) -> list[Cell]:
    section = sections[title.ref]
    cells = list(section["cells"])
    if not cells or section["images"]:
        return cells
    same_row_left_titles = [
        other
        for other in titles
        if other.row == title.row and other.start_col < title.start_col
    ]
    if not same_row_left_titles:
        return cells
    later_visual_titles = sorted(
        (
            other
            for other in titles
            if other.row > title.row
            and other.start_col < title.start_col
            and images_by_title.get(other.ref)
        ),
        key=lambda item: (item.row, item.start_col),
    )
    if not later_visual_titles:
        return cells
    first_later_visual_row = later_visual_titles[0].row
    if min(cell.row for cell in cells) >= first_later_visual_row:
        return cells
    return [cell for cell in cells if cell.row < first_later_visual_row]


def align_paired_rule_cells_to_section(section_cells: list[Cell], rule_cells: list[Cell]) -> list[Cell]:
    if not section_cells or not rule_cells:
        return rule_cells
    section_first_row = min(cell.row for cell in section_cells)
    rule_first_row = min(cell.row for cell in rule_cells)
    section_number_rows = {
        number: cell.row
        for cell in sorted(section_cells, key=lambda item: (item.row, item.start_col))
        if (number := first_number(cell.text)) is not None
    }
    rule_number_rows = [
        (cell.row, number)
        for cell in sorted(rule_cells, key=lambda item: (item.row, item.start_col))
        if (number := first_number(cell.text)) is not None
    ]
    if section_first_row != rule_first_row and section_number_rows and rule_number_rows:
        shifted: list[Cell] = []
        emitted_rows_end = section_first_row - 1
        cells_by_group: list[tuple[int, int | None, list[Cell]]] = []
        for index, (group_row, number) in enumerate(rule_number_rows):
            next_group_row = rule_number_rows[index + 1][0] if index + 1 < len(rule_number_rows) else 999999
            group_cells = [
                cell
                for cell in rule_cells
                if group_row <= cell.row < next_group_row
            ]
            cells_by_group.append((group_row, number, group_cells))
        leading_cells = [cell for cell in rule_cells if cell.row < rule_number_rows[0][0]]
        if leading_cells:
            cells_by_group.insert(0, (rule_first_row, None, leading_cells))
        for group_row, number, group_cells in cells_by_group:
            if not group_cells:
                continue
            target_row = section_number_rows.get(number, group_row + section_first_row - rule_first_row)
            if target_row <= emitted_rows_end:
                target_row = emitted_rows_end + 1
            offset = target_row - group_row
            shifted.extend(
                replace(
                    cell,
                    row=cell.row + offset,
                    start_row=cell.start_row + offset,
                    end_row=cell.end_row + offset,
                )
                for cell in group_cells
            )
            emitted_rows_end = max(cell.row + offset for cell in group_cells)
        return shifted
    offset = section_first_row - rule_first_row
    if offset == 0:
        return rule_cells
    return [
        replace(
            cell,
            row=cell.row + offset,
            start_row=cell.start_row + offset,
            end_row=cell.end_row + offset,
        )
        for cell in rule_cells
    ]


def row_texts(cells: list[Cell]) -> list[str]:
    rows: dict[int, list[Cell]] = {}
    for cell in cells:
        rows.setdefault(cell.row, []).append(cell)
    lines = [(row, cell_lines_for_row(row_cells)) for row, row_cells in sorted(rows.items())]
    lines = [(row, items) for row, items in lines if items]
    return [" ".join(item_text for _start, item_text in items).strip() for _row, items in lines]


def render_paired_rule_table(section_cells: list[Cell], rule_cells: list[Cell]) -> list[str] | None:
    left_rows = row_texts(section_cells)
    right_rows = row_texts(rule_cells)
    if not left_rows or not right_rows:
        return None

    table_rows: list[tuple[str, str]] = []
    width = max(len(left_rows), len(right_rows))
    for index in range(width):
        table_rows.append((
            left_rows[index] if index < len(left_rows) else "",
            right_rows[index] if index < len(right_rows) else "",
        ))

    if not table_rows:
        return None
    output = ["- |  |  |", "  | --- | --- |"]
    for left, right in table_rows:
        output.append(f"  | {left.replace('|', '\\|')} | {right.replace('|', '\\|')} |")
    return output


def render_paired_rule_section(
    flow_images: list[ImageItem],
    section_cells: list[Cell],
    rule_cells: list[Cell],
) -> list[str] | None:
    table = render_paired_rule_table(section_cells, rule_cells)
    if table is None:
        return None
    output: list[str] = []
    for row_group in image_groups(flow_images):
        append_list_image_item(output, " ".join(f"![image {item.index}](./{item.path})" for item in row_group))
        output.append("")
    output.extend(table)
    return output


def parse_one(path: Path, output_dir: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        _sheet_name, sheet_path = select_sheet(zf)
        shared = read_shared_strings(zf)
        styles = read_styles(zf)
        root = ET.fromstring(zf.read(sheet_path))
        cells = parse_cells(root, shared, read_merged_ranges(root))
        images = drawing_images(zf, sheet_path, output_dir)
        cells = filter_cells_hidden_by_images(cells, images)
        titles = detect_titles(cells, styles)
        lanes = title_lanes(titles)
        titles = filter_prefixed_titles(cells, titles, lanes)
        lanes = title_lanes(titles)
        titles = add_top_preamble_titles(cells, titles, styles, lanes, images)
        lanes = title_lanes(titles)

    title_key = lambda title: title.ref
    sections = {title_key(title): {"title": title, "cells": [], "images": []} for title in titles}
    images_by_title: dict[str, list[ImageItem]] = {title_key(title): [] for title in titles}
    image_title_by_index: dict[int, Cell] = {}
    for image in images:
        title = assign_to_title(image.from_row, image.from_col, image.to_col, titles, lanes)
        if title is not None:
            images_by_title[title_key(title)].append(image)
            image_title_by_index[image.index] = title
    for image in images:
        title = image_title_by_index.get(image.index)
        if title is None:
            continue
        visual_title = visual_section_for_auxiliary_image(image, title, titles, images_by_title)
        if visual_title is None or visual_title.ref == title.ref:
            continue
        images_by_title[title_key(title)] = [
            item for item in images_by_title[title_key(title)] if item.index != image.index
        ]
        images_by_title[title_key(visual_title)].append(image)
        image_title_by_index[image.index] = visual_title
    redirected_flows: dict[str, tuple[Cell, int]] = {}
    for cell in cells:
        if cell.is_title:
            continue
        if cell.row == 1:
            continue
        title = top_preamble_title_for_cell(cell, titles)
        top_preamble_assigned = title is not None
        if title is None:
            title = assign_to_title(cell.start_row, cell.start_col, cell.end_col, titles, lanes)
        if title is None:
            title = adjacent_left_prefix_title_for_cell(cell, titles)
        if title is None:
            title = same_row_left_prefix_title_for_cell(cell, cells, titles, lanes)
        if title is not None:
            right_content_title = (
                immediate_right_content_title_for_prefix(cell, cells, titles, lanes)
                if title.row == 1 and not top_preamble_assigned
                else None
            )
            if right_content_title is not None:
                title = right_content_title
            same_row_visual_title = None
            if right_content_title is None:
                same_row_visual_title = same_row_visual_table_title_for_cell(cell, cells, titles, lanes, images_by_title)
                if same_row_visual_title is not None:
                    title = same_row_visual_title
            right_title = adjacent_right_title_for_cell(cell, title, titles)
            if right_title is not None and same_row_visual_title is None and right_content_title is None:
                title = right_title
            if right_content_title is None:
                visual_title = visual_section_for_cell(cell, title, titles, images_by_title)
                if visual_title is not None:
                    title = visual_title
            title_index = titles.index(title)
            redirect = redirected_flows.get(title_key(title))
            if redirect and cell.start_col >= redirect[1]:
                title = redirect[0]
            elif title_index > 0:
                previous_title = titles[title_index - 1]
                previous_cells = sections[title_key(previous_title)]["cells"]
                if should_continue_previous_flow(
                    cell,
                    title,
                    previous_title,
                    previous_cells,
                    images_by_title.get(title_key(title), []),
                ):
                    previous_flow_cols = [
                        item.start_col
                        for item in previous_cells
                        if item.start_col > previous_title.start_col
                    ]
                    flow_col = min(previous_flow_cols) if previous_flow_cols else cell.start_col
                    redirected_flows[title_key(title)] = (previous_title, flow_col)
                    title = previous_title
            sections[title_key(title)]["cells"].append(cell)
    for title in titles:
        sections[title_key(title)]["images"].extend(images_by_title[title_key(title)])

    paired_rule_cells, consumed_rule_titles = paired_rule_cells_by_visual_title(titles, sections, images_by_title)
    ordered_titles = sorted(titles, key=lambda title: section_output_sort_key(title, titles, lanes))

    lines = [f"# {path.name}", ""]
    for title in ordered_titles:
        if title_key(title) in consumed_rule_titles:
            continue
        section = sections[title_key(title)]
        cell_images = [image for image in section["images"] if is_cell_image(image)]
        flow_images = [image for image in section["images"] if not is_cell_image(image)]
        cell_image_suffixes: dict[int, list[str]] = {}
        for image in cell_images:
            cell_image_suffixes.setdefault(image.from_row, []).append(f"![image {image.index}](./{image.path})")
        row_suffixes = {
            row: " ".join(items)
            for row, items in cell_image_suffixes.items()
        }
        rows: dict[int, list[Cell]] = {}
        rule_cells_for_title = paired_rule_cells.get(title_key(title), [])
        section_cells = trim_cells_before_later_visual_title(title, titles, sections, images_by_title)
        aligned_rule_cells = align_paired_rule_cells_to_section(section_cells, rule_cells_for_title)
        render_cells = [*section_cells, *aligned_rule_cells]
        for cell in render_cells:
            rows.setdefault(cell.row, []).append(cell)
        lines_by_row = [(row, cell_lines_for_row(row_cells)) for row, row_cells in sorted(rows.items())]
        lines_by_row = [(row, items) for row, items in lines_by_row if items]
        lines_by_row = append_row_suffixes(lines_by_row, row_suffixes)
        mixed_body = render_mixed_list_table_section(render_cells) if not flow_images else None
        paired_body = (
            render_paired_rule_section(flow_images, section_cells, rule_cells_for_title)
            if rule_cells_for_title and flow_images
            else None
        )
        primary_table_body = (
            primary_visual_table_section(
                lines_by_row,
                flow_images,
                [
                    (row, items)
                    for row, row_cells in sorted(
                        {
                            cell.row: [item for item in cells if item.row == cell.row]
                            for cell in cells
                            if not cell.is_title
                        }.items()
                    )
                    for items in [cell_lines_for_row(row_cells)]
                    if items
                ],
            )
            if flow_images
            else None
        )
        if primary_table_body is not None:
            body = primary_table_body
        elif paired_body is not None:
            body = paired_body
        elif mixed_body is not None:
            body = mixed_body
        elif use_card_rendering(render_cells, flow_images):
            body = render_card_section(render_cells, flow_images, styles)
        else:
            body = split_rows_for_images(lines_by_row, flow_images, (title.start_col, title.end_col), raw_rows=sorted(rows.items()))
        if body:
            lines.append(f"## {title.text}")
            lines.append("")
            lines.extend(body)
            lines.append("")
        elif not any(other.ref != title.ref and other.row == title.row for other in titles):
            lines.append(f"## {title.text}")
            lines.append("")
    while lines and lines[-1] == "":
        lines.pop()
    (output_dir / "extract.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return {
        "file": path.name,
        "xlsx": str(path),
        "output": str((output_dir / "extract.md").relative_to(ROOT)),
        "warnings": [],
    }


def main() -> int:
    query = sys.argv[1] if len(sys.argv) > 1 else None
    if len(sys.argv) > 2:
        print("usage: python generate_design_md.py [number-or-file-name]", file=sys.stderr)
        return 2
    all_files = parse_md_list()
    files = filter_files(all_files, query)
    if query and not files:
        print(f"no md-list item matched: {query}", file=sys.stderr)
        return 1
    if OUTPUT_ROOT.exists():
        shutil.rmtree(OUTPUT_ROOT, ignore_errors=True)
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

    items = []
    missing = []
    warnings = []
    for file_name in files:
        source = SOURCE_ROOT / file_name
        if not source.exists():
            missing.append(file_name)
            continue
        out_dir = OUTPUT_ROOT / safe_dir_name(file_name)
        out_dir.mkdir(parents=True, exist_ok=True)
        try:
            items.append(parse_one(source, out_dir))
        except Exception as exc:
            warnings.append(f"{file_name}: {exc}")

    (OUTPUT_ROOT / "outputs.txt").write_text(
        "\n".join(item["output"] for item in items) + ("\n" if items else ""),
        encoding="utf-8",
    )
    summary_lines = [
        "# Markdown 输出摘要",
        "",
        f"- 生成范围：{'匹配 ' + query if query else '全部清单'}",
        f"- 生成数量：{len(items)}",
        f"- 缺失数量：{len(missing)}",
        f"- 警告数量：{len(warnings)}",
        "",
        "## 输出文件",
        "",
    ]
    summary_lines.extend(f"- `{item['output']}`" for item in items)
    if missing:
        summary_lines.extend(["", "## 缺失文件", ""])
        summary_lines.extend(f"- `{item}`" for item in missing)
    if warnings:
        summary_lines.extend(["", "## 警告", ""])
        summary_lines.extend(f"- {item}" for item in warnings)
    (OUTPUT_ROOT / "summary.md").write_text("\n".join(summary_lines) + "\n", encoding="utf-8")
    print(f"generated={len(items)} missing={len(missing)} warnings={len(warnings)}")
    if warnings:
        print("\n".join(warnings), file=sys.stderr)
    return 1 if missing or warnings else 0


if __name__ == "__main__":
    configure_stdio()
    raise SystemExit(main())
