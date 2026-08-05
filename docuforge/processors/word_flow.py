"""Read the positioned intermediate DOCX used by regional reconstruction.

This module contains only source-page, positioned-line, column, and visual
island analysis helpers.  The former standard-flow Word builder was removed;
the supported compatibility postprocessor is ``word_region``.
"""

from __future__ import annotations

import io
import math
import re
import statistics
from collections import Counter, deque
from dataclasses import dataclass, field
from typing import Any, Iterable

from docuforge.models import MissingEngineError


@dataclass
class _FrameLine:
    paragraph: Any
    x: int
    y: int
    width: int
    height: int
    text: str
    font_size: float
    bold: bool

    @property
    def x1(self) -> int:
        return self.x + self.width

    @property
    def y1(self) -> int:
        return self.y + self.height

    @property
    def center_y(self) -> float:
        return self.y + self.height / 2


@dataclass
class _MergedLine:
    frames: list[_FrameLine]
    x: int
    y: int
    width: int
    height: int
    text: str
    font_size: float
    bold: bool
    column: str = "main"

    @property
    def x1(self) -> int:
        return self.x + self.width

    @property
    def y1(self) -> int:
        return self.y + self.height

    @property
    def center_y(self) -> float:
        return self.y + self.height / 2


@dataclass
class _VisualRegion:
    blob: bytes
    x: int
    y: int
    width: int
    height: int
    pixel_width: int
    pixel_height: int
    column: str = "main"

    @property
    def x1(self) -> int:
        return self.x + self.width

    @property
    def y1(self) -> int:
        return self.y + self.height

    @property
    def center_y(self) -> float:
        return self.y + self.height / 2


@dataclass
class _FlowPage:
    index: int
    width: int
    height: int
    frames: list[_FrameLine] = field(default_factory=list)
    background: bytes | None = None


_WORD_PATTERN = re.compile(r"[A-Za-z0-9]+(?:[-'][A-Za-z0-9]+)*|[\u3400-\u9fff]")
_REFERENCE_START = re.compile(r"^[A-Z][A-Za-z'’\-]+(?:,|\s+[A-Z]\.)")


def _integer_attribute(element: Any, name: str, default: int = 0) -> int:
    from docx.oxml.ns import qn

    try:
        return int(element.get(qn(name), str(default)))
    except (TypeError, ValueError):
        return default


def _section_page_size(section: Any | None) -> tuple[int, int]:
    from docx.oxml.ns import qn

    if section is None:
        return 12240, 15840
    page_size = section.find(qn("w:pgSz"))
    if page_size is None:
        return 12240, 15840
    width = max(1440, _integer_attribute(page_size, "w:w", 12240))
    height = max(1440, _integer_attribute(page_size, "w:h", 15840))
    return width, height


def _dominant_font_metrics(paragraph: Any) -> tuple[float, bool]:
    from docx.oxml.ns import qn

    candidates: list[tuple[int, float, bool]] = []
    for run in paragraph.runs:
        text_length = max(1, len(run.text.strip()))
        size = float(run.font.size.pt) if run.font.size is not None else 10.5
        properties = run._r.rPr
        if properties is not None:
            size_element = properties.find(qn("w:sz"))
            if size_element is not None:
                try:
                    size = float(size_element.get(qn("w:val"), "21")) / 2
                except (TypeError, ValueError):
                    pass
        candidates.append((text_length, max(5.0, min(72.0, size)), bool(run.bold)))
    if not candidates:
        return 10.5, False
    _length, size, bold = max(candidates, key=lambda item: (item[0], item[1]))
    return size, bold


def _paragraph_frame(paragraph: Any) -> _FrameLine | None:
    from docx.oxml.ns import qn

    properties = paragraph._p.pPr
    if properties is None:
        return None
    frame = properties.find(qn("w:framePr"))
    if frame is None:
        return None
    text = paragraph.text or ""
    if not text.strip():
        return None
    size, bold = _dominant_font_metrics(paragraph)
    x = _integer_attribute(frame, "w:x")
    y = _integer_attribute(frame, "w:y")
    width = max(20, _integer_attribute(frame, "w:w", 20))
    height = max(20, _integer_attribute(frame, "w:h", round(size * 20)))
    return _FrameLine(paragraph, x, y, width, height, text, size, bold)


def _largest_anchored_image(paragraph: Any, document: Any) -> bytes | None:
    from docx.oxml.ns import qn

    candidates: list[bytes] = []
    for blip in paragraph._p.xpath(".//wp:anchor//a:blip"):
        relationship_id = blip.get(qn("r:embed"))
        if not relationship_id:
            continue
        part = document.part.related_parts.get(relationship_id)
        blob = getattr(part, "blob", None)
        if isinstance(blob, bytes) and blob:
            candidates.append(blob)
    return max(candidates, key=len) if candidates else None


def _source_pages(document: Any) -> list[_FlowPage]:
    from docx.oxml.ns import qn

    paragraphs = {id(paragraph._p): paragraph for paragraph in document.paragraphs}
    body = document.element.body
    pages: list[_FlowPage] = []
    current_frames: list[_FrameLine] = []
    current_background: bytes | None = None
    page_index = 0

    def finish(section: Any | None) -> None:
        nonlocal current_frames, current_background, page_index
        width, height = _section_page_size(section)
        pages.append(
            _FlowPage(
                page_index,
                width,
                height,
                list(current_frames),
                current_background,
            )
        )
        page_index += 1
        current_frames = []
        current_background = None

    for element in body.iterchildren():
        if element.tag != qn("w:p"):
            continue
        paragraph = paragraphs.get(id(element))
        if paragraph is not None:
            frame = _paragraph_frame(paragraph)
            if frame is not None:
                current_frames.append(frame)
            background = _largest_anchored_image(paragraph, document)
            if background is not None and (
                current_background is None or len(background) > len(current_background)
            ):
                current_background = background
        properties = element.find(qn("w:pPr"))
        section = properties.find(qn("w:sectPr")) if properties is not None else None
        if section is not None:
            finish(section)

    final_section = body.find(qn("w:sectPr"))
    if current_frames or current_background is not None or not pages:
        finish(final_section)
    return pages


def _same_visual_row(first: _FrameLine, second: _FrameLine) -> bool:
    overlap = min(first.y1, second.y1) - max(first.y, second.y)
    tolerance = max(45, round(max(first.height, second.height) * 0.42))
    return overlap >= -tolerance and abs(first.center_y - second.center_y) <= tolerance


def _text_joiner(previous: str, current: str, gap: int, size: float) -> str:
    if not previous or not current:
        return ""
    if previous.endswith((" ", "\t", "-", "‐", "‑")) or current.startswith(
        (" ", "\t", ",", ".", ";", ":", ")", "]", "}", "%")
    ):
        return ""
    if previous.endswith(("(", "[", "{", "/")):
        return ""
    threshold = max(18, round(size * 20 * 0.16))
    return " " if gap >= threshold or previous[-1:].isalnum() else ""


def _merge_frame_rows(frames: Iterable[_FrameLine], page_width: int) -> list[_MergedLine]:
    rows: list[list[_FrameLine]] = []
    for frame in sorted(frames, key=lambda item: (item.y, item.x)):
        if rows and _same_visual_row(rows[-1][0], frame):
            rows[-1].append(frame)
        else:
            rows.append([frame])

    merged: list[_MergedLine] = []
    split_gap = max(220, round(page_width * 0.025))
    for row in rows:
        clusters: list[list[_FrameLine]] = []
        for frame in sorted(row, key=lambda item: item.x):
            if clusters and frame.x - max(item.x1 for item in clusters[-1]) <= split_gap:
                clusters[-1].append(frame)
            else:
                clusters.append([frame])
        for cluster in clusters:
            cluster.sort(key=lambda item: item.x)
            text = ""
            previous: _FrameLine | None = None
            for frame in cluster:
                if previous is not None:
                    text += _text_joiner(
                        text,
                        frame.text,
                        frame.x - previous.x1,
                        max(previous.font_size, frame.font_size),
                    )
                text += frame.text
                previous = frame
            x0 = min(item.x for item in cluster)
            y0 = min(item.y for item in cluster)
            x1 = max(item.x1 for item in cluster)
            y1 = max(item.y1 for item in cluster)
            dominant = max(
                cluster,
                key=lambda item: (len(item.text.strip()), item.font_size),
            )
            merged.append(
                _MergedLine(
                    cluster,
                    x0,
                    y0,
                    x1 - x0,
                    y1 - y0,
                    text,
                    dominant.font_size,
                    dominant.bold,
                )
            )
    return merged


def _background_color(image: Any) -> tuple[int, int, int]:
    from PIL import ImageStat

    width, height = image.size
    sample = max(2, min(width, height) // 50)
    swatches = [
        image.crop((0, 0, sample, sample)),
        image.crop((width - sample, 0, width, sample)),
        image.crop((0, height - sample, sample, height)),
        image.crop((width - sample, height - sample, width, height)),
    ]
    medians = [ImageStat.Stat(swatch).median for swatch in swatches]
    return tuple(int(statistics.median(item[channel] for item in medians)) for channel in range(3))


def _connected_boxes(mask: Any) -> list[tuple[int, int, int, int, int]]:
    pixels = mask.load()
    width, height = mask.size
    seen = bytearray(width * height)
    boxes: list[tuple[int, int, int, int, int]] = []
    for y in range(height):
        for x in range(width):
            index = y * width + x
            if seen[index] or pixels[x, y] == 0:
                continue
            queue: deque[tuple[int, int]] = deque([(x, y)])
            seen[index] = 1
            x0 = x1 = x
            y0 = y1 = y
            count = 0
            while queue:
                current_x, current_y = queue.popleft()
                count += 1
                x0 = min(x0, current_x)
                x1 = max(x1, current_x)
                y0 = min(y0, current_y)
                y1 = max(y1, current_y)
                for next_y in range(max(0, current_y - 1), min(height, current_y + 2)):
                    offset = next_y * width
                    for next_x in range(max(0, current_x - 1), min(width, current_x + 2)):
                        next_index = offset + next_x
                        if not seen[next_index] and pixels[next_x, next_y] != 0:
                            seen[next_index] = 1
                            queue.append((next_x, next_y))
            boxes.append((x0, y0, x1 + 1, y1 + 1, count))
    return boxes


def _box_gap(
    first: tuple[int, int, int, int],
    second: tuple[int, int, int, int],
) -> tuple[int, int]:
    horizontal = max(0, max(first[0], second[0]) - min(first[2], second[2]))
    vertical = max(0, max(first[1], second[1]) - min(first[3], second[3]))
    return horizontal, vertical


def _merge_visual_boxes(
    boxes: list[tuple[int, int, int, int]],
    width: int,
    height: int,
) -> list[tuple[int, int, int, int]]:
    merged = list(boxes)
    changed = True
    while changed:
        changed = False
        output: list[tuple[int, int, int, int]] = []
        while merged:
            current = merged.pop(0)
            index = 0
            while index < len(merged):
                candidate = merged[index]
                horizontal_gap, vertical_gap = _box_gap(current, candidate)
                vertical_overlap = min(current[3], candidate[3]) - max(current[1], candidate[1])
                horizontal_overlap = min(current[2], candidate[2]) - max(current[0], candidate[0])
                same_formula_row = bool(
                    vertical_overlap > -max(5, round(height * 0.006))
                    and horizontal_gap <= max(12, round(width * 0.035))
                )
                same_visual_block = bool(
                    horizontal_overlap > 0
                    and vertical_gap <= max(8, round(height * 0.012))
                    and min(current[2] - current[0], candidate[2] - candidate[0])
                    >= width * 0.08
                )
                formula_cluster = bool(
                    horizontal_gap <= max(16, round(width * 0.055))
                    and vertical_gap <= max(12, round(height * 0.025))
                    and current[2] - current[0] < width * 0.46
                    and candidate[2] - candidate[0] < width * 0.46
                    and current[3] - current[1] < height * 0.14
                    and candidate[3] - candidate[1] < height * 0.14
                )
                if same_formula_row or same_visual_block or formula_cluster:
                    current = (
                        min(current[0], candidate[0]),
                        min(current[1], candidate[1]),
                        max(current[2], candidate[2]),
                        max(current[3], candidate[3]),
                    )
                    merged.pop(index)
                    changed = True
                    index = 0
                    continue
                index += 1
            output.append(current)
        merged = output
    return merged


def _visual_regions(background: bytes, page_width: int, page_height: int) -> list[_VisualRegion]:
    try:
        from PIL import Image, ImageChops, ImageFilter
    except ImportError as exc:
        raise MissingEngineError("区域级 Word 视觉区域分析需要 Pillow") from exc

    try:
        image = Image.open(io.BytesIO(background)).convert("RGB")
    except Exception:
        return []
    scale = max(1, math.ceil(max(image.size) / 950))
    small = image.resize(
        (max(1, image.width // scale), max(1, image.height // scale)),
        Image.Resampling.BILINEAR,
    )
    background_color = _background_color(small)
    difference = ImageChops.difference(
        small,
        Image.new("RGB", small.size, background_color),
    ).convert("L")
    mask = difference.point(lambda value: 255 if value >= 9 else 0)
    mask = mask.filter(ImageFilter.MaxFilter(9))
    raw_boxes = [
        (x0, y0, x1, y1)
        for x0, y0, x1, y1, count in _connected_boxes(mask)
        if count >= 10 and x1 - x0 >= 3 and y1 - y0 >= 3
    ]
    boxes = _merge_visual_boxes(raw_boxes, small.width, small.height)

    regions: list[_VisualRegion] = []
    for x0, y0, x1, y1 in boxes:
        box_width = x1 - x0
        box_height = y1 - y0
        area_ratio = (box_width * box_height) / max(1, small.width * small.height)
        if (
            area_ratio < 0.00035
            and box_width < small.width * 0.08
            and box_height < small.height * 0.035
        ):
            continue
        if box_height < max(3, small.height * 0.006) and box_width < small.width * 0.4:
            continue
        if box_width / max(1, box_height) >= 8 and box_height < small.height * 0.02:
            continue
        margin = max(3, round(5 / scale))
        x0 = max(0, x0 - margin)
        y0 = max(0, y0 - margin)
        x1 = min(small.width, x1 + margin)
        y1 = min(small.height, y1 + margin)
        original_box = (
            max(0, x0 * scale),
            max(0, y0 * scale),
            min(image.width, x1 * scale),
            min(image.height, y1 * scale),
        )
        crop = image.crop(original_box)
        payload = io.BytesIO()
        crop.save(payload, format="PNG", optimize=True, compress_level=7)
        x_twips = round(original_box[0] * page_width / image.width)
        y_twips = round(original_box[1] * page_height / image.height)
        width_twips = max(20, round((original_box[2] - original_box[0]) * page_width / image.width))
        height_twips = max(20, round((original_box[3] - original_box[1]) * page_height / image.height))
        regions.append(
            _VisualRegion(
                payload.getvalue(),
                x_twips,
                y_twips,
                width_twips,
                height_twips,
                crop.width,
                crop.height,
            )
        )
    return sorted(regions, key=lambda item: (item.y, item.x))


def _two_column_layout(lines: list[_MergedLine], page_width: int) -> bool:
    midpoint = page_width / 2
    gutter = page_width * 0.04
    left = sum(1 for line in lines if line.x1 <= midpoint + gutter and line.x < midpoint)
    right = sum(1 for line in lines if line.x >= midpoint - gutter and line.x1 > midpoint)
    minimum = max(4, round(len(lines) * 0.14))
    return left >= minimum and right >= minimum


def _assign_column(item: Any, page_width: int, two_columns: bool) -> str:
    if not two_columns:
        return "main"
    midpoint = page_width / 2
    width = item.width
    crosses_midpoint = item.x < midpoint < item.x1
    if width >= page_width * 0.62 or (
        crosses_midpoint
        and item.x < midpoint - page_width * 0.12
        and item.x1 > midpoint + page_width * 0.12
    ):
        return "full"
    if item.x1 <= midpoint + page_width * 0.04:
        return "left"
    if item.x >= midpoint - page_width * 0.04:
        return "right"
    return "full"


def _heading_line(line: _MergedLine, median_size: float) -> bool:
    text = line.text.strip()
    return bool(
        text
        and len(text) <= 140
        and (
            line.font_size >= median_size * 1.18
            or (line.bold and len(text) <= 90)
        )
    )


def _can_merge_lines(
    previous: _MergedLine,
    current: _MergedLine,
    *,
    median_size: float,
) -> bool:
    if previous.column != current.column:
        return False
    if _heading_line(previous, median_size) or _heading_line(current, median_size):
        return False
    gap = current.y - previous.y1
    if gap < -max(previous.height, current.height) * 0.4:
        return False
    if gap > max(360, round(max(previous.height, current.height) * 1.35)):
        return False
    if abs(previous.font_size - current.font_size) > max(1.2, median_size * 0.16):
        return False
    indent_change = current.x - previous.x
    if (
        previous.text.rstrip().endswith((".", "!", "?", ":"))
        and indent_change > max(120, round(current.font_size * 12))
    ):
        return False
    if previous.text.rstrip().endswith(".") and _REFERENCE_START.match(current.text.strip()):
        return False
    return True


def _remove_last_hyphen(paragraph: Any) -> None:
    for run in reversed(paragraph.runs):
        text = run.text
        if not text:
            continue
        if text.endswith(("-", "‐", "‑")):
            run.text = text[:-1]
        return


def _line_break_hyphen_is_soft(
    previous: _MergedLine,
    current: _MergedLine,
    word_lexicon: set[str],
) -> bool:
    previous_match = re.search(r"([A-Za-z]{2,})[-‐‑]$", previous.text.rstrip())
    current_match = re.match(r"\s*([A-Za-z]{2,})", current.text)
    if previous_match is None or current_match is None:
        return False
    combined = (previous_match.group(1) + current_match.group(1)).casefold()
    return combined in word_lexicon


def _word_counter(text: str) -> Counter[str]:
    normalized = re.sub(
        r"(?<=[A-Za-z])[-‐‑]\s*(?=[A-Za-z])",
        "",
        text,
    )
    return Counter(token.casefold() for token in _WORD_PATTERN.findall(normalized))


__all__: list[str] = []
