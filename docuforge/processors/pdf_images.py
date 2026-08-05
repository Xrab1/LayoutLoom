"""High-fidelity extraction of bitmap images embedded in PDF documents.

PDFs may store a visible figure as one native image, as an inline image, as a
base image plus a soft mask, or as many neighbouring tiles.  The extraction
modes in this module make those trade-offs explicit instead of pretending that
one strategy is correct for every PDF.
"""

from __future__ import annotations

import io
import json
import math
import os
import warnings
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from typing import Any, Iterable

from ..models import MissingEngineError, ValidationError
from ..utils import unique_path
from .pdf import (
    PDFPasswordError,
    PDFProcessingWarning,
    PDFProcessorError,
    PageSpec,
    _atomic_output,
    _input_file,
    _output_directory,
    _parse_page_spec,
)

PathLike = str | os.PathLike[str]

_MODES = {"original", "visible", "smart", "both", "all"}
_FORMATS = {"auto", "png", "jpg"}
_MAX_RENDER_PIXELS = 100_000_000


@dataclass(frozen=True)
class _Placement:
    page_index: int
    occurrence: int
    bbox: tuple[float, float, float, float]
    transform: tuple[float, ...]
    width: int
    height: int
    xref: int
    smask: int
    digest: str
    colorspace: str
    bpc: int
    has_mask: bool


def _require_int(value: Any, name: str, *, minimum: int, maximum: int) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValidationError(f"{name} 必须是整数")
    if value < minimum or value > maximum:
        raise ValidationError(f"{name} 必须在 {minimum}–{maximum} 之间")
    return value


def _require_number(value: Any, name: str, *, minimum: float, maximum: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValidationError(f"{name} 必须是数字")
    result = float(value)
    if not math.isfinite(result) or result < minimum or result > maximum:
        raise ValidationError(f"{name} 必须在 {minimum:g}–{maximum:g} 之间")
    return result


def _normalize_extension(value: Any) -> str:
    extension = str(value or "png").lower().lstrip(".")
    aliases = {
        "jpeg": "jpg",
        "jpe": "jpg",
        "jpx": "jp2",
        "jp2k": "jp2",
        "tif": "tiff",
    }
    extension = aliases.get(extension, extension)
    if not extension.isalnum() or len(extension) > 8:
        return "bin"
    return extension


def _bytes_digest(payload: bytes) -> str:
    return sha256(payload).hexdigest()


def _bbox_tuple(value: Any) -> tuple[float, float, float, float]:
    try:
        left, top, right, bottom = (float(item) for item in value)
    except (TypeError, ValueError) as exc:
        raise PDFProcessorError("PDF 图片边界信息无效") from exc
    if not all(math.isfinite(item) for item in (left, top, right, bottom)):
        raise PDFProcessorError("PDF 图片边界信息无效")
    return (
        min(left, right),
        min(top, bottom),
        max(left, right),
        max(top, bottom),
    )


def _rect_dimensions(bbox: tuple[float, float, float, float]) -> tuple[float, float]:
    return max(0.0, bbox[2] - bbox[0]), max(0.0, bbox[3] - bbox[1])


def _rect_union(
    boxes: Iterable[tuple[float, float, float, float]],
) -> tuple[float, float, float, float]:
    items = list(boxes)
    return (
        min(item[0] for item in items),
        min(item[1] for item in items),
        max(item[2] for item in items),
        max(item[3] for item in items),
    )


def _overlap_ratio(start_a: float, end_a: float, start_b: float, end_b: float) -> float:
    overlap = max(0.0, min(end_a, end_b) - max(start_a, start_b))
    denominator = max(1e-6, min(end_a - start_a, end_b - start_b))
    return overlap / denominator


def _placements_are_neighbours(left: _Placement, right: _Placement, gap: float) -> bool:
    a = left.bbox
    b = right.bbox
    horizontal_gap = max(0.0, max(a[0], b[0]) - min(a[2], b[2]))
    vertical_gap = max(0.0, max(a[1], b[1]) - min(a[3], b[3]))
    if horizontal_gap == 0 and vertical_gap == 0:
        return True
    if horizontal_gap <= gap and _overlap_ratio(a[1], a[3], b[1], b[3]) >= 0.60:
        return True
    return vertical_gap <= gap and _overlap_ratio(a[0], a[2], b[0], b[2]) >= 0.60


def _cluster_placements(
    placements: list[_Placement], gap: float
) -> list[list[_Placement]]:
    if not placements:
        return []
    parents = list(range(len(placements)))

    def find(index: int) -> int:
        while parents[index] != index:
            parents[index] = parents[parents[index]]
            index = parents[index]
        return index

    def union(first: int, second: int) -> None:
        first_root = find(first)
        second_root = find(second)
        if first_root != second_root:
            parents[second_root] = first_root

    for index, placement in enumerate(placements):
        for earlier in range(index):
            if _placements_are_neighbours(placement, placements[earlier], gap):
                union(index, earlier)

    grouped: dict[int, list[_Placement]] = {}
    for index, placement in enumerate(placements):
        grouped.setdefault(find(index), []).append(placement)
    return sorted(
        grouped.values(),
        key=lambda group: min(item.occurrence for item in group),
    )


def _page_placements(page: Any, page_index: int) -> list[_Placement]:
    smasks: dict[int, int] = {}
    try:
        for item in page.get_images(full=True):
            xref = int(item[0] or 0)
            if xref > 0:
                smasks[xref] = int(item[1] or 0)
    except Exception:
        smasks = {}

    result: list[_Placement] = []
    for occurrence, info in enumerate(
        page.get_image_info(hashes=True, xrefs=True) or (), 1
    ):
        bbox = _bbox_tuple(info.get("bbox"))
        width, height = _rect_dimensions(bbox)
        if width <= 0 or height <= 0:
            continue
        digest = info.get("digest")
        digest_text = digest.hex() if isinstance(digest, bytes) else str(digest or "")
        xref = int(info.get("xref") or 0)
        transform = info.get("transform") or ()
        result.append(
            _Placement(
                page_index,
                occurrence,
                bbox,
                tuple(float(value) for value in transform),
                max(1, int(info.get("width") or 1)),
                max(1, int(info.get("height") or 1)),
                xref,
                smasks.get(xref, 0),
                digest_text,
                str(info.get("cs-name") or info.get("colorspace") or ""),
                int(info.get("bpc") or 0),
                bool(info.get("has-mask")) or smasks.get(xref, 0) > 0,
            )
        )
    return result


def _encode_pillow_image(image: Any, image_format: str, jpeg_quality: int) -> bytes:
    from PIL import Image

    output = io.BytesIO()
    if image_format == "jpg":
        if image.mode in {"RGBA", "LA"} or "transparency" in image.info:
            rgba = image.convert("RGBA")
            background = Image.new("RGB", rgba.size, "white")
            background.paste(rgba, mask=rgba.getchannel("A"))
            image = background
        elif image.mode != "RGB":
            image = image.convert("RGB")
        image.save(output, format="JPEG", quality=jpeg_quality, optimize=False)
    else:
        # Pillow cannot write some perfectly valid PDF source colour modes
        # (most notably CMYK JPEG) directly as PNG.  Preserve alpha whenever
        # present and convert only modes unsupported by the PNG encoder.
        png_modes = {"1", "L", "LA", "I", "P", "RGB", "RGBA"}
        if image.mode not in png_modes and not image.mode.startswith("I;16"):
            target_mode = (
                "RGBA"
                if "A" in image.getbands() or "transparency" in image.info
                else "RGB"
            )
            image = image.convert(target_mode)
        image.save(output, format="PNG", optimize=False, compress_level=6)
    return output.getvalue()


def _convert_image_payload(
    payload: bytes, image_format: str, jpeg_quality: int
) -> tuple[bytes, str, int, int]:
    try:
        from PIL import Image
    except ImportError as exc:  # pragma: no cover - mandatory packaged dependency
        raise MissingEngineError("PDF 图片提取需要 Pillow") from exc
    try:
        with Image.open(io.BytesIO(payload)) as image:
            image.load()
            width, height = image.size
            converted = _encode_pillow_image(image, image_format, jpeg_quality)
    except Exception as exc:
        raise PDFProcessorError(f"图片数据无法解码：{exc}") from exc
    return converted, image_format, width, height


def _compose_image_and_mask(
    image_payload: bytes,
    mask_payload: bytes,
    *,
    image_format: str,
    jpeg_quality: int,
) -> tuple[bytes, str, int, int]:
    try:
        from PIL import Image
    except ImportError as exc:  # pragma: no cover - mandatory packaged dependency
        raise MissingEngineError("PDF 图片提取需要 Pillow") from exc
    try:
        with Image.open(io.BytesIO(image_payload)) as source, Image.open(
            io.BytesIO(mask_payload)
        ) as mask:
            source.load()
            mask.load()
            if source.size != mask.size:
                raise PDFProcessorError(
                    f"透明蒙版尺寸 {mask.size} 与原图尺寸 {source.size} 不一致"
                )
            rgba = source.convert("RGBA")
            rgba.putalpha(mask.convert("L"))
            target_format = "png" if image_format == "auto" else image_format
            payload = _encode_pillow_image(rgba, target_format, jpeg_quality)
            return payload, target_format, rgba.width, rgba.height
    except PDFProcessorError:
        raise
    except Exception as exc:
        raise PDFProcessorError(f"透明蒙版合成失败：{exc}") from exc


def _pixmap_payload(
    pymupdf: Any,
    pixmap: Any,
    image_format: str,
    jpeg_quality: int,
) -> tuple[bytes, str, int, int]:
    target_format = "png" if image_format == "auto" else image_format
    if target_format == "jpg" and bool(getattr(pixmap, "alpha", False)):
        png = pixmap.tobytes("png")
        return _convert_image_payload(png, "jpg", jpeg_quality)
    if target_format == "jpg" and int(getattr(pixmap, "n", 0)) not in {1, 3}:
        pixmap = pymupdf.Pixmap(pymupdf.csRGB, pixmap)
    payload = pixmap.tobytes(target_format, jpg_quality=jpeg_quality)
    return payload, target_format, int(pixmap.width), int(pixmap.height)


def _xref_payloads(
    document: Any,
    pymupdf: Any,
    placement: _Placement,
    image_format: str,
    jpeg_quality: int,
) -> list[tuple[bytes, str, int, int, str]]:
    extracted = document.extract_image(placement.xref)
    base_payload = extracted.get("image")
    if not isinstance(base_payload, bytes) or not base_payload:
        raise PDFProcessorError(f"xref {placement.xref} 未返回有效图片数据")
    base_extension = _normalize_extension(extracted.get("ext"))
    width = int(extracted.get("width") or placement.width)
    height = int(extracted.get("height") or placement.height)

    if placement.smask > 0 or placement.has_mask:
        try:
            pixmap = pymupdf.Pixmap(document, placement.xref)
            if bool(getattr(pixmap, "alpha", False)):
                payload, extension, pix_width, pix_height = _pixmap_payload(
                    pymupdf, pixmap, image_format, jpeg_quality
                )
                return [(payload, extension, pix_width, pix_height, "composed")]
        except Exception:
            pass
        if placement.smask > 0:
            mask: dict[str, Any] = {}
            mask_payload: bytes | None = None
            try:
                mask = document.extract_image(placement.smask)
                mask_payload = mask.get("image")
                if not isinstance(mask_payload, bytes) or not mask_payload:
                    raise PDFProcessorError(
                        f"xref {placement.smask} 未返回有效蒙版数据"
                    )
                composed = _compose_image_and_mask(
                    base_payload,
                    mask_payload,
                    image_format=image_format,
                    jpeg_quality=jpeg_quality,
                )
                return [(*composed, "composed")]
            except Exception as exc:
                fallback_description = (
                    "已分别导出原图和蒙版"
                    if isinstance(mask_payload, bytes) and mask_payload
                    else "已保留可读取的原图"
                )
                warnings.warn(
                    f"第 {placement.page_index + 1} 页 xref {placement.xref} 的透明蒙版"
                    f"无法可靠合成，{fallback_description}（{exc}）",
                    PDFProcessingWarning,
                    stacklevel=3,
                )
                fallback = [
                    (base_payload, base_extension, width, height, "base"),
                ]
                if isinstance(mask_payload, bytes) and mask_payload:
                    fallback.append(
                        (
                            mask_payload,
                            _normalize_extension(mask.get("ext")),
                            int(mask.get("width") or width),
                            int(mask.get("height") or height),
                            "mask",
                        )
                    )
                return fallback

    if image_format == "auto":
        return [(base_payload, base_extension, width, height, "image")]
    converted = _convert_image_payload(base_payload, image_format, jpeg_quality)
    return [(*converted, "image")]


def _inline_blocks(page: Any) -> list[dict[str, Any]]:
    try:
        import pymupdf

        page_dict = page.get_text("dict", flags=pymupdf.TEXT_PRESERVE_IMAGES)
    except Exception:
        return []
    return [
        block
        for block in page_dict.get("blocks", ())
        if int(block.get("type", 0)) == 1 and isinstance(block.get("image"), bytes)
    ]


def _bbox_distance(
    left: tuple[float, float, float, float],
    right: tuple[float, float, float, float],
) -> float:
    return max(abs(first - second) for first, second in zip(left, right, strict=True))


def _find_inline_block(
    placement: _Placement,
    blocks: list[dict[str, Any]],
    used_indexes: set[int],
) -> tuple[int, dict[str, Any]] | None:
    """Match an xref-zero placement to its decoded text-dictionary block.

    BBox-only matching is ambiguous when multiple inline images are painted at
    the same coordinates.  Intrinsic dimensions provide a stable discriminator
    and, importantly, keep size filtering from selecting a skipped neighbour.
    """

    candidates: list[tuple[tuple[int, int, int], int, dict[str, Any]]] = []
    for index, block in enumerate(blocks):
        if index in used_indexes:
            continue
        try:
            if _bbox_distance(placement.bbox, _bbox_tuple(block.get("bbox"))) > 0.05:
                continue
            block_width = max(1, int(block.get("width") or 1))
            block_height = max(1, int(block.get("height") or 1))
        except (PDFProcessorError, TypeError, ValueError, OverflowError):
            continue
        dimension_delta = abs(block_width - placement.width) + abs(
            block_height - placement.height
        )
        candidates.append(
            (
                (
                    (
                        0
                        if (
                            block_width == placement.width
                            and block_height == placement.height
                        )
                        else 1
                    ),
                    dimension_delta,
                    index,
                ),
                index,
                block,
            )
        )
    if not candidates:
        return None
    _, index, block = min(candidates, key=lambda item: item[0])
    return index, block


def _inline_payloads(
    block: dict[str, Any], image_format: str, jpeg_quality: int
) -> list[tuple[bytes, str, int, int, str]]:
    payload = block["image"]
    extension = _normalize_extension(block.get("ext"))
    width = max(1, int(block.get("width") or 1))
    height = max(1, int(block.get("height") or 1))
    mask = block.get("mask")
    if isinstance(mask, bytes) and mask:
        try:
            composed = _compose_image_and_mask(
                payload,
                mask,
                image_format=image_format,
                jpeg_quality=jpeg_quality,
            )
            return [(*composed, "composed")]
        except Exception as exc:
            warnings.warn(
                f"内联图片透明蒙版无法可靠合成，已分别导出（{exc}）",
                PDFProcessingWarning,
                stacklevel=3,
            )
            return [
                (payload, extension, width, height, "base"),
                (mask, "png", width, height, "mask"),
            ]
    if image_format == "auto":
        return [(payload, extension, width, height, "image")]
    converted = _convert_image_payload(payload, image_format, jpeg_quality)
    return [(*converted, "image")]


def _target_path(
    directory: Path,
    *,
    page: int,
    occurrence: int,
    kind: str,
    extension: str,
    xref: int = 0,
    width: int = 0,
    height: int = 0,
    role: str = "image",
    overwrite: bool,
) -> Path:
    xref_part = f"_xref{xref}" if xref > 0 else ("_inline" if kind == "img" else "")
    role_part = "" if role in {"image", "composed"} else f"_{role}"
    dimensions = f"_{width}x{height}" if width > 0 and height > 0 else ""
    filename = (
        f"p{page:04d}_{kind}{occurrence:03d}{xref_part}{dimensions}{role_part}."
        f"{_normalize_extension(extension)}"
    )
    return unique_path(directory / filename, overwrite)


def _save_payload(
    payload: bytes,
    target: Path,
    *,
    overwrite: bool,
) -> Path:
    with _atomic_output(target, overwrite) as temporary:
        temporary.write_bytes(payload)
    return target


def _render_region_payload(
    page: Any,
    pymupdf: Any,
    bbox: tuple[float, float, float, float],
    *,
    dpi: int,
    image_format: str,
    jpeg_quality: int,
    padding: float,
    include_annotations: bool,
) -> tuple[bytes, str, int, int, tuple[float, float, float, float]]:
    page_rect = page.rect
    clip = (
        pymupdf.Rect(
            bbox[0] - padding,
            bbox[1] - padding,
            bbox[2] + padding,
            bbox[3] + padding,
        )
        & page_rect
    )
    if clip.is_empty or clip.width <= 0 or clip.height <= 0:
        raise PDFProcessorError("图片可见区域为空或位于页面外")
    expected_width = max(1, math.ceil(clip.width * dpi / 72.0))
    expected_height = max(1, math.ceil(clip.height * dpi / 72.0))
    if expected_width * expected_height > _MAX_RENDER_PIXELS:
        raise ValidationError(
            f"单个图片区域在 {dpi} DPI 下约为 {expected_width}×{expected_height} 像素，"
            "内存需求过高；请降低 DPI"
        )
    target_format = "png" if image_format == "auto" else image_format
    pixmap = page.get_pixmap(
        dpi=dpi,
        clip=clip,
        alpha=False,
        annots=include_annotations,
    )
    payload, extension, width, height = _pixmap_payload(
        pymupdf, pixmap, target_format, jpeg_quality
    )
    return payload, extension, width, height, tuple(float(value) for value in clip)


def extract_pdf_images(
    source: PathLike,
    output_dir: PathLike,
    *,
    mode: str = "original",
    pages: PageSpec | None = None,
    image_format: str = "auto",
    dpi: int = 300,
    jpeg_quality: int = 95,
    min_width: int = 1,
    min_height: int = 1,
    merge_gap: float = 4.0,
    region_padding: float = 2.0,
    deduplicate: bool = True,
    include_annotations: bool = False,
    write_manifest: bool = True,
    password: str | None = None,
    overwrite: bool = False,
) -> list[Path]:
    """Extract PDF bitmaps using native, visible, or tiled-figure strategies.

    ``original`` avoids resampling and keeps native encodings whenever possible.
    ``visible`` renders every placed occurrence exactly as it appears on the
    page.  ``smart`` first groups neighbouring image tiles and then renders each
    group.  ``both`` combines original and smart output; ``all`` emits all three.
    """

    from docuforge.runner import check_cancelled

    try:
        import pymupdf
        from PIL import Image as _PillowProbe  # noqa: F401
    except ImportError as exc:
        raise MissingEngineError("PDF 图片提取需要 PyMuPDF 和 Pillow") from exc

    source_path = _input_file(source, "PDF 文件")
    if source_path.suffix.lower() != ".pdf":
        raise ValidationError(f"不是 PDF 文件：{source_path.name}")
    mode_name = str(mode).strip().lower()
    if mode_name not in _MODES:
        raise ValidationError("mode 必须是 original、visible、smart、both 或 all")
    format_name = str(image_format).strip().lower().lstrip(".")
    if format_name == "jpeg":
        format_name = "jpg"
    if format_name not in _FORMATS:
        raise ValidationError("image_format 必须是 auto、png 或 jpg")
    dpi_value = _require_int(dpi, "DPI", minimum=72, maximum=1200)
    quality_value = _require_int(jpeg_quality, "JPG 质量", minimum=30, maximum=100)
    minimum_width = _require_int(min_width, "最小图片宽度", minimum=1, maximum=100_000)
    minimum_height = _require_int(
        min_height, "最小图片高度", minimum=1, maximum=100_000
    )
    merge_gap_value = _require_number(
        merge_gap, "碎片合并距离", minimum=0.0, maximum=72.0
    )
    padding_value = _require_number(
        region_padding, "区域留白", minimum=0.0, maximum=72.0
    )
    if not isinstance(deduplicate, bool):
        raise ValidationError("deduplicate 必须是布尔值")
    if not isinstance(include_annotations, bool):
        raise ValidationError("include_annotations 必须是布尔值")
    if not isinstance(write_manifest, bool):
        raise ValidationError("write_manifest 必须是布尔值")
    if not isinstance(overwrite, bool):
        raise ValidationError("overwrite 必须是布尔值")
    if password is not None and not isinstance(password, str):
        raise ValidationError("password 必须是字符串或 None")

    root = _output_directory(output_dir)
    outputs: list[Path] = []
    records: list[dict[str, Any]] = []
    extraction_warnings: list[str] = []
    seen_payloads: dict[str, str] = {}

    try:
        document = pymupdf.open(str(source_path))
    except Exception as exc:
        raise PDFProcessorError(f"无法打开 PDF：{source_path.name}（{exc}）") from exc
    try:
        if document.needs_pass:
            if not password or document.authenticate(password) <= 0:
                raise PDFPasswordError(f"PDF 密码错误或未提供：{source_path.name}")
        page_indexes = _parse_page_spec(pages, document.page_count)
        selected_modes = {
            "original": ("original",),
            "visible": ("visible",),
            "smart": ("smart",),
            "both": ("original", "smart"),
            "all": ("original", "visible", "smart"),
        }[mode_name]

        for page_index in page_indexes:
            check_cancelled("任务已取消；已提取的图片会保留")
            page = document[page_index]
            try:
                placements = _page_placements(page, page_index)
            except Exception as exc:
                message = f"第 {page_index + 1} 页图片目录无法读取：{exc}"
                extraction_warnings.append(message)
                warnings.warn(message, PDFProcessingWarning, stacklevel=2)
                continue

            if "original" in selected_modes:
                eligible_placements = [
                    placement
                    for placement in placements
                    if placement.width >= minimum_width
                    and placement.height >= minimum_height
                ]
                inline_blocks = (
                    _inline_blocks(page)
                    if any(placement.xref <= 0 for placement in eligible_placements)
                    else []
                )
                used_inline_blocks: set[int] = set()
                for placement in eligible_placements:
                    check_cancelled("任务已取消；已提取的图片会保留")
                    try:
                        if placement.xref > 0:
                            payloads = _xref_payloads(
                                document,
                                pymupdf,
                                placement,
                                format_name,
                                quality_value,
                            )
                        else:
                            match = _find_inline_block(
                                placement, inline_blocks, used_inline_blocks
                            )
                            if match is None:
                                raise PDFProcessorError("无法定位内联图片数据块")
                            block_index, block = match
                            used_inline_blocks.add(block_index)
                            payloads = _inline_payloads(
                                block, format_name, quality_value
                            )
                    except Exception as exc:
                        message = (
                            f"第 {page_index + 1} 页第 {placement.occurrence} 张原始图片"
                            f"提取失败：{exc}"
                        )
                        extraction_warnings.append(message)
                        warnings.warn(message, PDFProcessingWarning, stacklevel=2)
                        continue

                    original_dir = root / "原始资源"
                    for payload, extension, width, height, role in payloads:
                        digest = (
                            _bytes_digest(payload)
                            if deduplicate or write_manifest
                            else ""
                        )
                        duplicate_of = (
                            seen_payloads.get(digest) if deduplicate else None
                        )
                        relative_path: str | None = duplicate_of
                        if duplicate_of is None:
                            target = _target_path(
                                original_dir,
                                page=page_index + 1,
                                occurrence=placement.occurrence,
                                kind="img",
                                extension=extension,
                                xref=placement.xref,
                                width=width,
                                height=height,
                                role=role,
                                overwrite=overwrite,
                            )
                            _save_payload(payload, target, overwrite=overwrite)
                            outputs.append(target)
                            relative_path = target.relative_to(root).as_posix()
                            if deduplicate:
                                seen_payloads[digest] = relative_path
                        if write_manifest:
                            records.append(
                                {
                                    "mode": "original",
                                    "page": page_index + 1,
                                    "occurrence": placement.occurrence,
                                    "xref": placement.xref,
                                    "smask": placement.smask,
                                    "role": role,
                                    "native_size": [width, height],
                                    "bbox": list(placement.bbox),
                                    "transform": list(placement.transform),
                                    "colorspace": placement.colorspace,
                                    "bpc": placement.bpc,
                                    "sha256": digest,
                                    "output": relative_path,
                                    "duplicate_of": duplicate_of,
                                }
                            )

            for render_mode in ("visible", "smart"):
                if render_mode not in selected_modes:
                    continue
                groups = (
                    [[placement] for placement in placements]
                    if render_mode == "visible"
                    else _cluster_placements(placements, merge_gap_value)
                )
                render_dir = root / (
                    "可见位置" if render_mode == "visible" else "智能合并"
                )
                for group_index, group in enumerate(groups, 1):
                    check_cancelled("任务已取消；已提取的图片会保留")
                    bbox = _rect_union(item.bbox for item in group)
                    display_width, display_height = _rect_dimensions(bbox)
                    estimated_width = math.ceil(
                        (display_width + 2 * padding_value) * dpi_value / 72.0
                    )
                    estimated_height = math.ceil(
                        (display_height + 2 * padding_value) * dpi_value / 72.0
                    )
                    if (
                        estimated_width < minimum_width
                        or estimated_height < minimum_height
                    ):
                        continue
                    try:
                        payload, extension, width, height, rendered_bbox = (
                            _render_region_payload(
                                page,
                                pymupdf,
                                bbox,
                                dpi=dpi_value,
                                image_format=format_name,
                                jpeg_quality=quality_value,
                                padding=padding_value,
                                include_annotations=include_annotations,
                            )
                        )
                    except Exception as exc:
                        message = (
                            f"第 {page_index + 1} 页第 {group_index} 个{render_mode}区域"
                            f"渲染失败：{exc}"
                        )
                        extraction_warnings.append(message)
                        warnings.warn(message, PDFProcessingWarning, stacklevel=2)
                        continue
                    if width < minimum_width or height < minimum_height:
                        continue
                    digest = (
                        _bytes_digest(payload) if deduplicate or write_manifest else ""
                    )
                    duplicate_of = seen_payloads.get(digest) if deduplicate else None
                    relative_path = duplicate_of
                    if duplicate_of is None:
                        target = _target_path(
                            render_dir,
                            page=page_index + 1,
                            occurrence=group_index,
                            kind="region" if render_mode == "visible" else "group",
                            extension=extension,
                            xref=group[0].xref if len(group) == 1 else 0,
                            width=width,
                            height=height,
                            overwrite=overwrite,
                        )
                        _save_payload(payload, target, overwrite=overwrite)
                        outputs.append(target)
                        relative_path = target.relative_to(root).as_posix()
                        if deduplicate:
                            seen_payloads[digest] = relative_path
                    if write_manifest:
                        records.append(
                            {
                                "mode": render_mode,
                                "page": page_index + 1,
                                "region": group_index,
                                "occurrences": [item.occurrence for item in group],
                                "xrefs": sorted({item.xref for item in group}),
                                "bbox": list(rendered_bbox),
                                "output_size": [width, height],
                                "dpi": dpi_value,
                                "sha256": digest,
                                "output": relative_path,
                                "duplicate_of": duplicate_of,
                            }
                        )
    finally:
        document.close()

    image_outputs = list(outputs)
    if not image_outputs:
        detail = "；".join(extraction_warnings[:3])
        suffix = f"（{detail}）" if detail else ""
        raise PDFProcessorError(
            "未发现符合当前页码、尺寸和模式条件的 PDF 位图图片。"
            "纯文字或矢量图不会作为内嵌位图出现；可降低最小尺寸、切换“可见位置/"
            f"智能合并”，或使用“PDF 转图片”导出整页{suffix}"
        )

    if write_manifest:
        manifest_target = unique_path(root / "图片提取清单.json", overwrite)
        manifest = {
            "source": str(source_path),
            "mode": mode_name,
            "selected_pages": [index + 1 for index in page_indexes],
            "settings": {
                "image_format": format_name,
                "dpi": dpi_value,
                "jpeg_quality": quality_value,
                "min_width": minimum_width,
                "min_height": minimum_height,
                "merge_gap": merge_gap_value,
                "region_padding": padding_value,
                "deduplicate": deduplicate,
                "include_annotations": include_annotations,
            },
            "image_count": len(image_outputs),
            "records": records,
            "warnings": extraction_warnings,
        }
        with _atomic_output(manifest_target, overwrite) as temporary:
            temporary.write_text(
                json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        outputs.append(manifest_target)
    return outputs


extract_images = extract_pdf_images

__all__ = ["extract_images", "extract_pdf_images"]
