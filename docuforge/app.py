from __future__ import annotations

import ctypes
import contextlib
import json
import math
import os
import queue
import re
import threading
import time
import traceback
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, Callable, Iterable
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

import tkinter as tk
from tkinter import colorchooser, filedialog, font as tkfont, messagebox, ttk

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps, ImageTk

try:
    from tkinterdnd2 import COPY as DND_COPY
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # Buttons remain available if an older installation is reused.
    DND_COPY = "copy"
    DND_FILES = None
    _TkBase = tk.Tk
else:
    _TkBase = TkinterDnD.Tk

from .models import (
    CancelledError,
    Capability,
    DocuForgeError,
    Operation,
    ParameterSpec,
    TaskFailure,
    TaskResult,
)
from .runner import TaskRunner, progress_message

BG = "#F3F6FB"
PANEL = "#FFFFFF"
PANEL_ALT = "#F8FAFD"
SIDEBAR_BG = "#F7F9FD"
TEXT = "#172033"
MUTED = "#667085"
ACCENT = "#4F6BED"
ACCENT_DARK = "#3B55D9"
ACCENT_SOFT = "#EEF2FF"
SUCCESS = "#0F8F6B"
WARNING = "#C86B12"
DANGER = "#C83E4D"
BORDER = "#DDE4EF"
SHADOW = "#E8EDF5"

WPS_COMPATIBILITY_NOTICE = (
    "鉴于 WPS Office 在国内拥有庞大的用户基础，本软件主要围绕 WPS 进行定向优化。"
    "受开发时间限制，Microsoft Office 与 LibreOffice 的部分功能尚不完善；"
    "请安装桌面版 WPS Office，以获得最佳功能与版式体验。"
)


@dataclass(frozen=True)
class TaskResultPresentation:
    title: str
    subtitle: str
    icon: str
    gradient_start: str
    gradient_end: str
    accent: str


@dataclass(frozen=True)
class UnavailableOperationPresentation:
    title: str
    message: str


def unavailable_operation_presentation(
    operation: Operation, capability: Capability
) -> UnavailableOperationPresentation:
    """Build a clear dialog for an operation whose local engine is missing."""

    if capability.engine == "Office 渲染器":
        return UnavailableOperationPresentation(
            "缺少 Office 转换引擎",
            (
                f"“{operation.name}”当前无法运行。\n\n"
                f"检测结果：{capability.reason}\n\n"
                "本软件主要围绕 WPS Office 进行定向优化。请安装桌面版 WPS Office（推荐），"
                "或安装 Microsoft Office、自行安装 "
                "LibreOffice。若已经安装 WPS 仍看到此提示，请修复安装或重新安装桌面版 "
                "WPS，以确保 COM 自动化接口正确注册，然后彻底退出并重新启动页织工坊。"
            ),
        )
    return UnavailableOperationPresentation(
        "缺少处理组件",
        (
            f"“{operation.name}”当前无法运行。\n\n"
            f"检测结果：{capability.reason}\n\n"
            "请按上方功能说明安装或修复所需组件，然后重新启动页织工坊。"
        ),
    )


def task_result_presentation(result: TaskResult) -> TaskResultPresentation:
    """Map a structured task result to one unambiguous visual state."""

    completed = len(result.completed_inputs)
    failed = len(result.failed_inputs)
    unfinished = failed + len(result.cancelled_inputs)
    if result.outcome == "success":
        return TaskResultPresentation(
            "处理成功",
            f"全部处理完成，已生成 {len(result.outputs)} 个文件。",
            "✓",
            "#0F9D70",
            "#65D6A5",
            "#087C58",
        )
    if result.outcome == "partial":
        reason = "任务已由用户停止" if result.cancelled else "部分文件未能完成"
        return TaskResultPresentation(
            "部分完成",
            f"{reason}；成功 {completed} 个，未完成 {unfinished} 个。",
            "!",
            "#D79518",
            "#F4CE67",
            "#A76505",
        )
    if result.outcome == "cancelled":
        return TaskResultPresentation(
            "任务已取消",
            "任务已停止，未完成的临时输出已清理。",
            "!",
            "#D79518",
            "#F4CE67",
            "#A76505",
        )
    return TaskResultPresentation(
        "处理失败",
            f"没有文件处理成功，共 {unfinished} 个文件未完成。",
        "!",
        "#CF3F4C",
        "#F08A80",
        "#A92837",
    )

IMAGE_PREVIEW_OPERATION_IDS = frozenset(
    {
        "image.convert",
        "image.resize",
        "image.scale",
        "image.crop",
        "image.rotate",
        "image.flip",
        "image.compress",
        "image.adjust",
        "image.filter",
        "image.text_watermark",
        "image.image_watermark",
        "image.border",
        "image.mosaic",
        "image.overlay",
    }
)


@dataclass(frozen=True)
class InputCollectionResult:
    files: tuple[Path, ...]
    scanned_directories: int = 0
    duplicate_files: int = 0
    unsupported_files: int = 0
    missing_paths: tuple[str, ...] = ()
    errors: tuple[str, ...] = ()


def natural_path_key(path: Path) -> tuple[object, ...]:
    parts = re.split(r"(\d+)", str(path).casefold())
    return tuple(int(part) if part.isdigit() else part for part in parts)


def normalize_input_path(raw: str | os.PathLike[str]) -> Path:
    """Return a stable absolute path for dialog and desktop-drop inputs."""

    text = os.fspath(raw).strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {'"', "'"}:
        text = text[1:-1]
    parsed = urlparse(text)
    if parsed.scheme.casefold() == "file":
        uri_path = url2pathname(unquote(parsed.path))
        if parsed.netloc:
            uri_path = f"//{parsed.netloc}{uri_path}"
        text = uri_path
    text = os.path.expandvars(os.path.expanduser(text))
    # Keep normalization purely lexical. Path.resolve() may probe every path
    # component and can stall for seconds on unavailable UNC/network shares.
    # Existence and type are checked exactly once by the collection caller.
    return Path(os.path.abspath(os.path.normpath(text)))


def canonical_path_key(path: Path) -> str:
    return os.path.normcase(os.path.normpath(str(path)))


def parse_drop_payload(
    payload: object,
    splitlist: Callable[[str], Iterable[object]],
) -> tuple[str, ...]:
    """Decode a tkdnd file-list while preserving spaces and Unicode names."""

    if isinstance(payload, (tuple, list)):
        values = payload
    else:
        text = str(payload or "").strip()
        if not text:
            return ()
        try:
            values = tuple(splitlist(text))
        except Exception:
            values = (text,)
    return tuple(str(value).strip() for value in values if str(value).strip())


def collect_input_files(
    sources: Iterable[str | os.PathLike[str]],
    *,
    allowed_extensions: Iterable[str] = (),
    existing_paths: Iterable[Path] = (),
) -> InputCollectionResult:
    """Recursively expand dropped sources, normalize them and remove duplicates."""

    allowed = {
        (
            extension.casefold()
            if str(extension).startswith(".")
            else f".{str(extension).casefold()}"
        )
        for extension in allowed_extensions
        if str(extension).strip()
    }
    known = {canonical_path_key(normalize_input_path(path)) for path in existing_paths}
    files: list[Path] = []
    missing: list[str] = []
    errors: list[str] = []
    duplicate_files = 0
    unsupported_files = 0
    scanned_directories = 0

    def add_file(candidate: Path) -> None:
        nonlocal duplicate_files, unsupported_files
        try:
            normalized = normalize_input_path(candidate)
            is_file = normalized.is_file()
        except OSError as exc:
            errors.append(f"{candidate}：{exc}")
            return
        if not is_file:
            missing.append(str(normalized))
            return
        if allowed and normalized.suffix.casefold() not in allowed:
            unsupported_files += 1
            return
        key = canonical_path_key(normalized)
        if key in known:
            duplicate_files += 1
            return
        known.add(key)
        files.append(normalized)

    for raw_source in sources:
        try:
            source = normalize_input_path(raw_source)
            is_file = source.is_file()
            is_directory = source.is_dir()
        except (OSError, ValueError) as exc:
            errors.append(f"{raw_source}：{exc}")
            continue
        if is_file:
            add_file(source)
            continue
        if not is_directory:
            missing.append(str(source))
            continue

        scanned_directories += 1
        directory_errors: list[OSError] = []
        discovered: list[Path] = []

        def on_walk_error(exc: OSError) -> None:
            directory_errors.append(exc)

        try:
            for root, directories, names in os.walk(
                source,
                topdown=True,
                onerror=on_walk_error,
                followlinks=False,
            ):
                directories.sort(key=lambda name: natural_path_key(Path(name)))
                for name in sorted(
                    names,
                    key=lambda value: natural_path_key(Path(value)),
                ):
                    candidate = Path(root) / name
                    # Filter by suffix before normalizing/stat'ing every file in
                    # a mixed directory. This keeps large folder drops bounded
                    # by the number of files the selected operation can use.
                    if allowed and candidate.suffix.casefold() not in allowed:
                        unsupported_files += 1
                        continue
                    discovered.append(candidate)
        except OSError as exc:
            directory_errors.append(exc)
        for exc in directory_errors:
            location = getattr(exc, "filename", None) or source
            errors.append(f"{location}：{exc.strerror or exc}")
        for candidate in sorted(discovered, key=natural_path_key):
            add_file(candidate)

    return InputCollectionResult(
        files=tuple(files),
        scanned_directories=scanned_directories,
        duplicate_files=duplicate_files,
        unsupported_files=unsupported_files,
        missing_paths=tuple(dict.fromkeys(missing)),
        errors=tuple(dict.fromkeys(errors)),
    )


CATALOG_STRUCTURE: tuple[tuple[str, tuple[str, ...]], ...] = (
    (
        "文档工具",
        (
            "PDF 转换与生成",
            "PDF 页面整理",
            "PDF 压缩与安全",
            "PDF 水印与页码",
            "Office 格式转换",
            "Word 文档处理",
            "Excel 数据处理",
            "PowerPoint 演示处理",
            "兼容修复 / 高级工具",
        ),
    ),
    (
        "图片工具",
        (
            "格式转换与批量",
            "尺寸裁剪与几何",
            "压缩优化与隐私",
            "编辑增强与合成",
        ),
    ),
    ("视频工具", ("视频生成", "转码压缩与提取")),
)


def catalog_order_key(root_name: str, section_name: str = "") -> tuple[int, int, str]:
    """Return a deterministic navigation order for catalog roots and sections."""

    for root_index, (known_root, sections) in enumerate(CATALOG_STRUCTURE):
        if root_name != known_root:
            continue
        try:
            section_index = sections.index(section_name) if section_name else -1
        except ValueError:
            section_index = len(sections)
        return (root_index, section_index, section_name.casefold())
    return (len(CATALOG_STRUCTURE), 0, f"{root_name}\0{section_name}".casefold())


def operation_catalog_path(operation: Operation) -> tuple[str, str]:
    """Return the stable two-level navigation path for a public operation."""

    operation_id = operation.id
    if operation_id.startswith("video.") or operation.group == "视频生成":
        section = "视频生成" if operation.group == "视频生成" else "转码压缩与提取"
        return ("视频工具", section)

    if operation_id.startswith("pdf.") or operation_id == "image.to_pdf":
        if operation_id in {
            "pdf.merge",
            "pdf.split",
            "pdf.extract_pages",
            "pdf.delete_pages",
            "pdf.insert_pages",
            "pdf.rotate",
        }:
            return ("文档工具", "PDF 页面整理")
        if operation_id in {
            "pdf.compress",
            "pdf.compress_lossy",
            "pdf.encrypt",
            "pdf.decrypt",
        }:
            return ("文档工具", "PDF 压缩与安全")
        if operation_id in {"pdf.watermark", "pdf.header_footer"}:
            return ("文档工具", "PDF 水印与页码")
        return ("文档工具", "PDF 转换与生成")

    if operation_id.startswith(("word.", "excel.", "ppt.", "legacy.")):
        if operation.group == "兼容修复 / 高级工具":
            return ("文档工具", "兼容修复 / 高级工具")
        if operation.group == "Word 专项处理":
            return ("文档工具", "Word 文档处理")
        if operation.group == "Excel 专项处理":
            return ("文档工具", "Excel 数据处理")
        if operation.group == "PPT 专项处理":
            return ("文档工具", "PowerPoint 演示处理")
        return ("文档工具", "Office 格式转换")

    if operation_id.startswith("image."):
        sections = {
            "图片格式与批处理": "格式转换与批量",
            "图片尺寸与裁剪": "尺寸裁剪与几何",
            "图片压缩与优化": "压缩优化与隐私",
            "图片效果与编辑": "编辑增强与合成",
        }
        return ("图片工具", sections.get(operation.group, "其他图片处理"))

    return ("其他工具", operation.group)


def parameter_help_text(spec: ParameterSpec) -> str:
    """Build a concise, always-visible explanation for a parameter control."""

    details: list[str] = []
    if spec.help_text:
        details.append(spec.help_text.rstrip("。"))
    if spec.kind in {"integer", "number"}:
        if spec.minimum is not None and spec.maximum is not None:
            details.append(f"允许范围：{spec.minimum:g}–{spec.maximum:g}")
        elif spec.minimum is not None:
            details.append(f"最小值：{spec.minimum:g}")
        elif spec.maximum is not None:
            details.append(f"最大值：{spec.maximum:g}")
    if spec.default not in (None, "") and spec.kind != "password":
        default_value: object = spec.default
        if spec.kind == "choice":
            default_value = dict(spec.choices).get(str(spec.default), spec.default)
        elif spec.kind == "boolean":
            default_value = "开启" if bool(spec.default) else "关闭"
        details.append(f"默认：{default_value}")
    if spec.kind == "choice":
        details.append("请从下拉菜单选择")
    return "；".join(dict.fromkeys(details)) + ("。" if details else "")


def color_dialog_initial(value: object) -> str:
    """Return one Tk-compatible color from a possibly multi-color value."""

    for token in str(value or "").split(";"):
        candidate = token.strip()
        if not candidate:
            continue
        if re.fullmatch(r"#[0-9a-fA-F]{6}", candidate):
            return candidate
        match = re.fullmatch(
            r"\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*", candidate
        )
        if match:
            channels = tuple(int(channel) for channel in match.groups())
            if all(0 <= channel <= 255 for channel in channels):
                return "#" + "".join(f"{channel:02x}" for channel in channels)
    return "#000000"


def append_color_value(existing: object, selected: str) -> str:
    """Append a selected RGB color while preserving a compact unique list."""

    normalized = selected.strip()
    values = [token.strip() for token in str(existing or "").split(";") if token.strip()]
    if normalized and normalized.casefold() not in {item.casefold() for item in values}:
        values.append(normalized)
    return ";".join(values)


def fitted_media_rect(
    source_width: int,
    source_height: int,
    viewport_width: int,
    viewport_height: int,
) -> tuple[int, int, int, int]:
    """Return a centred aspect-fit rectangle inside a preview viewport."""

    if min(source_width, source_height, viewport_width, viewport_height) <= 0:
        return (0, 0, 0, 0)
    scale = min(viewport_width / source_width, viewport_height / source_height)
    width = max(1, int(round(source_width * scale)))
    height = max(1, int(round(source_height * scale)))
    left = (viewport_width - width) // 2
    top = (viewport_height - height) // 2
    return (left, top, left + width, top + height)


def canvas_point_to_source(
    x: float,
    y: float,
    display_rect: tuple[int, int, int, int],
    source_width: int,
    source_height: int,
) -> tuple[int, int] | None:
    """Map a preview click back to an exact source-frame pixel."""

    left, top, right, bottom = display_rect
    if right <= left or bottom <= top or not (left <= x < right and top <= y < bottom):
        return None
    source_x = int((x - left) * source_width / (right - left))
    source_y = int((y - top) * source_height / (bottom - top))
    return (
        min(max(0, source_x), max(0, source_width - 1)),
        min(max(0, source_y), max(0, source_height - 1)),
    )


def canvas_selection_to_percent_region(
    start: tuple[float, float],
    end: tuple[float, float],
    display_rect: tuple[int, int, int, int],
) -> tuple[float, float, float, float] | None:
    """Convert a clamped preview drag into x/y/width/height percentages."""

    left, top, right, bottom = display_rect
    width = right - left
    height = bottom - top
    if width <= 0 or height <= 0:
        return None
    x1 = min(max(min(start[0], end[0]), left), right)
    y1 = min(max(min(start[1], end[1]), top), bottom)
    x2 = min(max(max(start[0], end[0]), left), right)
    y2 = min(max(max(start[1], end[1]), top), bottom)
    if x2 - x1 < 3 or y2 - y1 < 3:
        return None
    return (
        100.0 * (x1 - left) / width,
        100.0 * (y1 - top) / height,
        100.0 * (x2 - x1) / width,
        100.0 * (y2 - y1) / height,
    )


def robust_frame_rgb(frame: Any, x: int, y: int, radius: int = 2) -> tuple[int, int, int]:
    """Sample a small BGR video patch and return its median RGB colour."""

    import numpy as np

    height, width = frame.shape[:2]
    left = max(0, int(x) - radius)
    right = min(width, int(x) + radius + 1)
    top = max(0, int(y) - radius)
    bottom = min(height, int(y) + radius + 1)
    patch = frame[top:bottom, left:right]
    if patch.size == 0:
        raise ValueError("取色位置不在视频画面内")
    blue, green, red = np.median(patch.reshape(-1, 3), axis=0)
    return int(round(red)), int(round(green)), int(round(blue))


def format_percent_region(region: tuple[float, float, float, float]) -> str:
    return ",".join(f"{value:.2f}".rstrip("0").rstrip(".") for value in region)


def supports_live_image_preview(operation_id: str) -> bool:
    """Return whether an operation has a truthful lightweight comparison."""

    return str(operation_id) in IMAGE_PREVIEW_OPERATION_IDS


def _preview_position(
    canvas_size: tuple[int, int],
    overlay_size: tuple[int, int],
    position: str,
    margin: int,
) -> tuple[int, int]:
    width, height = canvas_size
    overlay_width, overlay_height = overlay_size
    margin = max(0, int(margin))
    horizontal = {
        "top-left": margin,
        "center-left": margin,
        "bottom-left": margin,
        "top": (width - overlay_width) // 2,
        "center": (width - overlay_width) // 2,
        "bottom": (width - overlay_width) // 2,
        "top-right": width - overlay_width - margin,
        "center-right": width - overlay_width - margin,
        "bottom-right": width - overlay_width - margin,
    }
    vertical = {
        "top-left": margin,
        "top": margin,
        "top-right": margin,
        "center-left": (height - overlay_height) // 2,
        "center": (height - overlay_height) // 2,
        "center-right": (height - overlay_height) // 2,
        "bottom-left": height - overlay_height - margin,
        "bottom": height - overlay_height - margin,
        "bottom-right": height - overlay_height - margin,
    }
    key = str(position or "bottom-right")
    return (
        max(0, min(width - overlay_width, horizontal.get(key, margin))),
        max(0, min(height - overlay_height, vertical.get(key, margin))),
    )


def _preview_flatten_alpha(image: Image.Image, background: object) -> Image.Image:
    rgba = image.convert("RGBA")
    flattened = Image.new("RGBA", rgba.size, str(background or "white"))
    flattened.alpha_composite(rgba)
    rgba.close()
    result = flattened.convert("RGB")
    flattened.close()
    return result


def _preview_filter(image: Image.Image, name: str, intensity: float) -> Image.Image:
    name = str(name).strip().lower().replace("-", "_").replace(" ", "_")
    intensity = max(0.0, float(intensity))
    if name == "grayscale":
        return ImageOps.grayscale(image).convert("RGB")
    if name == "black_white":
        grayscale = ImageOps.grayscale(image)
        threshold = max(0, min(255, round(128 * max(intensity, 0.01))))
        result = grayscale.point(lambda value: 255 if value >= threshold else 0)
        grayscale.close()
        return result.convert("RGB")
    if name == "sepia":
        grayscale = ImageOps.grayscale(image)
        sepia = ImageOps.colorize(grayscale, "#2B170B", "#F4D7A1")
        grayscale.close()
        if intensity == 1.0:
            return sepia
        original = image.convert("RGB")
        result = Image.blend(original, sepia, min(1.0, intensity))
        original.close()
        sepia.close()
        return result
    if name in {"blur", "gaussian_blur"}:
        return image.filter(ImageFilter.GaussianBlur(radius=intensity))
    if name == "sharpen":
        return image.filter(
            ImageFilter.UnsharpMask(
                radius=2,
                percent=round(100 * intensity),
                threshold=3,
            )
        )
    filters = {
        "emboss": ImageFilter.EMBOSS,
        "find_edges": ImageFilter.FIND_EDGES,
        "smooth": ImageFilter.SMOOTH_MORE,
    }
    return image.filter(filters.get(name, ImageFilter.DETAIL))


def build_live_image_preview(
    source_path: Path,
    operation_id: str,
    parameters: dict[str, Any],
    *,
    max_dimension: int = 1400,
) -> tuple[Image.Image, Image.Image, str, str, tuple[int, int]]:
    """Build a bounded original/result pair without writing output files."""

    with Image.open(source_path) as opened:
        transposed = ImageOps.exif_transpose(opened)
        source = transposed.convert("RGBA" if "A" in transposed.getbands() else "RGB")
        if transposed is not opened:
            transposed.close()
    original_width, original_height = source.size
    source.thumbnail((max_dimension, max_dimension), Image.Resampling.LANCZOS)
    preview_width, preview_height = source.size
    scale_x = preview_width / max(1, original_width)
    scale_y = preview_height / max(1, original_height)
    source_display = source.copy()
    result = source.copy()
    expected_width, expected_height = original_width, original_height
    note = "参数变化后自动刷新，仅处理当前选中的预览图片。"

    def scaled_x(value: object) -> int:
        return round(float(value) * scale_x)

    def scaled_y(value: object) -> int:
        return round(float(value) * scale_y)

    def replace(new_image: Image.Image) -> None:
        nonlocal result
        old = result
        result = new_image
        if old is not source and old is not new_image:
            old.close()

    if operation_id == "image.resize":
        expected_width = max(1, int(parameters["width"]))
        expected_height = max(1, int(parameters["height"]))
        if bool(parameters.get("keep_aspect")):
            ratio = min(
                expected_width / max(1, original_width),
                expected_height / max(1, original_height),
            )
            expected_width = max(1, round(original_width * ratio))
            expected_height = max(1, round(original_height * ratio))
        preview_ratio = min(1.0, max_dimension / max(expected_width, expected_height))
        replace(
            result.resize(
                (
                    max(1, round(expected_width * preview_ratio)),
                    max(1, round(expected_height * preview_ratio)),
                ),
                Image.Resampling.LANCZOS,
            )
        )
    elif operation_id == "image.scale":
        ratio = max(0.01, float(parameters["percent"]) / 100.0)
        expected_width = max(1, round(original_width * ratio))
        expected_height = max(1, round(original_height * ratio))
        preview_ratio = min(1.0, max_dimension / max(expected_width, expected_height))
        replace(
            result.resize(
                (
                    max(1, round(expected_width * preview_ratio)),
                    max(1, round(expected_height * preview_ratio)),
                ),
                Image.Resampling.LANCZOS,
            )
        )
    elif operation_id == "image.crop":
        left = int(parameters["left"])
        top = int(parameters["top"])
        right = int(parameters["right"])
        bottom = int(parameters["bottom"])
        if right <= left or bottom <= top:
            raise ValueError("裁切范围需要满足：右边界大于左边界、下边界大于上边界")
        expected_width = right - left
        expected_height = bottom - top
        box = (scaled_x(left), scaled_y(top), scaled_x(right), scaled_y(bottom))
        replace(result.crop(box))
        overlay = source_display.convert("RGBA")
        draw = ImageDraw.Draw(overlay, "RGBA")
        clamped = (
            max(0, min(preview_width, box[0])),
            max(0, min(preview_height, box[1])),
            max(0, min(preview_width, box[2])),
            max(0, min(preview_height, box[3])),
        )
        draw.rectangle(clamped, outline=ACCENT, width=max(2, round(3 * scale_x)))
        source_display.close()
        source_display = overlay
        note = "左图蓝框表示保留范围；可直接在左图重新拖框，坐标会自动回填。"
    elif operation_id == "image.rotate":
        angle = float(parameters["angle"])
        expand = bool(parameters.get("expand", True))
        fill = str(parameters.get("background") or "white")
        replace(
            result.rotate(
                angle,
                Image.Resampling.BICUBIC,
                expand=expand,
                fillcolor=fill,
            )
        )
        if expand:
            radians = math.radians(angle % 180)
            expected_width = max(
                1,
                round(
                    abs(original_width * math.cos(radians))
                    + abs(original_height * math.sin(radians))
                ),
            )
            expected_height = max(
                1,
                round(
                    abs(original_width * math.sin(radians))
                    + abs(original_height * math.cos(radians))
                ),
            )
    elif operation_id == "image.flip":
        replace(
            ImageOps.mirror(result)
            if parameters.get("direction") == "horizontal"
            else ImageOps.flip(result)
        )
    elif operation_id == "image.adjust":
        brightness = ImageEnhance.Brightness(result).enhance(
            float(parameters["brightness"])
        )
        contrast = ImageEnhance.Contrast(brightness).enhance(
            float(parameters["contrast"])
        )
        brightness.close()
        adjusted = ImageEnhance.Color(contrast).enhance(
            float(parameters["saturation"])
        )
        contrast.close()
        replace(adjusted)
    elif operation_id == "image.filter":
        replace(
            _preview_filter(
                result,
                str(parameters["filter"]),
                float(parameters["intensity"]),
            )
        )
    elif operation_id == "image.border":
        width = max(0, int(parameters["width"]))
        scaled_width = max(0, round(width * (scale_x + scale_y) / 2))
        replace(
            ImageOps.expand(
                result,
                border=scaled_width,
                fill=str(parameters.get("color") or "black"),
            )
        )
        expected_width += width * 2
        expected_height += width * 2
    elif operation_id == "image.mosaic":
        left = int(parameters.get("left", 0))
        top = int(parameters.get("top", 0))
        right = int(parameters.get("right", 0))
        bottom = int(parameters.get("bottom", 0))
        if right <= left or bottom <= top:
            box = (0, 0, result.width, result.height)
        else:
            box = (scaled_x(left), scaled_y(top), scaled_x(right), scaled_y(bottom))
        box = (
            max(0, min(result.width, box[0])),
            max(0, min(result.height, box[1])),
            max(0, min(result.width, box[2])),
            max(0, min(result.height, box[3])),
        )
        if box[2] > box[0] and box[3] > box[1]:
            region = result.crop(box)
            block = max(2, round(int(parameters["block_size"]) * scale_x))
            tiny = region.resize(
                (
                    max(1, region.width // block),
                    max(1, region.height // block),
                ),
                Image.Resampling.BILINEAR,
            )
            pixelated = tiny.resize(region.size, Image.Resampling.NEAREST)
            result.paste(pixelated, box[:2])
            region.close()
            tiny.close()
            pixelated.close()
            draw = ImageDraw.Draw(source_display)
            draw.rectangle(
                (
                    box[0],
                    box[1],
                    max(box[0], box[2] - 1),
                    max(box[1], box[3] - 1),
                ),
                outline=ACCENT,
                width=max(2, round(3 * scale_x)),
            )
            note = "左图蓝框表示打码区域；也可以直接在左图重新拖框。"
    elif operation_id == "image.text_watermark":
        overlay = Image.new("RGBA", result.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        font_size = max(6, scaled_y(parameters["font_size"]))
        font_path = parameters.get("font_path")
        candidates = [
            Path(font_path) if font_path else None,
            Path("C:/Windows/Fonts/msyh.ttc"),
            Path("C:/Windows/Fonts/simhei.ttf"),
        ]
        font = None
        for candidate in candidates:
            if candidate and candidate.exists():
                with contextlib.suppress(OSError):
                    font = ImageFont.truetype(str(candidate), font_size)
                    break
        font = font or ImageFont.load_default()
        text_value = str(parameters.get("text") or "")
        bounds = draw.textbbox((0, 0), text_value, font=font)
        text_size = (max(1, bounds[2] - bounds[0]), max(1, bounds[3] - bounds[1]))
        position = _preview_position(
            result.size,
            text_size,
            str(parameters.get("position")),
            scaled_x(parameters.get("margin", 0)),
        )
        opacity = max(0, min(255, round(float(parameters["opacity"]) * 255)))
        try:
            from PIL import ImageColor

            red, green, blue = ImageColor.getrgb(str(parameters.get("color") or "white"))
        except (ValueError, TypeError):
            red, green, blue = (255, 255, 255)
        draw.text(position, text_value, font=font, fill=(red, green, blue, opacity))
        base = result.convert("RGBA")
        combined = Image.alpha_composite(base, overlay)
        base.close()
        overlay.close()
        replace(combined)
    elif operation_id in {"image.image_watermark", "image.overlay"}:
        auxiliary = parameters.get(
            "watermark_path" if operation_id == "image.image_watermark" else "overlay_path"
        )
        if not auxiliary or not Path(auxiliary).is_file():
            raise ValueError("请选择有效的水印或叠加图片后即可预览")
        with Image.open(auxiliary) as opened_overlay:
            transposed_overlay = ImageOps.exif_transpose(opened_overlay)
            watermark = transposed_overlay.convert("RGBA")
            if transposed_overlay is not opened_overlay:
                transposed_overlay.close()
        ratio = float(parameters.get("scale") or 0)
        if ratio > 0:
            target_width = max(1, round(result.width * ratio))
            target_height = max(1, round(watermark.height * target_width / watermark.width))
            resized = watermark.resize(
                (target_width, target_height),
                Image.Resampling.LANCZOS,
            )
            watermark.close()
            watermark = resized
        opacity = max(0.0, min(1.0, float(parameters.get("opacity", 1))))
        if opacity < 1:
            alpha = watermark.getchannel("A").point(lambda value: round(value * opacity))
            watermark.putalpha(alpha)
            alpha.close()
        position = _preview_position(
            result.size,
            watermark.size,
            str(parameters.get("position")),
            scaled_x(parameters.get("margin", 0)),
        )
        base = result.convert("RGBA")
        base.alpha_composite(watermark, position)
        watermark.close()
        replace(base)
    elif operation_id in {"image.convert", "image.compress"}:
        quality = int(parameters.get("quality", 90))
        target_format = str(parameters.get("format") or source_path.suffix.lstrip("."))
        if operation_id == "image.compress":
            target_format = source_path.suffix.lstrip(".")
            note = "右图模拟压缩画质趋势；最终文件大小仍以原尺寸编码结果为准。"
        target_format = "JPEG" if target_format.lower() in {"jpg", "jpeg"} else target_format.upper()
        if target_format in {"JPEG", "WEBP"}:
            encodable = (
                _preview_flatten_alpha(result, parameters.get("background", "white"))
                if target_format == "JPEG"
                else result
            )
            stream = BytesIO()
            encodable.save(stream, format=target_format, quality=quality)
            if encodable is not result:
                encodable.close()
            stream.seek(0)
            with Image.open(stream) as decoded:
                replace(decoded.convert("RGB"))

    original_info = f"原图：{original_width} × {original_height} px · {source_path.name}"
    result_info = f"预计结果：{expected_width} × {expected_height} px"
    if operation_id == "image.compress":
        result_info += f" · 质量 {int(parameters.get('quality', 90))}"
    source.close()
    return (
        source_display,
        result,
        original_info,
        f"{result_info}\n{note}",
        (original_width, original_height),
    )


class VideoFramePicker(tk.Toplevel):
    """Modal video-frame eyedropper and percentage-region picker."""

    def __init__(self, owner: tk.Misc, source: Path, *, mode: str) -> None:
        super().__init__(owner)
        import cv2

        self.cv2 = cv2
        self.source = source
        self.mode = mode
        self.result: str | None = None
        self.frame: Any | None = None
        self.photo: ImageTk.PhotoImage | None = None
        self.display_rect = (0, 0, 0, 0)
        self.drag_start: tuple[float, float] | None = None
        self.selection: tuple[float, float, float, float] | None = None
        self.marker: tuple[float, float] | None = None
        self._seek_job: str | None = None
        self._layout_job: str | None = None
        self._picker_layout = ""
        self.capture = cv2.VideoCapture(str(source))
        if not self.capture.isOpened():
            self.capture.release()
            self.destroy()
            raise OSError(f"无法读取视频：{source.name}")
        self.fps = float(self.capture.get(cv2.CAP_PROP_FPS) or 25.0)
        self.frame_count = max(1, int(self.capture.get(cv2.CAP_PROP_FRAME_COUNT) or 1))
        self.duration = max(0.0, (self.frame_count - 1) / max(0.001, self.fps))

        title = "从视频提取手写颜色" if mode == "color" else "在视频画面中框选区域"
        self.title(title)
        screen_width = max(1, int(self.winfo_screenwidth()))
        screen_height = max(1, int(self.winfo_screenheight()))
        try:
            self._display_scale = max(1.0, float(self.winfo_fpixels("1i")) / 96.0)
        except (tk.TclError, TypeError, ValueError):
            self._display_scale = 1.0
        logical_screen_width = max(1, round(screen_width / self._display_scale))
        logical_screen_height = max(1, round(screen_height / self._display_scale))
        logical_width, logical_height = fitted_dialog_size(
            logical_screen_width,
            logical_screen_height,
            preferred_width=1080,
            preferred_height=780,
            minimum_width=720,
            minimum_height=560,
        )
        initial_width = max(1, round(logical_width * self._display_scale))
        initial_height = max(1, round(logical_height * self._display_scale))
        left = max(0, (screen_width - initial_width) // 2)
        top = max(0, (screen_height - initial_height) // 2)
        self.geometry(f"{initial_width}x{initial_height}+{left}+{top}")
        self.minsize(
            min(screen_width, round(720 * self._display_scale)),
            min(screen_height, round(560 * self._display_scale)),
        )
        self.configure(bg=BG)
        self.transient(owner.winfo_toplevel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        shell = ttk.Frame(self, padding=16, style="Panel.TFrame")
        shell.pack(fill="both", expand=True, padx=10, pady=10)
        instruction = (
            "拖动时间轴找到手写笔迹最清晰的一帧，然后直接点击笔迹。软件会读取点击点附近 5×5 像素的中位 RGB，降低抗锯齿和压缩噪声影响。"
            if mode == "color"
            else "拖动时间轴找到目标清晰的一帧，然后按住鼠标拖出矩形。坐标会自动换算为画面百分比，无需手工填写。"
        )
        self.instruction_label = ttk.Label(
            shell,
            text=instruction,
            wraplength=980,
            justify="left",
            style="Subtle.TLabel",
        )
        self.instruction_label.pack(
            fill="x", pady=(0, 12)
        )
        preview_row = ttk.Frame(shell, style="Panel.TFrame")
        preview_row.pack(fill="both", expand=True)
        preview_row.columnconfigure(0, weight=1)
        preview_row.rowconfigure(0, weight=1)
        canvas_shell = tk.Frame(
            preview_row,
            bg=SHADOW,
            padx=1,
            pady=1,
            borderwidth=0,
        )
        canvas_shell.grid(row=0, column=0, sticky="nsew")
        self.canvas = tk.Canvas(
            canvas_shell,
            background="#111827",
            highlightthickness=0,
            borderwidth=0,
        )
        self.canvas.pack(fill="both", expand=True)
        self.picker_side = ttk.Frame(
            preview_row,
            width=206,
            padding=12,
            style="Soft.TFrame",
        )
        self.picker_side.grid(row=0, column=1, sticky="ns", padx=(12, 0))
        self.picker_side.grid_propagate(False)
        self.value_label = ttk.Label(
            self.picker_side,
            text="尚未取色" if mode == "color" else "尚未框选区域",
            wraplength=180,
            justify="left",
            style="CardSubtle.TLabel",
        )
        self.value_label.pack(fill="x", pady=(0, 10))
        self.swatch = tk.Label(
            self.picker_side,
            text="",
            background="#FFFFFF",
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        self.swatch.pack(fill="x", ipady=20, pady=(0, 10))
        self.loupe = tk.Label(
            self.picker_side,
            background="#111827",
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        self.loupe.pack(fill="x", ipady=4)
        self.use_button = ttk.Button(
            self.picker_side,
            text="使用此颜色" if mode == "color" else "使用此区域",
            command=self._accept,
            state="disabled",
            style="Accent.TButton",
        )
        self.use_button.pack(side="bottom", fill="x", pady=(8, 0))
        ttk.Button(
            self.picker_side,
            text="取消",
            command=self._cancel,
            style="Quiet.TButton",
        ).pack(
            side="bottom", fill="x", pady=(8, 0)
        )

        self.picker_timeline = ttk.Frame(shell, padding=(10, 8), style="Soft.TFrame")
        self.picker_timeline.pack(fill="x", pady=(12, 0))
        self.picker_timeline.columnconfigure(1, weight=1)
        self.previous_frame_button = ttk.Button(
            self.picker_timeline,
            text="−1 秒",
            command=lambda: self._step(-1.0),
            style="Quiet.TButton",
        )
        self.position = tk.DoubleVar(value=0.0)
        self.scale = ttk.Scale(
            self.picker_timeline,
            from_=0.0,
            to=max(0.01, self.duration),
            variable=self.position,
            command=self._schedule_seek,
        )
        self.next_frame_button = ttk.Button(
            self.picker_timeline,
            text="+1 秒",
            command=lambda: self._step(1.0),
            style="Quiet.TButton",
        )
        self.time_label = ttk.Label(
            self.picker_timeline,
            text="00:00.0 / 00:00.0",
            width=20,
            anchor="e",
            style="CardSubtle.TLabel",
        )

        self.canvas.bind("<Configure>", lambda _event: self._render())
        if mode == "color":
            self.canvas.bind("<Button-1>", self._pick_color)
            self.canvas.bind("<Motion>", self._hover_color)
        else:
            self.canvas.bind("<ButtonPress-1>", self._start_region)
            self.canvas.bind("<B1-Motion>", self._drag_region)
            self.canvas.bind("<ButtonRelease-1>", self._finish_region)
        self.bind("<Configure>", self._schedule_responsive_layout, add="+")
        self._apply_responsive_layout(logical_width, force=True)
        self.after_idle(lambda: self._load_frame(0.0))
        self.grab_set()
        self.focus_set()

    def _schedule_responsive_layout(self, event: tk.Event) -> None:
        if event.widget is not self:
            return
        logical_width = max(1, round(int(event.width) / self._display_scale))
        self.instruction_label.configure(wraplength=max(360, int(event.width) - 72))
        if self._layout_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._layout_job)
        self._layout_job = self.after(
            70,
            lambda value=logical_width: self._apply_responsive_layout(value),
        )

    def _apply_responsive_layout(self, logical_width: int, *, force: bool = False) -> None:
        self._layout_job = None
        layout = "compact" if int(logical_width) < 900 else "wide"
        if layout == self._picker_layout and not force:
            return
        self._picker_layout = layout
        for widget in (
            self.previous_frame_button,
            self.scale,
            self.next_frame_button,
            self.time_label,
        ):
            widget.grid_forget()
        self.previous_frame_button.grid(row=0, column=0, sticky="w")
        self.scale.grid(row=0, column=1, sticky="ew", padx=10)
        self.next_frame_button.grid(row=0, column=2, sticky="e")
        if layout == "compact":
            self.picker_side.configure(width=178)
            self.value_label.configure(wraplength=152)
            self.time_label.grid(
                row=1,
                column=0,
                columnspan=3,
                sticky="e",
                pady=(7, 0),
            )
        else:
            self.picker_side.configure(width=206)
            self.value_label.configure(wraplength=180)
            self.time_label.grid(row=0, column=3, sticky="e", padx=(10, 0))

    @staticmethod
    def _clock(seconds: float) -> str:
        minutes, remainder = divmod(max(0.0, seconds), 60.0)
        return f"{int(minutes):02d}:{remainder:04.1f}"

    def _schedule_seek(self, raw: object) -> None:
        if self._seek_job is not None:
            self.after_cancel(self._seek_job)
        try:
            timestamp = float(raw)
        except (TypeError, ValueError):
            timestamp = float(self.position.get())
        self._seek_job = self.after(80, lambda: self._load_frame(timestamp))

    def _step(self, seconds: float) -> None:
        target = min(max(0.0, float(self.position.get()) + seconds), self.duration)
        self.position.set(target)
        self._load_frame(target)

    def _load_frame(self, timestamp: float) -> None:
        self._seek_job = None
        timestamp = min(max(0.0, float(timestamp)), self.duration)
        self.capture.set(self.cv2.CAP_PROP_POS_MSEC, timestamp * 1000.0)
        ok, frame = self.capture.read()
        if not ok:
            return
        self.frame = frame
        self.position.set(timestamp)
        self.time_label.configure(
            text=f"{self._clock(timestamp)} / {self._clock(self.duration)}"
        )
        self.marker = None
        self.selection = None
        self.result = None
        self.value_label.configure(
            text="尚未取色" if self.mode == "color" else "尚未框选区域"
        )
        self.use_button.configure(state="disabled")
        self._render()

    def _render(self) -> None:
        if self.frame is None or not self.canvas.winfo_exists():
            return
        canvas_width = max(1, self.canvas.winfo_width())
        canvas_height = max(1, self.canvas.winfo_height())
        height, width = self.frame.shape[:2]
        self.display_rect = fitted_media_rect(width, height, canvas_width, canvas_height)
        left, top, right, bottom = self.display_rect
        rgb = self.cv2.cvtColor(self.frame, self.cv2.COLOR_BGR2RGB)
        image = Image.fromarray(rgb).resize(
            (right - left, bottom - top), Image.Resampling.LANCZOS
        )
        self.photo = ImageTk.PhotoImage(image)
        self.canvas.delete("all")
        self.canvas.create_image(left, top, anchor="nw", image=self.photo)
        if self.marker is not None:
            x, y = self.marker
            self.canvas.create_line(x - 10, y, x + 10, y, fill="#FF3B30", width=2)
            self.canvas.create_line(x, y - 10, x, y + 10, fill="#FF3B30", width=2)
        if self.selection is not None:
            x1, y1, x2, y2 = self.selection
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="#FF3B30", width=3)

    def _pick_color(self, event: tk.Event) -> None:
        self._preview_color(event, commit=True)

    def _hover_color(self, event: tk.Event) -> None:
        self._preview_color(event, commit=False)

    def _preview_color(self, event: tk.Event, *, commit: bool) -> None:
        if self.frame is None:
            return
        height, width = self.frame.shape[:2]
        point = canvas_point_to_source(
            event.x, event.y, self.display_rect, width, height
        )
        if point is None:
            return
        red, green, blue = robust_frame_rgb(self.frame, *point)
        selected = f"#{red:02X}{green:02X}{blue:02X}"
        if commit:
            self.result = selected
            self.marker = (float(event.x), float(event.y))
        self.value_label.configure(
            text=(
                ("已锁定颜色\n" if commit else "鼠标当前位置（单击锁定）\n")
                + f"RGB：{red}, {green}, {blue}\nHEX：{selected}\n"
                + f"源像素：{point[0]}, {point[1]}"
            )
        )
        self.swatch.configure(background=selected)
        x, y = point
        patch_left = max(0, x - 7)
        patch_top = max(0, y - 7)
        patch = self.frame[patch_top : y + 8, patch_left : x + 8]
        if patch.size:
            patch_rgb = self.cv2.cvtColor(patch, self.cv2.COLOR_BGR2RGB)
            loupe = Image.fromarray(patch_rgb).resize((150, 150), Image.Resampling.NEAREST)
            centre_x = min(149, max(0, (x - patch_left) * 10 + 5))
            centre_y = min(149, max(0, (y - patch_top) * 10 + 5))
            draw = ImageDraw.Draw(loupe)
            draw.line(
                (centre_x - 15, centre_y, centre_x + 15, centre_y),
                fill="#FF3B30",
                width=2,
            )
            draw.line(
                (centre_x, centre_y - 15, centre_x, centre_y + 15),
                fill="#FF3B30",
                width=2,
            )
            draw.rectangle(
                (centre_x - 5, centre_y - 5, centre_x + 5, centre_y + 5),
                outline="#FFFFFF",
                width=1,
            )
            self._loupe_photo = ImageTk.PhotoImage(loupe)
            self.loupe.configure(image=self._loupe_photo)
        if commit:
            self.use_button.configure(state="normal")
            self._render()

    def _start_region(self, event: tk.Event) -> None:
        self.drag_start = (float(event.x), float(event.y))
        self.selection = (*self.drag_start, *self.drag_start)
        self._render()

    def _drag_region(self, event: tk.Event) -> None:
        if self.drag_start is None:
            return
        self.selection = (*self.drag_start, float(event.x), float(event.y))
        self._render()

    def _finish_region(self, event: tk.Event) -> None:
        if self.drag_start is None:
            return
        self.selection = (*self.drag_start, float(event.x), float(event.y))
        region = canvas_selection_to_percent_region(
            self.drag_start, (float(event.x), float(event.y)), self.display_rect
        )
        self.drag_start = None
        if region is None:
            self.result = None
            self.value_label.configure(text="框选范围太小，请重新拖动")
            self.use_button.configure(state="disabled")
        else:
            self.result = format_percent_region(region)
            self.value_label.configure(
                text=(
                    f"左：{region[0]:.2f}%\n上：{region[1]:.2f}%\n"
                    f"宽：{region[2]:.2f}%\n高：{region[3]:.2f}%"
                )
            )
            self.use_button.configure(state="normal")
        self._render()

    def _accept(self) -> None:
        if self.result:
            self._close()

    def _cancel(self) -> None:
        self.result = None
        self._close()

    def _close(self) -> None:
        if self._seek_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._seek_job)
            self._seek_job = None
        if self._layout_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._layout_job)
            self._layout_job = None
        self.capture.release()
        with contextlib.suppress(tk.TclError):
            self.grab_release()
        self.destroy()


class VideoPptRepairWorkbench(tk.Toplevel):
    """Simple page-centred repair flow with no user-visible intermediate plan."""

    _METHOD_LABELS = {
        "自动多帧修复（推荐）": "temporal",
        "周边背景填充": "background",
        "指定纯色覆盖": "color",
    }

    def __init__(
        self,
        owner: tk.Misc,
        pptx_path: Path,
        video_path: Path,
        *,
        initial_plan: str = "",
    ) -> None:
        super().__init__(owner)
        import cv2

        from .processors import video_slide_repair

        self.cv2 = cv2
        self.engine = video_slide_repair
        self.pptx_path = pptx_path
        self.video_path = video_path
        self.result: str | None = None
        self.start_requested = False
        self.actions: list[dict[str, Any]] = []
        self.page_cache: dict[int, Any] = {}
        self.current_video_frame: Any | None = None
        self.result_preview: Any | None = None
        self.photos: dict[str, ImageTk.PhotoImage] = {}
        self.display_rects = {
            "page": (0, 0, 0, 0),
            "video": (0, 0, 0, 0),
            "result": (0, 0, 0, 0),
        }
        self.drag_source: str | None = None
        self.drag_start: tuple[float, float] | None = None
        self.selection: tuple[float, float, float, float] | None = None
        self.watermark_regions: list[tuple[float, float, float, float]] = []
        self._seek_job: str | None = None
        self._page_change_job: str | None = None
        self._layout_job: str | None = None
        self._last_preview_signature = ""
        self._advanced_visible = False
        self._preview_layout = ""
        self._suspend_page_trace = False
        self._active_mode = "repair_region"
        self._active_context: tuple[str, int] | None = None
        self._context_times: dict[tuple[str, int], float] = {}
        self._existing_page = 1
        self._insert_position = 1

        self.slide_count = video_slide_repair.ppt_slide_count(pptx_path)
        first = video_slide_repair.read_ppt_slide_image(pptx_path, 1)
        self.page_cache[1] = first
        self.target_height, self.target_width = first.shape[:2]
        self.capture = cv2.VideoCapture(str(video_path))
        if not self.capture.isOpened():
            self.capture.release()
            self.destroy()
            raise OSError(f"无法读取原视频：{video_path.name}")
        self.fps = float(self.capture.get(cv2.CAP_PROP_FPS) or 25.0)
        frame_count = max(1, int(self.capture.get(cv2.CAP_PROP_FRAME_COUNT) or 1))
        self.duration = max(0.0, (frame_count - 1) / max(0.001, self.fps))

        if initial_plan.strip():
            try:
                payload = json.loads(initial_plan)
                if payload.get("schema") == video_slide_repair.PLAN_SCHEMA:
                    raw_actions = payload.get("actions")
                    if isinstance(raw_actions, list):
                        self.actions = [
                            dict(item) for item in raw_actions if isinstance(item, dict)
                        ]
            except (json.JSONDecodeError, TypeError, AttributeError):
                self.actions = []
        if self.actions:
            # The visible page must include already confirmed actions from a
            # reopened plan, not the untouched source image cached above.
            self.page_cache.clear()

        self.title("视频 PPT 快速补修")
        screen_width = max(1, int(self.winfo_screenwidth()))
        screen_height = max(1, int(self.winfo_screenheight()))
        try:
            self._display_scale = max(1.0, float(self.winfo_fpixels("1i")) / 96.0)
        except (tk.TclError, TypeError, ValueError):
            self._display_scale = 1.0
        logical_screen_width = max(1, round(screen_width / self._display_scale))
        logical_screen_height = max(1, round(screen_height / self._display_scale))
        logical_width, logical_height = fitted_dialog_size(
            logical_screen_width,
            logical_screen_height,
            preferred_width=1500,
            preferred_height=900,
            minimum_width=860,
            minimum_height=620,
        )
        initial_width = max(1, round(logical_width * self._display_scale))
        initial_height = max(1, round(logical_height * self._display_scale))
        left = max(0, (screen_width - initial_width) // 2)
        top = max(0, (screen_height - initial_height) // 2)
        self.geometry(f"{initial_width}x{initial_height}+{left}+{top}")
        self.minsize(
            min(screen_width, round(860 * self._display_scale)),
            min(screen_height, round(620 * self._display_scale)),
        )
        self.configure(bg=BG)
        self.transient(owner.winfo_toplevel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        style = ttk.Style(self)
        style.configure(
            "RepairMode.TRadiobutton",
            background=PANEL_ALT,
            foreground=TEXT,
            padding=(12, 9),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "RepairMode.TRadiobutton",
            background=[("active", ACCENT_SOFT)],
            foreground=[("selected", ACCENT_DARK)],
        )
        style.configure(
            "Repair.TButton",
            padding=(16, 10),
            font=("Microsoft YaHei UI", 11, "bold"),
        )
        style.configure(
            "RepairQuiet.TButton",
            padding=(13, 9),
            font=("Microsoft YaHei UI", 10),
        )
        with contextlib.suppress(tk.TclError):
            style.layout("Repair.TButton", style.layout("Accent.TButton"))
            style.layout("RepairQuiet.TButton", style.layout("Quiet.TButton"))
        style.configure(
            "Repair.TSpinbox",
            padding=(10, 8),
            font=("Microsoft YaHei UI", 12, "bold"),
            arrowsize=18,
        )

        self.shell = ttk.Frame(self, padding=14, style="Panel.TFrame")
        self.shell.pack(fill="both", expand=True, padx=10, pady=10)
        self.status = tk.StringVar(
            value="选择原 PPT 页后，程序会自动定位该页的视频时间；在左图拖框即可。"
        )
        self.status_label = ttk.Label(
            self.shell,
            textvariable=self.status,
            justify="left",
            wraplength=1360,
            style="DropHint.TLabel",
        )
        self.status_label.pack(fill="x", pady=(0, 10))

        self.mode_row = ttk.Frame(self.shell, padding=(10, 4), style="Soft.TFrame")
        self.mode_row.pack(fill="x", pady=(0, 6))
        ttk.Label(
            self.mode_row,
            text="处理方式：",
            style="CardField.TLabel",
        ).pack(side="left", padx=(0, 4))
        self.mode = tk.StringVar(value="repair_region")
        ttk.Radiobutton(
            self.mode_row,
            text="修复已有页面",
            variable=self.mode,
            value="repair_region",
            command=self._mode_changed,
            style="RepairMode.TRadiobutton",
        ).pack(side="left")
        ttk.Radiobutton(
            self.mode_row,
            text="补插漏页",
            variable=self.mode,
            value="insert_page",
            command=self._mode_changed,
            style="RepairMode.TRadiobutton",
        ).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(
            self.mode_row,
            text="手动选最佳帧替换本页",
            variable=self.mode,
            value="replace_page_frame",
            command=self._mode_changed,
            style="RepairMode.TRadiobutton",
        ).pack(side="left", padx=(6, 0))

        self.page_row = ttk.Frame(self.shell, padding=(10, 5), style="Soft.TFrame")
        self.page_row.pack(fill="x", pady=(0, 10))
        self.page_label = ttk.Label(
            self.page_row,
            text="原 PPT 页码",
            style="CardField.TLabel",
        )
        self.page_label.pack(side="left")
        self.previous_page_button = ttk.Button(
            self.page_row,
            text="◀ 上一页",
            width=9,
            style="RepairQuiet.TButton",
            command=lambda: self._change_page(-1),
        )
        self.previous_page_button.pack(side="left", padx=(10, 6))
        self.page_number = tk.StringVar(value="1")
        self.page_spin = ttk.Spinbox(
            self.page_row,
            from_=1,
            to=self.slide_count,
            width=10,
            textvariable=self.page_number,
            command=self._page_changed,
            style="Repair.TSpinbox",
        )
        self.page_spin.pack(side="left")
        self.page_spin.bind("<Return>", lambda _event: self._page_changed())
        self.page_spin.bind("<FocusOut>", lambda _event: self._page_changed())
        self.page_number.trace_add("write", self._schedule_page_changed)
        self.next_page_button = ttk.Button(
            self.page_row,
            text="下一页 ▶",
            width=9,
            style="RepairQuiet.TButton",
            command=lambda: self._change_page(1),
        )
        self.next_page_button.pack(side="left", padx=(6, 0))
        self.position_note = ttk.Label(
            self.page_row,
            text=f"共 {self.slide_count} 页",
            style="CardSubtle.TLabel",
        )
        self.position_note.pack(side="left", padx=(10, 0))
        ttk.Button(
            self.page_row,
            text="清除全部框",
            width=12,
            style="RepairQuiet.TButton",
            command=self._clear_selection,
        ).pack(side="right")
        ttk.Button(
            self.page_row,
            text="撤销上个框",
            width=12,
            style="RepairQuiet.TButton",
            command=self._undo_last_region,
        ).pack(side="right", padx=(0, 6))

        self.previews = ttk.Frame(self.shell, style="Panel.TFrame")
        self.previews.pack(fill="both", expand=True)
        self.page_box = ttk.LabelFrame(
            self.previews,
            text="① 原 PPT 页面",
            padding=6,
            style="Card.TLabelframe",
        )
        self.video_box = ttk.LabelFrame(
            self.previews,
            text="② 对应视频画面",
            padding=6,
            style="Card.TLabelframe",
        )
        self.result_box = ttk.LabelFrame(
            self.previews,
            text="③ 处理结果",
            padding=6,
            style="Card.TLabelframe",
        )
        self.page_canvas = self._make_canvas(self.page_box, "page")
        self.video_canvas = self._make_canvas(self.video_box, "video")
        self.result_canvas = self._make_canvas(
            self.result_box, "result", selectable=False
        )

        timeline = ttk.Frame(self.shell, padding=(10, 7), style="Soft.TFrame")
        timeline.pack(fill="x", pady=(10, 0))
        for label, seconds in (("−1秒", -1.0), ("−0.1秒", -0.1)):
            ttk.Button(
                timeline,
                text=label,
                width=7,
                style="RepairQuiet.TButton",
                command=lambda amount=seconds: self._step_video(amount),
            ).pack(side="left")
        self.video_time = tk.DoubleVar(value=0.0)
        self.timeline_scale = ttk.Scale(
            timeline,
            from_=0.0,
            to=max(0.01, self.duration),
            variable=self.video_time,
            command=self._schedule_video_seek,
        )
        self.timeline_scale.pack(side="left", fill="x", expand=True, padx=8)
        for label, seconds in (("+0.1秒", 0.1), ("+1秒", 1.0)):
            ttk.Button(
                timeline,
                text=label,
                width=7,
                style="RepairQuiet.TButton",
                command=lambda amount=seconds: self._step_video(amount),
            ).pack(side="left")
        self.time_label = ttk.Label(
            timeline,
            text="00:00.0 / 00:00.0",
            width=21,
            style="CardSubtle.TLabel",
        )
        self.time_label.pack(side="left", padx=(8, 0))

        action_row = ttk.Frame(self.shell, style="Panel.TFrame")
        action_row.pack(fill="x", pady=(8, 0))
        self.selection_label = ttk.Label(
            action_row,
            text="请在左侧原 PPT 页面拖框标出水印。",
            style="Subtle.TLabel",
        )
        self.selection_label.pack(side="left", fill="x", expand=True)
        self.advanced_button = ttk.Button(
            action_row,
            text="高级方式 ▾",
            style="RepairQuiet.TButton",
            command=self._toggle_advanced,
        )
        self.advanced_button.pack(side="right", padx=(8, 0))
        self.primary_button = ttk.Button(
            action_row,
            text="自动校正稳定帧并预览",
            width=24,
            style="Accent.TButton",
            command=self._primary_action,
        )
        self.primary_button.pack(side="right")

        self.advanced_frame = ttk.LabelFrame(
            self.shell,
            text="高级修复方式",
            padding=10,
            style="Card.TLabelframe",
        )
        self.method = tk.StringVar(value=next(iter(self._METHOD_LABELS)))
        self.method_combo = ttk.Combobox(
            self.advanced_frame,
            textvariable=self.method,
            values=list(self._METHOD_LABELS),
            state="readonly",
            width=30,
        )
        self.method_combo.grid(row=0, column=0, sticky="ew")
        self.method_combo.bind(
            "<<ComboboxSelected>>", lambda _event: self._invalidate_preview()
        )
        self.fill_colour = tk.StringVar(value="#FFFFFF")
        self.fill_colour_entry = ttk.Entry(
            self.advanced_frame, textvariable=self.fill_colour, width=12
        )
        self.fill_colour_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.choose_fill_button = ttk.Button(
            self.advanced_frame,
            text="选择颜色…",
            style="Quiet.TButton",
            command=self._choose_fill_colour,
        )
        self.choose_fill_button.grid(row=0, column=2, sticky="e", padx=(6, 0))
        self.advanced_note = ttk.Label(
            self.advanced_frame,
            text="默认自动多帧最安全；背景/纯色会覆盖整个红框。",
            style="CardSubtle.TLabel",
            justify="left",
        )
        self.advanced_note.grid(
            row=1,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(7, 0),
        )
        self.advanced_frame.columnconfigure(0, weight=1)

        self.confirmed_frame = ttk.LabelFrame(
            self.shell,
            text="已确认页面",
            padding=8,
            style="Card.TLabelframe",
        )
        self.confirmed_frame.pack(fill="x", pady=(8, 0))
        self.action_list = tk.Listbox(
            self.confirmed_frame,
            height=3,
            activestyle="none",
            bg=PANEL_ALT,
            fg=TEXT,
            selectbackground="#E4E9FF",
            selectforeground=ACCENT_DARK,
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            font=("Microsoft YaHei UI", 9),
        )
        self.action_list.pack(side="left", fill="x", expand=True)
        ttk.Button(
            self.confirmed_frame,
            text="删除选中",
            width=12,
            style="RepairQuiet.TButton",
            command=self._delete_action,
        ).pack(side="right", padx=(8, 0))

        self.footer = ttk.Frame(self.shell, style="Panel.TFrame")
        self.footer.pack(fill="x", pady=(10, 0))
        ttk.Button(
            self.footer,
            text="取消",
            width=10,
            style="RepairQuiet.TButton",
            command=self._cancel,
        ).pack(side="right")
        ttk.Button(
            self.footer,
            text="完成并立即生成新 PPT",
            width=24,
            style="Accent.TButton",
            command=self._finish_and_start,
        ).pack(side="right", padx=(0, 8))

        self._refresh_action_list()
        self._mode_changed()
        self.bind("<Configure>", self._schedule_responsive_layout, add="+")
        self.after_idle(self._apply_responsive_layout)
        self.grab_set()
        self.focus_set()

    def _make_canvas(
        self, parent: tk.Misc, source: str, *, selectable: bool = True
    ) -> tk.Canvas:
        canvas = tk.Canvas(
            parent,
            background="#111827",
            width=1,
            height=1,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        canvas.pack(fill="both", expand=True)
        canvas.bind("<Configure>", lambda _event, key=source: self._render(key))
        if selectable:
            canvas.bind(
                "<ButtonPress-1>",
                lambda event, key=source: self._start_selection(event, key),
            )
            canvas.bind(
                "<B1-Motion>",
                lambda event, key=source: self._drag_selection(event, key),
            )
            canvas.bind(
                "<ButtonRelease-1>",
                lambda event, key=source: self._finish_selection(event, key),
            )
        return canvas

    @staticmethod
    def _clock(seconds: float) -> str:
        minutes, remainder = divmod(max(0.0, seconds), 60.0)
        return f"{int(minutes):02d}:{remainder:04.1f}"

    def _page_limit(self, mode: str | None = None) -> int:
        active_mode = mode or self.mode.get()
        return self.slide_count + (1 if active_mode == "insert_page" else 0)

    def _page_value(self, mode: str | None = None) -> int:
        try:
            value = int(self.page_number.get())
        except (TypeError, ValueError):
            value = 1
        return min(max(1, value), self._page_limit(mode))

    def _set_page_number(self, value: int) -> None:
        normalized = str(min(max(1, int(value)), self._page_limit()))
        if self.page_number.get() == normalized:
            return
        self._suspend_page_trace = True
        try:
            self.page_number.set(normalized)
        finally:
            self._suspend_page_trace = False

    def _schedule_page_changed(self, *_args: object) -> None:
        if self._suspend_page_trace or not hasattr(self, "page_canvas"):
            return
        if self._page_change_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._page_change_job)
        self._page_change_job = self.after(140, self._page_changed)

    def _change_page(self, delta: int) -> None:
        self._set_page_number(self._page_value() + int(delta))
        self._page_changed()

    def _remember_current_context(self) -> None:
        if self._active_context is not None and hasattr(self, "video_time"):
            self._context_times[self._active_context] = float(self.video_time.get())

    def _display_page_number(self) -> int:
        return min(self.slide_count, self._page_value())

    def _current_page_image(self) -> Any:
        page = self._display_page_number()
        if page not in self.page_cache:
            self.page_cache[page] = self.engine.render_page_after_actions(
                self.pptx_path,
                self.video_path,
                page,
                self.actions,
            )
        return self.page_cache[page]

    def _schedule_responsive_layout(self, event: tk.Event) -> None:
        if event.widget is not self:
            return
        if hasattr(self, "status_label"):
            self.status_label.configure(
                wraplength=max(420, int(event.width) - 36)
            )
        if self._layout_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._layout_job)
        self._layout_job = self.after(80, self._apply_responsive_layout)

    def _apply_responsive_layout(self) -> None:
        self._layout_job = None
        if not hasattr(self, "previews"):
            return
        physical_width = max(1, self.winfo_width())
        physical_height = max(1, self.winfo_height())
        width = max(1, round(physical_width / self._display_scale))
        height = max(1, round(physical_height / self._display_scale))
        # A two-row layout is useful only when the window has enough height;
        # otherwise three aspect-fitted canvases remain clearer and taller.
        layout = repair_preview_layout(width, height)
        if layout != self._preview_layout:
            for box in (self.page_box, self.video_box, self.result_box):
                box.grid_forget()
            for column in range(3):
                self.previews.columnconfigure(column, weight=0)
            for row in range(2):
                self.previews.rowconfigure(row, weight=0)
            if layout == "two_rows":
                self.previews.columnconfigure(0, weight=1, uniform="repair_preview")
                self.previews.columnconfigure(1, weight=1, uniform="repair_preview")
                self.previews.rowconfigure(0, weight=1, uniform="repair_preview_row")
                self.previews.rowconfigure(1, weight=1, uniform="repair_preview_row")
                self.page_box.grid(
                    row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4)
                )
                self.video_box.grid(
                    row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4)
                )
                self.result_box.grid(
                    row=1,
                    column=0,
                    columnspan=2,
                    sticky="nsew",
                    pady=(4, 0),
                )
            else:
                for column in range(3):
                    self.previews.columnconfigure(
                        column, weight=1, uniform="repair_preview"
                    )
                self.previews.rowconfigure(0, weight=1)
                self.page_box.grid(
                    row=0, column=0, sticky="nsew", padx=(0, 4)
                )
                self.video_box.grid(
                    row=0, column=1, sticky="nsew", padx=4
                )
                self.result_box.grid(
                    row=0, column=2, sticky="nsew", padx=(4, 0)
                )
            self._preview_layout = layout
        self.action_list.configure(
            height=1 if height < 760 else (2 if width < 1180 else 3)
        )
        self.shell.configure(padding=10 if width < 1040 else 14)
        self.status_label.configure(
            wraplength=max(360, physical_width - round(56 * self._display_scale))
        )
        self.advanced_note.configure(
            wraplength=max(320, physical_width - round(90 * self._display_scale))
        )
        self.after_idle(lambda: self._render("page"))
        self.after_idle(lambda: self._render("video"))
        self.after_idle(lambda: self._render("result"))

    def _mode_changed(self) -> None:
        if not hasattr(self, "page_spin"):
            return
        if self._page_change_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._page_change_job)
            self._page_change_job = None
        previous_mode = self._active_mode
        previous_value = self._page_value(previous_mode)
        if previous_mode == "insert_page":
            self._insert_position = previous_value
        else:
            self._existing_page = min(self.slide_count, previous_value)
        self._remember_current_context()
        mode = self.mode.get()
        inserting = mode == "insert_page"
        manual = mode == "replace_page_frame"
        self._active_mode = mode
        self.page_label.configure(text="插到第几页之前" if inserting else "原 PPT 页码")
        self.page_spin.configure(to=self.slide_count + (1 if inserting else 0))
        self._set_page_number(
            self._insert_position if inserting else self._existing_page
        )
        self.position_note.configure(
            text=(
                f"1–{self.slide_count + 1}，末尾为 {self.slide_count + 1}"
                if inserting
                else f"共 {self.slide_count} 页"
            )
        )
        self.selection_label.configure(
            text=(
                "先定位漏页画面；如需清水印，可直接在中间视频画面拖框。"
                if inserting
                else (
                    "在时间轴手动挑选文字最完整、手写最少的一帧；可在中间画面连续拖出多个水印框。"
                    if manual
                    else "请在左侧原 PPT 页面拖框标出水印。"
                )
            )
        )
        self.video_box.configure(
            text=(
                "② 漏页视频画面（可选框水印）"
                if inserting
                else (
                    "② 手动选择的最佳帧（可连续框水印）"
                    if manual
                    else "② 对应视频画面"
                )
            )
        )
        self.result_box.configure(text="③ 处理结果")
        self._page_changed()

    def _page_changed(self) -> None:
        if not hasattr(self, "page_canvas"):
            return
        if self._page_change_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._page_change_job)
            self._page_change_job = None
        mode = self.mode.get()
        selected_value = self._page_value(mode)
        self._set_page_number(selected_value)
        if mode == "insert_page":
            self._insert_position = selected_value
        else:
            self._existing_page = min(self.slide_count, selected_value)
        self._remember_current_context()
        context = (mode, selected_value)
        self._active_context = context
        self._clear_selection(update_status=False)
        page = self._display_page_number()
        inserting = mode == "insert_page"
        manual = mode == "replace_page_frame"
        if inserting:
            if self._page_value() > self.slide_count:
                self.page_box.configure(text="① 插入位置参考：当前末页")
            else:
                self.page_box.configure(text=f"① 将插在原第 {page} 页之前")
            self.status.set("在时间轴找到漏页，拖动后点击“自动校正稳定帧并预览”。")
            remembered = self._context_times.get(context)
            if remembered is not None:
                self._load_video_frame(remembered)
        else:
            self.page_box.configure(
                text=(
                    f"① 原 PPT 第 {page} 页（将被最佳帧替换）"
                    if manual
                    else f"① 原 PPT 第 {page} 页（在此拖框）"
                )
            )
            timestamp = self._context_times.get(context)
            from_report = timestamp is None
            if timestamp is None:
                timestamp = self.engine.companion_report_timestamp(self.pptx_path, page)
            if timestamp is not None:
                self.status.set(
                    (
                        f"已{'先定位' if from_report else '恢复'}第 {page} 页附近：{self._clock(timestamp)}。请自行前后移动，选出最优帧后在中图连续框选水印。"
                        if manual
                        else (
                            f"已根据提取报告自动定位第 {page} 页：{self._clock(timestamp)}。在左图拖框即可。"
                            if from_report
                            else f"已恢复第 {page} 页上次查看的位置：{self._clock(timestamp)}。在左图拖框即可。"
                        )
                    )
                )
                self._load_video_frame(timestamp)
            else:
                self.status.set("未找到相邻提取报告，请在时间轴手动定位这一页的视频画面。")
        self._render("page")
        self._render("video")

    def _schedule_video_seek(self, raw: object) -> None:
        if self._seek_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._seek_job)
        try:
            timestamp = float(raw)
        except (TypeError, ValueError):
            timestamp = float(self.video_time.get())
        self._seek_job = self.after(80, lambda: self._load_video_frame(timestamp))

    def _step_video(self, seconds: float) -> None:
        target = min(max(0.0, float(self.video_time.get()) + seconds), self.duration)
        self.video_time.set(target)
        self._load_video_frame(target)

    def _load_video_frame(self, timestamp: float) -> None:
        self._seek_job = None
        timestamp = min(max(0.0, float(timestamp)), self.duration)
        self.capture.set(self.cv2.CAP_PROP_POS_MSEC, timestamp * 1000.0)
        ok, frame = self.capture.read()
        if not ok:
            return
        self.current_video_frame = self.engine.align_video_frame(
            frame, self.target_width, self.target_height
        )
        self.video_time.set(timestamp)
        if self._active_context is not None:
            self._context_times[self._active_context] = timestamp
        self.time_label.configure(
            text=f"{self._clock(timestamp)} / {self._clock(self.duration)}"
        )
        self._invalidate_preview()
        self._render("video")

    def _choose_stable_frame(self) -> bool:
        try:
            stable = self.engine.find_stable_aligned_frame(
                self.video_path,
                float(self.video_time.get()),
                self.target_width,
                self.target_height,
                radius_seconds=1.5,
            )
        except (DocuForgeError, OSError, RuntimeError, ValueError) as exc:
            messagebox.showerror("无法选择稳定帧", str(exc), parent=self)
            return False
        self.current_video_frame = stable.image
        self.video_time.set(stable.timestamp)
        self.time_label.configure(
            text=f"{self._clock(stable.timestamp)} / {self._clock(self.duration)}"
        )
        self._invalidate_preview()
        self._render("video")
        return True

    def _source_image(self, source: str) -> Any | None:
        if source == "page":
            return self._current_page_image()
        if source == "video":
            return self.current_video_frame
        if source == "result":
            return self.result_preview
        return None

    def _canvas_for(self, source: str) -> tk.Canvas:
        return {
            "page": self.page_canvas,
            "video": self.video_canvas,
            "result": self.result_canvas,
        }[source]

    def _render(self, source: str) -> None:
        if not hasattr(self, "page_canvas"):
            return
        image = self._source_image(source)
        canvas = self._canvas_for(source)
        canvas.delete("all")
        if image is None:
            canvas.create_text(
                max(1, canvas.winfo_width()) // 2,
                max(1, canvas.winfo_height()) // 2,
                text="等待预览",
                fill="#E5E7EB",
            )
            return
        height, width = image.shape[:2]
        rect = fitted_media_rect(
            width,
            height,
            max(1, canvas.winfo_width()),
            max(1, canvas.winfo_height()),
        )
        self.display_rects[source] = rect
        left, top, right, bottom = rect
        rgb = self.cv2.cvtColor(image, self.cv2.COLOR_BGR2RGB)
        pil = Image.fromarray(rgb).resize(
            (right - left, bottom - top), Image.Resampling.LANCZOS
        )
        photo = ImageTk.PhotoImage(pil)
        self.photos[source] = photo
        canvas.create_image(left, top, anchor="nw", image=photo)
        expected = (
            "video"
            if self.mode.get() in {"insert_page", "replace_page_frame"}
            else "page"
        )
        if self.selection is not None and source == expected:
            x, y, box_width, box_height = self.selection
            display_width = right - left
            display_height = bottom - top
            canvas.create_rectangle(
                left + x * display_width / 100.0,
                top + y * display_height / 100.0,
                left + (x + box_width) * display_width / 100.0,
                top + (y + box_height) * display_height / 100.0,
                outline="#FF3B30",
                width=3,
            )
        if source == "video" and self.mode.get() == "replace_page_frame":
            display_width = right - left
            display_height = bottom - top
            for index, region in enumerate(self.watermark_regions, start=1):
                x, y, box_width, box_height = region
                x1 = left + x * display_width / 100.0
                y1 = top + y * display_height / 100.0
                x2 = left + (x + box_width) * display_width / 100.0
                y2 = top + (y + box_height) * display_height / 100.0
                canvas.create_rectangle(
                    x1, y1, x2, y2, outline="#FF3B30", width=3
                )
                canvas.create_text(
                    x1 + 9,
                    y1 + 9,
                    text=str(index),
                    fill="#FFFFFF",
                    anchor="center",
                    font=("Microsoft YaHei UI", 9, "bold"),
                )

    def _start_selection(self, event: tk.Event, source: str) -> None:
        expected = (
            "video"
            if self.mode.get() in {"insert_page", "replace_page_frame"}
            else "page"
        )
        if source != expected:
            return
        self.drag_source = source
        self.drag_start = (float(event.x), float(event.y))

    def _drag_selection(self, event: tk.Event, source: str) -> None:
        if self.drag_source != source or self.drag_start is None:
            return
        region = canvas_selection_to_percent_region(
            self.drag_start,
            (float(event.x), float(event.y)),
            self.display_rects[source],
        )
        if region is not None:
            self.selection = region
            self._render(source)

    def _finish_selection(self, event: tk.Event, source: str) -> None:
        if self.drag_source != source or self.drag_start is None:
            return
        region = canvas_selection_to_percent_region(
            self.drag_start,
            (float(event.x), float(event.y)),
            self.display_rects[source],
        )
        self.drag_source = None
        self.drag_start = None
        if region is None:
            self._clear_selection()
            return
        if self.mode.get() == "replace_page_frame":
            self.watermark_regions.append(region)
            self.selection = None
            self.selection_label.configure(
                text=f"已框选 {len(self.watermark_regions)} 个水印区域；可继续拖框。"
            )
        else:
            self.selection = region
            self.selection_label.configure(
                text=(
                    f"已框选：左 {region[0]:.1f}%、上 {region[1]:.1f}%、"
                    f"宽 {region[2]:.1f}%、高 {region[3]:.1f}%"
                )
            )
        self._invalidate_preview()
        self._render(source)

    def _clear_selection(self, *, update_status: bool = True) -> None:
        self.selection = None
        self.watermark_regions.clear()
        if hasattr(self, "selection_label"):
            self.selection_label.configure(
                text=(
                    "漏页无需清水印时可不框选；需要清理时在中间画面拖框。"
                    if self.mode.get() == "insert_page"
                    else (
                        "可在中间最佳帧画面连续拖出多个水印框。"
                        if self.mode.get() == "replace_page_frame"
                        else "请在左侧原 PPT 页面拖框标出水印。"
                    )
                )
            )
        self._invalidate_preview()
        if update_status and hasattr(self, "status"):
            self.status.set("红框已清除，可以重新框选。")
        if hasattr(self, "page_canvas"):
            self._render("page")
            self._render("video")

    def _undo_last_region(self) -> None:
        if self.mode.get() == "replace_page_frame" and self.watermark_regions:
            self.watermark_regions.pop()
            self.selection_label.configure(
                text=(
                    f"已框选 {len(self.watermark_regions)} 个水印区域；可继续拖框。"
                    if self.watermark_regions
                    else "可在中间最佳帧画面连续拖出多个水印框。"
                )
            )
            self._invalidate_preview()
            self._render("video")
            return
        self._clear_selection()

    def _toggle_advanced(self) -> None:
        if not self.advanced_frame.winfo_manager():
            # The confirmed-action list is intentionally hidden while empty,
            # so it cannot be used as a stable pack anchor.  The footer is
            # always visible and keeps the advanced panel in the right place.
            self.advanced_frame.pack(
                fill="x",
                pady=(8, 0),
                before=self.footer,
            )
            self._advanced_visible = True
            self.advanced_button.configure(text="高级方式 ▴")
        else:
            self.advanced_frame.pack_forget()
            self._advanced_visible = False
            self.advanced_button.configure(text="高级方式 ▾")

    def _choose_fill_colour(self) -> None:
        _rgb, selected = colorchooser.askcolor(
            color=self.fill_colour.get() or "#FFFFFF",
            title="选择框内填充色",
            parent=self,
        )
        if selected:
            self.fill_colour.set(selected)
            self._invalidate_preview()

    def _build_action(self) -> dict[str, Any]:
        kind = self.mode.get()
        if kind == "repair_region" and self.selection is None:
            raise ValueError("请先在左侧原 PPT 页面拖框标出需要修复的区域")
        action: dict[str, Any] = {
            "kind": kind,
            "timestamp": round(float(self.video_time.get()), 3),
            "method": self._METHOD_LABELS.get(self.method.get(), "temporal"),
            "colour": self.fill_colour.get() or "#FFFFFF",
        }
        if kind == "repair_region":
            action["page"] = self._page_value()
        elif kind == "replace_page_frame":
            action["page"] = self._page_value()
            action["regions"] = [
                [round(value, 4) for value in region]
                for region in self.watermark_regions
            ]
        else:
            action["position"] = self._page_value()
        if self.selection is not None:
            action["region"] = [round(value, 4) for value in self.selection]
        return action

    @staticmethod
    def _action_signature(action: dict[str, Any]) -> str:
        return json.dumps(
            action, ensure_ascii=False, sort_keys=True, separators=(",", ":")
        )

    def _invalidate_preview(self) -> None:
        self._last_preview_signature = ""
        self.result_preview = None
        if hasattr(self, "primary_button"):
            self.primary_button.configure(
                text=(
                    "使用当前手选帧并预览"
                    if self.mode.get() == "replace_page_frame"
                    else "自动校正稳定帧并预览"
                )
            )
        if hasattr(self, "result_canvas"):
            self._render("result")

    def _primary_action(self) -> None:
        try:
            action = self._build_action()
        except ValueError as exc:
            messagebox.showinfo("请先框选", str(exc), parent=self)
            return
        signature = self._action_signature(action)
        if self._last_preview_signature == signature and self.result_preview is not None:
            self.actions.append(action)
            if action["kind"] in {"repair_region", "replace_page_frame"}:
                self.page_cache[int(action["page"])] = self.result_preview.copy()
            self._refresh_action_list()
            self._clear_selection(update_status=False)
            self.status.set(
                "本页已确认，左侧已同步为最新结果。可继续处理本页、切换功能或选择其他页。"
            )
            return
        if action["kind"] != "replace_page_frame" and not self._choose_stable_frame():
            return
        try:
            action = self._build_action()
            _before, after = self.engine.preview_action(
                self.pptx_path,
                self.video_path,
                action,
                base_page=(
                    self._current_page_image()
                    if action["kind"] == "repair_region"
                    else None
                ),
            )
        except (DocuForgeError, OSError, RuntimeError, ValueError) as exc:
            messagebox.showerror("无法生成补修预览", str(exc), parent=self)
            return
        self.result_preview = after
        self._last_preview_signature = self._action_signature(action)
        self.result_box.configure(
            text=(
                "③ 修复后页面"
                if action["kind"] == "repair_region"
                else (
                    "③ 最佳帧重建结果"
                    if action["kind"] == "replace_page_frame"
                    else "③ 将要插入的页面"
                )
            )
        )
        self._render("result")
        self.primary_button.configure(text="确认此页并继续")
        self.status.set("请核对右侧结果；满意后再次点击“确认此页并继续”。")

    def _action_description(self, action: dict[str, Any]) -> str:
        if action.get("kind") == "repair_region":
            return (
                f"修复原第 {action.get('page')} 页 · "
                f"{self._clock(float(action.get('timestamp', 0.0)))}"
            )
        if action.get("kind") == "replace_page_frame":
            return (
                f"最佳帧替换原第 {action.get('page')} 页 · "
                f"{len(action.get('regions', []))} 个水印框 · "
                f"{self._clock(float(action.get('timestamp', 0.0)))}"
            )
        return (
            f"插到第 {action.get('position')} 页 · "
            f"{self._clock(float(action.get('timestamp', 0.0)))}"
        )

    def _refresh_action_list(self) -> None:
        if not hasattr(self, "action_list"):
            return
        self.action_list.delete(0, "end")
        for index, action in enumerate(self.actions, start=1):
            self.action_list.insert("end", f"{index}. {self._action_description(action)}")
        if not hasattr(self, "confirmed_frame"):
            return
        if self.actions:
            if not self.confirmed_frame.winfo_manager():
                self.confirmed_frame.pack(
                    fill="x",
                    pady=(8, 0),
                    before=self.footer,
                )
        else:
            self.confirmed_frame.pack_forget()

    def _delete_action(self) -> None:
        selected = list(self.action_list.curselection())
        if not selected:
            return
        affected_pages: set[int] = set()
        for index in reversed(selected):
            if 0 <= index < len(self.actions):
                action = self.actions[index]
                if action.get("kind") in {"repair_region", "replace_page_frame"}:
                    affected_pages.add(int(action.get("page", 0)))
                del self.actions[index]
        for page in affected_pages:
            self.page_cache.pop(page, None)
        self._refresh_action_list()
        self._invalidate_preview()
        self._render("page")
        self.status.set("已删除选中的确认项，页面预览已按剩余操作重新同步。")

    def _finish_and_start(self) -> None:
        # If the user is looking at a valid right-side preview and presses the
        # final button directly, include that page automatically.  Requiring a
        # second confirmation click here would recreate the old, cumbersome
        # workflow without adding safety because the preview is already exact.
        with contextlib.suppress(ValueError):
            current = self._build_action()
            if (
                self.result_preview is not None
                and self._last_preview_signature == self._action_signature(current)
            ):
                self.actions.append(current)
                self._refresh_action_list()
                self._last_preview_signature = ""
        if not self.actions:
            messagebox.showinfo("尚未确认页面", "请至少预览并确认一个修复页或漏页。", parent=self)
            return
        self.result = self.engine.make_plan(self.actions)
        self.start_requested = True
        self._close()

    def _cancel(self) -> None:
        self.result = None
        self.start_requested = False
        self._close()

    def _close(self) -> None:
        if self._seek_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._seek_job)
            self._seek_job = None
        if self._page_change_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._page_change_job)
            self._page_change_job = None
        if self._layout_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._layout_job)
            self._layout_job = None
        self.capture.release()
        with contextlib.suppress(tk.TclError):
            self.grab_release()
        self.destroy()


def responsive_wraplength(
    available_width: int,
    *,
    reserved_width: int = 0,
    padding: int = 16,
    minimum: int = 280,
) -> int:
    """Return a readable wrap width derived from the current container width."""

    width = max(0, int(available_width))
    reserved = max(0, int(reserved_width))
    inset = max(0, int(padding))
    floor = max(1, int(minimum))
    return max(floor, width - reserved - inset)


def mousewheel_scroll_units(delta: int | float) -> int:
    """Translate a wheel/touchpad delta into smooth Canvas scroll units."""

    try:
        value = float(delta)
    except (TypeError, ValueError, OverflowError):
        return 0
    if not math.isfinite(value) or value == 0:
        return 0
    magnitude = max(1, int(round(abs(value) / 120 * 3)))
    return -magnitude if value > 0 else magnitude


def ease_out_cubic(value: float) -> float:
    """Return a restrained ease-out curve used by lightweight UI motion."""

    progress = min(1.0, max(0.0, float(value)))
    return 1.0 - (1.0 - progress) ** 3


def interpolate_hex_colour(start: str, end: str, progress: float) -> str:
    """Interpolate two ``#RRGGBB`` colours without external UI dependencies."""

    eased = min(1.0, max(0.0, float(progress)))
    start_channels = tuple(int(start[index : index + 2], 16) for index in (1, 3, 5))
    end_channels = tuple(int(end[index : index + 2], 16) for index in (1, 3, 5))
    channels = tuple(
        round(first + (second - first) * eased)
        for first, second in zip(start_channels, end_channels)
    )
    return "#" + "".join(f"{channel:02X}" for channel in channels)


def responsive_layout_mode(window_width: int) -> str:
    """Select a three-tier layout that keeps controls useful at every width."""

    width = max(0, int(window_width))
    if width < 980:
        return "narrow"
    if width < 1320:
        return "compact"
    return "wide"


def catalog_sidebar_width(
    window_width: int,
    *,
    mode: str | None = None,
    preferred_width: int = 420,
    user_width: int | None = None,
) -> int:
    """Return a readable catalog width without starving the main work area."""

    width = max(1, int(window_width))
    layout_mode = mode or responsive_layout_mode(width)
    desired = max(1, int(user_width if user_width is not None else preferred_width))
    if layout_mode == "narrow":
        return max(240, width - 18)
    if layout_mode == "compact":
        minimum = 252
        maximum = min(360, max(minimum, width - 700), max(minimum, round(width * 0.31)))
    else:
        minimum = 285
        maximum = min(480, max(minimum, width - 820), max(minimum, round(width * 0.31)))
    return min(maximum, max(minimum, desired))


def initial_window_size(
    screen_width: int,
    screen_height: int,
) -> tuple[int, int]:
    """Choose a spacious default without overflowing small laptop displays."""

    screen_width = max(1, int(screen_width))
    screen_height = max(1, int(screen_height))
    width = min(screen_width, max(760, min(1480, screen_width - 96)))
    height = min(screen_height, max(560, min(920, screen_height - 96)))
    return width, height


def fitted_dialog_size(
    screen_width: int,
    screen_height: int,
    *,
    preferred_width: int,
    preferred_height: int,
    minimum_width: int,
    minimum_height: int,
    horizontal_margin: int = 72,
    vertical_margin: int = 88,
) -> tuple[int, int]:
    """Fit a logical-pixel dialog size to the display without tiny controls."""

    screen_width = max(1, int(screen_width))
    screen_height = max(1, int(screen_height))
    available_width = max(1, screen_width - max(0, int(horizontal_margin)))
    available_height = max(1, screen_height - max(0, int(vertical_margin)))
    width = min(max(1, int(preferred_width)), available_width)
    height = min(max(1, int(preferred_height)), available_height)
    width = min(screen_width, max(minimum_width, width))
    height = min(screen_height, max(minimum_height, height))
    return width, height


def repair_preview_layout(window_width: int, window_height: int) -> str:
    """Choose a complete three-preview arrangement for the repair workbench."""

    width = max(0, int(window_width))
    height = max(0, int(window_height))
    return "two_rows" if width < 1180 and height >= 760 else "three_columns"


def operation_description_text(
    description: str,
    notes: str,
    engine: str,
    reason: str,
    *,
    compact: bool,
    expanded: bool,
) -> str:
    """Keep the task header short while leaving precision details reachable."""

    summary = str(description).strip() or "暂无功能说明。"
    detail = summary
    if str(notes).strip():
        detail += f"\n精度说明：{str(notes).strip()}"
    detail += f"\n引擎：{str(engine).strip()} — {str(reason).strip()}"
    if expanded:
        return f"{detail}\n点击收起精度与引擎说明 ▴"
    return f"{summary}\n点击查看精度与引擎说明 ▾"


def progress_percent(value: object) -> float:
    """Convert a runner fraction to a finite percentage suitable for Tk."""

    try:
        fraction = float(value)
    except (TypeError, ValueError, OverflowError):
        return 0.0
    if not math.isfinite(fraction):
        return 0.0
    return min(100.0, max(0.0, fraction * 100.0))


def progress_status_text(
    message: str,
    elapsed_seconds: float,
    *,
    seconds_since_update: float = 0.0,
) -> str:
    """Add honest elapsed/activity feedback without inventing fake progress."""

    elapsed = max(0, int(elapsed_seconds))
    status = str(message).strip() or "阶段：处理中"
    status += f" · 已用时 {elapsed} 秒"
    if seconds_since_update >= 8:
        status += " · 当前阶段仍在运行"
    return status


def _enable_windows_dpi_awareness() -> None:
    """Ask Windows for crisp per-monitor rendering before Tk creates a window."""

    if os.name != "nt":
        return
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


class TaskResultDialog(tk.Toplevel):
    """Responsive result summary with success/partial/failure colour states."""

    def __init__(
        self,
        parent: tk.Misc,
        result: TaskResult,
        *,
        output_dir: str | Path,
        operation_name: str = "文件处理",
    ) -> None:
        super().__init__(parent)
        self.result = result
        self.output_dir = Path(output_dir).expanduser()
        self.output_paths = tuple(Path(path).expanduser() for path in result.outputs)
        self.presentation = task_result_presentation(result)
        self.title(self.presentation.title)
        self.configure(bg=BG)
        self.minsize(540, 460)
        self.resizable(True, True)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        try:
            parent.update_idletasks()
            width = min(760, max(580, int(parent.winfo_width() * 0.62)))
            height = min(680, max(500, int(parent.winfo_height() * 0.70)))
            left = parent.winfo_rootx() + max(0, (parent.winfo_width() - width) // 2)
            top = parent.winfo_rooty() + max(0, (parent.winfo_height() - height) // 2)
            self.geometry(f"{width}x{height}+{left}+{top}")
        except tk.TclError:
            self.geometry("660x560")

        shell = tk.Frame(self, bg=BG, padx=18, pady=18)
        shell.pack(fill="both", expand=True)
        card = tk.Frame(
            shell,
            bg=PANEL,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        card.pack(fill="both", expand=True)

        self.header = tk.Canvas(card, height=154, highlightthickness=0, bd=0)
        self.header.pack(fill="x")
        self.header.bind("<Configure>", self._draw_header)
        self.header.create_oval(
            34,
            40,
            76,
            82,
            fill="white",
            outline="white",
            tags=("header_icon_circle",),
        )
        self.header.create_text(
            55,
            61,
            anchor="center",
            text=self.presentation.icon,
            fill=self.presentation.accent,
            font=("Microsoft YaHei UI", 18, "bold"),
            tags=("header_icon",),
        )
        self.header.create_text(
            94,
            34,
            anchor="nw",
            text=self.presentation.title,
            fill="white",
            font=("Microsoft YaHei UI", 18, "bold"),
            tags=("header_title",),
        )
        self.header.create_text(
            95,
            86,
            anchor="nw",
            width=500,
            text=self.presentation.subtitle,
            fill="#FFFFFF",
            font=("Microsoft YaHei UI", 9),
            tags=("header_subtitle",),
        )

        body = tk.Frame(card, bg=PANEL, padx=22, pady=18)
        body.pack(fill="both", expand=True)
        operation_label = tk.Label(
            body,
            text=operation_name,
            bg=PANEL,
            fg=TEXT,
            anchor="w",
            justify="left",
            wraplength=620,
            font=("Microsoft YaHei UI", 11, "bold"),
        )
        operation_label.pack(fill="x")
        operation_label.bind(
            "<Configure>",
            lambda event: operation_label.configure(
                wraplength=max(260, event.width - 4)
            ),
        )

        stats = tk.Frame(body, bg=PANEL)
        stats.pack(fill="x", pady=(12, 12))
        completed = len(result.completed_inputs)
        failed = len(result.failed_inputs) + len(result.cancelled_inputs)
        unfinished_color = (
            WARNING if result.cancelled and not result.failed_inputs else DANGER
        )
        stat_values = (
            ("已完成输入", str(completed), SUCCESS),
            ("未完成输入", str(failed), unfinished_color if failed else MUTED),
            ("生成文件", str(len(result.outputs)), ACCENT),
        )
        for column, (label, value, color) in enumerate(stat_values):
            stats.grid_columnconfigure(column, weight=1, uniform="result-stat")
            panel = tk.Frame(
                stats,
                bg=PANEL_ALT,
                highlightthickness=1,
                highlightbackground=BORDER,
                padx=12,
                pady=10,
            )
            panel.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=(0 if column == 0 else 5, 0 if column == 2 else 5),
            )
            tk.Label(
                panel,
                text=value,
                bg=PANEL_ALT,
                fg=color,
                font=("Microsoft YaHei UI", 17, "bold"),
            ).pack(anchor="w")
            tk.Label(
                panel,
                text=label,
                bg=PANEL_ALT,
                fg=MUTED,
                font=("Microsoft YaHei UI", 9),
            ).pack(anchor="w", pady=(2, 0))

        location = tk.Label(
            body,
            text=f"输出位置：{self.output_dir}",
            bg=PANEL,
            fg=MUTED,
            anchor="w",
            justify="left",
            wraplength=620,
            font=("Microsoft YaHei UI", 9),
        )
        location.pack(fill="x", pady=(0, 10))
        location.bind(
            "<Configure>",
            lambda event: location.configure(wraplength=max(220, event.width - 4)),
        )

        details = self._detail_lines()
        if details:
            tk.Label(
                body,
                text=("停止情况与提示" if result.cancelled else "未完成文件与提示"),
                bg=PANEL,
                fg=TEXT,
                anchor="w",
                font=("Microsoft YaHei UI", 10, "bold"),
            ).pack(fill="x", pady=(2, 6))
            detail_shell = tk.Frame(
                body,
                bg=PANEL_ALT,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            detail_shell.pack(fill="both", expand=True)
            scrollbar = ttk.Scrollbar(detail_shell, orient="vertical")
            scrollbar.pack(side="right", fill="y")
            self.detail_text = tk.Text(
                detail_shell,
                height=7,
                wrap="word",
                bd=0,
                padx=12,
                pady=10,
                bg=PANEL_ALT,
                fg=TEXT,
                font=("Microsoft YaHei UI", 9),
                yscrollcommand=scrollbar.set,
            )
            self.detail_text.pack(side="left", fill="both", expand=True)
            scrollbar.configure(command=self.detail_text.yview)
            self.detail_text.insert("1.0", "\n\n".join(details))
            self.detail_text.tag_configure(
                "all",
                spacing1=2,
                spacing3=5,
                lmargin1=2,
                lmargin2=18,
            )
            self.detail_text.tag_add("all", "1.0", "end")
            self.detail_text.configure(state="disabled")
        else:
            tk.Label(
                body,
                text="所有输入均已按预期处理完成。",
                bg=PANEL,
                fg=SUCCESS,
                anchor="w",
                font=("Microsoft YaHei UI", 10),
            ).pack(fill="both", expand=True, pady=(8, 0))

        buttons = tk.Frame(body, bg=PANEL)
        buttons.pack(fill="x", pady=(16, 0))
        if result.failed_inputs or result.cancelled_inputs or result.warnings:
            self._result_button(
                buttons,
                text="复制未完成清单",
                command=self._copy_details,
            ).pack(side="left")
        self._result_button(
            buttons,
            text="关闭",
            command=self.destroy,
            primary=True,
        ).pack(side="right")
        if self.output_paths:
            self._result_button(
                buttons,
                text="打开文件" if len(self.output_paths) == 1 else "打开首个文件",
                command=self._open_primary_output,
            ).pack(side="right", padx=(0, 8))
        self._result_button(
            buttons,
            text="打开文件夹",
            command=self._open_output,
        ).pack(side="right", padx=(0, 8))

        self.after(20, self._fade_in, 0)
        self.after(60, self._activate_modal)

    def _result_button(
        self,
        parent: tk.Misc,
        *,
        text: str,
        command: Callable[[], None],
        primary: bool = False,
    ) -> tk.Button:
        background = self.presentation.accent if primary else PANEL_ALT
        foreground = "white" if primary else TEXT
        active_background = (
            self.presentation.gradient_start if primary else ACCENT_SOFT
        )
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=background,
            fg=foreground,
            activebackground=active_background,
            activeforeground=foreground,
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=16,
            pady=8,
            highlightthickness=1,
            highlightbackground=(background if primary else BORDER),
            highlightcolor=(background if primary else BORDER),
            font=("Microsoft YaHei UI", 9, "bold" if primary else "normal"),
        )

    @staticmethod
    def _hex_channels(value: str) -> tuple[int, int, int]:
        text = value.lstrip("#")
        return tuple(int(text[index : index + 2], 16) for index in (0, 2, 4))  # type: ignore[return-value]

    def _draw_header(self, event: tk.Event[tk.Canvas]) -> None:
        width = max(1, int(event.width))
        height = max(1, int(event.height))
        start = self._hex_channels(self.presentation.gradient_start)
        end = self._hex_channels(self.presentation.gradient_end)
        self.header.delete("gradient")
        steps = max(2, min(160, width // 4))
        for index in range(steps):
            ratio = index / max(1, steps - 1)
            channels = tuple(
                round(left + (right - left) * ratio)
                for left, right in zip(start, end)
            )
            color = "#" + "".join(f"{channel:02x}" for channel in channels)
            left = round(width * index / steps)
            right = round(width * (index + 1) / steps) + 1
            self.header.create_rectangle(
                left,
                0,
                right,
                height,
                fill=color,
                outline=color,
                tags=("gradient",),
            )
        self.header.tag_lower("gradient")
        self.header.itemconfigure("header_subtitle", width=max(260, width - 142))

    def _detail_lines(self) -> list[str]:
        lines = [
            f"{index}. {failure.input_path}\n   {failure.message}"
            for index, failure in enumerate(self.result.failed_inputs, start=1)
        ]
        start = len(lines) + 1
        lines.extend(
            f"{index}. {path}\n   用户已停止任务，该文件未完成处理。"
            for index, path in enumerate(self.result.cancelled_inputs, start=start)
        )
        lines.extend(f"提示：{warning}" for warning in self.result.warnings)
        return lines

    def _copy_details(self) -> None:
        details = self._detail_lines()
        if not details:
            return
        try:
            self.clipboard_clear()
            self.clipboard_append("\n\n".join(details))
            self.update_idletasks()
        except tk.TclError:
            return

    def _open_output(self) -> None:
        existing_output = next(
            (path for path in self.output_paths if path.exists()),
            None,
        )
        target_dir = (
            existing_output
            if existing_output is not None and existing_output.is_dir()
            else existing_output.parent
            if existing_output is not None
            else self.output_dir
        )
        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            os.startfile(target_dir)  # type: ignore[attr-defined]
        except (OSError, ValueError) as exc:
            messagebox.showerror("无法打开输出文件夹", str(exc), parent=self)

    def _open_primary_output(self) -> None:
        if not self.output_paths:
            return
        target = next(
            (path for path in self.output_paths if path.exists()),
            self.output_paths[0],
        )
        try:
            if not target.exists():
                raise FileNotFoundError(f"生成文件不存在或已被移动：{target}")
            os.startfile(target)  # type: ignore[attr-defined]
        except (OSError, ValueError) as exc:
            messagebox.showerror("无法打开生成文件", str(exc), parent=self)

    def _fade_in(self, step: int) -> None:
        if not self.winfo_exists():
            return
        try:
            if step == 0:
                self.attributes("-alpha", 0.88)
            alpha = min(1.0, 0.88 + step * 0.02)
            self.attributes("-alpha", alpha)
            if alpha < 1.0:
                self.after(18, self._fade_in, step + 1)
        except tk.TclError:
            pass

    def _activate_modal(self) -> None:
        try:
            self.grab_set()
            self.focus_force()
        except tk.TclError:
            pass


class DocuForgeApp(_TkBase):
    def __init__(self, operations: list[Operation]) -> None:
        super().__init__()
        self.title("页织工坊 · LayoutLoom")
        screen_width = max(1, int(self.winfo_screenwidth()))
        screen_height = max(1, int(self.winfo_screenheight()))
        try:
            self._display_scale = max(1.0, float(self.winfo_fpixels("1i")) / 96.0)
        except (tk.TclError, TypeError, ValueError):
            self._display_scale = 1.0
        logical_screen_width = max(1, round(screen_width / self._display_scale))
        logical_screen_height = max(1, round(screen_height / self._display_scale))
        logical_width, logical_height = initial_window_size(
            logical_screen_width,
            logical_screen_height,
        )
        initial_width = max(1, round(logical_width * self._display_scale))
        initial_height = max(1, round(logical_height * self._display_scale))
        self._initial_window_width = logical_width
        left = max(0, (screen_width - initial_width) // 2)
        top = max(0, (screen_height - initial_height) // 2)
        self.geometry(f"{initial_width}x{initial_height}+{left}+{top}")
        self.minsize(
            min(screen_width, round(760 * self._display_scale)),
            min(screen_height, round(560 * self._display_scale)),
        )
        self.configure(bg=BG)

        self.operations = operations
        self.operation_by_id = {item.id: item for item in operations}
        self.current_operation: Operation | None = None
        self.input_paths: list[Path] = []
        self.param_vars: dict[str, tk.Variable] = {}
        self.choice_maps: dict[str, dict[str, str]] = {}
        self.show_unavailable_var = tk.BooleanVar(value=False)
        self.worker_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.runner: TaskRunner | None = None
        self.worker: threading.Thread | None = None
        self.active_operation: Operation | None = None
        self.active_output_dir: str | None = None
        self.active_inputs: list[Path] = []
        self._closing = False
        self._setup_scroll_refresh_job: str | None = None
        self._drop_hint_reset_job: str | None = None
        self._progress_animation_job: str | None = None
        self._title_animation_job: str | None = None
        self._pending_input_scans = 0
        self._style_images: list[ImageTk.PhotoImage] = []
        self._brand_icon: ImageTk.PhotoImage | None = None
        self._sidebar_expanded = False
        self._sidebar_user_width: int | None = None
        self._sidebar_drag_start_x = 0
        self._sidebar_drag_start_width = 0
        self._catalog_preferred_width = 380
        self._catalog_tree_content_width = 360
        self.parameter_hint_labels: list[ttk.Label] = []
        self.parameter_rows: list[
            tuple[
                ParameterSpec,
                tk.Misc,
                tk.Misc,
                tk.Misc,
                tk.Misc | None,
                ttk.Label | None,
            ]
        ] = []
        self.parameter_section_frames: dict[str, ttk.LabelFrame] = {}
        self.parameter_section_order: list[str] = []
        self.advanced_parameters_frame: ttk.LabelFrame | None = None
        self.advanced_parameters_button: ttk.Button | None = None
        self.advanced_parameters_expanded = False
        self._layout_mode: str | None = None
        self._operation_description_parts = (
            "从左侧选择文档、图片或视频处理任务。",
            "",
            "等待选择",
            "",
        )
        self._operation_details_expanded = False
        self._last_progress_value = 0.0
        self._progress_started_at: float | None = None
        self._progress_last_update_at: float | None = None
        self._progress_base_message = "准备就绪"
        self._last_logged_progress_message = ""
        self._progress_indeterminate = False
        self._image_preview_job: str | None = None
        self._image_preview_generation = 0
        self._image_preview_index = 0
        self._image_preview_original: Image.Image | None = None
        self._image_preview_result: Image.Image | None = None
        self._image_preview_photos: dict[str, ImageTk.PhotoImage] = {}
        self._image_preview_layout = ""
        self._image_preview_display_rects = {
            "original": (0, 0, 0, 0),
            "result": (0, 0, 0, 0),
        }
        self._image_preview_source_size = (1, 1)
        self._direct_image_edit_start: tuple[int, int] | None = None
        self._direct_image_edit_canvas_start: tuple[float, float] | None = None

        self._configure_style()
        self._build_ui()
        self._rebuild_operation_tree()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Configure>", self._on_window_configure, add="+")
        self.after(120, self._poll_worker)
        self.after(1000, self._refresh_progress_elapsed)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        self.option_add("*Font", ("Microsoft YaHei UI", 10))
        self.option_add("*Menu.Font", ("Microsoft YaHei UI", 10))
        style.configure(".", background=BG, foreground=TEXT, bordercolor=BORDER)
        style.configure("Panel.TFrame", background=PANEL)
        style.configure("Soft.TFrame", background=PANEL_ALT)
        style.configure("Sidebar.TFrame", background=SIDEBAR_BG)
        style.configure(
            "Title.TLabel",
            background=PANEL,
            foreground=TEXT,
            font=("Microsoft YaHei UI", 18, "bold"),
        )
        style.configure(
            "Header.TLabel",
            background=BG,
            foreground=TEXT,
            font=("Microsoft YaHei UI", 20, "bold"),
        )
        style.configure(
            "Subtle.TLabel",
            background=PANEL,
            foreground=MUTED,
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "HeaderSubtle.TLabel",
            background=BG,
            foreground=MUTED,
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "SectionTitle.TLabel",
            background=PANEL,
            foreground=TEXT,
            font=("Microsoft YaHei UI", 11, "bold"),
        )
        style.configure(
            "CardField.TLabel",
            background=PANEL_ALT,
            foreground=TEXT,
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "CardSubtle.TLabel",
            background=PANEL_ALT,
            foreground=MUTED,
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "Card.TCheckbutton",
            background=PANEL_ALT,
            foreground=TEXT,
            font=("Microsoft YaHei UI", 9),
        )
        style.map("Card.TCheckbutton", background=[("active", PANEL_ALT)])
        style.configure(
            "DropHint.TLabel",
            background=ACCENT_SOFT,
            foreground="#44527A",
            padding=(16, 13),
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "DropActive.TLabel",
            background="#E0E7FF",
            foreground=ACCENT_DARK,
            padding=(16, 13),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.configure(
            "DropBusy.TLabel",
            background="#FFF7ED",
            foreground=WARNING,
            padding=(16, 13),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.configure(
            "DropError.TLabel",
            background="#FFF1F2",
            foreground=DANGER,
            padding=(16, 13),
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "SidebarTitle.TLabel",
            background=SIDEBAR_BG,
            foreground=TEXT,
            font=("Microsoft YaHei UI", 11, "bold"),
        )
        style.configure(
            "Sidebar.TCheckbutton",
            background=SIDEBAR_BG,
            foreground=MUTED,
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "Accent.TButton",
            background=ACCENT,
            foreground="white",
            borderwidth=0,
            focusthickness=0,
            padding=(20, 10),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[("active", ACCENT_DARK), ("disabled", "#B7C1D8")],
            foreground=[("disabled", "#F5F7FB")],
        )
        style.configure(
            "Quiet.TButton",
            background=PANEL,
            foreground=TEXT,
            bordercolor=BORDER,
            borderwidth=1,
            focusthickness=0,
            padding=(12, 8),
        )
        style.map(
            "Quiet.TButton",
            background=[("active", ACCENT_SOFT), ("pressed", "#E0E7FF")],
            bordercolor=[("active", "#C7D2FE")],
        )
        style.configure(
            "Danger.TButton",
            background="#FFF1F2",
            foreground=DANGER,
            bordercolor="#FECDD3",
            borderwidth=1,
            focusthickness=0,
            padding=(12, 8),
        )
        style.map("Danger.TButton", background=[("active", "#FFE4E6")])
        style.configure(
            "Nav.TButton",
            background=ACCENT_SOFT,
            foreground=ACCENT_DARK,
            bordercolor="#D7DEFF",
            borderwidth=1,
            focusthickness=0,
            padding=(12, 8),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map("Nav.TButton", background=[("active", "#E0E7FF")])
        style.configure(
            "Treeview",
            background=PANEL,
            fieldbackground=PANEL,
            foreground=TEXT,
            borderwidth=0,
            rowheight=32,
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "Treeview.Heading",
            background=PANEL_ALT,
            foreground=MUTED,
            relief="flat",
            padding=(8, 7),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "Treeview",
            background=[("selected", "#E4E9FF")],
            foreground=[("selected", ACCENT_DARK)],
        )
        style.configure(
            "TEntry",
            fieldbackground=PANEL,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            padding=(10, 8),
        )
        style.map("TEntry", bordercolor=[("focus", ACCENT)])
        style.configure(
            "TCombobox",
            fieldbackground=PANEL,
            background=PANEL,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            arrowsize=15,
            padding=(9, 7),
        )
        style.map("TCombobox", bordercolor=[("focus", ACCENT)])
        style.configure(
            "TSpinbox",
            fieldbackground=PANEL,
            bordercolor=BORDER,
            arrowsize=14,
            padding=(9, 7),
        )
        style.configure(
            "TNotebook",
            background=PANEL,
            borderwidth=0,
            tabmargins=(0, 0, 0, 8),
        )
        style.configure(
            "TNotebook.Tab",
            background=PANEL_ALT,
            foreground=MUTED,
            borderwidth=0,
            padding=(18, 9),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", ACCENT_SOFT), ("active", "#F0F3FA")],
            foreground=[("selected", ACCENT_DARK)],
        )
        style.configure(
            "Card.TLabelframe",
            background=PANEL_ALT,
            bordercolor=BORDER,
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=PANEL,
            foreground=TEXT,
            font=("Microsoft YaHei UI", 10, "bold"),
            padding=(6, 0),
        )
        style.configure(
            "Vertical.TScrollbar",
            background="#B8C2D4",
            troughcolor="#EDF1F7",
            borderwidth=0,
            arrowsize=13,
        )
        style.configure(
            "Horizontal.TScrollbar",
            background="#B8C2D4",
            troughcolor="#EDF1F7",
            borderwidth=0,
            arrowsize=13,
        )
        style.configure(
            "Horizontal.TProgressbar",
            background=ACCENT,
            troughcolor="#E8ECF5",
            borderwidth=0,
            thickness=10,
        )
        self._configure_rounded_button_images(style)
        self._configure_checkbutton_images(style)

    def _configure_rounded_button_images(self, style: ttk.Style) -> None:
        """Give the main buttons soft scalable corners with a safe ttk fallback."""

        def image(fill: str, border: str, *, radius: int = 10) -> ImageTk.PhotoImage:
            raster = Image.new("RGBA", (44, 38), (0, 0, 0, 0))
            draw = ImageDraw.Draw(raster)
            draw.rounded_rectangle(
                (1, 1, 42, 36),
                radius=radius,
                fill=fill,
                outline=border,
                width=1,
            )
            rendered = ImageTk.PhotoImage(raster, master=self)
            raster.close()
            self._style_images.append(rendered)
            return rendered

        definitions = (
            (
                "Accent.TButton",
                "DocuForgeAccent.background",
                (ACCENT, ACCENT),
                ("#5D77F0", "#5D77F0"),
                (ACCENT_DARK, ACCENT_DARK),
                ("#B8C2D8", "#B8C2D8"),
            ),
            (
                "Quiet.TButton",
                "DocuForgeQuiet.background",
                (PANEL, BORDER),
                (ACCENT_SOFT, "#C7D2FE"),
                ("#E4E9FF", "#B8C5F7"),
                ("#F4F6FA", "#E5E9F0"),
            ),
            (
                "Danger.TButton",
                "DocuForgeDanger.background",
                ("#FFF1F2", "#FECDD3"),
                ("#FFE4E6", "#FDA4AF"),
                ("#FECDD3", "#FB7185"),
                ("#F7F4F5", "#E8E2E4"),
            ),
            (
                "Nav.TButton",
                "DocuForgeNav.background",
                (ACCENT_SOFT, "#D7DEFF"),
                ("#E0E7FF", "#C7D2FE"),
                ("#D7DEFF", "#AFC0FA"),
                ("#F2F4F8", "#E3E7EF"),
            ),
        )
        for style_name, element_name, normal, active, pressed, disabled in definitions:
            try:
                normal_image = image(*normal)
                active_image = image(*active)
                pressed_image = image(*pressed)
                disabled_image = image(*disabled)
                style.element_create(
                    element_name,
                    "image",
                    normal_image,
                    ("disabled", disabled_image),
                    ("pressed", pressed_image),
                    ("active", active_image),
                    border=(12, 12, 12, 12),
                    sticky="nsew",
                )
                style.layout(
                    style_name,
                    [
                        (
                            element_name,
                            {
                                "sticky": "nsew",
                                "children": [
                                    (
                                        "Button.padding",
                                        {
                                            "sticky": "nsew",
                                            "children": [
                                                ("Button.label", {"sticky": "nsew"})
                                            ],
                                        },
                                    )
                                ],
                            },
                        )
                    ],
                )
            except tk.TclError:
                # Older Tk builds still receive the complete colour, spacing
                # and hover theme configured above.
                continue

    def _configure_checkbutton_images(self, style: ttk.Style) -> None:
        """Replace the ambiguous clam-theme cross with an explicit tick."""

        logical_box = 18
        scale = max(1.0, float(getattr(self, "_display_scale", 1.0)))
        box = max(18, round(logical_box * scale))
        gap = max(5, round(5 * scale))
        width = box + gap
        supersample = 3

        def indicator(
            fill: str,
            border: str,
            *,
            tick: str | None = None,
        ) -> ImageTk.PhotoImage:
            raster = Image.new(
                "RGBA",
                (width * supersample, box * supersample),
                (0, 0, 0, 0),
            )
            draw = ImageDraw.Draw(raster)
            inset = supersample
            right = box * supersample - supersample - 1
            bottom = box * supersample - supersample - 1
            draw.rounded_rectangle(
                (inset, inset, right, bottom),
                radius=max(4, round(box * 0.24)) * supersample,
                fill=fill,
                outline=border,
                width=max(1, round(1.2 * supersample)),
            )
            if tick is not None:
                points = (
                    (round(box * 0.27) * supersample, round(box * 0.52) * supersample),
                    (round(box * 0.43) * supersample, round(box * 0.68) * supersample),
                    (round(box * 0.76) * supersample, round(box * 0.32) * supersample),
                )
                draw.line(
                    points,
                    fill=tick,
                    width=max(3, round(box * 0.12) * supersample),
                    joint="curve",
                )
            rendered_raster = raster.resize(
                (width, box),
                Image.Resampling.LANCZOS,
            )
            rendered = ImageTk.PhotoImage(rendered_raster, master=self)
            raster.close()
            rendered_raster.close()
            self._style_images.append(rendered)
            return rendered

        normal = indicator("#FFFFFF", "#AEB8C8")
        active = indicator(ACCENT_SOFT, ACCENT)
        selected = indicator(ACCENT, ACCENT, tick="#FFFFFF")
        selected_active = indicator("#5D77F0", "#5D77F0", tick="#FFFFFF")
        disabled = indicator("#F2F4F7", "#D5DAE3")
        disabled_selected = indicator("#C7CEDC", "#C7CEDC", tick="#FFFFFF")
        element_name = "DocuForgeCheck.indicator"
        try:
            style.element_create(
                element_name,
                "image",
                normal,
                ("disabled", "selected", disabled_selected),
                ("disabled", disabled),
                ("active", "selected", selected_active),
                ("selected", selected),
                ("active", active),
                sticky="w",
            )
            layout = [
                (
                    "Checkbutton.padding",
                    {
                        "sticky": "nswe",
                        "children": [
                            (element_name, {"side": "left", "sticky": ""}),
                            (
                                "Checkbutton.focus",
                                {
                                    "side": "left",
                                    "sticky": "w",
                                    "children": [
                                        ("Checkbutton.label", {"sticky": "nswe"})
                                    ],
                                },
                            ),
                        ],
                    },
                )
            ]
            for style_name in (
                "TCheckbutton",
                "Card.TCheckbutton",
                "Sidebar.TCheckbutton",
            ):
                style.layout(style_name, layout)
        except tk.TclError:
            # Colour and font styling remain usable on older Tk versions.
            return

    def _create_brand_icon(self, size: int = 42) -> ImageTk.PhotoImage:
        raster = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(raster)
        draw.rounded_rectangle(
            (1, 1, size - 2, size - 2),
            radius=max(8, size // 4),
            fill=ACCENT,
        )
        left = round(size * 0.29)
        top = round(size * 0.20)
        right = round(size * 0.70)
        bottom = round(size * 0.78)
        fold = round(size * 0.13)
        draw.rounded_rectangle(
            (left, top, right, bottom),
            radius=max(2, size // 16),
            fill="#FFFFFF",
        )
        draw.polygon(
            (
                (right - fold, top),
                (right, top + fold),
                (right - fold, top + fold),
            ),
            fill="#DDE4FF",
        )
        line_left = left + round(size * 0.08)
        for ratio in (0.45, 0.58, 0.70):
            y = round(size * ratio)
            draw.rounded_rectangle(
                (line_left, y, right - round(size * 0.07), y + 2),
                radius=1,
                fill="#95A7F7",
            )
        icon = ImageTk.PhotoImage(raster, master=self)
        raster.close()
        return icon

    def _build_ui(self) -> None:
        self.header = ttk.Frame(self, padding=(24, 18, 24, 14))
        self.header.pack(fill="x")
        self.header.columnconfigure(0, weight=1)
        self.title_box = ttk.Frame(self.header)
        self.title_box.grid(row=0, column=0, sticky="ew")
        self.title_box.columnconfigure(1, weight=1)
        self._brand_icon = self._create_brand_icon()
        self.brand_icon_label = ttk.Label(
            self.title_box,
            image=self._brand_icon,
            background=BG,
        )
        self.brand_icon_label.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 12))
        self.header_title = ttk.Label(
            self.title_box, text="页织工坊 · LayoutLoom", style="Header.TLabel"
        )
        self.header_title.grid(row=0, column=1, sticky="sw")
        self.header_subtitle = ttk.Label(
            self.title_box,
            text="本地优先 · 批量处理 · 原文件默认不覆盖 · 按任务选择最合适引擎",
            style="HeaderSubtle.TLabel",
            justify="left",
        )
        self.header_subtitle.grid(row=1, column=1, sticky="nw", pady=(2, 0))
        ready = sum(op.capability().runnable for op in self.operations)
        self.engine_summary = ttk.Label(
            self.header,
            text=f"核心 {len(self.operations)} 项 · 本机可用 {ready} 项",
            background="#EAF8F3",
            foreground=SUCCESS,
            padding=(14, 9),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        self.engine_summary.grid(row=0, column=1, sticky="e", padx=(16, 0))
        self.sidebar_toggle_button = ttk.Button(
            self.header,
            text="☰  工具目录",
            style="Nav.TButton",
            command=self._toggle_sidebar,
        )

        self.wps_notice = tk.Frame(
            self,
            bg="#FFF8E6",
            highlightbackground="#F2C66D",
            highlightthickness=1,
            bd=0,
        )
        self.wps_notice.pack(fill="x", padx=20, pady=(0, 12))
        self.wps_notice.columnconfigure(1, weight=1)
        self.wps_notice_badge = tk.Label(
            self.wps_notice,
            text="WPS 最佳体验",
            bg="#D88916",
            fg="#FFFFFF",
            padx=11,
            pady=5,
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        self.wps_notice_badge.grid(
            row=0,
            column=0,
            sticky="w",
            padx=(12, 12),
            pady=9,
        )
        self.wps_notice_text = tk.Label(
            self.wps_notice,
            text=WPS_COMPATIBILITY_NOTICE,
            bg="#FFF8E6",
            fg="#704B12",
            justify="left",
            anchor="w",
            font=("Microsoft YaHei UI", 9),
        )
        self.wps_notice_text.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(0, 12),
            pady=9,
        )

        self.body = ttk.Frame(self, padding=(20, 0, 20, 20))
        self.body.pack(fill="both", expand=True)
        self.body.columnconfigure(1, weight=1)
        self.body.rowconfigure(0, weight=1)

        self.sidebar_shell = tk.Frame(
            self.body,
            bg=SHADOW,
            padx=1,
            pady=1,
            borderwidth=0,
            width=310,
        )
        self.sidebar_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.sidebar_shell.pack_propagate(False)
        self.sidebar = ttk.Frame(
            self.sidebar_shell,
            style="Sidebar.TFrame",
            padding=14,
        )
        self.sidebar.pack(fill="both", expand=True)
        self._measure_catalog_widths()
        self.sidebar.rowconfigure(3, weight=1)
        self.sidebar.columnconfigure(0, weight=1)
        ttk.Label(self.sidebar, text="处理工具", style="SidebarTitle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 9)
        )
        self.search_var = tk.StringVar()
        search = ttk.Entry(self.sidebar, textvariable=self.search_var)
        search.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        search.insert(0, "搜索功能…")
        search.bind("<FocusIn>", self._clear_search_hint)
        search.bind("<KeyRelease>", lambda _event: self._rebuild_operation_tree())
        self.show_unavailable_check = ttk.Checkbutton(
            self.sidebar,
            text="显示需要额外安装的核心功能",
            variable=self.show_unavailable_var,
            command=self._rebuild_operation_tree,
            style="Sidebar.TCheckbutton",
        )
        self.show_unavailable_check.grid(row=2, column=0, sticky="w", pady=(0, 10))
        self.operation_tree = ttk.Treeview(
            self.sidebar, show="tree", selectmode="browse"
        )
        self.operation_tree.column(
            "#0",
            width=self._catalog_tree_content_width,
            minwidth=180,
            stretch=True,
        )
        self.operation_tree.tag_configure(
            "catalog_root",
            foreground=TEXT,
            font=("Microsoft YaHei UI", 11, "bold"),
        )
        self.operation_tree.tag_configure(
            "catalog_section",
            foreground=MUTED,
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        self.operation_tree.grid(row=3, column=0, sticky="nsew")
        self.operation_tree.bind("<<TreeviewSelect>>", self._on_operation_selected)
        tree_scroll = ttk.Scrollbar(
            self.sidebar, orient="vertical", command=self.operation_tree.yview
        )
        tree_scroll.grid(row=3, column=1, sticky="ns")
        tree_x_scroll = ttk.Scrollbar(
            self.sidebar, orient="horizontal", command=self.operation_tree.xview
        )
        tree_x_scroll.grid(row=4, column=0, sticky="ew", pady=(3, 0))
        self.operation_tree.configure(
            yscrollcommand=tree_scroll.set,
            xscrollcommand=tree_x_scroll.set,
        )
        self.catalog_legend = ttk.Label(
            self.sidebar,
            text="● 内置可用   ◆ 外部引擎已就绪   ○ 需要安装",
            background=SIDEBAR_BG,
            foreground=MUTED,
            font=("Microsoft YaHei UI", 9),
            justify="left",
        )
        self.catalog_legend.grid(row=5, column=0, sticky="w", pady=(10, 0))

        self.sidebar_resize_handle = tk.Frame(
            self.sidebar_shell,
            bg=SHADOW,
            cursor="sb_h_double_arrow",
            width=8,
            borderwidth=0,
        )
        self.sidebar_resize_handle.place(
            relx=1.0,
            x=-8,
            y=0,
            width=8,
            relheight=1.0,
        )
        self.sidebar_resize_handle.bind(
            "<ButtonPress-1>", self._start_sidebar_resize
        )
        self.sidebar_resize_handle.bind("<B1-Motion>", self._drag_sidebar_resize)
        self.sidebar_resize_handle.bind(
            "<ButtonRelease-1>", self._finish_sidebar_resize
        )
        self.sidebar_resize_handle.bind(
            "<Double-Button-1>", self._reset_sidebar_width
        )
        self.sidebar_resize_handle.bind(
            "<Enter>", lambda _event: self.sidebar_resize_handle.configure(bg=ACCENT)
        )
        self.sidebar_resize_handle.bind(
            "<Leave>", lambda _event: self.sidebar_resize_handle.configure(bg=SHADOW)
        )

        self.content_shell = tk.Frame(
            self.body,
            bg=SHADOW,
            padx=1,
            pady=1,
            borderwidth=0,
        )
        self.content_shell.grid(row=0, column=1, sticky="nsew")
        self.content = ttk.Frame(
            self.content_shell,
            style="Panel.TFrame",
            padding=22,
        )
        self.content.pack(fill="both", expand=True)
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(2, weight=1)

        self.operation_header = ttk.Frame(self.content, style="Panel.TFrame")
        self.operation_header.grid(row=0, column=0, sticky="ew")
        self.operation_header.columnconfigure(0, weight=1)
        self.operation_title = ttk.Label(
            self.operation_header,
            text="请选择一个处理功能",
            style="Title.TLabel",
            justify="left",
        )
        self.operation_title.grid(row=0, column=0, sticky="w")
        self.capability_badge = tk.Label(
            self.operation_header,
            text="等待选择",
            bg="#EEF2F6",
            fg=MUTED,
            padx=10,
            pady=5,
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        self.capability_badge.grid(row=0, column=1, sticky="e")
        self.operation_description = ttk.Label(
            self.content,
            text="从左侧选择文档、图片或视频处理任务。",
            style="Subtle.TLabel",
            wraplength=360,
            justify="left",
        )
        self.operation_description.grid(row=1, column=0, sticky="ew", pady=(6, 14))
        self.operation_description.bind(
            "<Button-1>", self._toggle_operation_details, add="+"
        )
        self.content.bind("<Configure>", self._on_content_configure, add="+")

        self.notebook = ttk.Notebook(self.content)
        self.notebook.grid(row=2, column=0, sticky="nsew")
        self.setup_page = ttk.Frame(self.notebook, style="Panel.TFrame")
        self.setup_page.columnconfigure(0, weight=1)
        self.setup_page.rowconfigure(0, weight=1)
        self.setup_canvas = tk.Canvas(
            self.setup_page,
            bg=PANEL,
            highlightthickness=0,
            borderwidth=0,
            yscrollincrement=24,
        )
        self.setup_canvas.grid(row=0, column=0, sticky="nsew")
        self.setup_scroll = tk.Scrollbar(
            self.setup_page,
            orient="vertical",
            command=self.setup_canvas.yview,
            width=18,
            borderwidth=0,
            elementborderwidth=1,
            highlightthickness=0,
            relief="flat",
            background="#94A3B8",
            activebackground="#64748B",
            troughcolor="#E2E8F0",
        )
        self.setup_scroll.grid(row=0, column=1, sticky="ns", padx=(5, 0))
        self.setup_canvas.configure(yscrollcommand=self.setup_scroll.set)
        self.setup_tab = ttk.Frame(
            self.setup_canvas, style="Panel.TFrame", padding=(4, 14, 8, 4)
        )
        self.setup_window = self.setup_canvas.create_window(
            (0, 0), window=self.setup_tab, anchor="nw"
        )
        self.setup_tab.bind("<Configure>", self._on_setup_tab_configure, add="+")
        self.setup_canvas.bind("<Configure>", self._on_setup_canvas_configure, add="+")
        log_tab = ttk.Frame(self.notebook, style="Panel.TFrame", padding=(4, 14, 4, 4))
        self.notebook.add(self.setup_page, text="  文件与参数  ")
        self.notebook.add(log_tab, text="  运行日志  ")
        self.setup_tab.columnconfigure(0, weight=1)
        self.setup_tab.rowconfigure(1, weight=1)
        log_tab.columnconfigure(0, weight=1)
        log_tab.rowconfigure(0, weight=1)

        self.file_toolbar = ttk.Frame(self.setup_tab, style="Panel.TFrame")
        self.file_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.add_file_button = ttk.Button(
            self.file_toolbar,
            text="＋ 添加文件",
            style="Quiet.TButton",
            command=self._add_files,
        )
        self.add_folder_button = ttk.Button(
            self.file_toolbar,
            text="＋ 添加文件夹",
            style="Quiet.TButton",
            command=self._add_folder,
        )
        self.move_up_button = ttk.Button(
            self.file_toolbar,
            text="↑ 上移",
            style="Quiet.TButton",
            command=self._move_up,
        )
        self.move_down_button = ttk.Button(
            self.file_toolbar,
            text="↓ 下移",
            style="Quiet.TButton",
            command=self._move_down,
        )
        self.remove_button = ttk.Button(
            self.file_toolbar,
            text="− 移除选中",
            style="Quiet.TButton",
            command=self._remove_selected,
        )
        self.clear_button = ttk.Button(
            self.file_toolbar,
            text="清空列表",
            style="Danger.TButton",
            command=self._clear_inputs,
        )
        self.file_toolbar_buttons = (
            self.add_file_button,
            self.add_folder_button,
            self.move_up_button,
            self.move_down_button,
            self.remove_button,
            self.clear_button,
        )
        self.file_more_menu = tk.Menu(
            self.file_toolbar,
            tearoff=False,
            bg=PANEL,
            fg=TEXT,
            activebackground=ACCENT_SOFT,
            activeforeground=ACCENT_DARK,
            relief="flat",
            borderwidth=1,
        )
        self.file_more_menu.add_command(label="上移选中文件", command=self._move_up)
        self.file_more_menu.add_command(label="下移选中文件", command=self._move_down)
        self.file_more_menu.add_separator()
        self.file_more_menu.add_command(label="移除选中", command=self._remove_selected)
        self.file_more_menu.add_command(label="清空列表", command=self._clear_inputs)
        self.file_more_button = ttk.Button(
            self.file_toolbar,
            text="更多操作  ⋯",
            style="Quiet.TButton",
            command=self._show_file_more_menu,
        )
        self.file_count_label = ttk.Label(
            self.file_toolbar, text="尚未添加文件", style="Subtle.TLabel"
        )

        self.file_drop_frame = ttk.Frame(self.setup_tab, style="Panel.TFrame")
        self.file_drop_frame.grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="nsew",
        )
        self.file_drop_frame.columnconfigure(0, weight=1)
        self.file_drop_frame.rowconfigure(1, weight=1)
        self.drop_hint_label = ttk.Label(
            self.file_drop_frame,
            text=self._default_drop_hint_text(),
            style="DropHint.TLabel",
            anchor="center",
            justify="center",
        )
        self.drop_hint_label.grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(0, 6),
        )

        self.file_tree = ttk.Treeview(
            self.file_drop_frame,
            columns=("type", "size", "path"),
            show="headings",
            height=6,
        )
        self.file_tree.heading("type", text="格式")
        self.file_tree.heading("size", text="大小")
        self.file_tree.heading("path", text="文件路径")
        self.file_tree.column("type", width=72, anchor="center", stretch=False)
        self.file_tree.column("size", width=90, anchor="e", stretch=False)
        self.file_tree.column("path", width=520, minwidth=180, anchor="w")
        self.file_tree.tag_configure("even", background=PANEL)
        self.file_tree.tag_configure("odd", background=PANEL_ALT)
        self.file_tree.grid(row=1, column=0, sticky="nsew")
        self.file_tree.bind(
            "<<TreeviewSelect>>",
            self._on_file_tree_preview_selected,
            add="+",
        )
        file_scroll = ttk.Scrollbar(
            self.file_drop_frame,
            orient="vertical",
            command=self.file_tree.yview,
        )
        file_scroll.grid(row=1, column=1, sticky="ns")
        file_horizontal_scroll = ttk.Scrollbar(
            self.file_drop_frame,
            orient="horizontal",
            command=self.file_tree.xview,
        )
        file_horizontal_scroll.grid(row=2, column=0, sticky="ew")
        self.file_tree.configure(
            yscrollcommand=file_scroll.set,
            xscrollcommand=file_horizontal_scroll.set,
        )
        self._configure_file_drop_targets()

        self.output_frame = ttk.Frame(self.setup_tab, style="Panel.TFrame")
        self.output_frame.grid(row=2, column=0, sticky="ew", pady=(12, 8))
        self.output_label = ttk.Label(
            self.output_frame,
            text="输出文件夹",
            background=PANEL,
            foreground=TEXT,
        )
        self.output_var = tk.StringVar(
            value=str(Path.home() / "Documents" / "页织工坊 输出")
        )
        self.output_entry = ttk.Entry(self.output_frame, textvariable=self.output_var)
        self.output_browse_button = ttk.Button(
            self.output_frame,
            text="浏览…",
            style="Quiet.TButton",
            command=self._choose_output,
        )
        self.output_open_button = ttk.Button(
            self.output_frame,
            text="打开",
            style="Quiet.TButton",
            command=self._open_output,
        )

        ttk.Separator(self.setup_tab).grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=8
        )
        self.parameters_header = ttk.Frame(self.setup_tab, style="Panel.TFrame")
        self.parameters_header.grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=(0, 8)
        )
        self.parameters_header.columnconfigure(0, weight=1)
        self.parameters_title = ttk.Label(
            self.parameters_header,
            text="处理参数",
            style="SectionTitle.TLabel",
        )
        self.parameters_title.grid(row=0, column=0, sticky="w")
        self.parameters_scroll_hint = ttk.Label(
            self.parameters_header,
            text="滚动查看更多参数 ↓",
            style="Subtle.TLabel",
        )
        self.parameters_frame = ttk.Frame(self.setup_tab, style="Panel.TFrame")
        self.parameters_frame.grid(row=5, column=0, columnspan=2, sticky="ew")
        self.parameters_frame.columnconfigure(1, weight=1)
        self.parameters_frame.bind(
            "<Configure>", self._on_parameters_frame_configure, add="+"
        )

        self.image_preview_frame = ttk.LabelFrame(
            self.setup_tab,
            text="原图 / 效果预览",
            padding=12,
            style="Card.TLabelframe",
        )
        self.image_preview_frame.grid(
            row=6,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(12, 4),
        )
        self.image_preview_frame.grid_remove()
        self.image_preview_toolbar = ttk.Frame(
            self.image_preview_frame,
            style="Soft.TFrame",
            padding=(10, 7),
        )
        self.image_preview_toolbar.grid(row=0, column=0, sticky="ew")
        self.image_preview_toolbar.columnconfigure(1, weight=1)
        self.image_preview_previous = ttk.Button(
            self.image_preview_toolbar,
            text="‹ 上一张",
            style="Quiet.TButton",
            command=lambda: self._change_image_preview(-1),
        )
        self.image_preview_previous.grid(row=0, column=0, sticky="w")
        self.image_preview_status = ttk.Label(
            self.image_preview_toolbar,
            text="添加图片后自动生成对比预览",
            anchor="center",
            justify="center",
            style="CardSubtle.TLabel",
        )
        self.image_preview_status.grid(row=0, column=1, sticky="ew", padx=10)
        self.image_preview_next = ttk.Button(
            self.image_preview_toolbar,
            text="下一张 ›",
            style="Quiet.TButton",
            command=lambda: self._change_image_preview(1),
        )
        self.image_preview_next.grid(row=0, column=2, sticky="e")

        self.image_preview_panes = ttk.Frame(
            self.image_preview_frame,
            style="Soft.TFrame",
        )
        self.image_preview_panes.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.image_preview_frame.columnconfigure(0, weight=1)
        self.image_preview_original_box = ttk.LabelFrame(
            self.image_preview_panes,
            text="原图",
            padding=6,
            style="Card.TLabelframe",
        )
        self.image_preview_result_box = ttk.LabelFrame(
            self.image_preview_panes,
            text="参数效果",
            padding=6,
            style="Card.TLabelframe",
        )
        self.image_preview_original_canvas = tk.Canvas(
            self.image_preview_original_box,
            height=250,
            bg="#E9EEF6",
            highlightthickness=0,
            borderwidth=0,
        )
        self.image_preview_original_canvas.pack(fill="both", expand=True)
        self.image_preview_result_canvas = tk.Canvas(
            self.image_preview_result_box,
            height=250,
            bg="#E9EEF6",
            highlightthickness=0,
            borderwidth=0,
        )
        self.image_preview_result_canvas.pack(fill="both", expand=True)
        self.image_preview_original_canvas.bind(
            "<Configure>",
            lambda _event: self._render_image_preview_canvas("original"),
            add="+",
        )
        self.image_preview_result_canvas.bind(
            "<Configure>",
            lambda _event: self._render_image_preview_canvas("result"),
            add="+",
        )
        self.image_preview_original_canvas.bind(
            "<ButtonPress-1>",
            self._start_direct_image_edit,
            add="+",
        )
        self.image_preview_original_canvas.bind(
            "<B1-Motion>",
            self._drag_direct_image_edit,
            add="+",
        )
        self.image_preview_original_canvas.bind(
            "<ButtonRelease-1>",
            self._finish_direct_image_edit,
            add="+",
        )
        self.image_preview_details = ttk.Label(
            self.image_preview_frame,
            text="",
            style="CardSubtle.TLabel",
            justify="left",
            wraplength=900,
        )
        self.image_preview_details.grid(row=2, column=0, sticky="ew", pady=(8, 0))

        self.log_text = tk.Text(
            log_tab,
            bg="#101828",
            fg="#D0D5DD",
            insertbackground="white",
            relief="flat",
            padx=12,
            pady=12,
            font=("Consolas", 9),
            wrap="word",
            state="disabled",
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(
            log_tab, orient="vertical", command=self.log_text.yview
        )
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self.footer = ttk.Frame(self.content, style="Panel.TFrame")
        self.footer.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        self.footer.columnconfigure(0, weight=1)
        self.progress_box = ttk.Frame(self.footer, style="Panel.TFrame")
        self.progress_box.grid(row=0, column=0, sticky="ew", padx=(0, 14))
        self.progress_box.columnconfigure(0, weight=1)
        self.progress_var = tk.DoubleVar(value=0)
        self.progressbar = ttk.Progressbar(
            self.progress_box, variable=self.progress_var, maximum=100
        )
        self.progressbar.grid(row=0, column=0, sticky="ew")
        self.progress_percent_label = ttk.Label(
            self.progress_box,
            text="0%",
            style="Subtle.TLabel",
            width=6,
            anchor="e",
        )
        self.progress_percent_label.grid(row=0, column=1, padx=(8, 0), sticky="e")
        self.progress_label = ttk.Label(
            self.progress_box,
            text="准备就绪",
            style="Subtle.TLabel",
            justify="left",
        )
        self.progress_label.grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0)
        )
        self.cancel_button = ttk.Button(
            self.footer,
            text="停止任务",
            style="Quiet.TButton",
            command=self._cancel,
            state="disabled",
        )
        self.run_button = ttk.Button(
            self.footer,
            text="开始处理  →",
            style="Accent.TButton",
            command=self._start,
            state="disabled",
        )

        self._setup_scroll_bindtag = f"DocuForgeSetupScroll{id(self)}"
        self.bind_class(
            self._setup_scroll_bindtag,
            "<MouseWheel>",
            self._on_setup_mousewheel,
        )
        self.bind_class(
            self._setup_scroll_bindtag,
            "<Button-4>",
            self._on_setup_button_scroll,
        )
        self.bind_class(
            self._setup_scroll_bindtag,
            "<Button-5>",
            self._on_setup_button_scroll,
        )
        self._attach_setup_scroll_bindtag(self.setup_page)
        self._apply_responsive_layout(self._initial_window_width, force=True)
        self._schedule_setup_scroll_refresh()

    def _on_window_configure(self, event: tk.Event) -> None:
        if event.widget is not self:
            return
        self._apply_responsive_layout(self._logical_window_width(int(event.width)))

    def _logical_window_width(self, physical_width: int) -> int:
        return max(1, round(max(1, int(physical_width)) / self._display_scale))

    def _measure_catalog_widths(self) -> None:
        """Measure real Chinese labels so the catalog opens at a useful width."""

        normal_font = tkfont.Font(self, family="Microsoft YaHei UI", size=10)
        section_font = tkfont.Font(
            self, family="Microsoft YaHei UI", size=10, weight="bold"
        )
        root_font = tkfont.Font(
            self, family="Microsoft YaHei UI", size=11, weight="bold"
        )
        content_width = 220
        roots: set[str] = set()
        sections: set[str] = set()
        for operation in self.operations:
            root_name, section_name = operation_catalog_path(operation)
            roots.add(root_name)
            sections.add(section_name)
            content_width = max(
                content_width,
                68 + normal_font.measure(f"●  {operation.name}"),
            )
        for root_name in roots:
            content_width = max(content_width, 26 + root_font.measure(root_name))
        for section_name in sections:
            content_width = max(
                content_width,
                47 + section_font.measure(section_name),
            )
        self._catalog_tree_content_width = min(680, content_width + 18)
        self._catalog_preferred_width = min(
            380,
            max(350, self._catalog_tree_content_width + 24),
        )

    def _current_sidebar_width(self, window_width: int) -> int:
        preferred_width = min(
            self._catalog_preferred_width,
            280 if self._layout_mode == "compact" else 380,
        )
        return catalog_sidebar_width(
            window_width,
            mode=self._layout_mode,
            preferred_width=preferred_width,
            user_width=self._sidebar_user_width,
        )

    def _set_sidebar_width(self, width: int) -> None:
        if self._layout_mode == "narrow":
            return
        target = max(1, int(width))
        try:
            current = int(self.sidebar_shell.cget("width"))
        except (tk.TclError, TypeError, ValueError):
            current = 0
        if abs(current - target) >= 2:
            self.sidebar_shell.configure(width=target)

    def _start_sidebar_resize(self, event: tk.Event) -> str:
        if self._layout_mode == "narrow":
            return "break"
        self._sidebar_drag_start_x = int(event.x_root)
        self._sidebar_drag_start_width = max(1, self.sidebar_shell.winfo_width())
        return "break"

    def _drag_sidebar_resize(self, event: tk.Event) -> str:
        if self._layout_mode == "narrow":
            return "break"
        requested = self._sidebar_drag_start_width + (
            int(event.x_root) - self._sidebar_drag_start_x
        )
        window_width = self._logical_window_width(self.winfo_width())
        self._sidebar_user_width = catalog_sidebar_width(
            window_width,
            mode=self._layout_mode,
            preferred_width=self._catalog_preferred_width,
            user_width=requested,
        )
        self._set_sidebar_width(self._sidebar_user_width)
        return "break"

    def _finish_sidebar_resize(self, _event: tk.Event) -> str:
        self.sidebar_resize_handle.configure(bg=SHADOW)
        return "break"

    def _reset_sidebar_width(self, _event: tk.Event | None = None) -> str:
        self._sidebar_user_width = None
        self._set_sidebar_width(
            self._current_sidebar_width(
                self._logical_window_width(self.winfo_width())
            )
        )
        return "break"

    def _toggle_sidebar(self) -> None:
        if self._layout_mode != "narrow":
            return
        self._sidebar_expanded = not self._sidebar_expanded
        self._apply_responsive_layout(
            self._logical_window_width(self.winfo_width()),
            force=True,
        )

    def _apply_responsive_layout(
        self,
        window_width: int,
        *,
        force: bool = False,
    ) -> None:
        mode = responsive_layout_mode(window_width)
        narrow = mode == "narrow"
        compact = mode == "compact"
        stacked = mode != "wide"
        mode_changed = mode != self._layout_mode
        if mode_changed:
            self._sidebar_expanded = not narrow
        self._layout_mode = mode
        if mode_changed:
            self._refresh_operation_description()

        subtitle_width = max(220, int(window_width) - (70 if narrow else 440))
        self._set_label_wraplength(self.header_subtitle, subtitle_width)
        self._set_label_wraplength(
            self.wps_notice_text,
            max(220, int(window_width) - (34 if narrow else 190)),
        )
        self.wps_notice.pack_configure(
            padx=9 if narrow else (14 if compact else 20),
            pady=(0, 9 if narrow else 12),
        )
        self._set_label_wraplength(
            self.catalog_legend,
            210 if stacked else 280,
        )
        self._set_label_wraplength(
            self.operation_title,
            max(210, int(window_width * (0.68 if narrow else 0.52))),
        )

        if narrow:
            self.sidebar_resize_handle.place_forget()
        else:
            self.sidebar_resize_handle.place(
                relx=1.0,
                x=-8,
                y=0,
                width=8,
                relheight=1.0,
            )
            self.sidebar_resize_handle.lift()
            self._set_sidebar_width(self._current_sidebar_width(window_width))

        if not mode_changed and not force:
            return

        for column in range(2):
            self.body.columnconfigure(column, weight=0)
        for row in range(2):
            self.body.rowconfigure(row, weight=0)

        if narrow:
            self.wps_notice.columnconfigure(0, weight=1)
            self.wps_notice.columnconfigure(1, weight=0)
            self.wps_notice_badge.grid_configure(
                row=0,
                column=0,
                sticky="w",
                padx=11,
                pady=(8, 0),
            )
            self.wps_notice_text.grid_configure(
                row=1,
                column=0,
                sticky="ew",
                padx=11,
                pady=(5, 9),
            )
            self.header.configure(padding=(12, 11, 12, 9))
            self.header_subtitle.grid_remove()
            self.title_box.grid_configure(row=0, column=0, sticky="ew")
            self.sidebar_toggle_button.grid(
                row=1,
                column=0,
                sticky="w",
                pady=(9, 0),
            )
            self.engine_summary.grid_configure(
                row=1,
                column=1,
                sticky="e",
                padx=(8, 0),
                pady=(9, 0),
            )
            self.body.configure(padding=(9, 0, 9, 9))
            self.body.columnconfigure(0, weight=1)
            self.sidebar.configure(padding=12)
            self.content.configure(padding=12)
            self.show_unavailable_check.configure(text="显示需安装的功能")
            self.catalog_legend.configure(text="● 内置   ◆ 外部   ○ 需安装")

            if self._sidebar_expanded:
                sidebar_height = min(300, max(220, int(self.winfo_height() * 0.38)))
                self.sidebar_shell.configure(width=1, height=sidebar_height)
                self.sidebar_shell.grid(
                    row=0,
                    column=0,
                    sticky="ew",
                    padx=(0, 0),
                    pady=(0, 9),
                )
                self.content_shell.grid(row=1, column=0, sticky="nsew")
                self.body.rowconfigure(1, weight=1)
                self.sidebar_toggle_button.configure(text="▴  收起工具目录")
            else:
                self.sidebar_shell.grid_remove()
                self.content_shell.grid(row=0, column=0, sticky="nsew")
                self.body.rowconfigure(0, weight=1)
                self.sidebar_toggle_button.configure(text="☰  工具目录")

            self.operation_title.grid_configure(
                row=0,
                column=0,
                columnspan=2,
                sticky="w",
            )
            self.capability_badge.grid_configure(
                row=1,
                column=0,
                columnspan=2,
                sticky="w",
                pady=(7, 0),
            )

            for column in range(7):
                self.file_toolbar.columnconfigure(column, weight=0)
            for column in range(3):
                self.file_toolbar.columnconfigure(column, weight=1)
            for widget in (
                *self.file_toolbar_buttons,
                self.file_more_button,
                self.file_count_label,
            ):
                widget.grid_forget()
            for index, button in enumerate(
                (self.add_file_button, self.add_folder_button, self.file_more_button)
            ):
                button.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 6, 0))
            self.file_count_label.grid(
                row=1,
                column=0,
                columnspan=3,
                sticky="w",
                pady=(8, 1),
            )
            self.file_tree.configure(height=5)

            for column in range(4):
                self.output_frame.columnconfigure(column, weight=0)
            self.output_frame.columnconfigure(0, weight=1)
            for widget in (
                self.output_label,
                self.output_entry,
                self.output_browse_button,
                self.output_open_button,
            ):
                widget.grid_forget()
            self.output_label.grid(row=0, column=0, columnspan=3, sticky="w")
            self.output_entry.grid(row=1, column=0, sticky="ew", pady=(5, 0))
            self.output_browse_button.grid(row=1, column=1, padx=(8, 0), pady=(5, 0))
            self.output_open_button.grid(row=1, column=2, padx=(6, 0), pady=(5, 0))

            self.parameters_title.grid_configure(
                row=0, column=0, columnspan=2, sticky="w"
            )
            self.parameters_scroll_hint.grid_configure(
                row=1,
                column=0,
                columnspan=2,
                sticky="w",
                padx=(0, 0),
                pady=(3, 0),
            )

            self.progress_box.grid_configure(
                row=0,
                column=0,
                columnspan=2,
                sticky="ew",
                padx=(0, 0),
                pady=(0, 9),
            )
            self.cancel_button.grid_configure(
                row=1,
                column=0,
                sticky="ew",
                padx=(0, 6),
                pady=(0, 0),
            )
            self.run_button.grid_configure(
                row=1,
                column=1,
                sticky="ew",
                padx=(6, 0),
                pady=(0, 0),
            )
            self.footer.columnconfigure(0, weight=1)
            self.footer.columnconfigure(1, weight=1)
            self.footer.columnconfigure(2, weight=0)
        else:
            self.wps_notice.columnconfigure(0, weight=0)
            self.wps_notice.columnconfigure(1, weight=1)
            self.wps_notice_badge.grid_configure(
                row=0,
                column=0,
                sticky="w",
                padx=(12, 12),
                pady=9,
            )
            self.wps_notice_text.grid_configure(
                row=0,
                column=1,
                sticky="ew",
                padx=(0, 12),
                pady=9,
            )
            self.sidebar_toggle_button.grid_remove()
            self.sidebar_shell.grid(
                row=0,
                column=0,
                sticky="nsew",
                padx=(0, 11 if compact else 13),
                pady=(0, 0),
            )
            self.content_shell.grid(row=0, column=1, sticky="nsew")
            self.body.columnconfigure(1, weight=1)
            self.body.rowconfigure(0, weight=1)
            self.sidebar_shell.configure(
                width=self._current_sidebar_width(window_width),
                height=1,
            )

            if compact:
                self.header.configure(padding=(18, 14, 18, 11))
                self.header_subtitle.grid_remove()
                self.body.configure(padding=(14, 0, 14, 14))
                self.sidebar.configure(padding=12)
                self.content.configure(padding=16)
                self.show_unavailable_check.configure(text="显示需安装的功能")
                self.catalog_legend.configure(text="● 内置   ◆ 外部   ○ 需安装")
            else:
                self.header.configure(padding=(24, 18, 24, 14))
                self.header_subtitle.grid()
                self.body.configure(padding=(20, 0, 20, 20))
                self.sidebar.configure(padding=14)
                self.content.configure(padding=22)
                self.show_unavailable_check.configure(
                    text="显示需要额外安装的核心功能"
                )
                self.catalog_legend.configure(
                    text="● 内置可用   ◆ 外部引擎已就绪   ○ 需要安装"
                )

            self.title_box.grid_configure(row=0, column=0, sticky="ew")
            self.engine_summary.grid_configure(
                row=0,
                column=1,
                sticky="e",
                padx=(16, 0),
                pady=(0, 0),
            )
            self.operation_title.grid_configure(
                row=0,
                column=0,
                columnspan=1,
                sticky="w",
            )
            self.capability_badge.grid_configure(
                row=0,
                column=1,
                columnspan=1,
                sticky="e",
                pady=(0, 0),
            )

            for column in range(7):
                self.file_toolbar.columnconfigure(column, weight=0)
            for widget in (
                *self.file_toolbar_buttons,
                self.file_more_button,
                self.file_count_label,
            ):
                widget.grid_forget()
            if compact:
                for column in range(3):
                    self.file_toolbar.columnconfigure(column, weight=1)
                self.file_toolbar.columnconfigure(3, weight=1)
                for index, button in enumerate(
                    (
                        self.add_file_button,
                        self.add_folder_button,
                        self.file_more_button,
                    )
                ):
                    button.grid(
                        row=0,
                        column=index,
                        sticky="ew",
                        padx=(0 if index == 0 else 6, 0),
                    )
                self.file_count_label.grid(row=0, column=3, sticky="e", padx=(12, 0))
                self.file_tree.configure(height=6)
            else:
                self.file_toolbar.columnconfigure(6, weight=1)
                for index, button in enumerate(self.file_toolbar_buttons):
                    button.grid(
                        row=0,
                        column=index,
                        sticky="w",
                        padx=(0 if index == 0 else 6, 0),
                    )
                self.file_count_label.grid(row=0, column=6, sticky="e", padx=(12, 0))
                self.file_tree.configure(height=7)

            for column in range(4):
                self.output_frame.columnconfigure(column, weight=0)
            self.output_frame.columnconfigure(1, weight=1)
            for widget in (
                self.output_label,
                self.output_entry,
                self.output_browse_button,
                self.output_open_button,
            ):
                widget.grid_forget()
            self.output_label.grid(row=0, column=0, sticky="w", padx=(0, 10))
            self.output_entry.grid(row=0, column=1, sticky="ew")
            self.output_browse_button.grid(row=0, column=2, padx=(8, 0))
            self.output_open_button.grid(row=0, column=3, padx=(6, 0))

            self.parameters_title.grid_configure(
                row=0,
                column=0,
                columnspan=1,
                sticky="w",
            )
            self.parameters_scroll_hint.grid_configure(
                row=0,
                column=1,
                columnspan=1,
                sticky="e",
                padx=(12, 0),
                pady=(0, 0),
            )

            self.progress_box.grid_configure(
                row=0,
                column=0,
                columnspan=1,
                sticky="ew",
                padx=(0, 12),
                pady=(0, 0),
            )
            self.cancel_button.grid_configure(
                row=0,
                column=1,
                sticky="",
                padx=(0, 8),
                pady=(0, 0),
            )
            self.run_button.grid_configure(
                row=0,
                column=2,
                sticky="",
                padx=(0, 0),
                pady=(0, 0),
            )
            self.footer.columnconfigure(0, weight=1)
            self.footer.columnconfigure(1, weight=0)
            self.footer.columnconfigure(2, weight=0)

        self._layout_parameter_rows(stacked)
        self._layout_image_preview_panes(stacked)
        self._schedule_setup_scroll_refresh()

    @staticmethod
    def _set_label_wraplength(label: ttk.Label, wraplength: int) -> None:
        try:
            current = int(float(label.cget("wraplength")))
        except (tk.TclError, TypeError, ValueError):
            current = -1
        if current != wraplength:
            label.configure(wraplength=wraplength)

    def _on_content_configure(self, event: tk.Event) -> None:
        compact = self._layout_mode != "wide"
        wraplength = responsive_wraplength(
            event.width,
            padding=12,
            minimum=180 if compact else 320,
        )
        self._set_label_wraplength(self.operation_description, wraplength)
        self._set_label_wraplength(self.progress_label, wraplength)

    def _toggle_operation_details(self, _event: tk.Event | None = None) -> None:
        self._operation_details_expanded = not self._operation_details_expanded
        self._refresh_operation_description()
        self._schedule_setup_scroll_refresh()

    def _refresh_operation_description(self) -> None:
        description, notes, engine, reason = self._operation_description_parts
        compact = self._layout_mode != "wide"
        self.operation_description.configure(
            text=operation_description_text(
                description,
                notes,
                engine,
                reason,
                compact=compact,
                expanded=self._operation_details_expanded,
            ),
            cursor="hand2",
        )

    def _on_parameters_frame_configure(self, event: tk.Event) -> None:
        compact = self._layout_mode != "wide"
        wraplength = responsive_wraplength(
            event.width,
            reserved_width=0 if compact else 180,
            padding=12 if compact else 24,
            minimum=160 if compact else 280,
        )
        for label in tuple(self.parameter_hint_labels):
            try:
                if label.winfo_exists():
                    self._set_label_wraplength(label, wraplength)
            except tk.TclError:
                continue

    def _layout_image_preview_panes(self, stacked: bool) -> None:
        if not hasattr(self, "image_preview_panes"):
            return
        layout = "stacked" if stacked else "side_by_side"
        if layout == self._image_preview_layout:
            return
        self._image_preview_layout = layout
        for box in (self.image_preview_original_box, self.image_preview_result_box):
            box.grid_forget()
        for column in range(2):
            self.image_preview_panes.columnconfigure(column, weight=0)
        for row in range(2):
            self.image_preview_panes.rowconfigure(row, weight=0)
        if stacked:
            self.image_preview_panes.columnconfigure(0, weight=1)
            self.image_preview_original_box.grid(
                row=0,
                column=0,
                sticky="ew",
                pady=(0, 8),
            )
            self.image_preview_result_box.grid(
                row=1,
                column=0,
                sticky="ew",
            )
            height = max(180, round(205 * self._display_scale))
        else:
            self.image_preview_panes.columnconfigure(
                0,
                weight=1,
                uniform="image_preview",
            )
            self.image_preview_panes.columnconfigure(
                1,
                weight=1,
                uniform="image_preview",
            )
            self.image_preview_original_box.grid(
                row=0,
                column=0,
                sticky="nsew",
                padx=(0, 5),
            )
            self.image_preview_result_box.grid(
                row=0,
                column=1,
                sticky="nsew",
                padx=(5, 0),
            )
            height = max(210, round(250 * self._display_scale))
        self.image_preview_original_canvas.configure(height=height)
        self.image_preview_result_canvas.configure(height=height)
        self.after_idle(lambda: self._render_image_preview_canvas("original"))
        self.after_idle(lambda: self._render_image_preview_canvas("result"))

    def _image_preview_candidates(self) -> list[Path]:
        if self.current_operation is None:
            return []
        return [
            path
            for path in self.input_paths
            if path.is_file() and path.suffix.casefold() in {
                extension.casefold() for extension in self.current_operation.extensions
            }
        ]

    def _configure_image_preview(self) -> None:
        operation = self.current_operation
        if operation is None or not supports_live_image_preview(operation.id):
            self.image_preview_frame.grid_remove()
            self._image_preview_generation += 1
            self._clear_image_preview_images()
            self._schedule_setup_scroll_refresh()
            return
        self.image_preview_frame.grid()
        interactive = operation.id in {"image.crop", "image.mosaic"}
        self.image_preview_original_canvas.configure(
            cursor="crosshair" if interactive else "arrow"
        )
        self._layout_image_preview_panes(self._layout_mode != "wide")
        self._schedule_image_preview(delay=60)
        self._schedule_setup_scroll_refresh()

    def _initialize_image_parameters_from_source(self) -> None:
        """Replace misleading fixed crop defaults with the first image bounds."""

        operation = self.current_operation
        if operation is None or operation.id != "image.crop":
            return
        candidates = self._image_preview_candidates()
        if not candidates or "right" not in self.param_vars or "bottom" not in self.param_vars:
            return
        defaults = {spec.key: str(spec.default) for spec in operation.parameters}
        update_right = str(self.param_vars["right"].get()) == defaults.get("right")
        update_bottom = str(self.param_vars["bottom"].get()) == defaults.get("bottom")
        if not update_right and not update_bottom:
            return
        try:
            with Image.open(candidates[0]) as opened:
                width, height = opened.size
                orientation = int(opened.getexif().get(274, 1) or 1)
                if orientation in {5, 6, 7, 8}:
                    width, height = height, width
        except (OSError, ValueError, TypeError):
            return
        if update_right:
            self.param_vars["right"].set(str(width))
        if update_bottom:
            self.param_vars["bottom"].set(str(height))

    def _on_parameter_value_changed(self) -> None:
        self._update_parameter_visibility()
        self._schedule_image_preview()

    def _schedule_image_preview(self, *, delay: int = 260) -> None:
        if self._closing or not hasattr(self, "image_preview_frame"):
            return
        operation = self.current_operation
        if operation is None or not supports_live_image_preview(operation.id):
            return
        if self._image_preview_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._image_preview_job)
            self._image_preview_job = None
        candidates = self._image_preview_candidates()
        self._image_preview_index = min(
            max(0, self._image_preview_index),
            max(0, len(candidates) - 1),
        )
        self.image_preview_previous.configure(
            state="normal" if self._image_preview_index > 0 else "disabled"
        )
        self.image_preview_next.configure(
            state=(
                "normal"
                if self._image_preview_index + 1 < len(candidates)
                else "disabled"
            )
        )
        if not candidates:
            self._image_preview_generation += 1
            self._clear_image_preview_images()
            self.image_preview_status.configure(text="添加图片后自动生成对比预览")
            self.image_preview_details.configure(
                text=(
                    "添加图片后，可直接在左侧原图拖框选择区域。"
                    if operation.id in {"image.crop", "image.mosaic"}
                    else "预览不会生成文件，也不会修改原图。"
                )
            )
            for key in ("original", "result"):
                self._render_image_preview_canvas(key)
            return
        self.image_preview_status.configure(
            text=(
                f"准备预览第 {self._image_preview_index + 1} / {len(candidates)} 张 · "
                f"{candidates[self._image_preview_index].name}"
            )
        )
        self._image_preview_job = self.after(
            max(0, int(delay)),
            self._start_image_preview,
        )

    def _start_image_preview(self) -> None:
        self._image_preview_job = None
        operation = self.current_operation
        candidates = self._image_preview_candidates()
        if (
            operation is None
            or not supports_live_image_preview(operation.id)
            or not candidates
        ):
            return
        index = min(self._image_preview_index, len(candidates) - 1)
        source = candidates[index]
        raw_parameters = {
            spec.key: self._parameter_actual_value(spec.key)
            for spec in operation.parameters
        }
        self._image_preview_generation += 1
        generation = self._image_preview_generation
        self.image_preview_status.configure(
            text=f"正在更新第 {index + 1} / {len(candidates)} 张预览…"
        )

        def work() -> None:
            try:
                normalized = operation.normalize_parameters(raw_parameters)
                result = build_live_image_preview(
                    source,
                    operation.id,
                    normalized,
                )
            except Exception as exc:
                if self._closing:
                    return
                with contextlib.suppress(tk.TclError, RuntimeError):
                    self.after(
                        0,
                        lambda error=str(exc), token=generation: self._show_image_preview_error(
                            token,
                            error,
                        ),
                    )
                return
            if self._closing:
                result[0].close()
                result[1].close()
                return
            with contextlib.suppress(tk.TclError, RuntimeError):
                self.after(
                    0,
                    lambda payload=result, token=generation, current=index, total=len(
                        candidates
                    ): self._apply_image_preview(
                        token,
                        current,
                        total,
                        payload,
                    ),
                )

        threading.Thread(
            target=work,
            name="docuforge-image-preview",
            daemon=True,
        ).start()

    def _show_image_preview_error(self, generation: int, message: str) -> None:
        if generation != self._image_preview_generation or self._closing:
            return
        self._clear_image_preview_images()
        self.image_preview_status.configure(text="参数尚未满足预览条件")
        self.image_preview_details.configure(
            text=f"{message}\n继续调整参数后，预览会自动重试。"
        )
        self._render_image_preview_canvas("original")
        self._render_image_preview_canvas("result")

    def _apply_image_preview(
        self,
        generation: int,
        index: int,
        total: int,
        payload: tuple[
            Image.Image,
            Image.Image,
            str,
            str,
            tuple[int, int],
        ],
    ) -> None:
        original, result, original_info, result_info, source_size = payload
        if generation != self._image_preview_generation or self._closing:
            original.close()
            result.close()
            return
        self._clear_image_preview_images()
        self._image_preview_original = original
        self._image_preview_result = result
        self._image_preview_source_size = source_size
        self.image_preview_status.configure(
            text=f"第 {index + 1} / {total} 张 · 参数变化后自动刷新"
        )
        self.image_preview_details.configure(text=f"{original_info}\n{result_info}")
        self._render_image_preview_canvas("original")
        self._render_image_preview_canvas("result")
        self._schedule_setup_scroll_refresh()

    def _clear_image_preview_images(self) -> None:
        for image in (self._image_preview_original, self._image_preview_result):
            if image is not None:
                with contextlib.suppress(Exception):
                    image.close()
        self._image_preview_original = None
        self._image_preview_result = None
        self._image_preview_photos.clear()

    def _render_image_preview_canvas(self, key: str) -> None:
        if not hasattr(self, "image_preview_original_canvas"):
            return
        canvas = (
            self.image_preview_original_canvas
            if key == "original"
            else self.image_preview_result_canvas
        )
        source = (
            self._image_preview_original
            if key == "original"
            else self._image_preview_result
        )
        try:
            width = max(1, canvas.winfo_width())
            height = max(1, canvas.winfo_height())
            canvas.delete("all")
            if source is None:
                self._image_preview_display_rects[key] = (0, 0, 0, 0)
                canvas.create_text(
                    width // 2,
                    height // 2,
                    text="等待预览",
                    fill=MUTED,
                    font=("Microsoft YaHei UI", 10),
                )
                return
            display = source.copy()
            display.thumbnail(
                (max(1, width - 18), max(1, height - 18)),
                Image.Resampling.LANCZOS,
            )
            if "A" in display.getbands():
                background = Image.new("RGBA", display.size, "white")
                rgba_display = display.convert("RGBA")
                background.alpha_composite(rgba_display)
                rgba_display.close()
                display.close()
                display = background.convert("RGB")
                background.close()
            photo = ImageTk.PhotoImage(display, master=self)
            left = (width - display.width) // 2
            top = (height - display.height) // 2
            self._image_preview_display_rects[key] = (
                left,
                top,
                left + display.width,
                top + display.height,
            )
            display.close()
            self._image_preview_photos[key] = photo
            canvas.create_image(left, top, image=photo, anchor="nw")
        except tk.TclError:
            return

    def _direct_image_edit_point(self, event: tk.Event) -> tuple[int, int] | None:
        operation = self.current_operation
        if operation is None or operation.id not in {"image.crop", "image.mosaic"}:
            return None
        left, top, right, bottom = self._image_preview_display_rects["original"]
        if right <= left or bottom <= top:
            return None
        x = min(max(float(event.x), left), max(left, right - 1))
        y = min(max(float(event.y), top), max(top, bottom - 1))
        return canvas_point_to_source(
            x,
            y,
            (left, top, right, bottom),
            self._image_preview_source_size[0],
            self._image_preview_source_size[1],
        )

    def _start_direct_image_edit(self, event: tk.Event) -> None:
        point = self._direct_image_edit_point(event)
        if point is None:
            return
        left, top, right, bottom = self._image_preview_display_rects["original"]
        canvas_x = min(max(float(event.x), left), right - 1)
        canvas_y = min(max(float(event.y), top), bottom - 1)
        self._direct_image_edit_start = point
        self._direct_image_edit_canvas_start = (canvas_x, canvas_y)
        self.image_preview_original_canvas.delete("direct-image-edit")
        self.image_preview_original_canvas.create_rectangle(
            canvas_x,
            canvas_y,
            canvas_x,
            canvas_y,
            outline="#F0445E",
            width=max(2, round(2 * self._display_scale)),
            dash=(6, 3),
            tags=("direct-image-edit",),
        )

    def _drag_direct_image_edit(self, event: tk.Event) -> None:
        if self._direct_image_edit_start is None or self._direct_image_edit_canvas_start is None:
            return
        left, top, right, bottom = self._image_preview_display_rects["original"]
        canvas_x = min(max(float(event.x), left), right - 1)
        canvas_y = min(max(float(event.y), top), bottom - 1)
        start_x, start_y = self._direct_image_edit_canvas_start
        self.image_preview_original_canvas.coords(
            "direct-image-edit",
            start_x,
            start_y,
            canvas_x,
            canvas_y,
        )

    def _finish_direct_image_edit(self, event: tk.Event) -> None:
        start = self._direct_image_edit_start
        end = self._direct_image_edit_point(event)
        self._direct_image_edit_start = None
        self._direct_image_edit_canvas_start = None
        self.image_preview_original_canvas.delete("direct-image-edit")
        if start is None or end is None:
            return
        left = min(start[0], end[0])
        top = min(start[1], end[1])
        right = min(self._image_preview_source_size[0], max(start[0], end[0]) + 1)
        bottom = min(self._image_preview_source_size[1], max(start[1], end[1]) + 1)
        if right - left < 2 or bottom - top < 2:
            self.image_preview_status.configure(text="拖框范围太小，请重新拖动")
            return
        for key, value in (
            ("left", left),
            ("top", top),
            ("right", right),
            ("bottom", bottom),
        ):
            variable = self.param_vars.get(key)
            if variable is not None:
                variable.set(str(value))
        action = "裁切保留范围" if self.current_operation and self.current_operation.id == "image.crop" else "马赛克区域"
        self.image_preview_status.configure(
            text=f"已从原图框选{action}：{left}, {top} → {right}, {bottom}"
        )

    def _change_image_preview(self, delta: int) -> None:
        candidates = self._image_preview_candidates()
        if not candidates:
            return
        self._image_preview_index = min(
            max(0, self._image_preview_index + int(delta)),
            len(candidates) - 1,
        )
        if str(self._image_preview_index) in self.file_tree.get_children(""):
            self.file_tree.selection_set(str(self._image_preview_index))
            self.file_tree.see(str(self._image_preview_index))
        self._schedule_image_preview(delay=20)

    def _on_file_tree_preview_selected(self, _event: tk.Event | None = None) -> None:
        selection = self.file_tree.selection()
        if not selection or not str(selection[0]).isdigit():
            return
        selected_path_index = int(selection[0])
        if not 0 <= selected_path_index < len(self.input_paths):
            return
        selected_path = self.input_paths[selected_path_index]
        candidates = self._image_preview_candidates()
        with contextlib.suppress(ValueError):
            self._image_preview_index = candidates.index(selected_path)
            self._schedule_image_preview(delay=20)

    def _on_setup_canvas_configure(self, event: tk.Event) -> None:
        try:
            self.setup_canvas.itemconfigure(
                self.setup_window, width=max(1, int(event.width))
            )
        except tk.TclError:
            return
        self._schedule_setup_scroll_refresh()

    def _on_setup_tab_configure(self, _event: tk.Event) -> None:
        self._schedule_setup_scroll_refresh()

    def _schedule_setup_scroll_refresh(self) -> None:
        if self._closing:
            return
        if self._setup_scroll_refresh_job is not None:
            try:
                self.after_cancel(self._setup_scroll_refresh_job)
            except tk.TclError:
                pass
        try:
            self._setup_scroll_refresh_job = self.after_idle(
                self._refresh_setup_scrollregion
            )
        except tk.TclError:
            self._setup_scroll_refresh_job = None

    def _refresh_setup_scrollregion(self) -> None:
        self._setup_scroll_refresh_job = None
        try:
            bounds = self.setup_canvas.bbox("all")
            if bounds is None:
                return
            visible_width = max(1, self.setup_canvas.winfo_width())
            visible_height = max(1, self.setup_canvas.winfo_height())
            content_width = max(visible_width, int(bounds[2]))
            content_height = max(visible_height, int(bounds[3]) + 4)
            self.setup_canvas.configure(
                scrollregion=(0, 0, content_width, content_height)
            )
        except tk.TclError:
            return

    def _reset_setup_scroll_to_top(self) -> None:
        self._refresh_setup_scrollregion()
        try:
            self.setup_canvas.yview_moveto(0.0)
        except tk.TclError:
            return

    def _attach_setup_scroll_bindtag(self, widget: tk.Misc) -> None:
        if widget is self.file_tree:
            return
        tags = widget.bindtags()
        if self._setup_scroll_bindtag not in tags:
            widget.bindtags((self._setup_scroll_bindtag, *tags))
        for child in widget.winfo_children():
            self._attach_setup_scroll_bindtag(child)

    def _on_setup_mousewheel(self, event: tk.Event) -> str | None:
        try:
            if self.notebook.select() != str(self.setup_page):
                return None
            units = mousewheel_scroll_units(getattr(event, "delta", 0))
            if units == 0:
                return None
            self.setup_canvas.yview_scroll(units, "units")
        except tk.TclError:
            return None
        return "break"

    def _on_setup_button_scroll(self, event: tk.Event) -> str | None:
        number = int(getattr(event, "num", 0) or 0)
        if number not in {4, 5}:
            return None
        try:
            if self.notebook.select() != str(self.setup_page):
                return None
            self.setup_canvas.yview_scroll(-3 if number == 4 else 3, "units")
        except tk.TclError:
            return None
        return "break"

    def _clear_search_hint(self, _event: tk.Event) -> None:
        if self.search_var.get() == "搜索功能…":
            self.search_var.set("")

    def _rebuild_operation_tree(self) -> None:
        selected_id = self.current_operation.id if self.current_operation else None
        self.operation_tree.delete(*self.operation_tree.get_children())
        query = self.search_var.get().strip().lower()
        if query == "搜索功能…":
            query = ""
        catalog: dict[str, dict[str, list[Operation]]] = {}
        for operation in self.operations:
            capability = operation.capability()
            if not self.show_unavailable_var.get() and not capability.runnable:
                continue
            root_name, section_name = operation_catalog_path(operation)
            haystack = (
                f"{root_name} {section_name} {operation.group} "
                f"{operation.name} {operation.description}".lower()
            )
            if query and query not in haystack:
                continue
            catalog.setdefault(root_name, {}).setdefault(section_name, []).append(
                operation
            )
        ordered_roots = sorted(
            catalog.items(), key=lambda item: catalog_order_key(item[0])
        )
        for root_index, (root_name, sections) in enumerate(ordered_roots):
            root_contains_selection = any(
                operation.id == selected_id
                for operations in sections.values()
                for operation in operations
            )
            root = self.operation_tree.insert(
                "",
                "end",
                text=root_name,
                open=bool(query) or root_contains_selection or root_index == 0,
                tags=("catalog_root",),
            )
            ordered_sections = sorted(
                sections.items(),
                key=lambda item: catalog_order_key(root_name, item[0]),
            )
            for section_index, (section_name, operations) in enumerate(
                ordered_sections
            ):
                section_contains_selection = any(
                    item.id == selected_id for item in operations
                )
                section = self.operation_tree.insert(
                    root,
                    "end",
                    text=section_name,
                    open=(
                        bool(query)
                        or section_contains_selection
                        or (root_index == 0 and section_index == 0)
                    ),
                    tags=("catalog_section",),
                )
                for operation in operations:
                    capability = operation.capability()
                    marker = {
                        "ready": "●",
                        "external": "◆",
                        "unavailable": "○",
                    }[capability.status]
                    self.operation_tree.insert(
                        section,
                        "end",
                        iid=operation.id,
                        text=f"{marker}  {operation.name}",
                    )
        if selected_id and self.operation_tree.exists(selected_id):
            self.operation_tree.selection_set(selected_id)
            self.operation_tree.focus(selected_id)
            self.operation_tree.see(selected_id)
            return
        first_operation = self._first_operation_tree_item()
        if first_operation:
            self.operation_tree.selection_set(first_operation)
            self.operation_tree.focus(first_operation)
            self.operation_tree.see(first_operation)
            self._on_operation_selected()

    def _first_operation_tree_item(self, parent: str = "") -> str | None:
        for item in self.operation_tree.get_children(parent):
            if item in self.operation_by_id:
                return item
            nested = self._first_operation_tree_item(item)
            if nested:
                return nested
        return None

    def _on_operation_selected(self, _event: tk.Event | None = None) -> None:
        selection = self.operation_tree.selection()
        if not selection or selection[0] not in self.operation_by_id:
            return
        operation = self.operation_by_id[selection[0]]
        self.current_operation = operation
        capability = operation.capability()
        self.operation_title.configure(text=operation.name)
        self._animate_operation_title()
        self._operation_description_parts = (
            operation.description,
            operation.notes,
            capability.engine,
            capability.reason,
        )
        self._operation_details_expanded = False
        self._refresh_operation_description()
        badge_config = {
            "ready": ("可直接运行", "#EAF8F3", SUCCESS),
            "external": ("外部引擎已就绪", "#FFF7ED", WARNING),
            "unavailable": ("需要安装组件", "#FFF1F2", DANGER),
        }[capability.status]
        self.capability_badge.configure(
            text=badge_config[0], bg=badge_config[1], fg=badge_config[2]
        )
        self._build_parameters(operation.parameters)
        self._filter_existing_inputs()
        self._configure_image_preview()
        self._reset_drop_hint()
        worker_running = bool(self.worker and self.worker.is_alive())
        can_explain_missing_office = (
            not capability.runnable and capability.engine == "Office 渲染器"
        )
        self.run_button.configure(
            state=(
                "normal"
                if not worker_running
                and (capability.runnable or can_explain_missing_office)
                else "disabled"
            )
        )
        self.after_idle(self._reset_setup_scroll_to_top)
        if _event is not None and self._layout_mode == "narrow" and self._sidebar_expanded:
            self._sidebar_expanded = False
            self.after(
                90,
                lambda: self._apply_responsive_layout(
                    self._logical_window_width(self.winfo_width()),
                    force=True,
                ),
            )

    def _animate_operation_title(self) -> None:
        if self._title_animation_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._title_animation_job)
            self._title_animation_job = None
        frames = 9

        def step(index: int) -> None:
            progress = ease_out_cubic(index / frames)
            try:
                self.operation_title.configure(
                    foreground=interpolate_hex_colour(ACCENT, TEXT, progress)
                )
            except tk.TclError:
                self._title_animation_job = None
                return
            if index < frames:
                self._title_animation_job = self.after(18, step, index + 1)
            else:
                self._title_animation_job = None

        step(0)

    def _build_parameters(self, specs: tuple[ParameterSpec, ...]) -> None:
        for child in self.parameters_frame.winfo_children():
            child.destroy()
        self.param_vars.clear()
        self.choice_maps.clear()
        self.parameter_hint_labels.clear()
        self.parameter_rows.clear()
        self.parameter_section_frames.clear()
        self.parameter_section_order.clear()
        self.advanced_parameters_frame = None
        self.advanced_parameters_button = None
        self.advanced_parameters_expanded = False
        if not specs:
            ttk.Label(
                self.parameters_frame,
                text="此任务无需额外参数。",
                style="Subtle.TLabel",
            ).grid(row=0, column=0, sticky="w")
            self._attach_setup_scroll_bindtag(self.parameters_frame)
            self._schedule_setup_scroll_refresh()
            return
        use_sections = any(spec.section for spec in specs)
        advanced_specs = any(spec.advanced for spec in specs)
        if use_sections:
            self.parameters_frame.columnconfigure(0, weight=1)
            for spec in specs:
                section = spec.section or "其他设置"
                if section in self.parameter_section_frames:
                    continue
                frame = ttk.LabelFrame(
                    self.parameters_frame,
                    text=section,
                    padding=(14, 10),
                    style="Card.TLabelframe",
                )
                frame.columnconfigure(1, weight=1)
                self.parameter_section_frames[section] = frame
                self.parameter_section_order.append(section)
            normal_sections = [
                section
                for section in self.parameter_section_order
                if not all(
                    item.advanced for item in specs if (item.section or "其他设置") == section
                )
            ]
            advanced_sections = [
                section for section in self.parameter_section_order if section not in normal_sections
            ]
            grid_row = 0
            for section in normal_sections:
                self.parameter_section_frames[section].grid(
                    row=grid_row, column=0, sticky="ew", pady=(0, 10)
                )
                grid_row += 1
            if advanced_sections:
                self.advanced_parameters_button = ttk.Button(
                    self.parameters_frame,
                    text="▶ 显示高级精准修复（通常无需填写）",
                    style="Quiet.TButton",
                    command=self._toggle_advanced_parameters,
                )
                self.advanced_parameters_button.grid(
                    row=grid_row, column=0, sticky="ew", pady=(0, 8)
                )
                grid_row += 1
                # The video tool uses one advanced section.  Supporting more is
                # harmless: they are stacked and toggled together.
                for section in advanced_sections:
                    frame = self.parameter_section_frames[section]
                    frame.grid(row=grid_row, column=0, sticky="ew", pady=(0, 10))
                    frame.grid_remove()
                    if self.advanced_parameters_frame is None:
                        self.advanced_parameters_frame = frame
                    grid_row += 1

        section_counts: dict[str, int] = {}
        for row, spec in enumerate(specs):
            section = spec.section or "其他设置"
            parent: tk.Misc = (
                self.parameter_section_frames[section]
                if use_sections
                else self.parameters_frame
            )
            local_index = section_counts.get(section, 0)
            section_counts[section] = local_index + 1
            control_row = local_index * 2
            label_widget = ttk.Label(
                parent,
                text=spec.label,
                foreground=TEXT,
                style="CardField.TLabel" if use_sections else "TLabel",
            )
            label_widget.grid(
                row=control_row,
                column=0,
                sticky="nw",
                padx=(0, 14),
                pady=(7, 2),
            )
            is_video_ppt_repair_plan = (
                self.current_operation is not None
                and self.current_operation.id == "video.repair_slides_ppt"
                and spec.key == "repair_plan"
            )
            if is_video_ppt_repair_plan:
                variable = tk.StringVar(
                    value="" if spec.default is None else str(spec.default)
                )
                control = ttk.Label(
                    parent,
                    text="无需填写参数：在快速补修窗口中直接选择页面、框选并预览。",
                    style="CardSubtle.TLabel" if use_sections else "Subtle.TLabel",
                    wraplength=520,
                    justify="left",
                )
            elif spec.kind == "boolean":
                variable: tk.Variable = tk.BooleanVar(value=bool(spec.default))
                control = ttk.Checkbutton(
                    parent,
                    variable=variable,
                    text=spec.help_text or "启用",
                    style="Card.TCheckbutton" if use_sections else "TCheckbutton",
                )
            elif spec.kind == "choice":
                value_to_label = {value: label for value, label in spec.choices}
                label_to_value = {label: value for value, label in spec.choices}
                self.choice_maps[spec.key] = label_to_value
                default_label = value_to_label.get(
                    str(spec.default), next(iter(label_to_value), "")
                )
                variable = tk.StringVar(value=default_label)
                control = ttk.Combobox(
                    parent,
                    textvariable=variable,
                    values=list(label_to_value),
                    state="readonly",
                )
            elif spec.kind in {"integer", "number"} and (
                spec.minimum is not None or spec.maximum is not None
            ):
                variable = tk.StringVar(
                    value="" if spec.default is None else str(spec.default)
                )
                control = ttk.Spinbox(
                    parent,
                    textvariable=variable,
                    from_=spec.minimum if spec.minimum is not None else -1_000_000_000,
                    to=spec.maximum if spec.maximum is not None else 1_000_000_000,
                    increment=1 if spec.kind == "integer" else 0.1,
                )
            else:
                variable = tk.StringVar(
                    value="" if spec.default is None else str(spec.default)
                )
                control = ttk.Entry(
                    parent,
                    textvariable=variable,
                    show="•" if spec.kind == "password" else "",
                )
            control.grid(row=control_row, column=1, sticky="ew", pady=(7, 2))
            self.param_vars[spec.key] = variable
            action_button: tk.Misc | None = None
            if spec.kind == "path":
                action_button = ttk.Button(
                    parent,
                    text="选择…",
                    style="Quiet.TButton",
                    command=lambda key=spec.key: self._choose_parameter_path(key),
                )
                action_button.grid(row=control_row, column=2, padx=(10, 0), pady=(7, 2))
            elif is_video_ppt_repair_plan:
                action_button = ttk.Button(
                    parent,
                    text="打开快速补修…",
                    command=lambda key=spec.key: self._open_video_ppt_repair_workbench(key),
                )
            elif (
                spec.kind == "colors"
                and self.current_operation is not None
                and self.current_operation.id == "video.extract_slides_ppt"
                and spec.key == "annotation_colors"
            ):
                action_button = ttk.Frame(parent)
                ttk.Button(
                    action_button,
                    text="从视频取色…",
                    style="Quiet.TButton",
                    command=lambda key=spec.key: self._choose_video_color(key),
                ).pack(side="left")
                ttk.Button(
                    action_button,
                    text="调色板…",
                    style="Quiet.TButton",
                    command=lambda key=spec.key: self._choose_parameter_color(
                        key, append=True
                    ),
                ).pack(side="left", padx=(6, 0))
            elif spec.kind in {"color", "colors"}:
                action_button = ttk.Button(
                    parent,
                    text="添加颜色…" if spec.kind == "colors" else "选颜色…",
                    style="Quiet.TButton",
                    command=lambda key=spec.key, append=spec.kind
                    == "colors": self._choose_parameter_color(key, append=append),
                )
                action_button.grid(row=control_row, column=2, padx=(10, 0), pady=(7, 2))
            elif spec.kind == "region":
                append = spec.key == "fixed_watermark_regions"
                action_button = ttk.Button(
                    parent,
                    text="添加框选…" if append else "在视频中框选…",
                    style="Quiet.TButton",
                    command=lambda key=spec.key, append=append: self._choose_video_region(
                        key, append=append
                    ),
                )
                action_button.grid(row=control_row, column=2, padx=(10, 0), pady=(7, 2))
            hint = parameter_help_text(spec)
            hint_label: ttk.Label | None = None
            if hint:
                hint_label = ttk.Label(
                    parent,
                    text=hint,
                    style="CardSubtle.TLabel" if use_sections else "Subtle.TLabel",
                    wraplength=responsive_wraplength(
                        self.parameters_frame.winfo_width(),
                        reserved_width=180,
                        padding=24,
                        minimum=320,
                    ),
                    justify="left",
                )
                hint_label.grid(
                    row=control_row + 1,
                    column=1,
                    columnspan=2,
                    sticky="ew",
                    pady=(0, 6),
                )
                self.parameter_hint_labels.append(hint_label)
            self.parameter_rows.append(
                (spec, parent, label_widget, control, action_button, hint_label)
            )
            variable.trace_add("write", lambda *_args: self._on_parameter_value_changed())
        self._layout_parameter_rows(self._layout_mode != "wide")
        self._update_parameter_visibility()
        self._attach_setup_scroll_bindtag(self.parameters_frame)
        self._schedule_setup_scroll_refresh()

    def _layout_parameter_rows(self, compact: bool) -> None:
        if not self.parameter_rows:
            return
        parents = {parent for _spec, parent, *_widgets in self.parameter_rows}
        for parent in parents:
            for column in range(3):
                parent.columnconfigure(column, weight=0)
            parent.columnconfigure(0 if compact else 1, weight=1)
        visible_index: dict[tk.Misc, int] = {}
        for spec, parent, label, control, action, hint in self.parameter_rows:
            for widget in (label, control, action, hint):
                if widget is not None:
                    widget.grid_forget()
            if not self._parameter_is_visible(spec):
                continue
            index = visible_index.get(parent, 0)
            visible_index[parent] = index + 1
            if compact:
                base_row = index * 3
                label.grid(
                    row=base_row,
                    column=0,
                    columnspan=2,
                    sticky="w",
                    pady=(8, 1),
                )
                control.grid(
                    row=base_row + 1,
                    column=0,
                    columnspan=1 if action is not None else 2,
                    sticky="ew",
                    pady=(2, 2),
                )
                if action is not None:
                    action.grid(
                        row=base_row + 1,
                        column=1,
                        sticky="e",
                        padx=(8, 0),
                        pady=(2, 2),
                    )
                if hint is not None:
                    hint.grid(
                        row=base_row + 2,
                        column=0,
                        columnspan=2,
                        sticky="ew",
                        pady=(0, 6),
                    )
            else:
                base_row = index * 2
                label.grid(
                    row=base_row,
                    column=0,
                    sticky="nw",
                    padx=(0, 14),
                    pady=(7, 2),
                )
                control.grid(row=base_row, column=1, sticky="ew", pady=(7, 2))
                if action is not None:
                    action.grid(
                        row=base_row,
                        column=2,
                        padx=(10, 0),
                        pady=(7, 2),
                    )
                if hint is not None:
                    hint.grid(
                        row=base_row + 1,
                        column=1,
                        columnspan=2,
                        sticky="ew",
                        pady=(0, 6),
                    )

    def _parameter_actual_value(self, key: str) -> str:
        variable = self.param_vars.get(key)
        if variable is None:
            return ""
        raw = str(variable.get())
        return self.choice_maps.get(key, {}).get(raw, raw)

    def _parameter_is_visible(self, spec: ParameterSpec) -> bool:
        if spec.visible_when is None:
            return True
        dependency, allowed = spec.visible_when
        return self._parameter_actual_value(dependency) in set(allowed)

    def _update_parameter_visibility(self) -> None:
        if not self.parameter_rows:
            return
        self._layout_parameter_rows(self._layout_mode != "wide")
        self._schedule_setup_scroll_refresh()

    def _toggle_advanced_parameters(self) -> None:
        advanced_frames = {
            parent
            for spec, parent, *_widgets in self.parameter_rows
            if spec.advanced
        }
        self.advanced_parameters_expanded = not self.advanced_parameters_expanded
        for frame in advanced_frames:
            if self.advanced_parameters_expanded:
                frame.grid()
            else:
                frame.grid_remove()
        if self.advanced_parameters_button is not None:
            self.advanced_parameters_button.configure(
                text=(
                    "▼ 隐藏高级精准修复"
                    if self.advanced_parameters_expanded
                    else "▶ 显示高级精准修复（通常无需填写）"
                )
            )
        self._schedule_setup_scroll_refresh()

    def _choose_parameter_path(self, key: str) -> None:
        path = filedialog.askopenfilename(title="选择辅助文件")
        if path:
            self.param_vars[key].set(path)

    def _choose_parameter_color(self, key: str, *, append: bool = False) -> None:
        initial = color_dialog_initial(self.param_vars[key].get())
        try:
            _rgb, selected = colorchooser.askcolor(color=initial, title="选择颜色")
        except tk.TclError:
            _rgb, selected = colorchooser.askcolor(color="#000000", title="选择颜色")
        if selected:
            if append:
                selected = append_color_value(self.param_vars[key].get(), selected)
            self.param_vars[key].set(selected)

    def _first_video_input(self) -> Path | None:
        video_extensions = {
            ".mp4",
            ".mkv",
            ".mov",
            ".avi",
            ".webm",
            ".wmv",
            ".m4v",
            ".flv",
            ".mpeg",
            ".mpg",
            ".ts",
            ".mts",
            ".m2ts",
            ".3gp",
        }
        return next(
            (
                path
                for path in self.input_paths
                if path.is_file() and path.suffix.casefold() in video_extensions
            ),
            None,
        )

    def _video_ppt_repair_inputs(self) -> tuple[Path | None, Path | None]:
        pptx = next(
            (
                path
                for path in self.input_paths
                if path.is_file() and path.suffix.casefold() == ".pptx"
            ),
            None,
        )
        return pptx, self._first_video_input()

    def _open_video_ppt_repair_workbench(self, key: str) -> None:
        pptx, video = self._video_ppt_repair_inputs()
        if pptx is None or video is None:
            messagebox.showinfo(
                "请先添加两个文件",
                "请先在上方文件列表中添加一份由视频提取功能生成的 PPTX，以及它所对应的原视频。",
                parent=self,
            )
            return
        try:
            workbench = VideoPptRepairWorkbench(
                self,
                pptx,
                video,
                initial_plan=str(self.param_vars[key].get()),
            )
        except (DocuForgeError, OSError, RuntimeError, ValueError) as exc:
            messagebox.showerror("无法打开人工补修工作台", str(exc), parent=self)
            return
        self.wait_window(workbench)
        if workbench.result:
            self.param_vars[key].set(workbench.result)
            if workbench.start_requested:
                self.after_idle(self._start)

    def _open_video_picker(self, mode: str) -> str | None:
        source = self._first_video_input()
        if source is None:
            messagebox.showinfo(
                "请先添加视频",
                "请先把要处理的视频拖入上方文件列表，再使用视频取色或画面框选。",
                parent=self,
            )
            return None
        try:
            picker = VideoFramePicker(self, source, mode=mode)
        except (OSError, RuntimeError, ValueError) as exc:
            messagebox.showerror("无法打开视频预览", str(exc), parent=self)
            return None
        self.wait_window(picker)
        return picker.result

    def _set_choice_parameter_value(self, key: str, value: str) -> None:
        variable = self.param_vars.get(key)
        if variable is None:
            return
        label = next(
            (
                displayed
                for displayed, actual in self.choice_maps.get(key, {}).items()
                if actual == value
            ),
            value,
        )
        variable.set(label)

    def _choose_video_color(self, key: str) -> None:
        selected = self._open_video_picker("color")
        if not selected:
            return
        existing = str(self.param_vars[key].get()).strip()
        was_manual = self._parameter_actual_value("annotation_color_mode") == "manual"
        if not was_manual or existing.casefold() == "#00aeef":
            combined = selected
        else:
            combined = append_color_value(existing, selected)
        self.param_vars[key].set(combined)
        self._set_choice_parameter_value("annotation_color_mode", "manual")

    def _choose_video_region(self, key: str, *, append: bool = False) -> None:
        selected = self._open_video_picker("region")
        if not selected:
            return
        if append:
            current = str(self.param_vars[key].get()).strip()
            selected = f"{current};{selected}" if current else selected
        self.param_vars[key].set(selected)
        controlling_choice = {
            "crop_rect": ("crop_mode", "custom"),
            "watermark_rect": ("watermark_search", "custom"),
            "presenter_rect": ("presenter_policy", "custom"),
        }.get(key)
        if controlling_choice is not None:
            self._set_choice_parameter_value(*controlling_choice)

    def _default_drop_hint_text(self) -> str:
        if DND_FILES is None:
            return "批量添加文件或文件夹\n拖放组件未安装，请重新运行安装脚本"
        allowed = sorted(self._allowed_extensions())
        type_hint = ""
        if allowed:
            preview = "、".join(
                extension.lstrip(".").upper() for extension in allowed[:6]
            )
            if len(allowed) > 6:
                preview += " 等"
            type_hint = f" · 支持 {preview}"
        return "把文件或文件夹拖到这里\n自动展开文件夹、过滤类型并去重" + type_hint

    def _configure_file_drop_targets(self) -> None:
        if DND_FILES is None:
            self.drop_hint_label.configure(style="DropError.TLabel")
            return
        try:
            for widget in (
                self.file_drop_frame,
                self.drop_hint_label,
                self.file_tree,
            ):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<DropEnter>>", self._on_drop_enter)
                widget.dnd_bind("<<DropLeave>>", self._on_drop_leave)
                widget.dnd_bind("<<Drop>>", self._on_drop_files)
        except (AttributeError, tk.TclError) as exc:
            self.drop_hint_label.configure(
                text=f"拖放初始化失败：{exc}；仍可使用上方按钮批量添加",
                style="DropError.TLabel",
            )

    def _set_drop_hint(
        self,
        text: str,
        *,
        style: str = "DropHint.TLabel",
        reset_after_ms: int | None = None,
    ) -> None:
        if self._drop_hint_reset_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._drop_hint_reset_job)
            self._drop_hint_reset_job = None
        self.drop_hint_label.configure(text=text, style=style)
        if reset_after_ms is not None:
            self._drop_hint_reset_job = self.after(
                max(500, int(reset_after_ms)),
                self._reset_drop_hint,
            )

    def _reset_drop_hint(self) -> None:
        if self._drop_hint_reset_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._drop_hint_reset_job)
            self._drop_hint_reset_job = None
        if not hasattr(self, "drop_hint_label"):
            return
        if self._pending_input_scans:
            self.drop_hint_label.configure(
                text=f"正在后台扫描拖入内容…（{self._pending_input_scans} 批）",
                style="DropBusy.TLabel",
            )
            return
        self.drop_hint_label.configure(
            text=self._default_drop_hint_text(),
            style="DropHint.TLabel" if DND_FILES is not None else "DropError.TLabel",
        )

    def _on_drop_enter(self, event: Any) -> str:
        self._set_drop_hint(
            "松开鼠标即可加入 · 文件夹将递归扫描，重复和不支持的文件会跳过",
            style="DropActive.TLabel",
        )
        # Inputs are only read by LayoutLoom. Always advertise COPY so a source
        # file can never be moved or deleted by the desktop drag operation.
        return str(DND_COPY)

    def _on_drop_leave(self, event: Any) -> str:
        self._reset_drop_hint()
        return str(DND_COPY)

    def _on_drop_files(self, event: Any) -> str:
        sources = parse_drop_payload(
            getattr(event, "data", ""),
            self.tk.splitlist,
        )
        if not sources:
            self._set_drop_hint(
                "未识别到可用路径，请重试或使用“添加文件/添加文件夹”按钮",
                style="DropError.TLabel",
                reset_after_ms=6000,
            )
            if not self._closing:
                messagebox.showwarning(
                    "拖放失败", "未能从拖放内容中识别文件或文件夹路径。"
                )
            return str(DND_COPY)
        self._start_input_collection(sources, source_label="拖放")
        return str(DND_COPY)

    def _start_input_collection(
        self,
        sources: Iterable[str | os.PathLike[str]],
        *,
        source_label: str,
    ) -> None:
        source_values = tuple(
            os.fspath(source) for source in sources if os.fspath(source)
        )
        if not source_values:
            return
        allowed = tuple(self._allowed_extensions())
        existing = tuple(self.input_paths)
        self._pending_input_scans += 1
        self._set_drop_hint(
            f"正在扫描 {len(source_values)} 个{source_label}项目…",
            style="DropBusy.TLabel",
        )

        def work() -> None:
            try:
                result = collect_input_files(
                    source_values,
                    allowed_extensions=allowed,
                    existing_paths=existing,
                )
            except Exception as exc:
                result = InputCollectionResult(
                    files=(),
                    errors=(f"扫描输入时发生异常：{type(exc).__name__}: {exc}",),
                )
            self.worker_queue.put(("input_scan", (source_label, result)))

        threading.Thread(
            target=work,
            name="docuforge-input-scan",
            daemon=True,
        ).start()

    def _finish_input_collection(
        self,
        source_label: str,
        result: InputCollectionResult,
    ) -> None:
        self._pending_input_scans = max(0, self._pending_input_scans - 1)
        # Recheck against the current operation and any batch that completed while
        # this directory was being scanned.
        current = collect_input_files(
            result.files,
            allowed_extensions=self._allowed_extensions(),
            existing_paths=self.input_paths,
        )
        candidate_count = len(current.files)
        if candidate_count > 5000 and not messagebox.askyesno(
            "文件数量较多",
            f"共找到 {candidate_count} 个匹配文件，加入列表可能需要一些时间。是否继续？",
        ):
            self._set_drop_hint(
                f"已取消加入 {candidate_count} 个文件",
                style="DropError.TLabel",
                reset_after_ms=5000,
            )
            return

        self.input_paths.extend(current.files)
        self._refresh_file_tree()
        duplicate_count = result.duplicate_files + current.duplicate_files
        unsupported_count = result.unsupported_files + current.unsupported_files
        missing_paths = tuple(
            dict.fromkeys((*result.missing_paths, *current.missing_paths))
        )
        errors = tuple(dict.fromkeys((*result.errors, *current.errors)))
        added_count = len(current.files)
        details = [f"加入 {added_count} 个文件"]
        if result.scanned_directories:
            details.append(f"递归扫描 {result.scanned_directories} 个文件夹")
        if duplicate_count:
            details.append(f"跳过重复 {duplicate_count} 个")
        if unsupported_count:
            details.append(f"跳过格式不支持 {unsupported_count} 个")
        if missing_paths:
            details.append(f"路径不存在 {len(missing_paths)} 个")
        if errors:
            details.append(f"读取失败 {len(errors)} 处")
        summary = "；".join(details)
        self._log(f"[{source_label}添加输入] {summary}")

        has_hard_problem = bool(missing_paths or errors)
        if added_count:
            self._set_drop_hint(
                summary,
                style=(
                    "DropActive.TLabel" if not has_hard_problem else "DropBusy.TLabel"
                ),
                reset_after_ms=7000,
            )
        else:
            self._set_drop_hint(
                summary,
                style="DropError.TLabel",
                reset_after_ms=7000,
            )

        if not self._closing and (has_hard_problem or (added_count == 0 and summary)):
            problem_lines: list[str] = []
            if unsupported_count:
                problem_lines.append(
                    f"• {unsupported_count} 个文件格式不受当前任务支持"
                )
            if duplicate_count:
                problem_lines.append(f"• {duplicate_count} 个文件已在列表中")
            for path in missing_paths[:3]:
                problem_lines.append(f"• 路径不存在：{path}")
            for error in errors[:3]:
                problem_lines.append(f"• 无法读取：{error}")
            remaining = (
                len(missing_paths)
                + len(errors)
                - min(3, len(missing_paths))
                - min(3, len(errors))
            )
            if remaining > 0:
                problem_lines.append(f"• 另有 {remaining} 项，请查看运行日志")
            messagebox.showwarning(
                "部分输入未加入" if added_count else "没有加入文件",
                summary + ("\n\n" + "\n".join(problem_lines) if problem_lines else ""),
            )

    def _filetypes(self) -> list[tuple[str, str]]:
        if not self.current_operation or not self.current_operation.extensions:
            return [("所有文件", "*.*")]
        patterns = " ".join(
            f"*{ext if ext.startswith('.') else '.' + ext}"
            for ext in self.current_operation.extensions
        )
        return [("当前任务支持的文件", patterns), ("所有文件", "*.*")]

    def _show_file_more_menu(self) -> None:
        try:
            self.file_more_menu.tk_popup(
                self.file_more_button.winfo_rootx(),
                self.file_more_button.winfo_rooty()
                + self.file_more_button.winfo_height()
                + 4,
            )
        finally:
            with contextlib.suppress(tk.TclError):
                self.file_more_menu.grab_release()

    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择要处理的文件", filetypes=self._filetypes()
        )
        if paths:
            self._start_input_collection(paths, source_label="按钮")

    def _add_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择包含文件的文件夹")
        if not folder:
            return
        self._start_input_collection((folder,), source_label="文件夹")

    def _allowed_extensions(self) -> set[str]:
        if not self.current_operation:
            return set()
        return {
            item.lower() if item.startswith(".") else f".{item.lower()}"
            for item in self.current_operation.extensions
        }

    def _append_inputs(self, paths: Any) -> None:
        result = collect_input_files(
            paths,
            allowed_extensions=self._allowed_extensions(),
            existing_paths=self.input_paths,
        )
        self.input_paths.extend(result.files)
        self._refresh_file_tree()

    def _filter_existing_inputs(self) -> None:
        allowed = self._allowed_extensions()
        if allowed:
            self.input_paths = [
                path for path in self.input_paths if path.suffix.lower() in allowed
            ]
        self._refresh_file_tree()

    def _remove_selected(self) -> None:
        indexes = sorted(
            (int(item) for item in self.file_tree.selection()), reverse=True
        )
        for index in indexes:
            if 0 <= index < len(self.input_paths):
                del self.input_paths[index]
        self._refresh_file_tree()

    def _move_up(self) -> None:
        self._move_selected(-1)

    def _move_down(self) -> None:
        self._move_selected(1)

    def _move_selected(self, direction: int) -> None:
        selected = {
            int(item)
            for item in self.file_tree.selection()
            if str(item).isdigit() and 0 <= int(item) < len(self.input_paths)
        }
        if not selected:
            return
        if direction < 0:
            for index in range(1, len(self.input_paths)):
                if index in selected and index - 1 not in selected:
                    self.input_paths[index - 1], self.input_paths[index] = (
                        self.input_paths[index],
                        self.input_paths[index - 1],
                    )
                    selected.remove(index)
                    selected.add(index - 1)
        else:
            for index in range(len(self.input_paths) - 2, -1, -1):
                if index in selected and index + 1 not in selected:
                    self.input_paths[index], self.input_paths[index + 1] = (
                        self.input_paths[index + 1],
                        self.input_paths[index],
                    )
                    selected.remove(index)
                    selected.add(index + 1)
        self._refresh_file_tree()
        self.file_tree.selection_set(*(str(index) for index in sorted(selected)))

    def _clear_inputs(self) -> None:
        self.input_paths.clear()
        self._refresh_file_tree()

    def _refresh_file_tree(self) -> None:
        self.file_tree.delete(*self.file_tree.get_children())
        for index, path in enumerate(self.input_paths):
            try:
                size = self._format_size(path.stat().st_size)
            except OSError:
                size = "?"
            self.file_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(path.suffix.upper().lstrip("."), size, str(path)),
                tags=("even" if index % 2 == 0 else "odd",),
            )
        self.file_count_label.configure(
            text=(
                f"已添加 {len(self.input_paths)} 个文件"
                if self.input_paths
                else "尚未添加文件"
            )
        )
        self._initialize_image_parameters_from_source()
        self._schedule_image_preview(delay=50)

    @staticmethod
    def _natural_sort_key(path: Path) -> tuple[object, ...]:
        return natural_path_key(path)

    @staticmethod
    def _format_size(size: int) -> str:
        value = float(size)
        for unit in ("B", "KB", "MB", "GB"):
            if value < 1024 or unit == "GB":
                return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
            value /= 1024
        return f"{value:.1f} GB"

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.output_var.set(path)

    def _open_output(self) -> None:
        path = Path(self.output_var.get()).expanduser()
        try:
            path.mkdir(parents=True, exist_ok=True)
            os.startfile(path)  # type: ignore[attr-defined]
        except (OSError, ValueError) as exc:
            messagebox.showerror("无法打开", str(exc))

    def _collect_parameters(self) -> dict[str, Any]:
        values: dict[str, Any] = {}
        for key, variable in self.param_vars.items():
            value = variable.get()
            if key in self.choice_maps:
                value = self.choice_maps[key].get(str(value), value)
            values[key] = value
        return values

    def _start(self) -> None:
        operation = self.current_operation
        if operation is None:
            return
        capability = operation.capability()
        if not capability.runnable:
            presentation = unavailable_operation_presentation(operation, capability)
            messagebox.showerror(
                presentation.title,
                presentation.message,
                parent=self,
            )
            return
        if not self.input_paths:
            messagebox.showwarning("还没有文件", "请先添加要处理的文件。")
            return
        if self.worker and self.worker.is_alive():
            return
        self._log(f"\n[{operation.name}] 开始，输入 {len(self.input_paths)} 个文件")
        self._last_progress_value = 0.0
        self._progress_started_at = time.monotonic()
        self._progress_last_update_at = self._progress_started_at
        self._last_logged_progress_message = ""
        self._set_progress(
            0.0,
            progress_message("启动任务", total_files=len(self.input_paths)),
            write_log=False,
        )
        self.run_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.runner = TaskRunner()
        parameters = self._collect_parameters()
        inputs = list(self.input_paths)
        output = self.output_var.get()
        self.active_operation = operation
        self.active_output_dir = output
        self.active_inputs = inputs

        def progress(value: float, message: str) -> None:
            self.worker_queue.put(("progress", (value, message)))

        def work() -> None:
            try:
                assert self.runner is not None
                result = self.runner.run(
                    operation, inputs, output, parameters, progress
                )
                self.worker_queue.put(("success", result))
            except Exception as exc:
                self.worker_queue.put(("error", (exc, traceback.format_exc())))

        self.worker = threading.Thread(
            target=work, name="docuforge-worker", daemon=False
        )
        self.worker.start()

    def _cancel(self) -> None:
        if self.runner:
            self.runner.cancel()
            self.cancel_button.configure(state="disabled")
            self._progress_base_message = (
                "阶段：正在立即停止处理引擎 · 正在清理未完成输出 · 已完成输出会保留"
            )
            self._progress_last_update_at = time.monotonic()
            self._refresh_progress_label()
            self._log("已请求取消任务")

    def _set_progress(
        self,
        value: object,
        message: str,
        *,
        write_log: bool = True,
    ) -> None:
        self._set_progressbar_indeterminate(False)
        percent = progress_percent(value)
        if self._progress_started_at is not None:
            percent = max(self._last_progress_value, percent)
        self._last_progress_value = percent
        self._animate_progress_to(percent)
        self.progress_percent_label.configure(text=f"{percent:.0f}%")
        self._progress_base_message = str(message).strip() or "阶段：处理中"
        self._progress_last_update_at = time.monotonic()
        self._refresh_progress_label()
        if (
            write_log
            and self._progress_base_message != self._last_logged_progress_message
        ):
            self._log(self._progress_base_message)
            self._last_logged_progress_message = self._progress_base_message

    def _cancel_progress_animation(self) -> None:
        if self._progress_animation_job is None:
            return
        with contextlib.suppress(tk.TclError):
            self.after_cancel(self._progress_animation_job)
        self._progress_animation_job = None

    def _animate_progress_to(self, target: float) -> None:
        self._cancel_progress_animation()
        try:
            start = float(self.progress_var.get())
        except (tk.TclError, TypeError, ValueError):
            start = target
        target = min(100.0, max(0.0, float(target)))
        if abs(target - start) < 0.2 or not self.winfo_viewable():
            self.progress_var.set(target)
            return
        frames = 10

        def step(index: int) -> None:
            progress = ease_out_cubic(index / frames)
            value = start + (target - start) * progress
            try:
                self.progress_var.set(value)
            except tk.TclError:
                self._progress_animation_job = None
                return
            if index < frames:
                self._progress_animation_job = self.after(16, step, index + 1)
            else:
                self.progress_var.set(target)
                self._progress_animation_job = None

        step(1)

    def _refresh_progress_label(self) -> None:
        if self._progress_started_at is None:
            self._set_progressbar_indeterminate(False)
            self.progress_label.configure(text=self._progress_base_message)
            return
        now = time.monotonic()
        last_update = self._progress_last_update_at or self._progress_started_at
        seconds_since_update = now - last_update
        # Keep the bar strictly determinate and monotonic. Long conversion
        # stages may legitimately report no intermediate percentage; changing
        # to ttk's indeterminate mode made the indicator sweep backwards and
        # then jump forward again whenever a real progress event arrived.
        self._set_progressbar_indeterminate(False)
        self.progress_label.configure(
            text=progress_status_text(
                self._progress_base_message,
                now - self._progress_started_at,
                seconds_since_update=seconds_since_update,
            )
        )

    def _set_progressbar_indeterminate(self, active: bool) -> None:
        # Retained as a compatibility shim for callers that finish/reset the
        # task. The UI now always shows truthful one-way determinate progress.
        if self._progress_indeterminate:
            self.progressbar.stop()
        self._progress_indeterminate = False
        self.progressbar.configure(mode="determinate", maximum=100)

    def _refresh_progress_elapsed(self) -> None:
        try:
            if self._progress_started_at is not None:
                self._refresh_progress_label()
            if self.winfo_exists():
                self.after(1000, self._refresh_progress_elapsed)
        except tk.TclError:
            return

    def _poll_worker(self) -> None:
        pending_progress: tuple[object, str] | None = None
        processed = 0

        def flush_progress() -> None:
            nonlocal pending_progress
            if pending_progress is None:
                return
            value, message = pending_progress
            self._set_progress(value, message)
            pending_progress = None

        try:
            # Coalesce dense progress bursts and cap work per UI tick. This
            # keeps dragging, scrolling and the cancel button responsive even
            # when a fast batch produces thousands of status events.
            while processed < 250:
                kind, payload = self.worker_queue.get_nowait()
                processed += 1
                if kind == "progress":
                    pending_progress = payload
                    continue
                flush_progress()
                if kind == "input_scan":
                    source_label, result = payload
                    if not self._closing:
                        self._finish_input_collection(source_label, result)
                elif kind == "success":
                    self._finish_worker()
                    elapsed = payload.details.get("elapsed_seconds", "?")
                    self._progress_started_at = None
                    self._progress_last_update_at = None
                    state_text = {
                        "success": "处理成功",
                        "partial": "部分完成",
                        "failure": "处理失败",
                        "cancelled": "任务已取消",
                    }.get(payload.outcome, "处理完成")
                    self._set_progress(
                        1.0,
                        f"阶段：{state_text} · 批量任务已结束 · 用时 {elapsed} 秒",
                        write_log=False,
                    )
                    for path in payload.outputs:
                        self._log(f"输出：{path}")
                    for failure in payload.failed_inputs:
                        self._log(
                            f"未完成：{failure.input_path} · {failure.message}"
                        )
                    for warning in payload.warnings:
                        self._log(f"警告：{warning}")
                    if not self._closing:
                        TaskResultDialog(
                            self,
                            payload,
                            output_dir=(
                                self.active_output_dir or self.output_var.get()
                            ),
                            operation_name=(
                                self.active_operation.name
                                if self.active_operation
                                else "文件处理"
                            ),
                        )
                    self.active_operation = None
                    self.active_output_dir = None
                    self.active_inputs = []
                elif kind == "error":
                    self._finish_worker()
                    exc, trace = payload
                    self._progress_started_at = None
                    self._progress_last_update_at = None
                    if isinstance(exc, CancelledError):
                        self._progress_base_message = (
                            "任务已取消 · 未完成输出已清理 · 已完成输出已保留"
                        )
                        self._refresh_progress_label()
                        self._log(str(exc))
                        cancelled_result = exc.result or TaskResult(cancelled=True)
                        cancelled_result.cancelled = True
                        for path in cancelled_result.outputs:
                            self._log(f"已保留输出：{path}")
                        if not self._closing:
                            TaskResultDialog(
                                self,
                                cancelled_result,
                                output_dir=(
                                    self.active_output_dir or self.output_var.get()
                                ),
                                operation_name=(
                                    self.active_operation.name
                                    if self.active_operation
                                    else "文件处理"
                                ),
                            )
                    else:
                        self._progress_base_message = "处理失败 · 请查看运行日志"
                        self._refresh_progress_label()
                        self._log(trace)
                        friendly = (
                            str(exc)
                            if isinstance(exc, DocuForgeError)
                            else f"{type(exc).__name__}: {exc}"
                        )
                        if not self._closing:
                            failure_inputs = self.active_inputs or [Path("当前任务")]
                            failed_result = TaskResult(
                                failed_inputs=[
                                    TaskFailure(
                                        input_path=path,
                                        error_type=type(exc).__name__,
                                        message=friendly,
                                    )
                                    for path in failure_inputs
                                ]
                            )
                            TaskResultDialog(
                                self,
                                failed_result,
                                output_dir=(
                                    self.active_output_dir or self.output_var.get()
                                ),
                                operation_name=(
                                    self.active_operation.name
                                    if self.active_operation
                                    else "文件处理"
                                ),
                            )
                    self.active_operation = None
                    self.active_output_dir = None
                    self.active_inputs = []
        except queue.Empty:
            pass
        finally:
            flush_progress()
            self.after(30 if processed >= 250 else 120, self._poll_worker)

    def _finish_worker(self) -> None:
        self._set_progressbar_indeterminate(False)
        self.cancel_button.configure(state="disabled")
        if self._closing:
            return
        capability = (
            self.current_operation.capability() if self.current_operation else None
        )
        can_explain_missing_office = bool(
            capability
            and not capability.runnable
            and capability.engine == "Office 渲染器"
        )
        self.run_button.configure(
            state=(
                "normal"
                if capability
                and (capability.runnable or can_explain_missing_office)
                else "disabled"
            )
        )

    def _dispose_image_preview_resources(self) -> None:
        self._image_preview_generation += 1
        if self._image_preview_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._image_preview_job)
            self._image_preview_job = None
        self._clear_image_preview_images()

    def _on_close(self) -> None:
        if not self.worker or not self.worker.is_alive():
            self._closing = True
            self._dispose_image_preview_resources()
            self.destroy()
            return
        if self._closing:
            return
        if not messagebox.askokcancel(
            "任务仍在运行",
            "现在关闭会先请求取消，并等待当前文件或外部引擎安全退出。是否继续？",
        ):
            return
        self._closing = True
        self._dispose_image_preview_resources()
        if self.runner:
            self.runner.cancel()
        self.run_button.configure(state="disabled")
        self.cancel_button.configure(state="disabled")
        self._progress_base_message = "阶段：正在取消并等待处理引擎安全退出"
        self._progress_last_update_at = time.monotonic()
        self._refresh_progress_label()
        self._log("关闭请求已收到；正在等待当前处理引擎退出")
        self.after(150, self._wait_for_close)

    def _wait_for_close(self) -> None:
        if self.worker and self.worker.is_alive():
            self.after(150, self._wait_for_close)
            return
        self._dispose_image_preview_resources()
        self.destroy()

    def _log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message.rstrip() + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


def launch() -> None:
    from .registry import get_operations

    _enable_windows_dpi_awareness()
    app = DocuForgeApp(get_operations())
    app.mainloop()


if __name__ == "__main__":
    launch()
