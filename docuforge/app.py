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

UI_THEME_LABELS = {
    "tech": "科技风",
    "original": "原版画风",
    "cream": "奶油画风",
}

# UI themes deliberately contain colours only.  Geometry, padding, radii,
# fonts, catalogue structure and motion timings stay shared by every theme.
UI_THEME_PALETTES: dict[str, dict[str, str]] = {
    "tech": {
        "BG": "#07101D",
        "PANEL": "#0D192A",
        "PANEL_ALT": "#112238",
        "SIDEBAR_BG": "#091626",
        "TEXT": "#EAF2FF",
        "MUTED": "#91A5BF",
        "ACCENT": "#4C91FF",
        "ACCENT_DARK": "#2F6FDB",
        "ACCENT_SOFT": "#173A61",
        "PINK": "#8B7CF6",
        "CYAN": "#35D3E6",
        "ORANGE": "#F4B95F",
        "LIME": "#43D5A0",
        "SUCCESS": "#39D98A",
        "WARNING": "#F4BF5F",
        "DANGER": "#FF6B7B",
        "BORDER": "#203752",
        "SHADOW": "#030710",
        "SAGE": "#18283D",
        "CREAM_BLUE": "#102B46",
        "CREAM_YELLOW": "#172236",
        "DUSTY_BLUE": "#58B8FF",
        "CARD_BLUE": "#10243A",
        "CARD_SAGE": "#102A30",
        "CARD_YELLOW": "#121E2E",
        "DROP_SURFACE": "#0C2135",
        "INPUT_SURFACE": "#081523",
        "HOVER_SURFACE": "#18314D",
        "PRESSED_SURFACE": "#10263F",
        "DISABLED_SURFACE": "#101B2B",
        "DISABLED_TEXT": "#62748B",
        "DANGER_SURFACE": "#321A28",
        "DANGER_HOVER": "#462131",
        "SUCCESS_SURFACE": "#103126",
        "WARNING_SURFACE": "#382A18",
        "GRID_LINE": "#102239",
        "GLOW_BLUE": "#0B1D33",
        "GLOW_PURPLE": "#151B38",
        "ON_ACCENT": "#EAF2FF",
        "SELECTED_ACTIVE_SURFACE": "#1D4B78",
        "ACCENT_PRESSED": "#245CBD",
        "DANGER_BORDER": "#6B2A3B",
        "DANGER_PRESSED": "#25131E",
        "DANGER_PULSE_BORDER": "#7B3146",
        "ACCENT_PULSE_BORDER": "#9BEAF3",
        "SELECTED_PULSE_BORDER": "#89E8F2",
        "GLOW_THIRD": "#0A2231",
        "ORBIT_LINE": "#17344F",
        "ORBIT_ACCENT": "#214B6C",
        "DROP_BORDER": "#2A567A",
        "TREE_ODD": "#102237",
        "DROP_PORTAL": "#2A638A",
        "DROP_SCAN": "#12304A",
        "DROP_RING": "#173F5F",
        "PARTICLE_MOON": "#F6D77B",
        "PARTICLE_STAR": "#FFD36A",
        "SCROLL_THUMB": "#58B8FF",
        "SCROLL_ACTIVE": "#2F6FDB",
        "SCROLL_TROUGH": "#18283D",
        "PROGRESS_TROUGH": "#18283D",
        "TREE_SELECTED": "#173A61",
        "NAV_BORDER": "#203752",
    },
    "original": {
        "BG": "#F3F6FB",
        "PANEL": "#FFFFFF",
        "PANEL_ALT": "#F8FAFD",
        "SIDEBAR_BG": "#F7F9FD",
        "TEXT": "#172033",
        "MUTED": "#667085",
        "ACCENT": "#4F6BED",
        "ACCENT_DARK": "#3B55D9",
        "ACCENT_SOFT": "#EEF2FF",
        "PINK": "#7E8FEA",
        "CYAN": "#5D77F0",
        "ORANGE": "#D79518",
        "LIME": "#0F9D70",
        "SUCCESS": "#0F8F6B",
        "WARNING": "#C86B12",
        "DANGER": "#C83E4D",
        "BORDER": "#DDE4EF",
        "SHADOW": "#E8EDF5",
        "SAGE": "#EDF1F7",
        "CREAM_BLUE": "#EEF2FF",
        "CREAM_YELLOW": "#F8FAFD",
        "DUSTY_BLUE": "#95A7F7",
        "CARD_BLUE": "#F4F6FA",
        "CARD_SAGE": "#EAF8F3",
        "CARD_YELLOW": "#FFFFFF",
        "DROP_SURFACE": "#F8FAFD",
        "INPUT_SURFACE": "#FFFFFF",
        "HOVER_SURFACE": "#EEF2FF",
        "PRESSED_SURFACE": "#E0E7FF",
        "DISABLED_SURFACE": "#F2F4F8",
        "DISABLED_TEXT": "#98A2B3",
        "DANGER_SURFACE": "#FFF1F2",
        "DANGER_HOVER": "#FFE4E6",
        "SUCCESS_SURFACE": "#EAF8F3",
        "WARNING_SURFACE": "#FFF7ED",
        "GRID_LINE": "#E8EDF5",
        "GLOW_BLUE": "#EAF0FF",
        "GLOW_PURPLE": "#F1EEFF",
        "ON_ACCENT": "#FFFFFF",
        "SELECTED_ACTIVE_SURFACE": "#E0E7FF",
        "ACCENT_PRESSED": "#3B55D9",
        "DANGER_BORDER": "#FECDD3",
        "DANGER_PRESSED": "#FDA4AF",
        "DANGER_PULSE_BORDER": "#FB7185",
        "ACCENT_PULSE_BORDER": "#B8C5F7",
        "SELECTED_PULSE_BORDER": "#AFC0FA",
        "GLOW_THIRD": "#EDF1F7",
        "ORBIT_LINE": "#DDE4EF",
        "ORBIT_ACCENT": "#C7D2FE",
        "DROP_BORDER": "#D7DEFF",
        "TREE_ODD": "#F8FAFD",
        "DROP_PORTAL": "#95A7F7",
        "DROP_SCAN": "#E4E9FF",
        "DROP_RING": "#C7D2FE",
        "PARTICLE_MOON": "#D79518",
        "PARTICLE_STAR": "#E8B84E",
        "SCROLL_THUMB": "#B8C2D4",
        "SCROLL_ACTIVE": "#94A3B8",
        "SCROLL_TROUGH": "#EDF1F7",
        "PROGRESS_TROUGH": "#E8ECF5",
        "TREE_SELECTED": "#E4E9FF",
        "NAV_BORDER": "#D7DEFF",
    },
    "cream": {
        "BG": "#F7F0E7",
        "PANEL": "#FFFDFC",
        "PANEL_ALT": "#FAF1E8",
        "SIDEBAR_BG": "#F5E8DD",
        "TEXT": "#3E312D",
        "MUTED": "#796862",
        "ACCENT": "#D47C6A",
        "ACCENT_DARK": "#B75F52",
        "ACCENT_SOFT": "#F7DED4",
        "PINK": "#D69CB0",
        "CYAN": "#568F85",
        "ORANGE": "#D99650",
        "LIME": "#79A97D",
        "SUCCESS": "#3C8B6F",
        "WARNING": "#A96729",
        "DANGER": "#B94F5D",
        "BORDER": "#E5D1C1",
        "SHADOW": "#E9DDD2",
        "SAGE": "#EEE0D3",
        "CREAM_BLUE": "#E7F1EE",
        "CREAM_YELLOW": "#FFF0D2",
        "DUSTY_BLUE": "#79A9A1",
        "CARD_BLUE": "#F5EBE2",
        "CARD_SAGE": "#EDF3ED",
        "CARD_YELLOW": "#FFF8EE",
        "DROP_SURFACE": "#FFF6ED",
        "INPUT_SURFACE": "#FFFCF8",
        "HOVER_SURFACE": "#FBE8DE",
        "PRESSED_SURFACE": "#F4D4C8",
        "DISABLED_SURFACE": "#EFE5DC",
        "DISABLED_TEXT": "#9B867D",
        "DANGER_SURFACE": "#FCE8E8",
        "DANGER_HOVER": "#F8D9DC",
        "SUCCESS_SURFACE": "#E8F4ED",
        "WARNING_SURFACE": "#FFF0D8",
        "GRID_LINE": "#ECDCCD",
        "GLOW_BLUE": "#F4E3D7",
        "GLOW_PURPLE": "#F0E1E9",
        "ON_ACCENT": "#FFFDFC",
        "SELECTED_ACTIVE_SURFACE": "#E7A193",
        "ACCENT_PRESSED": "#A94F45",
        "DANGER_BORDER": "#E7ABB1",
        "DANGER_PRESSED": "#A84351",
        "DANGER_PULSE_BORDER": "#E28C96",
        "ACCENT_PULSE_BORDER": "#F4C2B5",
        "SELECTED_PULSE_BORDER": "#EFB0A4",
        "GLOW_THIRD": "#F1E6D9",
        "ORBIT_LINE": "#E7CFC0",
        "ORBIT_ACCENT": "#D8B5A1",
        "DROP_BORDER": "#DFC1AD",
        "TREE_ODD": "#FCF4EA",
        "DROP_PORTAL": "#D89B8F",
        "DROP_SCAN": "#EBD5C4",
        "DROP_RING": "#E0B9A4",
        "PARTICLE_MOON": "#E6B85C",
        "PARTICLE_STAR": "#F0C56A",
        "SCROLL_THUMB": "#C99F8E",
        "SCROLL_ACTIVE": "#B75F52",
        "SCROLL_TROUGH": "#F1E6DC",
        "PROGRESS_TROUGH": "#EEDFD4",
        "TREE_SELECTED": "#F4DDD3",
        "NAV_BORDER": "#E7CABD",
    },
}
UI_THEME_PALETTE_KEYS = tuple(UI_THEME_PALETTES["tech"])
CURRENT_UI_THEME = "tech"


def normalize_ui_theme(value: object) -> str:
    """Normalize persisted or user-selected UI theme identifiers safely."""

    theme = str(value or "tech").strip().casefold()
    return theme if theme in UI_THEME_PALETTES else "tech"


def _activate_ui_theme(value: object) -> str:
    """Expose one palette through the legacy colour names used by the UI."""

    global CURRENT_UI_THEME
    theme = normalize_ui_theme(value)
    globals().update(UI_THEME_PALETTES[theme])
    CURRENT_UI_THEME = theme
    return theme


def _style_combobox_popdown(widget: ttk.Combobox) -> None:
    """Apply the active palette whenever an existing combobox is posted."""

    try:
        popdown = widget.tk.call("ttk::combobox::PopdownWindow", widget._w)
        listbox = f"{popdown}.f.l"
        widget.tk.call(
            listbox,
            "configure",
            "-background",
            INPUT_SURFACE,
            "-foreground",
            TEXT,
            "-selectbackground",
            ACCENT_SOFT,
            "-selectforeground",
            TEXT,
            "-highlightbackground",
            BORDER,
            "-highlightcolor",
            ACCENT,
            "-highlightthickness",
            1,
            "-relief",
            "flat",
            "-borderwidth",
            0,
        )
    except tk.TclError:
        return


def ui_preferences_path() -> Path:
    """Return a per-user writable preferences file outside the portable app."""

    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "LayoutLoom" / "preferences.json"
    return Path.home() / ".layoutloom" / "preferences.json"


def normalize_particle_effects_enabled(value: object) -> bool:
    """Normalize a persisted particle toggle without treating ``"false"`` as true."""

    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value != 0
    normalized = str(value if value is not None else "on").strip().casefold()
    if normalized in {"0", "false", "off", "no", "disabled", "关闭", "关"}:
        return False
    return True


def load_ui_preferences(path: str | Path | None = None) -> dict[str, object]:
    """Load supported UI preferences and recover safely from damaged JSON."""

    target = Path(path) if path is not None else ui_preferences_path()
    try:
        payload = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        payload = {}
    if not isinstance(payload, dict):
        payload = {}
    return {
        "ui_theme": normalize_ui_theme(payload.get("ui_theme")),
        "particle_effects": normalize_particle_effects_enabled(
            payload.get("particle_effects")
        ),
    }


def save_ui_preferences(
    preferences: dict[str, object],
    path: str | Path | None = None,
) -> Path:
    """Persist supported UI preferences atomically."""

    target = Path(path) if path is not None else ui_preferences_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.{os.getpid()}.tmp")
    payload = {
        "ui_theme": normalize_ui_theme(preferences.get("ui_theme")),
        "particle_effects": normalize_particle_effects_enabled(
            preferences.get("particle_effects")
        ),
    }
    try:
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        os.replace(temporary, target)
    finally:
        with contextlib.suppress(OSError):
            if temporary.exists():
                temporary.unlink()
    return target


_activate_ui_theme("tech")
MOTION_MODE_LABELS = {
    "rich": "完整界面动效",
    "light": "省资源界面动效",
    "off": "关闭界面动效",
}

OFFICE_COMPATIBILITY_NOTICE = (
    "已对桌面版 WPS Office 与 Microsoft Office 的真实 COM 引擎完成独立识别和定向适配；"
    "自动模式按 WPS → Microsoft Office → LibreOffice 尝试，也可锁定单一引擎。"
    "WPS 与 Microsoft Office 会保留各自原生排版特性；LibreOffice 作为兼容回退，"
    "复杂版式建议优先选择本机表现更好的 WPS 或 Microsoft Office。"
)
# Backwards-compatible public name retained for older integrations/tests.
WPS_COMPATIBILITY_NOTICE = OFFICE_COMPATIBILITY_NOTICE

OFFICE_ENGINE_VALUES = (
    "auto",
    "wps",
    "microsoft_office",
    "libreoffice",
)
OFFICE_ENGINE_PARAMETER_KEYS = (
    "engine",
    "renderer",
    "verification_engine",
)
OFFICE_ENGINE_MENU_LABELS = {
    "auto": "自动（WPS → Microsoft Office → LibreOffice）",
    "wps": "WPS Office COM",
    "microsoft_office": "Microsoft Office COM",
    "libreoffice": "LibreOffice",
}
OFFICE_ENGINE_SHORT_LABELS = {
    "auto": "自动",
    "wps": "WPS",
    "microsoft_office": "Microsoft",
    "libreoffice": "LibreOffice",
}


def office_engine_parameter_spec(operation: Operation | None) -> ParameterSpec | None:
    """Return a genuine Office engine selector without touching name collisions."""

    if operation is None:
        return None
    by_key = {spec.key: spec for spec in operation.parameters}
    required_values = set(OFFICE_ENGINE_VALUES)
    for key in OFFICE_ENGINE_PARAMETER_KEYS:
        spec = by_key.get(key)
        if spec is None or spec.kind != "choice":
            continue
        values = {str(value) for value, _label in spec.choices}
        if required_values.issubset(values):
            return spec
    return None


def office_engine_button_text(
    value: str,
    *,
    compact: bool = False,
    active: bool = False,
) -> str:
    label = OFFICE_ENGINE_SHORT_LABELS.get(str(value), str(value) or "自动")
    if compact:
        return f"引擎：{label}  ▾"
    return f"{'当前' if active else '默认'}引擎：{label}  ▾"


@dataclass(frozen=True)
class TaskResultPresentation:
    title: str
    subtitle: str
    icon: str
    gradient_start: str
    gradient_end: str
    accent: str


@dataclass(frozen=True)
class MotionTiming:
    """A deterministic animation budget consumed by Tk ``after`` jobs."""

    frames: int
    step_ms: int
    enabled: bool


@dataclass(frozen=True)
class TransitionFrame:
    """One deterministic frame of a local sliding indicator."""

    progress: float
    scale: float


@dataclass(frozen=True)
class ParticleFrame:
    """One frame of the slower three-stage click-particle motion."""

    spread: float
    opacity: float
    scale: float
    curve: float


@dataclass(frozen=True)
class ParticleSpec:
    """One deterministic symbol in a themed click-particle burst."""

    kind: str
    tangent: float
    drift: float
    colour_role: str
    size: float


@dataclass(frozen=True)
class DoodlePlacement:
    """A deterministic decorative mark restricted to a background edge band."""

    kind: str
    box: tuple[float, float, float, float]
    colour: str


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
                "页织工坊已分别适配桌面版 WPS Office 与 Microsoft Office 的真实 COM "
                "引擎；也可使用自行安装的 LibreOffice 作为兼容回退。请至少安装其中一种。"
                "若已经安装 WPS 或 Microsoft Office 仍看到此提示，请修复对应安装，确认 "
                "COM 自动化接口已正确注册，然后彻底退出并重新启动页织工坊。"
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


def result_dialog_geometry(
    parent_width: int,
    parent_height: int,
) -> tuple[int, int]:
    """Choose a result-dialog size that leaves room for all action labels."""

    width = min(820, max(650, round(max(1, int(parent_width)) * 0.66)))
    height = min(700, max(520, round(max(1, int(parent_height)) * 0.72)))
    return width, height

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
            style="Timeline.Horizontal.TScale",
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
            style="Timeline.Horizontal.TScale",
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
            postcommand=lambda: _style_combobox_popdown(self.method_combo),
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
            selectbackground=ACCENT_SOFT,
            selectforeground=TEXT,
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


def ease_out_back(value: float, overshoot: float = 1.35) -> float:
    """Return a small spring-like overshoot for short, non-distracting motion."""

    progress = min(1.0, max(0.0, float(value)))
    strength = max(0.0, float(overshoot))
    shifted = progress - 1.0
    return 1.0 + (strength + 1.0) * shifted**3 + strength * shifted**2


def normalize_motion_mode(value: object) -> str:
    """Normalize persisted or user-provided motion preferences safely."""

    mode = str(value or "rich").strip().casefold()
    return mode if mode in MOTION_MODE_LABELS else "rich"


def motion_frame_delay(
    mode: object,
    *,
    busy: bool = False,
    minimized: bool = False,
) -> int:
    """Choose a conservative frame interval without taxing conversion tasks."""

    normalized = normalize_motion_mode(mode)
    if normalized == "off":
        return 0
    if minimized:
        return 650
    if normalized == "light":
        return 280 if busy else 160
    return 180 if busy else 72


def motion_effect_timing(
    mode: object,
    effect: str,
    *,
    busy: bool = False,
    minimized: bool = False,
) -> MotionTiming:
    """Return a small bounded budget for one finite or ambient UI effect."""

    normalized = normalize_motion_mode(mode)
    effect_name = str(effect or "ambient").strip().casefold()
    if normalized == "off":
        return MotionTiming(0, 0, False)
    if minimized:
        # Finite feedback should snap while hidden.  Ambient work only needs a
        # slow wake-up check and must never burn frames in the background.
        if effect_name == "ambient":
            return MotionTiming(1, 650, False)
        return MotionTiming(0, 0, False)
    effective_mode = "light" if busy and normalized == "rich" else normalized
    plans = {
        "click": {
            "rich": MotionTiming(5, 26, True),
            "light": MotionTiming(3, 30, True),
        },
        "transition": {
            "rich": MotionTiming(10, 16, True),
            "light": MotionTiming(5, 22, True),
        },
        "title": {
            "rich": MotionTiming(9, 18, True),
            "light": MotionTiming(5, 24, True),
        },
        "progress": {
            "rich": MotionTiming(10, 15, True),
            "light": MotionTiming(6, 22, True),
        },
        "dialog": {
            "rich": MotionTiming(7, 18, True),
            "light": MotionTiming(3, 24, True),
        },
        "ambient": {
            "rich": MotionTiming(8, 110, True),
            "light": MotionTiming(1, 320, False),
        },
    }
    selected = plans.get(effect_name, plans["ambient"])
    return selected[effective_mode]


def particle_effect_button_text(value: object, *, compact: bool = False) -> str:
    """Return an explicit two-state label for the user-facing particle toggle."""

    enabled = normalize_particle_effects_enabled(value)
    if compact:
        return f"粒子：{'开' if enabled else '关'}"
    return f"粒子动效：{'开' if enabled else '关'}"


def motion_button_text(value: object, *, compact: bool = False) -> str:
    """Backward-compatible alias for integrations using the former helper name."""

    return particle_effect_button_text(value, compact=compact)


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


def composite_hex_colour(background: str, foreground: str, alpha: object) -> str:
    """Simulate Canvas transparency by pre-mixing two ``#RRGGBB`` colours."""

    try:
        opacity = float(alpha)
    except (TypeError, ValueError, OverflowError):
        opacity = 0.0
    if not math.isfinite(opacity):
        opacity = 0.0
    return interpolate_hex_colour(background, foreground, min(1.0, max(0.0, opacity)))


def result_dialog_entrance_plan(mode: object) -> tuple[float, ...]:
    """Return a monotonic, mode-aware opacity plan for the result dialog."""

    normalized = normalize_motion_mode(mode)
    if normalized == "off":
        return (1.0,)
    timing = motion_effect_timing(normalized, "dialog")
    start = 0.88 if normalized == "rich" else 0.94
    values = tuple(
        start + (1.0 - start) * ease_out_cubic(index / max(1, timing.frames))
        for index in range(timing.frames + 1)
    )
    return values[:-1] + (1.0,)


def q_bounce_transition_plan(mode: object, *, busy: bool = False) -> tuple[TransitionFrame, ...]:
    """Return a smooth, monotonic plan for page/category indicators.

    The historical function name is retained for compatibility.  The former
    overshoot made the indicator jump past its destination and then reverse,
    which looked like a stalled page switch rather than a soft transition.
    """

    normalized = normalize_motion_mode(mode)
    if normalized == "off":
        return (TransitionFrame(1.0, 1.0),)
    frame_count = 5 if busy or normalized == "light" else 10
    pulse = 0.0 if busy or normalized == "light" else 0.028
    frames: list[TransitionFrame] = []
    for index in range(frame_count + 1):
        linear = index / max(1, frame_count)
        progress = linear * linear * (3.0 - 2.0 * linear)
        scale = 1.0 + math.sin(math.pi * linear) * pulse
        frames.append(TransitionFrame(progress, scale))
    frames[-1] = TransitionFrame(1.0, 1.0)
    return tuple(frames)


def click_particle_specs(variant: int = 0) -> tuple[ParticleSpec, ...]:
    """Return one of three deterministic, theme-aware particle arrangements."""

    variants = (
        (
            ParticleSpec("star", -0.92, -0.22, "ORANGE", 6.8),
            ParticleSpec("sparkle", -0.72, 0.18, "CYAN", 5.8),
            ParticleSpec("dot", -0.50, 0.48, "PINK", 3.8),
            ParticleSpec("moon", -0.28, -0.34, "PARTICLE_MOON", 8.2),
            ParticleSpec("diamond", -0.05, 0.26, "ACCENT", 5.6),
            ParticleSpec("star", 0.20, -0.48, "PARTICLE_STAR", 6.4),
            ParticleSpec("ring", 0.42, 0.22, "CYAN", 6.8),
            ParticleSpec("sparkle", 0.65, -0.18, "LIME", 5.8),
            ParticleSpec("hex", 0.86, -0.52, "ACCENT", 5.8),
            ParticleSpec("comet", -0.62, 0.62, "PINK", 5.2),
            ParticleSpec("dot", 0.92, 0.48, "ORANGE", 3.4),
        ),
        (
            ParticleSpec("ring", -0.90, -0.32, "CYAN", 7.0),
            ParticleSpec("hex", -0.70, 0.42, "ACCENT", 6.0),
            ParticleSpec("sparkle", -0.46, -0.52, "PARTICLE_STAR", 6.2),
            ParticleSpec("diamond", -0.24, 0.18, "PINK", 5.6),
            ParticleSpec("dot", 0.00, 0.58, "LIME", 3.6),
            ParticleSpec("comet", 0.24, -0.38, "CYAN", 5.4),
            ParticleSpec("hex", 0.46, 0.30, "ORANGE", 5.8),
            ParticleSpec("star", 0.68, -0.58, "PARTICLE_STAR", 6.4),
            ParticleSpec("sparkle", 0.88, 0.08, "PINK", 5.8),
            ParticleSpec("moon", -0.58, 0.66, "PARTICLE_MOON", 7.8),
            ParticleSpec("dot", 0.76, 0.56, "ACCENT", 3.5),
        ),
        (
            ParticleSpec("comet", -0.92, 0.12, "ORANGE", 5.8),
            ParticleSpec("sparkle", -0.70, -0.50, "CYAN", 6.0),
            ParticleSpec("moon", -0.46, 0.44, "PARTICLE_MOON", 8.0),
            ParticleSpec("dot", -0.22, -0.28, "LIME", 3.6),
            ParticleSpec("star", 0.00, 0.56, "PARTICLE_STAR", 6.6),
            ParticleSpec("ring", 0.24, -0.46, "PINK", 6.8),
            ParticleSpec("diamond", 0.46, 0.22, "ACCENT", 5.8),
            ParticleSpec("comet", 0.68, -0.24, "CYAN", 5.4),
            ParticleSpec("hex", 0.88, 0.50, "ORANGE", 5.8),
            ParticleSpec("sparkle", -0.62, 0.68, "PINK", 5.6),
            ParticleSpec("dot", 0.64, 0.64, "LIME", 3.5),
        ),
    )
    return variants[int(variant) % len(variants)]


def click_particle_frame_plan(frame_count: int = 19) -> tuple[ParticleFrame, ...]:
    """Build a visible reveal, outward drift and fully transparent settle plan."""

    count = max(3, int(frame_count))
    frames: list[ParticleFrame] = []
    for index in range(count):
        progress = index / max(1, count - 1)
        if progress <= 0.18:
            local = progress / 0.18
            spread = 0.08 * ease_out_cubic(local)
            opacity = 0.38 + 0.62 * ease_out_cubic(local)
            scale = 0.64 + 0.48 * ease_out_back(local, overshoot=0.52)
        elif progress <= 0.72:
            local = (progress - 0.18) / 0.54
            spread = 0.08 + 0.72 * (local * local * (3.0 - 2.0 * local))
            opacity = 1.0 - 0.18 * local
            scale = 1.08 - 0.14 * local
        else:
            local = (progress - 0.72) / 0.28
            spread = 0.80 + 0.20 * ease_out_cubic(local)
            opacity = 0.82 * (1.0 - local) ** 2
            scale = max(0.28, 0.94 * (1.0 - 0.64 * ease_out_cubic(local)))
        frames.append(
            ParticleFrame(
                min(1.0, max(0.0, spread)),
                min(1.0, max(0.0, opacity)),
                max(0.24, scale),
                math.sin(math.pi * progress),
            )
        )
    frames[-1] = ParticleFrame(1.0, 0.0, 0.24, 0.0)
    return tuple(frames)


def background_watermark_plan(
    width: int,
    height: int,
    layout_mode: str,
) -> tuple[DoodlePlacement, ...]:
    """Build a fixed low-density plan that never enters the central workspace."""

    w = max(0, int(width))
    h = max(0, int(height))
    if w < 320 or h < 240:
        return ()
    mode = layout_mode if layout_mode in {"narrow", "compact", "wide"} else "narrow"
    count = {"narrow": 4, "compact": 9, "wide": 15}[mode]
    kinds = ("star", "file", "node", "star", "satellite", "file")
    colours = (
        composite_hex_colour(BG, CYAN, 0.075),
        composite_hex_colour(BG, ACCENT, 0.085),
        composite_hex_colour(BG, PINK, 0.070),
    )
    placements: list[DoodlePlacement] = []
    golden = 0.61803398875
    for index in range(count):
        t = (0.137 + index * golden) % 1.0
        size = 14 + (index * 5) % (14 if mode == "wide" else 10)
        edge = index % 4
        margin = 18 + (index * 7) % 30
        if edge == 0:
            x = margin + t * max(1, w - size - margin * 2)
            y = 8 + (index % 3) * 11
        elif edge == 1:
            x = w - size - 8 - (index % 2) * 13
            y = margin + t * max(1, h - size - margin * 2)
        elif edge == 2:
            x = margin + t * max(1, w - size - margin * 2)
            y = h - size - 8 - (index % 2) * 13
        else:
            x = 8 + (index % 2) * 13
            y = margin + t * max(1, h - size - margin * 2)
        placements.append(
            DoodlePlacement(
                kinds[index % len(kinds)],
                (float(x), float(y), float(x + size), float(y + size)),
                colours[index % len(colours)],
            )
        )
    return tuple(placements)


def _canvas_doodle_line_colour(
    background: str,
    ink: str | None = None,
    *,
    strength: float = 0.14,
) -> str:
    """Return calm simulated-transparent line art suitable for Tk Canvas."""

    return composite_hex_colour(background, ink or CYAN, strength)


def draw_rocket_doodle(
    canvas: tk.Canvas,
    box: tuple[float, float, float, float],
    *,
    colour: str,
    accent: str | None = None,
    tags: tuple[str, ...] = ("doodle",),
) -> None:
    """Draw a tiny native Canvas rocket inside an already reserved safe box."""

    x1, y1, x2, y2 = box
    width = max(1.0, x2 - x1)
    height = max(1.0, y2 - y1)

    def point(rx: float, ry: float) -> tuple[float, float]:
        return (x1 + width * rx, y1 + height * ry)

    body = (
        *point(0.50, 0.04),
        *point(0.30, 0.28),
        *point(0.36, 0.72),
        *point(0.43, 0.84),
        *point(0.57, 0.84),
        *point(0.64, 0.72),
        *point(0.70, 0.28),
    )
    canvas.create_polygon(
        body,
        fill="",
        outline=colour,
        width=1,
        smooth=True,
        splinesteps=20,
        tags=tags,
    )
    canvas.create_polygon(
        (*point(0.35, 0.56), *point(0.14, 0.79), *point(0.39, 0.73)),
        fill="",
        outline=colour,
        width=1,
        tags=tags,
    )
    canvas.create_polygon(
        (*point(0.65, 0.56), *point(0.86, 0.79), *point(0.61, 0.73)),
        fill="",
        outline=colour,
        width=1,
        tags=tags,
    )
    canvas.create_oval(
        x1 + width * 0.41,
        y1 + height * 0.31,
        x1 + width * 0.59,
        y1 + height * 0.49,
        outline=accent or colour,
        width=1,
        tags=tags,
    )
    canvas.create_line(
        *point(0.44, 0.84),
        *point(0.38, 0.98),
        fill=accent or colour,
        width=1,
        smooth=True,
        tags=tags,
    )
    canvas.create_line(
        *point(0.56, 0.84),
        *point(0.62, 0.98),
        fill=accent or colour,
        width=1,
        smooth=True,
        tags=tags,
    )


def draw_laptop_doodle(
    canvas: tk.Canvas,
    box: tuple[float, float, float, float],
    *,
    colour: str,
    accent: str | None = None,
    tags: tuple[str, ...] = ("doodle",),
) -> None:
    """Draw a minimal laptop glyph inside a non-content margin."""

    x1, y1, x2, y2 = box
    width = max(1.0, x2 - x1)
    height = max(1.0, y2 - y1)
    canvas.create_rectangle(
        x1 + width * 0.08,
        y1 + height * 0.08,
        x1 + width * 0.92,
        y1 + height * 0.65,
        outline=colour,
        width=1,
        tags=tags,
    )
    code_colour = accent or colour
    for y_ratio, right_ratio in ((0.26, 0.64), (0.38, 0.79), (0.50, 0.58)):
        canvas.create_line(
            x1 + width * 0.24,
            y1 + height * y_ratio,
            x1 + width * right_ratio,
            y1 + height * y_ratio,
            fill=code_colour,
            width=1,
            tags=tags,
        )
    canvas.create_line(
        x1 + width * 0.45,
        y1 + height * 0.65,
        x1 + width * 0.55,
        y1 + height * 0.80,
        fill=colour,
        width=1,
        tags=tags,
    )
    canvas.create_line(
        x1 + width * 0.28,
        y1 + height * 0.82,
        x1 + width * 0.72,
        y1 + height * 0.82,
        fill=colour,
        width=1,
        tags=tags,
    )


def draw_file_doodle(
    canvas: tk.Canvas,
    box: tuple[float, float, float, float],
    *,
    colour: str,
    tags: tuple[str, ...] = ("doodle",),
) -> None:
    x1, y1, x2, y2 = box
    width = max(1.0, x2 - x1)
    height = max(1.0, y2 - y1)
    fold = min(width, height) * 0.24
    canvas.create_polygon(
        x1,
        y1,
        x2 - fold,
        y1,
        x2,
        y1 + fold,
        x2,
        y2,
        x1,
        y2,
        fill="",
        outline=colour,
        width=1,
        tags=tags,
    )
    canvas.create_line(
        x2 - fold,
        y1,
        x2 - fold,
        y1 + fold,
        x2,
        y1 + fold,
        fill=colour,
        width=1,
        tags=tags,
    )
    for ratio in (0.52, 0.70):
        canvas.create_line(
            x1 + width * 0.22,
            y1 + height * ratio,
            x2 - width * 0.18,
            y1 + height * ratio,
            fill=colour,
            width=1,
            tags=tags,
        )


def draw_satellite_doodle(
    canvas: tk.Canvas,
    box: tuple[float, float, float, float],
    *,
    colour: str,
    tags: tuple[str, ...] = ("doodle",),
) -> None:
    x1, y1, x2, y2 = box
    width = max(1.0, x2 - x1)
    height = max(1.0, y2 - y1)
    canvas.create_rectangle(
        x1 + width * 0.37,
        y1 + height * 0.32,
        x1 + width * 0.63,
        y1 + height * 0.68,
        outline=colour,
        width=1,
        tags=tags,
    )
    for left, right in ((0.03, 0.34), (0.66, 0.97)):
        canvas.create_rectangle(
            x1 + width * left,
            y1 + height * 0.24,
            x1 + width * right,
            y1 + height * 0.76,
            outline=colour,
            width=1,
            tags=tags,
        )
    canvas.create_line(
        x1 + width * 0.50,
        y1 + height * 0.32,
        x1 + width * 0.50,
        y1 + height * 0.08,
        fill=colour,
        width=1,
        tags=tags,
    )
    canvas.create_oval(
        x1 + width * 0.45,
        y1,
        x1 + width * 0.55,
        y1 + height * 0.10,
        fill=colour,
        outline="",
        tags=tags,
    )


def draw_starlet(
    canvas: tk.Canvas,
    box: tuple[float, float, float, float],
    *,
    colour: str,
    tags: tuple[str, ...] = ("doodle",),
) -> None:
    x1, y1, x2, y2 = box
    cx = (x1 + x2) / 2
    cy = (y1 + y2) / 2
    width = max(2.0, x2 - x1)
    height = max(2.0, y2 - y1)
    canvas.create_line(cx, y1, cx, y2, fill=colour, width=1, tags=tags)
    canvas.create_line(x1, cy, x2, cy, fill=colour, width=1, tags=tags)
    canvas.create_line(
        cx - width * 0.24,
        cy - height * 0.24,
        cx + width * 0.24,
        cy + height * 0.24,
        fill=colour,
        width=1,
        tags=tags,
    )
    canvas.create_line(
        cx + width * 0.24,
        cy - height * 0.24,
        cx - width * 0.24,
        cy + height * 0.24,
        fill=colour,
        width=1,
        tags=tags,
    )


def responsive_layout_mode(window_width: int) -> str:
    """Select a three-tier layout that keeps controls useful at every width."""

    width = max(0, int(window_width))
    if width < 980:
        return "narrow"
    if width < 1320:
        return "compact"
    return "wide"


def short_window_layout(window_height: int) -> bool:
    """Return whether vertical space needs the scroll-first compact layout."""

    return max(1, int(window_height)) < 680


def smooth_progress_step(
    current: float,
    target: float,
    elapsed_ms: float,
    *,
    time_constant_ms: float = 135.0,
) -> float:
    """Move determinate progress monotonically towards its latest target."""

    current = min(100.0, max(0.0, float(current)))
    target = min(100.0, max(0.0, float(target)))
    if target <= current or target - current <= 0.06:
        return target
    elapsed_ms = max(0.0, float(elapsed_ms))
    time_constant_ms = max(1.0, float(time_constant_ms))
    alpha = 1.0 - math.exp(-elapsed_ms / time_constant_ms)
    value = current + (target - current) * alpha
    return min(target, max(current, value))


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
        minimum = 244
        maximum = min(360, max(minimum, width - 700), max(minimum, round(width * 0.30)))
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
    width = (
        screen_width
        if screen_width <= 760
        else min(screen_width, max(680, min(1480, screen_width - 96)))
    )
    height = (
        screen_height
        if screen_height <= 560
        else min(screen_height, max(500, min(920, screen_height - 96)))
    )
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
        detail += f"\n更多说明：{str(notes).strip()}"
    technical = " · ".join(
        part for part in (str(engine).strip(), str(reason).strip()) if part
    )
    if technical:
        detail += f"\n技术信息：{technical}"
    if expanded:
        return f"{detail}\n收起说明 ▴"
    return f"{summary}\n了解更多 ▾"


def operation_display_name(name: object) -> str:
    """Return a concise catalog label while keeping details in the task body."""

    value = str(name or "").strip()
    shortened = re.sub(r"\s*[（(][^）)]*[）)]\s*$", "", value).strip()
    return shortened or value


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


def windows_redraw_flags() -> int:
    """Flags for a non-erasing native invalidation of the Tk child tree."""

    # RDW_INVALIDATE | RDW_ALLCHILDREN.  Do not use RDW_UPDATENOW: Tk handles
    # WM_PAINT by queuing Expose work, so synchronous painting can expose an
    # incomplete native-child frame during Map/resize transactions.
    return 0x0081


def _force_windows_window_redraw(window: tk.Misc) -> bool:
    """Queue paint events for the full native Tk tree without erasing it."""

    if os.name != "nt":
        return False
    try:
        user32 = ctypes.windll.user32
        hwnd_type = ctypes.c_void_p
        get_ancestor = user32.GetAncestor
        get_ancestor.argtypes = [hwnd_type, ctypes.c_uint]
        get_ancestor.restype = hwnd_type
        widget_hwnd = hwnd_type(int(window.winfo_id()))
        root_hwnd = get_ancestor(widget_hwnd, 2) or widget_hwnd  # GA_ROOT
        return bool(user32.RedrawWindow(root_hwnd, None, None, windows_redraw_flags()))
    except (AttributeError, OSError, TypeError, ValueError, tk.TclError):
        return False


class RoundedCard(tk.Canvas):
    """A responsive, genuinely rounded Tk surface with a normal child frame."""

    def __init__(
        self,
        master: tk.Misc,
        *,
        fill: str,
        background: str,
        outline: str | None = None,
        shadow: str | None = None,
        radius: int = 24,
        inset: int = 9,
        auto_height: bool = False,
        **kwargs: Any,
    ) -> None:
        super().__init__(
            master,
            bg=background,
            highlightthickness=0,
            borderwidth=0,
            **kwargs,
        )
        self.card_fill = fill
        self.card_outline = outline or BORDER
        self.card_shadow = shadow or SHADOW
        self.card_radius = max(8, int(radius))
        self.card_inset = max(4, int(inset))
        self.auto_height = bool(auto_height)
        self._last_drawn_size = (0, 0)
        self._pending_draw_size = (0, 0)
        self._redraw_job: str | None = None
        self._height_fit_job: str | None = None
        self.inner = tk.Frame(self, bg=fill, borderwidth=0, highlightthickness=0)
        self._inner_window = self.create_window(
            self.card_inset,
            self.card_inset,
            window=self.inner,
            anchor="nw",
            tags=("card_inner",),
        )
        self.bind("<Configure>", self._schedule_redraw, add="+")
        if self.auto_height:
            self.inner.bind(
                "<Configure>",
                self._schedule_fit_requested_height,
                add="+",
            )

    def _schedule_fit_requested_height(
        self,
        _event: tk.Event | None = None,
    ) -> None:
        """Coalesce nested layout changes instead of forcing synchronous relayouts."""

        if self._height_fit_job is not None:
            return
        with contextlib.suppress(tk.TclError):
            self._height_fit_job = self.after_idle(self._run_scheduled_height_fit)

    def _run_scheduled_height_fit(self) -> None:
        self._height_fit_job = None
        self._fit_requested_height()

    def _fit_requested_height(self, _event: tk.Event | None = None) -> None:
        try:
            grid_box = self.inner.grid_bbox()
            content_height = max(
                self.inner.winfo_reqheight(),
                grid_box[1] + grid_box[3] if grid_box else 1,
            )
            requested = content_height + self.card_inset * 2 + 6
            if requested > 1 and abs(int(float(self.cget("height"))) - requested) > 1:
                self.configure(height=requested)
        except tk.TclError:
            return

    def destroy(self) -> None:
        if self._redraw_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._redraw_job)
            self._redraw_job = None
        if self._height_fit_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._height_fit_job)
            self._height_fit_job = None
        super().destroy()

    def _schedule_redraw(self, event: tk.Event | None = None) -> None:
        """Record the newest size and repaint only after native resizing settles."""

        if event is None:
            width = max(1, self.winfo_width())
            height = max(1, self.winfo_height())
        else:
            width = max(1, int(event.width))
            height = max(1, int(event.height))
        self._pending_draw_size = (width, height)
        if self._redraw_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._redraw_job)
            self._redraw_job = None
        if self._root_is_resizing():
            # Keep the prior child-frame geometry while the user drags.  The
            # outer Canvas supplies a stable theme-coloured surface for newly
            # exposed pixels; the full nested HWND tree is resized once later.
            return
        with contextlib.suppress(tk.TclError):
            self._redraw_job = self.after(28, self._run_scheduled_redraw)

    def _run_scheduled_redraw(self) -> None:
        self._redraw_job = None
        if self._root_is_resizing():
            return
        width, height = self._pending_draw_size
        self._resize_inner_window(max(1, width), max(1, height))
        self._draw_card_surface(max(1, width), max(1, height))

    def _root_is_resizing(self) -> bool:
        try:
            root = self.winfo_toplevel()
            if hasattr(root, "_window_mapped") and not bool(root._window_mapped):
                return True
            return bool(getattr(root, "_window_resizing", False))
        except tk.TclError:
            return False

    def _resize_inner_window(self, width: int, height: int) -> None:
        inset = self.card_inset
        with contextlib.suppress(tk.TclError):
            self.coords(self._inner_window, inset, inset)
            self.itemconfigure(
                self._inner_window,
                width=max(1, width - inset * 2 - 4),
                height=max(1, height - inset * 2 - 6),
            )

    def _redraw(self, _event: tk.Event | None = None) -> None:
        """Force an immediate complete redraw (theme changes and final settle)."""

        width = max(1, self.winfo_width())
        height = max(1, self.winfo_height())
        self._pending_draw_size = (width, height)
        self._resize_inner_window(width, height)
        if self._redraw_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._redraw_job)
            self._redraw_job = None
        self._draw_card_surface(width, height)

    def _draw_card_surface(self, width: int, height: int) -> None:
        if (width, height) == self._last_drawn_size:
            return
        self._last_drawn_size = (width, height)
        inset = self.card_inset
        if width < inset * 2 + 8 or height < inset * 2 + 8:
            for item_name in ("_shadow_surface", "_fill_surface"):
                item_id = getattr(self, item_name, None)
                if item_id is not None:
                    with contextlib.suppress(tk.TclError):
                        self.itemconfigure(item_id, state="hidden")
            return
        shadow_id = getattr(self, "_shadow_surface", None)
        fill_id = getattr(self, "_fill_surface", None)
        shadow_points = self._round_rect_points(
            4,
            6,
            width - 2,
            height - 2,
            radius=self.card_radius,
        )
        fill_points = self._round_rect_points(
            1,
            1,
            width - 5,
            height - 6,
            radius=self.card_radius,
        )
        if shadow_id is None or fill_id is None:
            self._shadow_surface = self.create_round_rect(
                4,
                6,
                width - 2,
                height - 2,
                radius=self.card_radius,
                fill=self.card_shadow,
                outline="",
                tags=("card_surface",),
            )
            self._fill_surface = self.create_round_rect(
                1,
                1,
                width - 5,
                height - 6,
                radius=self.card_radius,
                fill=self.card_fill,
                outline=self.card_outline,
                width=1,
                tags=("card_surface",),
            )
        else:
            self.coords(shadow_id, *shadow_points)
            self.itemconfigure(
                shadow_id,
                fill=self.card_shadow,
                outline="",
                state="normal",
            )
            self.coords(fill_id, *fill_points)
            self.itemconfigure(
                fill_id,
                fill=self.card_fill,
                outline=self.card_outline,
                width=1,
                state="normal",
            )
        self.tag_lower("card_surface")
        self._resize_inner_window(width, height)

    def set_palette(
        self,
        *,
        fill: str | None = None,
        background: str | None = None,
        outline: str | None = None,
        shadow: str | None = None,
    ) -> None:
        """Update card colours without rebuilding its child widgets."""

        if fill is not None:
            self.card_fill = fill
            self.inner.configure(bg=fill)
        if background is not None:
            self.configure(bg=background)
        if outline is not None:
            self.card_outline = outline
        if shadow is not None:
            self.card_shadow = shadow
        self._last_drawn_size = (0, 0)
        self._redraw()

    def create_round_rect(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        *,
        radius: float,
        **kwargs: Any,
    ) -> int:
        points = self._round_rect_points(x1, y1, x2, y2, radius=radius)
        return self.create_polygon(points, smooth=True, splinesteps=24, **kwargs)

    @staticmethod
    def _round_rect_points(
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        *,
        radius: float,
    ) -> tuple[float, ...]:
        radius = max(2.0, min(float(radius), (x2 - x1) / 2, (y2 - y1) / 2))
        return (
            x1 + radius,
            y1,
            x2 - radius,
            y1,
            x2,
            y1,
            x2,
            y1 + radius,
            x2,
            y2 - radius,
            x2,
            y2,
            x2 - radius,
            y2,
            x1 + radius,
            y2,
            x1,
            y2,
            x1,
            y2 - radius,
            x1,
            y1 + radius,
            x1,
            y1,
        )


class TaskResultDialog(tk.Toplevel):
    """Responsive result summary with success/partial/failure colour states."""

    def __init__(
        self,
        parent: tk.Misc,
        result: TaskResult,
        *,
        output_dir: str | Path,
        operation_name: str = "文件处理",
        motion_mode: str = "rich",
    ) -> None:
        super().__init__(parent)
        self.result = result
        self.output_dir = Path(output_dir).expanduser()
        self.output_paths = tuple(Path(path).expanduser() for path in result.outputs)
        self.presentation = task_result_presentation(result)
        self.motion_mode = normalize_motion_mode(motion_mode)
        self._entrance_plan = result_dialog_entrance_plan(self.motion_mode)
        self._entrance_job: str | None = None
        self.title(self.presentation.title)
        self.configure(bg=BG)
        self.minsize(520, 440)
        self.resizable(True, True)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        try:
            parent.update_idletasks()
            width, height = result_dialog_geometry(
                parent.winfo_width(),
                parent.winfo_height(),
            )
            left = parent.winfo_rootx() + max(0, (parent.winfo_width() - width) // 2)
            top = parent.winfo_rooty() + max(0, (parent.winfo_height() - height) // 2)
            self.geometry(f"{width}x{height}+{left}+{top}")
        except tk.TclError:
            self.geometry("660x560")

        shell = tk.Frame(self, bg=BG, padx=16, pady=16)
        shell.pack(fill="both", expand=True)
        self.result_card = RoundedCard(
            shell,
            fill=PANEL,
            background=BG,
            outline=BORDER,
            shadow=SHADOW,
            radius=30,
            inset=3,
        )
        self.result_card.pack(fill="both", expand=True)
        card = self.result_card.inner

        self.header = tk.Canvas(
            card,
            height=154,
            bg=PANEL,
            highlightthickness=0,
            bd=0,
        )
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
        body.rowconfigure(4, weight=1)
        body.columnconfigure(0, weight=1)
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
        operation_label.grid(row=0, column=0, sticky="ew")
        operation_label.bind(
            "<Configure>",
            lambda event: operation_label.configure(
                wraplength=max(260, event.width - 4)
            ),
        )

        stats = tk.Frame(body, bg=PANEL)
        stats.grid(row=1, column=0, sticky="ew", pady=(12, 12))
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
        location.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        location.bind(
            "<Configure>",
            lambda event: location.configure(wraplength=max(220, event.width - 4)),
        )

        details = self._detail_lines()
        if details:
            detail_title = tk.Label(
                body,
                text=("停止情况与提示" if result.cancelled else "未完成文件与提示"),
                bg=PANEL,
                fg=TEXT,
                anchor="w",
                font=("Microsoft YaHei UI", 10, "bold"),
            )
            detail_title.grid(row=3, column=0, sticky="ew", pady=(2, 6))
            detail_shell = tk.Frame(
                body,
                bg=PANEL_ALT,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            detail_shell.grid(row=4, column=0, sticky="nsew")
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
                selectbackground=ACCENT_SOFT,
                selectforeground=TEXT,
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
            success_message = tk.Label(
                body,
                text="所有输入均已按预期处理完成。",
                bg=PANEL,
                fg=SUCCESS,
                anchor="w",
                font=("Microsoft YaHei UI", 10),
            )
            success_message.grid(row=3, column=0, sticky="ew", pady=(8, 0))
            tk.Frame(body, bg=PANEL, height=8).grid(
                row=4,
                column=0,
                sticky="nsew",
            )

        buttons = tk.Frame(body, bg=PANEL)
        buttons.grid(row=5, column=0, sticky="ew", pady=(16, 0))
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=0)
        left_actions = tk.Frame(buttons, bg=PANEL)
        left_actions.grid(row=0, column=0, sticky="w")
        right_actions = tk.Frame(buttons, bg=PANEL)
        right_actions.grid(row=0, column=1, sticky="e")
        if result.failed_inputs or result.cancelled_inputs or result.warnings:
            self._result_button(
                left_actions,
                text="复制未完成清单",
                command=self._copy_details,
            ).pack(side="left")
        self._result_button(
            right_actions,
            text="关闭",
            command=self.destroy,
            primary=True,
        ).pack(side="right")
        if self.output_paths:
            self._result_button(
                right_actions,
                text="打开文件" if len(self.output_paths) == 1 else "打开首个文件",
                command=self._open_primary_output,
            ).pack(side="right", padx=(0, 8))
        self._result_button(
            right_actions,
            text="打开文件夹",
            command=self._open_output,
        ).pack(side="right", padx=(0, 8))

        self.after(20, self._start_entrance_animation)
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
        self.header.delete("result_decor")
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
        # The right-hand illustration owns a strict reserved strip and is never
        # allowed to overlap the title/subtitle.  Narrow dialogs hide it.
        decor_reserved = 176 if width >= 660 else 78 if width >= 580 else 0
        text_right = width - decor_reserved - 28
        self.header.itemconfigure(
            "header_subtitle",
            width=max(220, text_right - 95),
        )
        title_bbox = self.header.bbox("header_title")
        subtitle_bbox = self.header.bbox("header_subtitle")
        text_guard_right = max(
            (bbox[2] for bbox in (title_bbox, subtitle_bbox) if bbox),
            default=0,
        ) + 14
        if width >= 660 and text_guard_right < width - 166:
            base = self.presentation.gradient_end
            line = composite_hex_colour(base, "#FFFFFF", 0.22)
            accent = composite_hex_colour(base, "#FFFFFF", 0.36)
            draw_laptop_doodle(
                self.header,
                (width - 148, 29, width - 62, 112),
                colour=line,
                accent=accent,
                tags=("result_decor",),
            )
            self.header.create_arc(
                width - 164,
                17,
                width - 29,
                137,
                start=214,
                extent=156,
                style="arc",
                outline=line,
                width=1,
                dash=(4, 6),
                tags=("result_decor",),
            )
            for x, y, radius in (
                (width - 151, 46, 3),
                (width - 49, 80, 2),
                (width - 123, 126, 2),
            ):
                self.header.create_oval(
                    x - radius,
                    y - radius,
                    x + radius,
                    y + radius,
                    fill=accent,
                    outline="",
                    tags=("result_decor",),
                )
        elif width >= 580 and text_guard_right < width - 74:
            line = composite_hex_colour(self.presentation.gradient_end, "#FFFFFF", 0.30)
            for index, y in enumerate((46, 76, 106)):
                x = width - 47 - index * 9
                self.header.create_oval(
                    x - 3,
                    y - 3,
                    x + 3,
                    y + 3,
                    fill=line,
                    outline="",
                    tags=("result_decor",),
                )
        self.header.tag_raise("result_decor", "gradient")
        self.header.tag_raise("header_icon_circle")
        self.header.tag_raise("header_icon")
        self.header.tag_raise("header_title")
        self.header.tag_raise("header_subtitle")

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

    def _start_entrance_animation(self) -> None:
        if not self.winfo_exists():
            return
        if len(self._entrance_plan) <= 1:
            with contextlib.suppress(tk.TclError):
                self.attributes("-alpha", 1.0)
            return
        self._fade_in(0)

    def _fade_in(self, step: int) -> None:
        if not self.winfo_exists():
            return
        try:
            alpha = self._entrance_plan[min(step, len(self._entrance_plan) - 1)]
            self.attributes("-alpha", alpha)
            if step < len(self._entrance_plan) - 1:
                timing = motion_effect_timing(self.motion_mode, "dialog")
                self._entrance_job = self.after(
                    timing.step_ms,
                    self._fade_in,
                    step + 1,
                )
            else:
                self._entrance_job = None
        except tk.TclError:
            pass

    def destroy(self) -> None:
        if self._entrance_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._entrance_job)
            self._entrance_job = None
        super().destroy()

    def _activate_modal(self) -> None:
        try:
            self.grab_set()
            self.focus_force()
        except tk.TclError:
            pass


class DocuForgeApp(_TkBase):
    def __init__(self, operations: list[Operation]) -> None:
        super().__init__()
        self._preferences_path = ui_preferences_path()
        self._ui_preferences = load_ui_preferences(self._preferences_path)
        initial_theme = _activate_ui_theme(self._ui_preferences["ui_theme"])
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
        self._initial_window_height = logical_height
        left = max(0, (screen_width - initial_width) // 2)
        top = max(0, (screen_height - initial_height) // 2)
        self.geometry(f"{initial_width}x{initial_height}+{left}+{top}")
        self.minsize(
            min(screen_width, round(680 * self._display_scale)),
            min(screen_height, round(500 * self._display_scale)),
        )
        self.configure(bg=BG)

        self.operations = operations
        self.operation_by_id = {item.id: item for item in operations}
        self.current_operation: Operation | None = None
        self.input_paths: list[Path] = []
        self.param_vars: dict[str, tk.Variable] = {}
        self.choice_maps: dict[str, dict[str, str]] = {}
        self.office_engine_preference = tk.StringVar(value="auto")
        # General layout/progress transitions stay lightweight and automatic.
        # The user-facing switch controls click particles only.
        self.motion_mode_var = tk.StringVar(value="rich")
        self.particle_effects_var = tk.BooleanVar(
            value=normalize_particle_effects_enabled(
                self._ui_preferences.get("particle_effects")
            )
        )
        self.ui_theme_var = tk.StringVar(value=initial_theme)
        self.catalog_root_var = tk.StringVar(value="文档工具")
        self._syncing_office_engine = False
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
        self._progress_animation_last_at: float | None = None
        self._progress_target = 0.0
        self._progress_shimmer_job: str | None = None
        self._progress_track_resize_job: str | None = None
        self._title_animation_job: str | None = None
        self._button_motion_jobs: dict[str, str] = {}
        self._workspace_transition_job: str | None = None
        self._content_transition_job: str | None = None
        self._content_transition_generation = 0
        self._content_transition_overlay: tk.Canvas | None = None
        self._catalog_transition_job: str | None = None
        self._indicator_realign_job: str | None = None
        self._particle_windows: list[tk.Toplevel] = []
        self._particle_jobs: set[str] = set()
        self._particle_variant = 0
        self._last_particle_spawn_at = 0.0
        self._progress_shimmer_phase = 0.0
        self._progress_track_signature: tuple[int, int, str] | None = None
        self._progress_track_items: dict[str, Any] = {}
        self._background_resize_job: str | None = None
        self._background_last_signature: tuple[int, int, str] | None = None
        self._empty_drop_redraw_job: str | None = None
        self._empty_drop_last_signature: tuple[int, int] | None = None
        self._preview_resize_jobs: dict[str, str] = {}
        self._preview_last_sizes: dict[str, tuple[int, int]] = {}
        self._header_art_last_signature: tuple[int, int, str | None] | None = None
        self._setup_canvas_width = 0
        self._window_layout_job: str | None = None
        self._window_resize_finish_job: str | None = None
        self._window_restore_job: str | None = None
        self._window_restore_finalize_job: str | None = None
        self._window_resizing = False
        self._window_mapped = False
        self._window_restoring = False
        self._window_restore_attempts = 0
        self._last_window_configure_size: tuple[int, int] | None = None
        self._pending_window_width: int | None = None
        self._pending_window_height: int | None = None
        self._setup_canvas_resize_job: str | None = None
        self._pending_setup_canvas_width: int | None = None
        self._pending_input_scans = 0
        self._style_images: list[ImageTk.PhotoImage] = []
        self._style_element_failures: set[str] = set()
        self._brand_icon: ImageTk.PhotoImage | None = None
        self._sidebar_expanded = False
        self._sidebar_user_width: int | None = None
        self._sidebar_drag_start_x = 0
        self._sidebar_drag_start_width = 0
        self._catalog_preferred_width = 320
        self._catalog_tree_content_width = 300
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
        self.parameter_section_frames: dict[str, RoundedCard] = {}
        self.parameter_section_order: list[str] = []
        self.advanced_parameters_frame: RoundedCard | None = None
        self.simple_parameters_card: RoundedCard | None = None
        self.advanced_parameters_button: ttk.Button | None = None
        self.advanced_parameters_expanded = False
        self._layout_mode: str | None = None
        self._layout_short: bool | None = None
        self._last_narrow_sidebar_height = 0
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
        self._build_background_layer()
        self._build_ui()
        self._rebuild_operation_tree()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Configure>", self._on_window_configure, add="+")
        self.bind("<Map>", self._on_window_map, add="+")
        self.bind("<Unmap>", self._on_window_unmap, add="+")
        self.after(120, self._poll_worker)
        self.after(1000, self._refresh_progress_elapsed)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if style.theme_use() != "clam":
            style.theme_use("clam")
        self.option_add("*Font", ("Microsoft YaHei UI", 10))
        self.option_add("*Menu.Font", ("Microsoft YaHei UI", 10))
        self.option_add("*Menu.background", PANEL)
        self.option_add("*Menu.foreground", TEXT)
        self.option_add("*Menu.activeBackground", ACCENT_SOFT)
        self.option_add("*Menu.activeForeground", TEXT)
        self.option_add("*Menu.selectColor", CYAN)
        self.option_add("*Menu.relief", "flat")
        self.option_add("*Menu.borderWidth", 1)
        self.option_add("*Menu.activeBorderWidth", 0)
        self.option_add("*TCombobox*Listbox.background", INPUT_SURFACE)
        self.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.option_add("*TCombobox*Listbox.selectBackground", ACCENT_SOFT)
        self.option_add("*TCombobox*Listbox.selectForeground", TEXT)
        self.option_add("*TCombobox*Listbox.relief", "flat")
        self.option_add("*TCombobox*Listbox.borderWidth", 1)
        style.configure(".", background=BG, foreground=TEXT, bordercolor=BORDER)
        style.configure("Panel.TFrame", background=PANEL)
        style.configure("Workspace.TFrame", background=PANEL)
        style.configure("Soft.TFrame", background=PANEL_ALT)
        style.configure("Sidebar.TFrame", background=SIDEBAR_BG)
        style.configure("Header.TFrame", background=BG)
        style.configure("Main.TFrame", background=BG)
        style.configure("Canvas.TFrame", background=PANEL)
        style.configure("Footer.TFrame", background=PANEL_ALT)
        style.configure(
            "CanvasField.TLabel",
            background=PANEL,
            foreground=TEXT,
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "CanvasSubtle.TLabel",
            background=PANEL,
            foreground=MUTED,
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "StepTitle.TLabel",
            background=PANEL,
            foreground=CYAN,
            font=("Microsoft YaHei UI", 13, "bold"),
            padding=(2, 4),
        )
        style.configure(
            "FooterSubtle.TLabel",
            background=PANEL_ALT,
            foreground=MUTED,
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "Category.TButton",
            background=CARD_YELLOW,
            foreground=MUTED,
            bordercolor=BORDER,
            borderwidth=1,
            focusthickness=0,
            padding=(12, 9),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "Category.TButton",
            background=[("active", HOVER_SURFACE), ("pressed", PRESSED_SURFACE)],
            foreground=[("active", TEXT), ("pressed", CYAN)],
            bordercolor=[("active", ACCENT), ("pressed", CYAN)],
        )
        style.configure(
            "CategoryActive.TButton",
            background=ACCENT_SOFT,
            foreground=TEXT,
            bordercolor=ACCENT,
            borderwidth=1,
            focusthickness=0,
            padding=(12, 9),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "CategoryActive.TButton",
            background=[
                ("active", SELECTED_ACTIVE_SURFACE),
                ("pressed", ACCENT_DARK),
            ],
            foreground=[("active", TEXT), ("pressed", TEXT)],
            bordercolor=[("active", CYAN), ("pressed", CYAN)],
        )
        style.configure(
            "WorkspaceTab.TButton",
            background=CARD_YELLOW,
            foreground=MUTED,
            bordercolor=BORDER,
            borderwidth=1,
            focusthickness=0,
            padding=(9, 6),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "WorkspaceTab.TButton",
            background=[("active", HOVER_SURFACE), ("pressed", PRESSED_SURFACE)],
            foreground=[("active", TEXT), ("pressed", CYAN)],
            bordercolor=[("active", ACCENT), ("pressed", CYAN)],
        )
        style.configure(
            "WorkspaceTabActive.TButton",
            background=ACCENT_SOFT,
            foreground=TEXT,
            bordercolor=ACCENT,
            borderwidth=1,
            focusthickness=0,
            padding=(9, 6),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "WorkspaceTabActive.TButton",
            background=[
                ("active", SELECTED_ACTIVE_SURFACE),
                ("pressed", ACCENT_DARK),
            ],
            foreground=[("active", TEXT), ("pressed", TEXT)],
            bordercolor=[("active", CYAN), ("pressed", CYAN)],
        )
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
            background=DROP_SURFACE,
            foreground=CYAN,
            padding=(16, 13),
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "DropActive.TLabel",
            background=ACCENT_SOFT,
            foreground=TEXT,
            padding=(16, 13),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.configure(
            "DropBusy.TLabel",
            background=DROP_SURFACE,
            foreground=WARNING,
            padding=(16, 13),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.configure(
            "DropError.TLabel",
            background=DANGER_SURFACE,
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
            background=ACCENT_DARK,
            foreground=ON_ACCENT,
            borderwidth=0,
            focusthickness=0,
            padding=(20, 10),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[
                ("active", ACCENT),
                ("pressed", ACCENT_PRESSED),
                ("disabled", DISABLED_SURFACE),
            ],
            foreground=[("disabled", DISABLED_TEXT)],
        )
        style.configure(
            "Quiet.TButton",
            background=CARD_YELLOW,
            foreground=TEXT,
            bordercolor=BORDER,
            borderwidth=1,
            focusthickness=0,
            padding=(12, 8),
        )
        style.map(
            "Quiet.TButton",
            background=[("active", HOVER_SURFACE), ("pressed", PRESSED_SURFACE)],
            foreground=[("active", TEXT), ("disabled", DISABLED_TEXT)],
            bordercolor=[("active", ACCENT), ("pressed", CYAN)],
        )
        style.configure(
            "ParticleOn.TButton",
            background=ACCENT_SOFT,
            foreground=CYAN,
            bordercolor=ACCENT,
            borderwidth=1,
            focusthickness=0,
            padding=(12, 8),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "ParticleOn.TButton",
            background=[
                ("active", SELECTED_ACTIVE_SURFACE),
                ("pressed", ACCENT_DARK),
            ],
            foreground=[("active", TEXT), ("pressed", ON_ACCENT)],
            bordercolor=[("active", CYAN), ("pressed", CYAN)],
        )
        style.configure(
            "Danger.TButton",
            background=DANGER_SURFACE,
            foreground=DANGER,
            bordercolor=DANGER_BORDER,
            borderwidth=1,
            focusthickness=0,
            padding=(12, 8),
        )
        style.map(
            "Danger.TButton",
            background=[("active", DANGER_HOVER), ("pressed", DANGER_PRESSED)],
            foreground=[("disabled", DISABLED_TEXT)],
        )
        style.configure(
            "Nav.TButton",
            background=CARD_YELLOW,
            foreground=TEXT,
            bordercolor=NAV_BORDER,
            borderwidth=1,
            focusthickness=0,
            padding=(12, 8),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "Nav.TButton",
            background=[("active", HOVER_SURFACE), ("pressed", ACCENT_SOFT)],
            foreground=[("active", CYAN), ("pressed", TEXT)],
            bordercolor=[("active", ACCENT), ("pressed", CYAN)],
        )
        style.configure(
            "Treeview",
            background=CARD_YELLOW,
            fieldbackground=CARD_YELLOW,
            foreground=TEXT,
            borderwidth=0,
            rowheight=42,
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            background=PANEL_ALT,
            foreground=TEXT,
            relief="flat",
            padding=(8, 7),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "Treeview",
            background=[("selected", TREE_SELECTED)],
            foreground=[("selected", TEXT)],
        )
        style.map(
            "Treeview.Heading",
            background=[("active", HOVER_SURFACE)],
            foreground=[("active", CYAN)],
        )
        style.configure(
            "TEntry",
            fieldbackground=INPUT_SURFACE,
            foreground=TEXT,
            insertcolor=TEXT,
            selectbackground=ACCENT_SOFT,
            selectforeground=TEXT,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            padding=(10, 8),
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", ACCENT), ("disabled", BORDER)],
            fieldbackground=[("disabled", DISABLED_SURFACE)],
            foreground=[("disabled", DISABLED_TEXT)],
        )
        style.configure(
            "TCombobox",
            fieldbackground=INPUT_SURFACE,
            background=INPUT_SURFACE,
            foreground=TEXT,
            arrowcolor=CYAN,
            selectbackground=ACCENT_SOFT,
            selectforeground=TEXT,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            arrowsize=15,
            padding=(9, 7),
        )
        style.map(
            "TCombobox",
            bordercolor=[("focus", ACCENT)],
            fieldbackground=[("readonly", INPUT_SURFACE), ("disabled", DISABLED_SURFACE)],
            foreground=[("readonly", TEXT), ("disabled", DISABLED_TEXT)],
            arrowcolor=[("active", TEXT), ("disabled", DISABLED_TEXT)],
        )
        style.configure(
            "TSpinbox",
            fieldbackground=INPUT_SURFACE,
            foreground=TEXT,
            insertcolor=TEXT,
            arrowcolor=CYAN,
            bordercolor=BORDER,
            arrowsize=14,
            padding=(9, 7),
        )
        style.map(
            "TSpinbox",
            bordercolor=[("focus", ACCENT)],
            fieldbackground=[("disabled", DISABLED_SURFACE)],
            foreground=[("disabled", DISABLED_TEXT)],
        )
        style.configure(
            "Workspace.TNotebook",
            background=PANEL,
            borderwidth=0,
            tabmargins=(0, 0, 0, 8),
        )
        style.configure(
            "Workspace.TNotebook.Tab",
            background=CARD_SAGE,
            foreground=MUTED,
            borderwidth=0,
            padding=(22, 11),
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        style.map(
            "Workspace.TNotebook.Tab",
            background=[("selected", ACCENT_SOFT), ("active", HOVER_SURFACE)],
            foreground=[("selected", TEXT), ("active", CYAN)],
        )
        style.configure(
            "HiddenTabs.TNotebook",
            background=PANEL,
            borderwidth=0,
            tabmargins=(0, 0, 0, 0),
        )
        style.layout("HiddenTabs.TNotebook.Tab", [])
        style.configure(
            "Card.TLabelframe",
            background=CARD_YELLOW,
            bordercolor=BORDER,
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=CREAM_YELLOW,
            foreground=TEXT,
            font=("Microsoft YaHei UI", 10, "bold"),
            padding=(6, 0),
        )
        for scrollbar_style in (
            "Vertical.TScrollbar",
            "Horizontal.TScrollbar",
            "Workspace.Vertical.TScrollbar",
        ):
            style.configure(
                scrollbar_style,
                background=SCROLL_THUMB,
                troughcolor=SCROLL_TROUGH,
                bordercolor=SCROLL_TROUGH,
                lightcolor=SCROLL_THUMB,
                darkcolor=SCROLL_THUMB,
                arrowcolor=TEXT,
                borderwidth=0,
                arrowsize=13,
                width=18 if scrollbar_style.startswith("Workspace.") else 14,
            )
            style.map(
                scrollbar_style,
                background=[
                    ("disabled", SCROLL_THUMB),
                    ("pressed", SCROLL_ACTIVE),
                    ("active", SCROLL_ACTIVE),
                ],
                troughcolor=[
                    ("disabled", SCROLL_TROUGH),
                    ("pressed", SCROLL_TROUGH),
                    ("active", SCROLL_TROUGH),
                ],
                arrowcolor=[
                    ("disabled", MUTED),
                    ("pressed", ON_ACCENT),
                    ("active", ON_ACCENT),
                ],
            )
        style.configure(
            "Timeline.Horizontal.TScale",
            background=SCROLL_THUMB,
            troughcolor=SCROLL_TROUGH,
            bordercolor=SCROLL_TROUGH,
            lightcolor=SCROLL_THUMB,
            darkcolor=SCROLL_THUMB,
            sliderrelief="flat",
            borderwidth=0,
            sliderlength=22,
        )
        style.map(
            "Timeline.Horizontal.TScale",
            background=[
                ("disabled", DISABLED_TEXT),
                ("pressed", SCROLL_ACTIVE),
                ("active", SCROLL_ACTIVE),
            ],
            troughcolor=[
                ("disabled", SCROLL_TROUGH),
                ("pressed", SCROLL_TROUGH),
                ("active", SCROLL_TROUGH),
            ],
        )
        style.configure(
            "Horizontal.TProgressbar",
            background=CYAN,
            troughcolor=PROGRESS_TROUGH,
            borderwidth=0,
            thickness=10,
        )
        self._configure_rounded_button_images(style)
        self._configure_checkbutton_images(style)

    def _build_background_layer(self) -> None:
        self.background_canvas = tk.Canvas(
            self,
            bg=BG,
            highlightthickness=0,
            borderwidth=0,
        )
        self.background_canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.tk.call("lower", self.background_canvas._w)
        self.background_canvas.bind(
            "<Configure>",
            self._schedule_background_redraw,
            add="+",
        )

    def _schedule_background_redraw(
        self,
        _event: tk.Event | None = None,
        *,
        delay: int | None = None,
        force: bool = False,
    ) -> None:
        if self._closing or not self.winfo_exists():
            return
        if force:
            self._background_last_signature = None
        if self._background_resize_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._background_resize_job)
        wait_ms = (
            150
            if self._window_resizing
            else (105 if self.worker and self.worker.is_alive() else 65)
        ) if delay is None else max(0, int(delay))
        self._background_resize_job = self.after(
            wait_ms,
            self._redraw_background_layer,
        )

    def _redraw_background_layer(self) -> None:
        self._background_resize_job = None
        if self._closing or not self.winfo_exists():
            return
        try:
            canvas = self.background_canvas
            width = max(1, canvas.winfo_width())
            height = max(1, canvas.winfo_height())
            layout_mode = responsive_layout_mode(
                self._logical_window_width(width)
                if hasattr(self, "_display_scale")
                else width
            )
            signature = (width, height, layout_mode)
            if signature == self._background_last_signature:
                return
            previous = self._background_last_signature
            if (
                previous is not None
                and previous[2] == layout_mode
                and abs(previous[0] - width) < 6
                and abs(previous[1] - height) < 6
            ):
                return
            next_tag = "decor_next"
            canvas.delete(next_tag)

            # A sparse technical grid adds depth without turning the window into
            # a bright RGB dashboard.  The lines stay deliberately close to BG.
            grid_step = max(72, round(min(width, height) * 0.095))
            for x in range(grid_step // 2, width, grid_step):
                canvas.create_line(
                    x,
                    0,
                    x,
                    height,
                    fill=GRID_LINE,
                    width=1,
                    tags=next_tag,
                )
            for y in range(grid_step // 2, height, grid_step):
                canvas.create_line(
                    0,
                    y,
                    width,
                    y,
                    fill=GRID_LINE,
                    width=1,
                    tags=next_tag,
                )
            canvas.create_oval(
                -round(width * 0.18),
                -round(height * 0.25),
                round(width * 0.34),
                round(height * 0.30),
                fill=GLOW_BLUE,
                outline="",
                tags=next_tag,
            )
            canvas.create_oval(
                round(width * 0.75),
                -round(height * 0.20),
                round(width * 1.15),
                round(height * 0.28),
                fill=GLOW_PURPLE,
                outline="",
                tags=next_tag,
            )
            canvas.create_oval(
                round(width * 0.72),
                round(height * 0.72),
                round(width * 1.16),
                round(height * 1.18),
                fill=GLOW_THIRD,
                outline="",
                tags=next_tag,
            )
            orbit_colour = ORBIT_LINE
            orbit_accent = ORBIT_ACCENT
            canvas.create_arc(
                round(width * 0.72),
                -round(height * 0.11),
                round(width * 1.06),
                round(height * 0.28),
                start=205,
                extent=128,
                style="arc",
                outline=orbit_colour,
                width=2,
                tags=next_tag,
            )
            canvas.create_arc(
                -round(width * 0.10),
                round(height * 0.72),
                round(width * 0.22),
                round(height * 1.10),
                start=18,
                extent=126,
                style="arc",
                outline=orbit_colour,
                width=2,
                tags=next_tag,
            )
            for ratio_x, ratio_y, radius in (
                (0.82, 0.075, 3),
                (0.90, 0.16, 2),
                (0.085, 0.82, 3),
                (0.145, 0.91, 2),
            ):
                x = round(width * ratio_x)
                y = round(height * ratio_y)
                canvas.create_oval(
                    x - radius,
                    y - radius,
                    x + radius,
                    y + radius,
                    fill=CYAN if radius == 3 else PINK,
                    outline=orbit_accent,
                    width=1,
                    tags=next_tag,
                )

            # Many small marks make the backdrop feel inhabited while the
            # entire layer remains behind every real widget and all text.
            for placement in background_watermark_plan(
                width,
                height,
                layout_mode,
            ):
                tags = (next_tag, "micro_watermark_next")
                if placement.kind == "file":
                    draw_file_doodle(
                        canvas,
                        placement.box,
                        colour=placement.colour,
                        tags=tags,
                    )
                elif placement.kind == "satellite":
                    draw_satellite_doodle(
                        canvas,
                        placement.box,
                        colour=placement.colour,
                        tags=tags,
                    )
                elif placement.kind == "node":
                    x1, y1, x2, y2 = placement.box
                    canvas.create_oval(
                        x1,
                        y1,
                        x2,
                        y2,
                        outline=placement.colour,
                        width=1,
                        tags=tags,
                    )
                    canvas.create_oval(
                        (x1 + x2) / 2 - 1.5,
                        (y1 + y2) / 2 - 1.5,
                        (x1 + x2) / 2 + 1.5,
                        (y1 + y2) / 2 + 1.5,
                        fill=placement.colour,
                        outline="",
                        tags=tags,
                    )
                else:
                    draw_starlet(
                        canvas,
                        placement.box,
                        colour=placement.colour,
                        tags=tags,
                    )
            canvas.tag_lower(next_tag)
            canvas.delete("decor")
            canvas.addtag_withtag("decor", next_tag)
            canvas.addtag_withtag("micro_watermark", "micro_watermark_next")
            canvas.dtag(next_tag, next_tag)
            canvas.dtag("micro_watermark_next", "micro_watermark_next")
            self._background_last_signature = signature
        except tk.TclError:
            return

    def _configure_rounded_button_images(self, style: ttk.Style) -> None:
        """Give the main buttons soft scalable corners with a safe ttk fallback."""

        theme_key = normalize_ui_theme(self.ui_theme_var.get())
        existing_elements = set(style.element_names())

        def image(fill: str, border: str, *, radius: int = 10) -> ImageTk.PhotoImage:
            raster = Image.new("RGBA", (54, 42), (0, 0, 0, 0))
            draw = ImageDraw.Draw(raster)
            draw.rounded_rectangle(
                (1, 1, 52, 40),
                radius=max(radius, 14),
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
                (ACCENT_DARK, ACCENT_DARK),
                (ACCENT, CYAN),
                (ACCENT_PRESSED, ACCENT_DARK),
                (DISABLED_SURFACE, DISABLED_SURFACE),
                (ACCENT, ACCENT_PULSE_BORDER),
            ),
            (
                "Quiet.TButton",
                "DocuForgeQuiet.background",
                (CARD_YELLOW, BORDER),
                (HOVER_SURFACE, ACCENT),
                (PRESSED_SURFACE, CYAN),
                (DISABLED_SURFACE, BORDER),
                (ACCENT_SOFT, CYAN),
            ),
            (
                "ParticleOn.TButton",
                "DocuForgeParticleOn.background",
                (ACCENT_SOFT, ACCENT),
                (SELECTED_ACTIVE_SURFACE, CYAN),
                (ACCENT_DARK, CYAN),
                (DISABLED_SURFACE, BORDER),
                (SELECTED_ACTIVE_SURFACE, SELECTED_PULSE_BORDER),
            ),
            (
                "Danger.TButton",
                "DocuForgeDanger.background",
                (DANGER_SURFACE, DANGER_BORDER),
                (DANGER_HOVER, DANGER),
                (DANGER_PRESSED, DANGER_PULSE_BORDER),
                (DISABLED_SURFACE, BORDER),
                (DANGER_HOVER, DANGER),
            ),
            (
                "Nav.TButton",
                "DocuForgeNav.background",
                (CARD_YELLOW, BORDER),
                (HOVER_SURFACE, ACCENT),
                (ACCENT_SOFT, CYAN),
                (DISABLED_SURFACE, BORDER),
                (ACCENT_SOFT, CYAN),
            ),
            (
                "Category.TButton",
                "DocuForgeCategory.background",
                (CARD_YELLOW, BORDER),
                (HOVER_SURFACE, ACCENT),
                (PRESSED_SURFACE, CYAN),
                (DISABLED_SURFACE, BORDER),
                (ACCENT_SOFT, CYAN),
            ),
            (
                "CategoryActive.TButton",
                "DocuForgeCategoryActive.background",
                (ACCENT_SOFT, ACCENT),
                (SELECTED_ACTIVE_SURFACE, CYAN),
                (ACCENT_DARK, CYAN),
                (DISABLED_SURFACE, BORDER),
                (SELECTED_ACTIVE_SURFACE, SELECTED_PULSE_BORDER),
            ),
            (
                "WorkspaceTab.TButton",
                "DocuForgeWorkspaceTab.background",
                (CARD_YELLOW, BORDER),
                (HOVER_SURFACE, ACCENT),
                (PRESSED_SURFACE, CYAN),
                (DISABLED_SURFACE, BORDER),
                (ACCENT_SOFT, CYAN),
            ),
            (
                "WorkspaceTabActive.TButton",
                "DocuForgeWorkspaceTabActive.background",
                (ACCENT_SOFT, ACCENT),
                (SELECTED_ACTIVE_SURFACE, CYAN),
                (ACCENT_DARK, CYAN),
                (DISABLED_SURFACE, BORDER),
                (SELECTED_ACTIVE_SURFACE, SELECTED_PULSE_BORDER),
            ),
        )
        for (
            style_name,
            element_base,
            normal,
            active,
            pressed,
            disabled,
            pulse,
        ) in definitions:
            element_name = f"{element_base}.{theme_key}"
            if element_name in getattr(self, "_style_element_failures", set()):
                continue
            try:
                if element_name not in existing_elements:
                    normal_image = image(*normal)
                    active_image = image(*active)
                    pressed_image = image(*pressed)
                    disabled_image = image(*disabled)
                    pulse_image = image(*pulse)
                    style.element_create(
                        element_name,
                        "image",
                        normal_image,
                        ("disabled", disabled_image),
                        ("pressed", pressed_image),
                        ("alternate", pulse_image),
                        ("active", active_image),
                        border=(16, 16, 16, 16),
                        sticky="nsew",
                    )
                    existing_elements.add(element_name)
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
                self._style_element_failures.add(element_name)
                continue

    def _style_popup_menu(self, menu: tk.Menu) -> None:
        """Apply one calm floating-panel treatment to every native Tk menu."""

        with contextlib.suppress(tk.TclError):
            menu.configure(
                tearoff=False,
                bg=PANEL_ALT,
                fg=TEXT,
                activebackground=ACCENT_SOFT,
                activeforeground=TEXT,
                selectcolor=CYAN,
                disabledforeground=DISABLED_TEXT,
                relief="flat",
                borderwidth=1,
                activeborderwidth=0,
                font=("Microsoft YaHei UI", 10),
                cursor="hand2",
            )

    @staticmethod
    def _reset_menu_anchor(button: ttk.Button) -> None:
        with contextlib.suppress(tk.TclError):
            button.state(["!pressed", "!alternate"])

    def _popup_menu_below(self, menu: tk.Menu, button: ttk.Button, *, gap: int = 5) -> None:
        """Open a menu at a stable anchor and always clear stale button states."""

        self._reset_menu_anchor(button)
        try:
            menu.tk_popup(
                button.winfo_rootx(),
                button.winfo_rooty() + button.winfo_height() + gap,
            )
        finally:
            with contextlib.suppress(tk.TclError):
                menu.grab_release()
            self._reset_menu_anchor(button)

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

        theme_key = normalize_ui_theme(self.ui_theme_var.get())
        element_name = f"DocuForgeCheck.{theme_key}.indicator"
        if element_name in getattr(self, "_style_element_failures", set()):
            return
        try:
            if element_name not in set(style.element_names()):
                normal = indicator(CARD_YELLOW, DUSTY_BLUE)
                active = indicator(ACCENT_SOFT, ACCENT)
                selected = indicator(ACCENT, ACCENT, tick="#FFFFFF")
                selected_active = indicator(
                    ACCENT_DARK,
                    ACCENT_DARK,
                    tick="#FFFFFF",
                )
                disabled = indicator(SAGE, BORDER)
                disabled_selected = indicator(
                    DISABLED_TEXT,
                    DISABLED_TEXT,
                    tick=TEXT,
                )
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
            self._style_element_failures.add(element_name)
            return

    def _create_brand_icon(self, size: int = 42) -> ImageTk.PhotoImage:
        raster = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        gradient = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        gradient_draw = ImageDraw.Draw(gradient)
        for x in range(size):
            position = x / max(1, size - 1)
            colour = interpolate_hex_colour(ACCENT, CYAN, position)
            gradient_draw.line((x, 0, x, size), fill=colour)
        mask = Image.new("L", (size, size), 0)
        ImageDraw.Draw(mask).rounded_rectangle(
            (1, 1, size - 2, size - 2),
            radius=max(8, size // 4),
            fill=255,
        )
        raster.alpha_composite(Image.composite(gradient, raster, mask))
        draw = ImageDraw.Draw(raster)
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
            fill=CREAM_YELLOW,
        )
        line_left = left + round(size * 0.08)
        for ratio in (0.45, 0.58, 0.70):
            y = round(size * ratio)
            draw.rounded_rectangle(
                (line_left, y, right - round(size * 0.07), y + 2),
                radius=1,
                fill=DUSTY_BLUE,
            )
        icon = ImageTk.PhotoImage(raster, master=self)
        gradient.close()
        mask.close()
        raster.close()
        return icon

    def _render_header_art(self, _event: tk.Event | None = None) -> None:
        """Render decorative line art only inside the header's reserved cell."""

        if not hasattr(self, "header_art_canvas"):
            return
        try:
            canvas = self.header_art_canvas
            width = max(1, canvas.winfo_width())
            height = max(1, canvas.winfo_height())
            signature = (width, height, self._layout_mode)
            if signature == self._header_art_last_signature:
                return
            self._header_art_last_signature = signature
            canvas.delete("all")
            if self._layout_mode != "wide" or width < 96 or height < 40:
                return
            line = _canvas_doodle_line_colour(BG, CYAN, strength=0.13)
            accent = _canvas_doodle_line_colour(BG, ACCENT, strength=0.23)
            canvas.create_arc(
                6,
                5,
                width - 7,
                height - 5,
                start=198,
                extent=146,
                style="arc",
                outline=line,
                width=1,
                dash=(4, 7),
                tags=("header_doodle",),
            )
            draw_rocket_doodle(
                canvas,
                (width * 0.41, 5, width * 0.67, height - 4),
                colour=line,
                accent=accent,
                tags=("header_doodle",),
            )
            for x, y, radius in (
                (width * 0.18, height * 0.67, 2),
                (width * 0.79, height * 0.27, 3),
            ):
                canvas.create_oval(
                    x - radius,
                    y - radius,
                    x + radius,
                    y + radius,
                    fill=accent,
                    outline="",
                    tags=("header_doodle",),
                )
        except tk.TclError:
            return

    def _build_ui(self) -> None:
        self.header_card = RoundedCard(
            self,
            fill=BG,
            background=BG,
            outline=BG,
            shadow=BG,
            radius=30,
            inset=12,
            height=max(88, round(92 * self._display_scale)),
        )
        self.header_card.pack(fill="x", padx=20, pady=(16, 10))
        self.header = self.header_card.inner
        self.header.columnconfigure(0, weight=1)
        self.title_box = tk.Frame(self.header, bg=BG)
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
            text="把文件拖进来，选择功能，然后开始处理",
            style="HeaderSubtle.TLabel",
            justify="left",
        )
        self.header_subtitle.grid(row=1, column=1, sticky="nw", pady=(2, 0))
        self.header_art_canvas = tk.Canvas(
            self.header,
            width=142,
            height=58,
            bg=BG,
            highlightthickness=0,
            borderwidth=0,
            takefocus=0,
        )
        self.header_art_canvas.grid(row=0, column=1, sticky="e", padx=(12, 4))
        self.header_art_canvas.bind(
            "<Configure>",
            self._render_header_art,
            add="+",
        )
        self.header_actions = tk.Frame(self.header, bg=BG)
        self.header_actions.grid(row=0, column=2, sticky="e", padx=(16, 0))
        self.engine_select_button = ttk.Button(
            self.header_actions,
            text=office_engine_button_text("auto"),
            style="Nav.TButton",
            command=self._show_office_engine_menu,
        )
        self.office_engine_menu = tk.Menu(
            self,
            tearoff=False,
            bg=PANEL,
            fg=TEXT,
            activebackground=ACCENT_SOFT,
            activeforeground=TEXT,
            selectcolor=CYAN,
            disabledforeground=MUTED,
            font=("Microsoft YaHei UI", 10),
        )
        self._style_popup_menu(self.office_engine_menu)
        for value in OFFICE_ENGINE_VALUES:
            self.office_engine_menu.add_radiobutton(
                label=OFFICE_ENGINE_MENU_LABELS[value],
                variable=self.office_engine_preference,
                value=value,
                command=lambda selected=value: self._select_office_engine(selected),
            )
        self.particle_effect_button = ttk.Button(
            self.header_actions,
            text=particle_effect_button_text(self.particle_effects_var.get()),
            style=(
                "ParticleOn.TButton"
                if self.particle_effects_var.get()
                else "Quiet.TButton"
            ),
            command=self._toggle_particle_effects,
        )
        self.engine_select_button.pack(side="left", padx=(0, 7))
        self.particle_effect_button.pack(side="left", padx=(0, 7))
        self.ui_theme_menu = tk.Menu(
            self,
            tearoff=False,
            bg=PANEL,
            fg=TEXT,
            activebackground=CREAM_BLUE,
            activeforeground=TEXT,
            selectcolor=CYAN,
            disabledforeground=MUTED,
            font=("Microsoft YaHei UI", 10),
        )
        self._style_popup_menu(self.ui_theme_menu)
        for value, label in UI_THEME_LABELS.items():
            self.ui_theme_menu.add_radiobutton(
                label=label,
                variable=self.ui_theme_var,
                value=value,
                command=lambda selected=value: self._select_ui_theme(selected),
            )
        self.preferences_button = ttk.Button(
            self.header_actions,
            text="偏好设置  ⚙",
            style="Nav.TButton",
            command=self._show_preferences_menu,
        )
        self.preferences_button.pack(side="left")
        self.preferences_menu = tk.Menu(
            self,
            tearoff=False,
            bg=PANEL,
            fg=TEXT,
            activebackground=CREAM_BLUE,
            activeforeground=TEXT,
            selectcolor=CYAN,
            disabledforeground=MUTED,
            font=("Microsoft YaHei UI", 10),
        )
        self._style_popup_menu(self.preferences_menu)
        self.preferences_menu.add_cascade(
            label="Office 默认引擎",
            menu=self.office_engine_menu,
        )
        self.preferences_menu.add_cascade(
            label="界面画风",
            menu=self.ui_theme_menu,
        )
        self.preferences_menu.add_checkbutton(
            label="粒子动效",
            variable=self.particle_effects_var,
            command=self._apply_particle_effects_preference,
        )
        self.preferences_menu.add_separator()
        self.preferences_menu.add_checkbutton(
            label="显示需要额外安装的功能",
            variable=self.show_unavailable_var,
            command=self._rebuild_operation_tree,
        )
        self.preferences_menu.add_command(
            label="Office 兼容与引擎说明",
            command=self._show_office_compatibility_info,
        )
        self.sidebar_toggle_button = ttk.Button(
            self.header,
            text="☰  工具目录",
            style="Nav.TButton",
            command=self._toggle_sidebar,
        )

        self.body = ttk.Frame(self, padding=(20, 4, 20, 20), style="Main.TFrame")
        self.body.pack(fill="both", expand=True)
        self.body.columnconfigure(1, weight=1)
        self.body.rowconfigure(0, weight=1)

        self.sidebar_shell = RoundedCard(
            self.body,
            fill=SIDEBAR_BG,
            background=BG,
            outline=BORDER,
            shadow=SHADOW,
            radius=30,
            inset=12,
            width=310,
        )
        self.sidebar_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.sidebar = self.sidebar_shell.inner
        self._measure_catalog_widths()
        self.sidebar.rowconfigure(4, weight=1)
        self.sidebar.columnconfigure(0, weight=1)
        ttk.Label(self.sidebar, text="选择功能", style="SidebarTitle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 9)
        )
        self.category_bar = tk.Frame(self.sidebar, bg=SIDEBAR_BG)
        self.category_bar.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.category_bar.rowconfigure(1, minsize=5)
        self.category_buttons: dict[str, ttk.Button] = {}
        for index, (root_name, label) in enumerate(
            (("文档工具", "文档"), ("图片工具", "图片"), ("视频工具", "视频"))
        ):
            self.category_bar.columnconfigure(index, weight=1, uniform="catalog_root")
            button = ttk.Button(
                self.category_bar,
                text=label,
                style="Category.TButton",
                command=lambda selected=root_name: self._select_catalog_root(selected),
            )
            button.grid(
                row=0,
                column=index,
                sticky="ew",
                padx=(0 if index == 0 else 4, 0),
            )
            self.category_buttons[root_name] = button
        self.category_indicator = tk.Canvas(
            self.category_bar,
            height=5,
            bg=SIDEBAR_BG,
            highlightthickness=0,
            borderwidth=0,
        )
        self.category_indicator.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        self.search_var = tk.StringVar()
        search = ttk.Entry(self.sidebar, textvariable=self.search_var)
        search.grid(row=2, column=0, sticky="ew", pady=(0, 10))
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
            foreground=CYAN,
            background=SIDEBAR_BG,
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        self.operation_tree.tag_configure("catalog_ready", foreground=SUCCESS)
        self.operation_tree.tag_configure("catalog_external", foreground=ACCENT)
        self.operation_tree.tag_configure("catalog_unavailable", foreground=ORANGE)
        self.operation_tree.grid(row=4, column=0, sticky="nsew")
        self.operation_tree.bind("<<TreeviewSelect>>", self._on_operation_selected)
        tree_scroll = ttk.Scrollbar(
            self.sidebar, orient="vertical", command=self.operation_tree.yview
        )
        tree_scroll.grid(row=4, column=1, sticky="ns")
        tree_x_scroll = ttk.Scrollbar(
            self.sidebar, orient="horizontal", command=self.operation_tree.xview
        )
        tree_x_scroll.grid(row=5, column=0, sticky="ew", pady=(3, 0))
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
        self.catalog_legend.grid_remove()

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

        self.content_shell = RoundedCard(
            self.body,
            fill=PANEL,
            background=BG,
            outline=BORDER,
            shadow=SHADOW,
            radius=30,
            inset=16,
        )
        self.content_shell.grid(row=0, column=1, sticky="nsew")
        self.content = self.content_shell.inner
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
            bg=CARD_SAGE,
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

        self.workspace_card = RoundedCard(
            self.content,
            fill=PANEL,
            background=PANEL,
            outline=PANEL,
            shadow=PANEL,
            radius=22,
            inset=4,
        )
        self.workspace_card.grid(row=2, column=0, sticky="nsew")
        self.workspace_card.inner.rowconfigure(0, weight=0)
        self.workspace_card.inner.columnconfigure(0, weight=1)
        self.workspace_tabs = tk.Frame(self.workspace_card.inner, bg=PANEL)
        self.workspace_tabs.grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.setup_tab_button = ttk.Button(
            self.workspace_tabs,
            text="处理工作台",
            style="WorkspaceTabActive.TButton",
            command=lambda: self._select_workspace_tab("setup"),
        )
        self.setup_tab_button.grid(row=0, column=0, padx=(0, 2))
        self.log_tab_button = ttk.Button(
            self.workspace_tabs,
            text="处理记录",
            style="WorkspaceTab.TButton",
            command=lambda: self._select_workspace_tab("log"),
        )
        self.log_tab_button.grid(row=0, column=1, padx=(0, 0))
        self.workspace_indicator = tk.Canvas(
            self.workspace_tabs,
            width=1,
            height=4,
            bg=PANEL,
            highlightthickness=0,
            borderwidth=0,
        )
        self.workspace_indicator.grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(2, 0),
        )
        self.notebook_style_name = "HiddenTabs.TNotebook"
        self.notebook = ttk.Notebook(
            self.workspace_card.inner,
            style=self.notebook_style_name,
        )
        self.notebook.grid(row=1, column=0, sticky="nsew")
        self.workspace_card.inner.rowconfigure(1, weight=1)
        self.setup_page = ttk.Frame(self.notebook, style="Workspace.TFrame")
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
        self.setup_scroll = ttk.Scrollbar(
            self.setup_page,
            orient="vertical",
            command=self.setup_canvas.yview,
            style="Workspace.Vertical.TScrollbar",
            takefocus=False,
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
        log_tab = ttk.Frame(self.notebook, style="Canvas.TFrame", padding=(4, 14, 4, 4))
        self.notebook.add(self.setup_page, text="  开始处理  ")
        self.notebook.add(log_tab, text="  处理记录  ")
        self.log_tab = log_tab
        self.notebook.bind("<<NotebookTabChanged>>", self._on_workspace_tab_changed, add="+")
        self.setup_tab.columnconfigure(0, weight=1)
        self.setup_tab.rowconfigure(2, weight=1)
        log_tab.columnconfigure(0, weight=1)
        log_tab.rowconfigure(0, weight=1)

        self.file_step_label = ttk.Label(
            self.setup_tab,
            text="●  1   添加文件",
            style="StepTitle.TLabel",
        )
        self.file_step_label.grid(row=0, column=0, sticky="w", pady=(0, 7))
        self.file_toolbar = ttk.Frame(self.setup_tab, style="Panel.TFrame")
        self.file_toolbar.grid(row=1, column=0, sticky="ew", pady=(0, 8))
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
            activeforeground=TEXT,
            selectcolor=CYAN,
            disabledforeground=MUTED,
            relief="flat",
            borderwidth=1,
        )
        self._style_popup_menu(self.file_more_menu)
        self.file_more_menu.add_command(label="添加文件夹", command=self._add_folder)
        self.file_more_menu.add_separator()
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
            self.file_toolbar, text="尚未添加文件", style="CanvasSubtle.TLabel"
        )

        self.file_card = RoundedCard(
            self.setup_tab,
            fill=DROP_SURFACE,
            background=PANEL,
            outline=DROP_BORDER,
            shadow=SHADOW,
            radius=26,
            inset=12,
        )
        self.file_card.grid(
            row=2,
            column=0,
            columnspan=2,
            sticky="nsew",
        )
        self.file_drop_frame = self.file_card.inner
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
        self.empty_drop_canvas = tk.Canvas(
            self.file_drop_frame,
            bg=DROP_SURFACE,
            highlightthickness=0,
            borderwidth=0,
            height=210,
            cursor="hand2",
        )
        self.empty_drop_canvas.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.empty_drop_canvas.bind(
            "<Configure>", self._schedule_empty_drop_redraw, add="+"
        )
        self.empty_drop_canvas.bind("<Button-1>", lambda _event: self._add_files())

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
        self.file_tree.tag_configure("even", background=CARD_YELLOW)
        self.file_tree.tag_configure("odd", background=TREE_ODD)
        self.file_tree.grid(row=1, column=0, sticky="nsew")
        self.file_tree.bind(
            "<<TreeviewSelect>>",
            self._on_file_tree_preview_selected,
            add="+",
        )
        self.file_scroll = ttk.Scrollbar(
            self.file_drop_frame,
            orient="vertical",
            command=self.file_tree.yview,
        )
        self.file_scroll.grid(row=1, column=1, sticky="ns")
        self.file_horizontal_scroll = ttk.Scrollbar(
            self.file_drop_frame,
            orient="horizontal",
            command=self.file_tree.xview,
        )
        self.file_horizontal_scroll.grid(row=2, column=0, sticky="ew")
        self.file_tree.configure(
            yscrollcommand=self.file_scroll.set,
            xscrollcommand=self.file_horizontal_scroll.set,
        )
        self._configure_file_drop_targets()

        self.output_card = RoundedCard(
            self.setup_tab,
            fill=PANEL_ALT,
            background=PANEL,
            outline=BORDER,
            shadow=SHADOW,
            radius=24,
            inset=11,
            height=max(58, round(62 * self._display_scale)),
        )
        self.output_card.grid(row=3, column=0, sticky="ew", pady=(12, 8))
        self.output_frame = self.output_card.inner
        self.output_label = ttk.Label(
            self.output_frame,
            text="输出文件夹",
            background=PANEL_ALT,
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

        self.settings_divider = tk.Frame(
            self.setup_tab,
            bg=PANEL,
            height=1,
            borderwidth=0,
        )
        self.settings_divider.grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=9
        )
        self.parameters_header = ttk.Frame(self.setup_tab, style="Panel.TFrame")
        self.parameters_header.grid(
            row=5, column=0, columnspan=2, sticky="ew", pady=(0, 8)
        )
        self.parameters_header.columnconfigure(0, weight=1)
        self.parameters_title = ttk.Label(
            self.parameters_header,
            text="●  2   选择设置",
            style="StepTitle.TLabel",
        )
        self.parameters_title.grid(row=0, column=0, sticky="w")
        self.parameters_scroll_hint = ttk.Label(
            self.parameters_header,
            text="滚动查看更多参数 ↓",
            style="CanvasSubtle.TLabel",
        )
        self.parameters_frame = ttk.Frame(self.setup_tab, style="Panel.TFrame")
        self.parameters_frame.grid(row=6, column=0, columnspan=2, sticky="ew")
        self.parameters_frame.columnconfigure(1, weight=1)
        self.parameters_frame.bind(
            "<Configure>", self._on_parameters_frame_configure, add="+"
        )

        self.image_preview_card = RoundedCard(
            self.setup_tab,
            fill=CARD_BLUE,
            background=PANEL,
            outline=BORDER,
            shadow=SHADOW,
            radius=22,
            inset=12,
            auto_height=True,
            height=120,
        )
        self.image_preview_card.grid(
            row=7,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(12, 4),
        )
        self.image_preview_card.grid_remove()
        self.image_preview_frame = self.image_preview_card.inner
        self.image_preview_title = tk.Label(
            self.image_preview_frame,
            text="原图 / 效果预览",
            bg=CARD_BLUE,
            fg=TEXT,
            font=("Microsoft YaHei UI", 10, "bold"),
            anchor="w",
        )
        self.image_preview_title.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.image_preview_toolbar = ttk.Frame(
            self.image_preview_frame,
            style="Soft.TFrame",
            padding=(10, 7),
        )
        self.image_preview_toolbar.grid(row=1, column=0, sticky="ew")
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
        self.image_preview_panes.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.image_preview_frame.columnconfigure(0, weight=1)
        self.image_preview_original_box = RoundedCard(
            self.image_preview_panes,
            fill=CREAM_YELLOW,
            background=PANEL_ALT,
            outline=BORDER,
            shadow=SHADOW,
            radius=18,
            inset=10,
            height=294,
        )
        self.image_preview_result_box = RoundedCard(
            self.image_preview_panes,
            fill=CREAM_YELLOW,
            background=PANEL_ALT,
            outline=BORDER,
            shadow=SHADOW,
            radius=18,
            inset=10,
            height=294,
        )
        self.image_preview_original_title = tk.Label(
            self.image_preview_original_box.inner,
            text="原图",
            bg=CREAM_YELLOW,
            fg=TEXT,
            font=("Microsoft YaHei UI", 10, "bold"),
            anchor="w",
        )
        self.image_preview_original_title.pack(fill="x", pady=(0, 7))
        self.image_preview_result_title = tk.Label(
            self.image_preview_result_box.inner,
            text="参数效果",
            bg=CREAM_YELLOW,
            fg=TEXT,
            font=("Microsoft YaHei UI", 10, "bold"),
            anchor="w",
        )
        self.image_preview_result_title.pack(fill="x", pady=(0, 7))
        self.image_preview_original_canvas = tk.Canvas(
            self.image_preview_original_box.inner,
            height=250,
            bg=CREAM_BLUE,
            highlightthickness=0,
            borderwidth=0,
        )
        self.image_preview_original_canvas.pack(fill="both", expand=True)
        self.image_preview_result_canvas = tk.Canvas(
            self.image_preview_result_box.inner,
            height=250,
            bg=CREAM_BLUE,
            highlightthickness=0,
            borderwidth=0,
        )
        self.image_preview_result_canvas.pack(fill="both", expand=True)
        self.image_preview_original_canvas.bind(
            "<Configure>",
            lambda _event: self._schedule_image_preview_redraw("original"),
            add="+",
        )
        self.image_preview_result_canvas.bind(
            "<Configure>",
            lambda _event: self._schedule_image_preview_redraw("result"),
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
        self.image_preview_details.grid(row=3, column=0, sticky="ew", pady=(8, 0))

        self.log_text = tk.Text(
            log_tab,
            bg=INPUT_SURFACE,
            fg=TEXT,
            insertbackground=TEXT,
            selectbackground=ACCENT_SOFT,
            selectforeground=TEXT,
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

        self.footer_card = RoundedCard(
            self.content,
            fill=PANEL_ALT,
            background=PANEL,
            outline=BORDER,
            shadow=SHADOW,
            radius=26,
            inset=12,
            height=max(72, round(76 * self._display_scale)),
        )
        self.footer_card.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        self.footer = self.footer_card.inner
        self.footer.columnconfigure(0, weight=1)
        self.progress_box = tk.Frame(self.footer, bg=PANEL_ALT)
        self.progress_box.grid(row=0, column=0, sticky="ew", padx=(0, 14))
        self.progress_box.columnconfigure(0, weight=1)
        self.progress_var = tk.DoubleVar(value=0)
        self.progressbar = ttk.Progressbar(
            self.progress_box, variable=self.progress_var, maximum=100
        )
        self.progress_track = tk.Canvas(
            self.progress_box,
            height=max(12, round(12 * self._display_scale)),
            bg=PANEL_ALT,
            highlightthickness=0,
            borderwidth=0,
        )
        self.progress_track.grid(row=0, column=0, sticky="ew")
        self.progress_track.bind(
            "<Configure>", self._schedule_progress_track_redraw, add="+"
        )
        self.progress_var.trace_add("write", lambda *_args: self._render_progress_track())
        self.progress_percent_label = ttk.Label(
            self.progress_box,
            text="0%",
            style="FooterSubtle.TLabel",
            width=6,
            anchor="e",
        )
        self.progress_percent_label.grid(row=0, column=1, padx=(8, 0), sticky="e")
        self.progress_label = ttk.Label(
            self.progress_box,
            text="准备就绪",
            style="FooterSubtle.TLabel",
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
            text="3   开始处理  →",
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
        self.bind_class("TButton", "<Enter>", self._on_button_enter, add="+")
        self.bind_class("TButton", "<ButtonRelease-1>", self._on_button_release, add="+")
        self._refresh_category_buttons()
        self._apply_responsive_layout(
            self._initial_window_width,
            window_height=self._initial_window_height,
            force=True,
        )
        self._refresh_file_tree()
        self._start_progress_shimmer()
        self._schedule_setup_scroll_refresh()

    def _motion_mode(self) -> str:
        return normalize_motion_mode(self.motion_mode_var.get())

    def _run_content_transition(self, callback: Callable[[], None]) -> None:
        """Hide synchronous reflow behind a short left-to-right reveal."""

        if self._window_resizing or not self._window_mapped:
            callback()
            return
        if self._content_transition_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._content_transition_job)
            self._content_transition_job = None
        self._content_transition_generation += 1
        generation = self._content_transition_generation
        timing = motion_effect_timing(
            self._motion_mode(),
            "transition",
            busy=bool(self.worker and self.worker.is_alive()),
            minimized=self.state() in {"iconic", "withdrawn"},
        )
        if not timing.enabled or not self.winfo_viewable():
            callback()
            return

        if self._content_transition_overlay is None:
            self._content_transition_overlay = tk.Canvas(
                self.workspace_card.inner,
                bg=PANEL,
                highlightthickness=0,
                borderwidth=0,
                takefocus=False,
            )
        overlay = self._content_transition_overlay
        overlay.configure(bg=PANEL)
        overlay.delete("all")
        width = max(1, self.notebook.winfo_width())
        height = max(1, self.notebook.winfo_height())
        origin_x = self.notebook.winfo_x()
        origin_y = self.notebook.winfo_y()
        overlay.create_rectangle(0, 0, width, 4, fill=ACCENT, outline="")
        overlay.create_rectangle(
            0,
            4,
            width,
            min(height, 8),
            fill=interpolate_hex_colour(PANEL, CYAN, 0.22),
            outline="",
        )
        overlay.place(x=origin_x, y=origin_y, width=width, height=height)
        overlay.tk.call("raise", overlay._w)

        def rebuild() -> None:
            if generation != self._content_transition_generation or self._closing:
                return
            try:
                callback()
            except Exception:
                overlay.place_forget()
                self._content_transition_job = None
                raise

            def reveal(index: int = 0) -> None:
                if generation != self._content_transition_generation or self._closing:
                    return
                progress = ease_out_cubic(index / max(1, timing.frames))
                current_width = max(1, self.notebook.winfo_width())
                current_height = max(1, self.notebook.winfo_height())
                left = round(current_width * progress)
                remaining = max(1, current_width - left)
                try:
                    overlay.place_configure(
                        x=self.notebook.winfo_x() + left,
                        y=self.notebook.winfo_y(),
                        width=remaining,
                        height=current_height,
                    )
                    overlay.tk.call("raise", overlay._w)
                except tk.TclError:
                    self._content_transition_job = None
                    return
                if index < timing.frames:
                    self._content_transition_job = self.after(
                        timing.step_ms,
                        reveal,
                        index + 1,
                    )
                else:
                    overlay.place_forget()
                    self._content_transition_job = None

            self._content_transition_job = self.after_idle(reveal)

        # Let Tk paint the cover before the parameter tree performs its reflow.
        self._content_transition_job = self.after(18, rebuild)

    def _select_workspace_tab(self, tab: str) -> None:
        target = self.setup_page if tab == "setup" else self.log_tab
        with contextlib.suppress(tk.TclError):
            if self.notebook.select() == str(target):
                return

        def select_target() -> None:
            with contextlib.suppress(tk.TclError):
                self.notebook.select(target)
            self._pulse_workspace_tabs()

        self._run_content_transition(select_target)

    def _on_workspace_tab_changed(self, _event: tk.Event | None = None) -> None:
        setup_selected = self.notebook.select() == str(self.setup_page)
        self.setup_tab_button.configure(
            style=(
                "WorkspaceTabActive.TButton"
                if setup_selected
                else "WorkspaceTab.TButton"
            )
        )
        self.log_tab_button.configure(
            style=(
                "WorkspaceTab.TButton"
                if setup_selected
                else "WorkspaceTabActive.TButton"
            )
        )

    @staticmethod
    def _indicator_target_geometry(
        button: ttk.Button,
        container: tk.Misc,
    ) -> tuple[float, float]:
        left = button.winfo_rootx() - container.winfo_rootx() + 8
        width = max(22.0, button.winfo_width() - 16.0)
        return float(left), float(width)

    def _position_indicator(
        self,
        canvas: tk.Canvas,
        button: ttk.Button,
        container: tk.Misc,
    ) -> None:
        try:
            left, width = self._indicator_target_geometry(button, container)
            height = max(4, canvas.winfo_height())
            items = canvas.find_withtag("indicator")
            if items:
                item = items[0]
                canvas.coords(item, left, height / 2, left + width, height / 2)
                canvas.itemconfigure(
                    item,
                    fill=CYAN,
                    width=max(3, height - 2),
                )
                for extra in items[1:]:
                    canvas.delete(extra)
            else:
                canvas.create_line(
                    left,
                    height / 2,
                    left + width,
                    height / 2,
                    fill=CYAN,
                    width=max(3, height - 2),
                    capstyle="round",
                    tags=("indicator",),
                )
        except tk.TclError:
            return

    def _animate_indicator(
        self,
        canvas: tk.Canvas,
        button: ttk.Button,
        container: tk.Misc,
        *,
        job_attribute: str,
    ) -> None:
        current_job = getattr(self, job_attribute, None)
        if current_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(current_job)
            setattr(self, job_attribute, None)
        try:
            target_left, target_width = self._indicator_target_geometry(button, container)
            coords = canvas.coords("indicator")
            start_left = float(coords[0]) if len(coords) >= 4 else target_left
            start_width = max(1.0, float(coords[2]) - float(coords[0])) if len(coords) >= 4 else target_width
        except (tk.TclError, TypeError, ValueError):
            return
        busy = bool(self.worker and self.worker.is_alive())
        frames = q_bounce_transition_plan(self._motion_mode(), busy=busy)
        timing = motion_effect_timing(
            self._motion_mode(),
            "transition",
            busy=busy,
            minimized=self.state() in {"iconic", "withdrawn"},
        )
        if not timing.enabled or len(frames) <= 1:
            self._position_indicator(canvas, button, container)
            return

        def step(index: int) -> None:
            frame = frames[index]
            left = start_left + (target_left - start_left) * frame.progress
            base_width = start_width + (target_width - start_width) * min(1.0, frame.progress)
            width = base_width * frame.scale
            centre = left + base_width / 2
            height = max(4, canvas.winfo_height())
            try:
                items = canvas.find_withtag("indicator")
                if items:
                    item = items[0]
                else:
                    item = canvas.create_line(
                        0,
                        height / 2,
                        0,
                        height / 2,
                        capstyle="round",
                        tags=("indicator",),
                    )
                canvas.coords(
                    item,
                    centre - width / 2,
                    height / 2,
                    centre + width / 2,
                    height / 2,
                )
                canvas.itemconfigure(
                    item,
                    fill=CYAN,
                    width=max(3, height - 2),
                )
                for extra in items[1:]:
                    canvas.delete(extra)
            except tk.TclError:
                setattr(self, job_attribute, None)
                return
            if index + 1 < len(frames):
                setattr(
                    self,
                    job_attribute,
                    self.after(timing.step_ms, step, index + 1),
                )
            else:
                setattr(self, job_attribute, None)
                self._position_indicator(canvas, button, container)

        step(0)

    def _pulse_workspace_tabs(self) -> None:
        """Add a local, non-overlay transition without moving page contents."""

        if self._workspace_transition_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._workspace_transition_job)
            self._workspace_transition_job = None
        timing = motion_effect_timing(
            self._motion_mode(),
            "transition",
            busy=bool(self.worker and self.worker.is_alive()),
            minimized=self.state() in {"iconic", "withdrawn"},
        )
        if not timing.enabled:
            return
        selected = (
            self.setup_tab_button
            if self.notebook.select() == str(self.setup_page)
            else self.log_tab_button
        )
        self._animate_indicator(
            self.workspace_indicator,
            selected,
            self.workspace_tabs,
            job_attribute="_workspace_transition_job",
        )

    def _realign_indicators_after_layout(self) -> None:
        """Snap idle indicators to their controls after a responsive reflow."""

        self._indicator_realign_job = None
        if self._catalog_transition_job is None:
            selected_category = self.category_buttons.get(
                self.catalog_root_var.get()
            )
            if selected_category is not None:
                self._position_indicator(
                    self.category_indicator,
                    selected_category,
                    self.category_bar,
                )
        if self._workspace_transition_job is None:
            selected_tab = (
                self.setup_tab_button
                if self.notebook.select() == str(self.setup_page)
                else self.log_tab_button
            )
            self._position_indicator(
                self.workspace_indicator,
                selected_tab,
                self.workspace_tabs,
            )

    def _schedule_indicator_realign(self) -> None:
        if self._indicator_realign_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._indicator_realign_job)
        self._indicator_realign_job = self.after_idle(
            self._realign_indicators_after_layout
        )

    def _motion_delay(self) -> int:
        minimized = False
        with contextlib.suppress(tk.TclError):
            minimized = self.state() in {"iconic", "withdrawn"}
        return motion_frame_delay(
            self._motion_mode(),
            busy=bool(self.worker and self.worker.is_alive()),
            minimized=minimized,
        )

    def _schedule_progress_track_redraw(
        self,
        _event: tk.Event | None = None,
        *,
        delay: int | None = None,
    ) -> None:
        if self._closing or not hasattr(self, "progress_track"):
            return
        if self._progress_track_resize_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._progress_track_resize_job)
        wait_ms = (48 if self._window_resizing else 18) if delay is None else max(0, int(delay))

        def redraw() -> None:
            self._progress_track_resize_job = None
            self._render_progress_track()

        with contextlib.suppress(tk.TclError):
            self._progress_track_resize_job = self.after(wait_ms, redraw)

    def _render_progress_track(self) -> None:
        if not hasattr(self, "progress_track"):
            return
        try:
            width = max(1, self.progress_track.winfo_width())
            height = max(8, self.progress_track.winfo_height())
            percent = min(100.0, max(0.0, float(self.progress_var.get())))
            completed_width = max(0, round(width * percent / 100.0))
            inset = 1
            radius = max(4, (height - inset * 2) // 2)
            signature = (width, height, CURRENT_UI_THEME)
            if signature != self._progress_track_signature:
                self.progress_track.delete("all")
                middle_right = max(radius, width - radius)
                background = (
                    self.progress_track.create_rectangle(
                        radius,
                        inset,
                        middle_right,
                        height - inset,
                        fill=SAGE,
                        outline="",
                    ),
                    self.progress_track.create_oval(
                        inset,
                        inset,
                        min(width - inset, inset + radius * 2),
                        height - inset,
                        fill=SAGE,
                        outline="",
                    ),
                    self.progress_track.create_oval(
                        max(inset, width - inset - radius * 2),
                        inset,
                        width - inset,
                        height - inset,
                        fill=SAGE,
                        outline="",
                    ),
                )
                colours = (ACCENT_DARK, ACCENT, CYAN)
                segment_items: list[int] = []
                segment_count = 30
                for index in range(segment_count):
                    progress = index / max(1, segment_count - 1)
                    if progress < 0.55:
                        colour = interpolate_hex_colour(
                            colours[0], colours[1], progress / 0.55
                        )
                    else:
                        colour = interpolate_hex_colour(
                            colours[1], colours[2], (progress - 0.55) / 0.45
                        )
                    segment_items.append(
                        self.progress_track.create_rectangle(
                            0,
                            inset,
                            0,
                            height - inset,
                            fill=colour,
                            outline="",
                            state="hidden",
                        )
                    )
                self._progress_track_items = {
                    "background": background,
                    "segments": tuple(segment_items),
                    "left_cap": self.progress_track.create_oval(
                        0,
                        inset,
                        0,
                        height - inset,
                        fill=ACCENT_DARK,
                        outline="",
                        state="hidden",
                    ),
                    "right_cap": self.progress_track.create_oval(
                        0,
                        inset,
                        0,
                        height - inset,
                        fill=ACCENT,
                        outline="",
                        state="hidden",
                    ),
                    "shimmer": self.progress_track.create_rectangle(
                        0,
                        inset + 2,
                        0,
                        height - inset - 2,
                        fill=CREAM_YELLOW,
                        outline="",
                        stipple="gray50",
                        state="hidden",
                    ),
                }
                self._progress_track_signature = signature

            segment_items = self._progress_track_items.get("segments", ())
            left_cap = self._progress_track_items.get("left_cap")
            right_cap = self._progress_track_items.get("right_cap")
            shimmer = self._progress_track_items.get("shimmer")
            if completed_width <= 1:
                for item in (*segment_items, left_cap, right_cap, shimmer):
                    if item is not None:
                        self.progress_track.itemconfigure(item, state="hidden")
                return

            cap_right = min(completed_width, inset + radius * 2)
            self.progress_track.coords(
                left_cap, inset, inset, cap_right, height - inset
            )
            self.progress_track.itemconfigure(left_cap, state="normal")
            fill_left = min(completed_width, radius)
            fill_right = max(fill_left, completed_width - radius)
            segment_count = len(segment_items)
            for index, item in enumerate(segment_items):
                left = round(fill_left + (fill_right - fill_left) * index / segment_count)
                right = round(
                    fill_left + (fill_right - fill_left) * (index + 1) / segment_count
                ) + 1
                self.progress_track.coords(
                    item,
                    left,
                    inset,
                    min(completed_width, right),
                    height - inset,
                )
                self.progress_track.itemconfigure(
                    item,
                    state="normal" if right > left and fill_right > fill_left else "hidden",
                )
            if completed_width > radius * 2:
                self.progress_track.coords(
                    right_cap,
                    completed_width - radius * 2,
                    inset,
                    completed_width,
                    height - inset,
                )
                self.progress_track.itemconfigure(
                    right_cap,
                    fill=CYAN if percent > 70 else ACCENT,
                    state="normal",
                )
            else:
                self.progress_track.itemconfigure(right_cap, state="hidden")

            active_shimmer = (
                self._motion_mode() != "off"
                and self.worker is not None
                and self.worker.is_alive()
            )
            if active_shimmer and shimmer is not None:
                shimmer_width = max(18, round(width * 0.06))
                centre = (
                    round(
                        (completed_width + shimmer_width)
                        * self._progress_shimmer_phase
                    )
                    - shimmer_width
                )
                shimmer_left = max(0, centre - shimmer_width)
                shimmer_right = min(completed_width, centre + shimmer_width)
                if shimmer_right > shimmer_left:
                    self.progress_track.coords(
                        shimmer,
                        shimmer_left,
                        inset + 2,
                        shimmer_right,
                        height - inset - 2,
                    )
                    self.progress_track.itemconfigure(shimmer, state="normal")
                else:
                    self.progress_track.itemconfigure(shimmer, state="hidden")
            elif shimmer is not None:
                self.progress_track.itemconfigure(shimmer, state="hidden")
        except (tk.TclError, TypeError, ValueError):
            return

    def _start_progress_shimmer(self) -> None:
        if self._progress_shimmer_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._progress_shimmer_job)
            self._progress_shimmer_job = None
        self._progress_shimmer_phase = 0.0

        def tick() -> None:
            if self._closing:
                self._progress_shimmer_job = None
                return
            if self._motion_mode() != "off" and self.worker and self.worker.is_alive():
                self._progress_shimmer_phase = (self._progress_shimmer_phase + 0.032) % 1.0
                self._render_progress_track()
            delay = max(48, self._motion_delay()) if self._motion_mode() != "off" else 400
            self._progress_shimmer_job = self.after(delay, tick)

        self._progress_shimmer_job = self.after(180, tick)

    def _show_preferences_menu(self) -> None:
        self._popup_menu_below(self.preferences_menu, self.preferences_button)

    def _select_ui_theme(self, value: str) -> None:
        """Apply a colour-only theme immediately and remember the choice."""

        theme = normalize_ui_theme(value)
        self.ui_theme_var.set(theme)
        changed = theme != CURRENT_UI_THEME
        if changed:
            _activate_ui_theme(theme)
            self._refresh_ui_theme_colours()
        self._ui_preferences["ui_theme"] = theme
        try:
            save_ui_preferences(self._ui_preferences, self._preferences_path)
        except OSError as exc:
            self._log(f"界面画风已切换，但偏好保存失败：{exc}")
        else:
            if changed:
                self._log(f"界面画风已切换为：{UI_THEME_LABELS[theme]}。")

    @staticmethod
    def _set_widget_colours(widget: tk.Misc, **colours: str) -> None:
        with contextlib.suppress(tk.TclError):
            widget.configure(**colours)

    def _refresh_ui_theme_colours(self) -> None:
        """Recolour existing widgets without rebuilding or moving any of them."""

        self._configure_style()
        self._set_widget_colours(self, bg=BG)

        for _spec, _parent, _label, control, _action, _hint in self.parameter_rows:
            if isinstance(control, ttk.Combobox):
                _style_combobox_popdown(control)

        card_specs = (
            ("header_card", BG, BG, BG, BG),
            ("sidebar_shell", SIDEBAR_BG, BG, BORDER, SHADOW),
            ("content_shell", PANEL, BG, BORDER, SHADOW),
            ("workspace_card", PANEL, PANEL, PANEL, PANEL),
            ("file_card", DROP_SURFACE, PANEL, DROP_BORDER, SHADOW),
            ("output_card", PANEL_ALT, PANEL, BORDER, SHADOW),
            ("image_preview_card", CARD_BLUE, PANEL, BORDER, SHADOW),
            (
                "image_preview_original_box",
                CREAM_YELLOW,
                PANEL_ALT,
                BORDER,
                SHADOW,
            ),
            (
                "image_preview_result_box",
                CREAM_YELLOW,
                PANEL_ALT,
                BORDER,
                SHADOW,
            ),
            ("footer_card", PANEL_ALT, PANEL, BORDER, SHADOW),
        )
        for attribute, fill, background, outline, shadow in card_specs:
            card = getattr(self, attribute, None)
            if isinstance(card, RoundedCard):
                card.set_palette(
                    fill=fill,
                    background=background,
                    outline=outline,
                    shadow=shadow,
                )
        if isinstance(self.simple_parameters_card, RoundedCard):
            self.simple_parameters_card.set_palette(
                fill=PANEL,
                background=PANEL,
                outline=PANEL,
                shadow=PANEL,
            )
        for index, section in enumerate(self.parameter_section_order):
            card = self.parameter_section_frames.get(section)
            if not isinstance(card, RoundedCard):
                continue
            card.set_palette(
                fill=PANEL,
                background=PANEL,
                outline=PANEL,
                shadow=PANEL,
            )
            for child in card.inner.winfo_children():
                if not isinstance(child, tk.Label):
                    continue
                with contextlib.suppress(tk.TclError):
                    if str(child.cget("text")).startswith("●"):
                        child.configure(
                            bg=PANEL,
                            fg=ACCENT_DARK if index % 2 else SUCCESS,
                        )

        native_widgets = (
            ("background_canvas", {"bg": BG}),
            ("title_box", {"bg": BG}),
            ("header_art_canvas", {"bg": BG}),
            ("header_actions", {"bg": BG}),
            ("category_bar", {"bg": SIDEBAR_BG}),
            ("category_indicator", {"bg": SIDEBAR_BG}),
            ("sidebar_resize_handle", {"bg": SHADOW}),
            ("workspace_tabs", {"bg": PANEL}),
            ("workspace_indicator", {"bg": PANEL}),
            ("setup_canvas", {"bg": PANEL}),
            ("settings_divider", {"bg": PANEL}),
            ("empty_drop_canvas", {"bg": DROP_SURFACE}),
            ("image_preview_title", {"bg": CARD_BLUE, "fg": TEXT}),
            (
                "image_preview_original_title",
                {"bg": CREAM_YELLOW, "fg": TEXT},
            ),
            (
                "image_preview_result_title",
                {"bg": CREAM_YELLOW, "fg": TEXT},
            ),
            ("image_preview_original_canvas", {"bg": CREAM_BLUE}),
            ("image_preview_result_canvas", {"bg": CREAM_BLUE}),
            ("progress_box", {"bg": PANEL_ALT}),
            ("progress_track", {"bg": PANEL_ALT}),
        )
        for attribute, colours in native_widgets:
            widget = getattr(self, attribute, None)
            if widget is not None:
                self._set_widget_colours(widget, **colours)

        if hasattr(self, "brand_icon_label"):
            self._brand_icon = self._create_brand_icon()
            self.brand_icon_label.configure(
                image=self._brand_icon,
                background=BG,
            )
        if hasattr(self, "catalog_legend"):
            self.catalog_legend.configure(
                background=SIDEBAR_BG,
                foreground=MUTED,
            )
        if hasattr(self, "output_label"):
            self.output_label.configure(background=PANEL_ALT, foreground=TEXT)
        if hasattr(self, "log_text"):
            self.log_text.configure(
                bg=INPUT_SURFACE,
                fg=TEXT,
                insertbackground=TEXT,
                selectbackground=ACCENT_SOFT,
                selectforeground=TEXT,
            )
        if hasattr(self, "operation_title"):
            self.operation_title.configure(foreground=TEXT)

        for _spec, parent, label, _control, _action, hint in self.parameter_rows:
            with contextlib.suppress(tk.TclError):
                label.configure(background=parent.cget("bg"), foreground=TEXT)
            if hint is not None:
                with contextlib.suppress(tk.TclError):
                    hint.configure(background=parent.cget("bg"), foreground=MUTED)

        for menu_name in (
            "office_engine_menu",
            "ui_theme_menu",
            "preferences_menu",
            "file_more_menu",
        ):
            menu = getattr(self, menu_name, None)
            if isinstance(menu, tk.Menu):
                self._style_popup_menu(menu)

        if hasattr(self, "operation_tree"):
            self.operation_tree.tag_configure(
                "catalog_root",
                foreground=TEXT,
            )
            self.operation_tree.tag_configure(
                "catalog_section",
                foreground=CYAN,
                background=SIDEBAR_BG,
            )
            self.operation_tree.tag_configure("catalog_ready", foreground=SUCCESS)
            self.operation_tree.tag_configure(
                "catalog_external",
                foreground=ACCENT,
            )
            self.operation_tree.tag_configure(
                "catalog_unavailable",
                foreground=ORANGE,
            )
        if hasattr(self, "file_tree"):
            self.file_tree.tag_configure("even", background=CARD_YELLOW)
            self.file_tree.tag_configure("odd", background=TREE_ODD)

        if self.current_operation is None:
            if hasattr(self, "capability_badge"):
                self.capability_badge.configure(bg=CARD_SAGE, fg=MUTED)
        else:
            capability = self.current_operation.capability()
            badge_colours = {
                "ready": (SUCCESS_SURFACE, SUCCESS),
                "external": (CREAM_BLUE, ACCENT),
                "unavailable": (DANGER_SURFACE, DANGER),
            }[capability.status]
            self.capability_badge.configure(
                bg=badge_colours[0],
                fg=badge_colours[1],
            )

        self._background_last_signature = None
        self._header_art_last_signature = None
        self._empty_drop_last_signature = None
        self._redraw_background_layer()
        self._render_header_art()
        if not self.input_paths and hasattr(self, "empty_drop_canvas"):
            self._schedule_empty_drop_redraw(delay=0, force=True)
        self._render_progress_track()
        for key in ("original", "result"):
            self._schedule_image_preview_redraw(key, delay=0, force=True)
        self._refresh_category_buttons()
        self._on_workspace_tab_changed()
        self.after_idle(
            lambda: self._position_indicator(
                self.category_indicator,
                self.category_buttons[self.catalog_root_var.get()],
                self.category_bar,
            )
        )
        selected_tab_button = (
            self.setup_tab_button
            if self.notebook.select() == str(self.setup_page)
            else self.log_tab_button
        )
        self.after_idle(
            lambda button=selected_tab_button: self._position_indicator(
                self.workspace_indicator,
                button,
                self.workspace_tabs,
            )
        )

    def _show_office_compatibility_info(self) -> None:
        messagebox.showinfo(
            "Office 兼容与引擎说明",
            (
                f"{OFFICE_COMPATIBILITY_NOTICE}\n\n"
                "WPS Office、Microsoft Office 与 LibreOffice 均不随软件分发，"
                "页织工坊只调用用户本机已经合法安装的程序。"
            ),
            parent=self,
        )

    def _select_catalog_root(self, root_name: str) -> None:
        if root_name not in {"文档工具", "图片工具", "视频工具"}:
            return
        if self.catalog_root_var.get() == root_name:
            return
        self.catalog_root_var.set(root_name)
        self._refresh_category_buttons(position_indicator=False)
        self._rebuild_operation_tree()
        self._pulse_catalog_root(root_name)

    def _pulse_catalog_root(self, root_name: str) -> None:
        if self._catalog_transition_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._catalog_transition_job)
            self._catalog_transition_job = None
        button = self.category_buttons.get(root_name)
        if button is None:
            return
        timing = motion_effect_timing(
            self._motion_mode(),
            "transition",
            busy=bool(self.worker and self.worker.is_alive()),
            minimized=self.state() in {"iconic", "withdrawn"},
        )
        if not timing.enabled:
            self._position_indicator(self.category_indicator, button, self.category_bar)
            return
        self._animate_indicator(
            self.category_indicator,
            button,
            self.category_bar,
            job_attribute="_catalog_transition_job",
        )

    def _refresh_category_buttons(self, *, position_indicator: bool = True) -> None:
        selected = self.catalog_root_var.get()
        for root_name, button in self.category_buttons.items():
            button.configure(
                style=(
                    "CategoryActive.TButton"
                    if root_name == selected
                    else "Category.TButton"
                )
            )
        selected_button = self.category_buttons.get(selected)
        if selected_button is not None and position_indicator:
            self.after_idle(
                lambda: self._position_indicator(
                    self.category_indicator,
                    selected_button,
                    self.category_bar,
                )
            )

    def _toggle_particle_effects(self) -> None:
        self._set_particle_effects(not self.particle_effects_var.get())

    def _apply_particle_effects_preference(self) -> None:
        self._set_particle_effects(self.particle_effects_var.get(), announce=True)

    def _set_particle_effects(self, enabled: object, *, announce: bool = False) -> None:
        normalized = normalize_particle_effects_enabled(enabled)
        changed = normalized != bool(self.particle_effects_var.get())
        self.particle_effects_var.set(normalized)
        self._refresh_particle_effect_button()
        if not normalized:
            self._dispose_particle_resources()
        self._ui_preferences["particle_effects"] = normalized
        try:
            save_ui_preferences(self._ui_preferences, self._preferences_path)
        except OSError as exc:
            self._log(f"粒子动效已切换，但偏好保存失败：{exc}")
        else:
            if changed or announce:
                self._log(f"粒子动效已{'开启' if normalized else '关闭'}。")

    def _refresh_particle_effect_button(self) -> None:
        if hasattr(self, "particle_effect_button"):
            enabled = bool(self.particle_effects_var.get())
            self.particle_effect_button.configure(
                text=particle_effect_button_text(
                    enabled,
                    compact=self._layout_mode == "narrow",
                ),
                style="ParticleOn.TButton" if enabled else "Quiet.TButton",
            )

    def _on_button_enter(self, event: tk.Event) -> None:
        with contextlib.suppress(tk.TclError):
            event.widget.configure(cursor="hand2")

    def _on_button_release(self, event: tk.Event) -> None:
        widget = event.widget
        with contextlib.suppress(tk.TclError):
            if "disabled" in widget.state():
                return
        busy = bool(self.worker and self.worker.is_alive())
        timing = motion_effect_timing(
            self._motion_mode(),
            "click",
            busy=busy,
            minimized=self.state() in {"iconic", "withdrawn"},
        )
        if not timing.enabled:
            return
        if self.particle_effects_var.get() and not busy:
            click_x = getattr(event, "x", None)
            click_y = getattr(event, "y", None)
            self.after_idle(
                lambda target=widget, x=click_x, y=click_y: self._spawn_click_particles(
                    target,
                    x,
                    y,
                )
            )
        key = str(widget)
        existing = self._button_motion_jobs.pop(key, None)
        if existing is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(existing)
        with contextlib.suppress(tk.TclError):
            widget.state(["pressed", "alternate"])

        def rebound() -> None:
            self._button_motion_jobs.pop(key, None)
            with contextlib.suppress(tk.TclError):
                widget.state(["!pressed", "!alternate"])

        self._button_motion_jobs[key] = self.after(
            min(180, timing.frames * timing.step_ms),
            rebound,
        )

    @staticmethod
    def _set_click_through_window(window: tk.Toplevel) -> bool:
        """Make a tiny Windows particle layer ignore all mouse interaction."""

        if os.name != "nt":
            return False
        try:
            window.update_idletasks()
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            child = int(window.winfo_id())
            get_ancestor = user32.GetAncestor
            get_ancestor.argtypes = (ctypes.c_void_p, ctypes.c_uint)
            get_ancestor.restype = ctypes.c_void_p
            root_handle = get_ancestor(ctypes.c_void_p(child), 2)
            if not root_handle:
                return False
            hwnd = int(root_handle)
            pointer_type = ctypes.c_ssize_t
            get_long = getattr(user32, "GetWindowLongPtrW", user32.GetWindowLongW)
            set_long = getattr(user32, "SetWindowLongPtrW", user32.SetWindowLongW)
            get_long.argtypes = (ctypes.c_void_p, ctypes.c_int)
            set_long.argtypes = (ctypes.c_void_p, ctypes.c_int, pointer_type)
            get_long.restype = pointer_type
            set_long.restype = pointer_type
            ctypes.set_last_error(0)
            ex_style = get_long(hwnd, -20)
            if ex_style == 0 and ctypes.get_last_error() != 0:
                return False
            # WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
            ctypes.set_last_error(0)
            previous = set_long(
                hwnd,
                -20,
                ex_style | 0x80000 | 0x20 | 0x80 | 0x08000000,
            )
            if previous == 0 and ctypes.get_last_error() != 0:
                return False
            set_window_pos = user32.SetWindowPos
            set_window_pos.argtypes = (
                ctypes.c_void_p,
                ctypes.c_void_p,
                ctypes.c_int,
                ctypes.c_int,
                ctypes.c_int,
                ctypes.c_int,
                ctypes.c_uint,
            )
            set_window_pos.restype = ctypes.c_int
            positioned = set_window_pos(
                ctypes.c_void_p(hwnd),
                ctypes.c_void_p(-1),
                0,
                0,
                0,
                0,
                0x2 | 0x1 | 0x10 | 0x20,
            )
            return bool(positioned)
        except (AttributeError, OSError, tk.TclError, ValueError):
            return False

    @staticmethod
    def _draw_particle_symbol(
        canvas: tk.Canvas,
        kind: str,
        x: float,
        y: float,
        size: float,
        colour: str,
    ) -> None:
        line_width = max(1, round(size * 0.18))
        if kind == "moon":
            canvas.create_arc(
                x - size,
                y - size,
                x + size,
                y + size,
                start=70,
                extent=220,
                style="arc",
                outline=colour,
                width=line_width,
            )
            canvas.create_arc(
                x - size * 0.45,
                y - size * 0.92,
                x + size * 1.10,
                y + size * 0.92,
                start=112,
                extent=146,
                style="arc",
                outline=colour,
                width=line_width,
            )
            return
        if kind == "star":
            points: list[float] = []
            for index in range(10):
                angle = -math.pi / 2 + index * math.pi / 5
                radius = size if index % 2 == 0 else size * 0.43
                points.extend((x + math.cos(angle) * radius, y + math.sin(angle) * radius))
            canvas.create_polygon(points, fill=colour, outline="")
            return
        if kind == "sparkle":
            canvas.create_line(
                x,
                y - size,
                x,
                y + size,
                fill=colour,
                width=line_width,
                capstyle="round",
            )
            canvas.create_line(
                x - size,
                y,
                x + size,
                y,
                fill=colour,
                width=line_width,
                capstyle="round",
            )
            diagonal = size * 0.55
            canvas.create_line(
                x - diagonal,
                y - diagonal,
                x + diagonal,
                y + diagonal,
                fill=colour,
                width=max(1, line_width - 1),
                capstyle="round",
            )
            canvas.create_line(
                x - diagonal,
                y + diagonal,
                x + diagonal,
                y - diagonal,
                fill=colour,
                width=max(1, line_width - 1),
                capstyle="round",
            )
            return
        if kind == "diamond":
            canvas.create_polygon(
                x,
                y - size,
                x + size * 0.72,
                y,
                x,
                y + size,
                x - size * 0.72,
                y,
                fill=colour,
                outline="",
            )
            return
        if kind == "pulse":
            # Four short detached rays read as a light click response rather
            # than a loading indicator.  Keeping a gap around the centre also
            # prevents the pulse from swallowing the smaller starlet.
            ray_width = max(1, round(size * 0.11))
            inner_radius = size * 0.66
            for angle in (
                -math.pi / 4,
                math.pi / 4,
                3 * math.pi / 4,
                5 * math.pi / 4,
            ):
                canvas.create_line(
                    x + math.cos(angle) * inner_radius,
                    y + math.sin(angle) * inner_radius,
                    x + math.cos(angle) * size,
                    y + math.sin(angle) * size,
                    fill=colour,
                    width=ray_width,
                    capstyle="round",
                )
            return
        if kind == "ring":
            # A broken, slightly flattened orbit is visually lighter than a
            # closed circle and no longer resembles a progress/loading ring.
            orbit_height = size * 0.72
            orbit_width = max(1, round(size * 0.12))
            canvas.create_arc(
                x - size,
                y - orbit_height,
                x + size,
                y + orbit_height,
                start=18,
                extent=118,
                style="arc",
                outline=colour,
                width=orbit_width,
            )
            canvas.create_arc(
                x - size,
                y - orbit_height,
                x + size,
                y + orbit_height,
                start=202,
                extent=104,
                style="arc",
                outline=colour,
                width=orbit_width,
            )
            dot_radius = max(0.8, size * 0.13)
            dot_x = x + math.cos(math.radians(18)) * size
            dot_y = y - math.sin(math.radians(18)) * orbit_height
            canvas.create_oval(
                dot_x - dot_radius,
                dot_y - dot_radius,
                dot_x + dot_radius,
                dot_y + dot_radius,
                fill=colour,
                outline="",
            )
            return
        if kind == "comet":
            canvas.create_line(
                x - size * 1.35,
                y + size * 0.70,
                x + size * 0.15,
                y - size * 0.12,
                fill=colour,
                width=line_width,
                capstyle="round",
            )
            canvas.create_oval(
                x - size * 0.18,
                y - size * 0.46,
                x + size * 0.55,
                y + size * 0.27,
                fill=colour,
                outline="",
            )
            return
        if kind == "hex":
            points: list[float] = []
            for index in range(6):
                angle = -math.pi / 2 + index * math.pi / 3
                points.extend(
                    (x + math.cos(angle) * size, y + math.sin(angle) * size)
                )
            canvas.create_polygon(
                points,
                fill="",
                outline=colour,
                width=line_width,
            )
            return
        canvas.create_oval(
            x - size * 0.34,
            y - size * 0.34,
            x + size * 0.34,
            y + size * 0.34,
            fill=colour,
            outline="",
        )

    def _spawn_click_particles(
        self,
        widget: tk.Misc,
        click_x: float | None = None,
        click_y: float | None = None,
    ) -> None:
        """Burst varied, DPI-aware particles just outside the clicked button."""

        if (
            self._closing
            or self._window_resizing
            or not self._window_mapped
            or not self.particle_effects_var.get()
            or self.state() in {"iconic", "withdrawn"}
        ):
            return
        now = time.monotonic()
        if now - self._last_particle_spawn_at < 0.075:
            return
        self._last_particle_spawn_at = now
        while len(self._particle_windows) >= 2:
            oldest = self._particle_windows.pop(0)
            for job in tuple(getattr(oldest, "_particle_job_ids", ())):
                self._particle_jobs.discard(job)
                with contextlib.suppress(tk.TclError):
                    self.after_cancel(job)
            with contextlib.suppress(tk.TclError):
                oldest.destroy()
        virtual_bounds: tuple[int, int, int, int] | None = None
        try:
            widget_width = max(1, widget.winfo_width())
            widget_height = max(1, widget.winfo_height())
            particle_scale = min(1.6, max(1.0, float(self._display_scale)))
            safe_padding = round(76 * particle_scale)
            width = max(96, widget_width + safe_padding * 2)
            height = max(92, widget_height + safe_padding * 2)
            horizontal_padding = (width - widget_width) / 2
            vertical_padding = (height - widget_height) / 2
            widget_root_x = widget.winfo_rootx()
            widget_root_y = widget.winfo_rooty()
            left = round(widget_root_x - horizontal_padding)
            top = round(widget_root_y - vertical_padding)
            if os.name == "nt":
                user32 = ctypes.windll.user32
                virtual_left = int(user32.GetSystemMetrics(76))
                virtual_top = int(user32.GetSystemMetrics(77))
                virtual_width = int(user32.GetSystemMetrics(78))
                virtual_height = int(user32.GetSystemMetrics(79))
                if virtual_width > 0 and virtual_height > 0:
                    virtual_bounds = (
                        virtual_left,
                        virtual_top,
                        virtual_left + virtual_width,
                        virtual_top + virtual_height,
                    )
                    left = min(
                        max(left, virtual_left),
                        virtual_left + virtual_width - width,
                    )
                    top = min(
                        max(top, virtual_top),
                        virtual_top + virtual_height - height,
                    )
        except (AttributeError, OSError, tk.TclError, TypeError, ValueError):
            return
        overlay = tk.Toplevel(self)
        overlay.withdraw()
        overlay.overrideredirect(True)
        overlay.configure(bg="#010203")
        alpha_supported = False
        with contextlib.suppress(tk.TclError):
            overlay.attributes("-topmost", True)
        try:
            overlay.attributes("-transparentcolor", "#010203")
        except tk.TclError:
            overlay.destroy()
            return
        try:
            overlay.attributes("-alpha", 0.0)
            alpha_supported = True
        except tk.TclError:
            alpha_supported = False
        left_offset = f"+{left}" if left >= 0 else str(left)
        top_offset = f"+{top}" if top >= 0 else str(top)
        overlay.geometry(f"{width}x{height}{left_offset}{top_offset}")
        canvas = tk.Canvas(
            overlay,
            width=width,
            height=height,
            bg="#010203",
            highlightthickness=0,
            borderwidth=0,
            takefocus=0,
        )
        canvas.pack(fill="both", expand=True)
        if not self._set_click_through_window(overlay):
            overlay.destroy()
            return
        overlay.deiconify()
        self._particle_windows.append(overlay)
        local_x = min(
            widget_width,
            max(0.0, widget_width / 2 if click_x is None else float(click_x)),
        )
        local_y = min(
            widget_height,
            max(0.0, widget_height / 2 if click_y is None else float(click_y)),
        )
        distances = {
            "left": local_x,
            "right": widget_width - local_x,
            "top": local_y,
            "bottom": widget_height - local_y,
        }
        if virtual_bounds is None:
            edge = min(distances, key=distances.get)
        else:
            virtual_left, virtual_top, virtual_right, virtual_bottom = virtual_bounds
            available_space = {
                "left": widget_root_x - virtual_left,
                "right": virtual_right - (widget_root_x + widget_width),
                "top": widget_root_y - virtual_top,
                "bottom": virtual_bottom - (widget_root_y + widget_height),
            }
            desired_space = 64.0 * particle_scale
            edge = min(
                distances,
                key=lambda name: distances[name]
                + max(0.0, desired_space - available_space[name]) * 2.5,
            )
        normal_x, normal_y = {
            "left": (-1.0, 0.0),
            "right": (1.0, 0.0),
            "top": (0.0, -1.0),
            "bottom": (0.0, 1.0),
        }[edge]
        tangent_x, tangent_y = -normal_y, normal_x
        click_overlay_x = widget_root_x + local_x - left
        click_overlay_y = widget_root_y + local_y - top
        edge_distance = distances[edge]
        origin_x = click_overlay_x + normal_x * (
            edge_distance + 4.0 * particle_scale
        )
        origin_y = click_overlay_y + normal_y * (
            edge_distance + 4.0 * particle_scale
        )
        particles = click_particle_specs(self._particle_variant)
        self._particle_variant = (self._particle_variant + 1) % 3
        frame_plan = click_particle_frame_plan()
        frame_count = len(frame_plan)
        step_ms = 27

        def dispose_overlay() -> None:
            if getattr(overlay, "_particle_disposed", False):
                return
            setattr(overlay, "_particle_disposed", True)
            for job in tuple(getattr(overlay, "_particle_job_ids", ())):
                self._particle_jobs.discard(job)
                with contextlib.suppress(tk.TclError):
                    self.after_cancel(job)
            setattr(overlay, "_particle_job_ids", set())
            with contextlib.suppress(ValueError):
                self._particle_windows.remove(overlay)
            with contextlib.suppress(tk.TclError):
                overlay.destroy()

        def frame(index: int) -> None:
            try:
                exists = bool(overlay.winfo_exists())
            except tk.TclError:
                exists = False
            if not exists:
                dispose_overlay()
                return
            state = frame_plan[index]
            if index + 1 >= frame_count:
                if alpha_supported:
                    with contextlib.suppress(tk.TclError):
                        overlay.attributes("-alpha", 0.0)
                with contextlib.suppress(tk.TclError):
                    canvas.delete("all")
                dispose_overlay()
                return
            try:
                if alpha_supported:
                    overlay.attributes(
                        "-alpha",
                        min(1.0, max(0.04, state.opacity)),
                    )
                canvas.delete("all")
                if index <= 5:
                    flash_progress = index / 5
                    flash_size = particle_scale * (
                        3.6 + 6.4 * ease_out_cubic(flash_progress)
                    )
                    self._draw_particle_symbol(
                        canvas,
                        "pulse",
                        origin_x,
                        origin_y,
                        flash_size,
                        PARTICLE_STAR,
                    )
                    if index <= 3:
                        self._draw_particle_symbol(
                            canvas,
                            "sparkle",
                            origin_x,
                            origin_y,
                            particle_scale * (3.8 + index * 0.7),
                            PARTICLE_STAR,
                        )
                for particle_index, particle in enumerate(particles):
                    outward = particle_scale * (
                        5.0
                        + state.spread * (46.0 + abs(particle.drift) * 11.0)
                    )
                    side = particle.tangent * particle_scale * (
                        7.0 + state.spread * 24.0
                    )
                    curve = (
                        (1.0 if particle_index % 2 == 0 else -1.0)
                        * state.curve
                        * 4.8
                        * particle_scale
                    )
                    x = (
                        origin_x
                        + normal_x * outward
                        + tangent_x * (side + curve)
                    )
                    y = (
                        origin_y
                        + normal_y * outward
                        + tangent_y * (side + curve)
                    )
                    colour = globals().get(particle.colour_role, CYAN)
                    self._draw_particle_symbol(
                        canvas,
                        particle.kind,
                        x,
                        y,
                        max(
                            1.5 * particle_scale,
                            particle.size * particle_scale * state.scale,
                        ),
                        str(colour),
                    )
            except tk.TclError:
                dispose_overlay()
                return
            if index + 1 < frame_count:
                job_holder: dict[str, str] = {}

                def continue_frame() -> None:
                    job = job_holder.get("id")
                    if job is not None:
                        self._particle_jobs.discard(job)
                        jobs = getattr(overlay, "_particle_job_ids", None)
                        if jobs is not None:
                            jobs.discard(job)
                    frame(index + 1)

                job = self.after(step_ms, continue_frame)
                job_holder["id"] = job
                self._particle_jobs.add(job)
                jobs = getattr(overlay, "_particle_job_ids", None)
                if jobs is None:
                    jobs = set()
                    setattr(overlay, "_particle_job_ids", jobs)
                jobs.add(job)
            else:
                dispose_overlay()

        frame(0)

    def _dispose_particle_resources(self) -> None:
        """Immediately stop every particle overlay without touching other UI motion."""

        for job in tuple(self._particle_jobs):
            with contextlib.suppress(tk.TclError):
                self.after_cancel(job)
        self._particle_jobs.clear()
        for window in tuple(self._particle_windows):
            for job in tuple(getattr(window, "_particle_job_ids", ())):
                with contextlib.suppress(tk.TclError):
                    self.after_cancel(job)
            setattr(window, "_particle_job_ids", set())
            setattr(window, "_particle_disposed", True)
            with contextlib.suppress(tk.TclError):
                window.destroy()
        self._particle_windows.clear()

    def _dispose_motion_resources(self) -> None:
        for job in (
            self._progress_shimmer_job,
            self._progress_track_resize_job,
            self._background_resize_job,
            self._empty_drop_redraw_job,
            self._progress_animation_job,
            self._title_animation_job,
            self._setup_scroll_refresh_job,
            self._drop_hint_reset_job,
            self._image_preview_job,
            self._workspace_transition_job,
            self._content_transition_job,
            self._catalog_transition_job,
            self._indicator_realign_job,
            self._window_layout_job,
            self._window_resize_finish_job,
            self._window_restore_job,
            self._window_restore_finalize_job,
            self._setup_canvas_resize_job,
            *self._button_motion_jobs.values(),
            *self._preview_resize_jobs.values(),
        ):
            if job is not None:
                with contextlib.suppress(tk.TclError):
                    self.after_cancel(job)
        self._progress_shimmer_job = None
        self._progress_track_resize_job = None
        self._background_resize_job = None
        self._empty_drop_redraw_job = None
        self._progress_animation_job = None
        self._progress_animation_last_at = None
        self._title_animation_job = None
        self._setup_scroll_refresh_job = None
        self._drop_hint_reset_job = None
        self._image_preview_job = None
        self._workspace_transition_job = None
        self._content_transition_job = None
        self._content_transition_generation += 1
        self._catalog_transition_job = None
        self._indicator_realign_job = None
        self._window_layout_job = None
        self._window_resize_finish_job = None
        self._window_restore_job = None
        self._window_restore_finalize_job = None
        self._setup_canvas_resize_job = None
        self._window_resizing = False
        self._window_mapped = False
        self._window_restoring = False
        self._window_restore_attempts = 0
        self._pending_window_width = None
        self._pending_window_height = None
        self._pending_setup_canvas_width = None
        self._button_motion_jobs.clear()
        self._preview_resize_jobs.clear()
        if self._content_transition_overlay is not None:
            with contextlib.suppress(tk.TclError):
                self._content_transition_overlay.place_forget()
        self._dispose_particle_resources()

    def _on_window_configure(self, event: tk.Event) -> None:
        if event.widget is not self or self._closing:
            return
        if not self._window_mapped:
            return
        try:
            if self.state() in {"iconic", "withdrawn"}:
                return
        except tk.TclError:
            return
        physical_size = (max(1, int(event.width)), max(1, int(event.height)))
        previous_size = self._last_window_configure_size
        self._last_window_configure_size = physical_size
        self._pending_window_width = self._logical_window_width(physical_size[0])
        self._pending_window_height = self._logical_window_height(physical_size[1])
        # Map is a separate transaction.  Configure events emitted while the
        # native child tree is being restored may update the final size, but
        # must never cancel the already queued two-stage restoration.
        if self._window_restoring:
            return
        # Tk sends Configure for a pure move as well.  Repainting the complete
        # child HWND tree for those events caused needless flashes.
        if previous_size is None or physical_size == previous_size:
            return
        self._window_resizing = True
        for attribute in ("_window_restore_job", "_window_restore_finalize_job"):
            job = getattr(self, attribute)
            if job is not None:
                with contextlib.suppress(tk.TclError):
                    self.after_cancel(job)
                setattr(self, attribute, None)
        if self._window_resize_finish_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._window_resize_finish_job)
        self._window_resize_finish_job = self.after(158, self._finish_window_resize)

    def _on_window_map(self, event: tk.Event) -> None:
        if event.widget is not self or self._closing:
            return
        self._window_mapped = True
        self._window_restoring = True
        self._window_restore_attempts = 0
        self._window_resizing = True
        for attribute in ("_window_restore_job", "_window_restore_finalize_job"):
            job = getattr(self, attribute)
            if job is not None:
                with contextlib.suppress(tk.TclError):
                    self.after_cancel(job)
                setattr(self, attribute, None)
        # Keep the last complete frame visible.  RDW_UPDATENOW must not run in
        # the Map callback: Windows is still mapping child HWNDs and Tk has not
        # repainted their Canvas/Frame surfaces yet.
        try:
            self._window_restore_job = self.after_idle(self._restore_window_after_map)
        except tk.TclError:
            self._window_restore_job = None
            self._window_restoring = False
            self._window_resizing = False

    def _on_window_unmap(self, event: tk.Event) -> None:
        if event.widget is not self:
            return
        self._window_mapped = False
        self._window_restoring = False
        self._window_restore_attempts = 0
        self._window_resizing = False
        for attribute in (
            "_window_layout_job",
            "_window_resize_finish_job",
            "_window_restore_job",
            "_window_restore_finalize_job",
            "_setup_canvas_resize_job",
            "_background_resize_job",
            "_empty_drop_redraw_job",
            "_progress_track_resize_job",
            "_setup_scroll_refresh_job",
        ):
            job = getattr(self, attribute, None)
            if job is not None:
                with contextlib.suppress(tk.TclError):
                    self.after_cancel(job)
                setattr(self, attribute, None)
        for job in tuple(self._preview_resize_jobs.values()):
            with contextlib.suppress(tk.TclError):
                self.after_cancel(job)
        self._preview_resize_jobs.clear()
        if self._content_transition_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._content_transition_job)
            self._content_transition_job = None
        self._content_transition_generation += 1
        if self._content_transition_overlay is not None:
            with contextlib.suppress(tk.TclError):
                self._content_transition_overlay.place_forget()
        self._dispose_particle_resources()

    def _window_is_renderable(self) -> bool:
        if self._closing or not self._window_mapped:
            return False
        try:
            return bool(self.winfo_exists()) and self.state() not in {
                "iconic",
                "withdrawn",
            }
        except tk.TclError:
            return False

    def _restore_window_after_map(self) -> None:
        self._window_restore_job = None
        if not self._window_is_renderable():
            if self._window_mapped and self._window_restore_attempts < 3:
                self._window_restore_attempts += 1
                try:
                    self._window_restore_job = self.after(
                        26,
                        self._restore_window_after_map,
                    )
                    return
                except tk.TclError:
                    self._window_restore_job = None
            self._window_restoring = False
            self._window_resizing = False
            return
        physical_size = (max(1, self.winfo_width()), max(1, self.winfo_height()))
        self._last_window_configure_size = physical_size
        self._pending_window_width = self._logical_window_width(physical_size[0])
        self._pending_window_height = self._logical_window_height(physical_size[1])
        self._window_resizing = True
        self._flush_window_layout(force=True)
        self._flush_setup_canvas_width(force=True)
        for widget in self.winfo_children():
            self._redraw_rounded_descendants(widget)
        try:
            self._window_restore_finalize_job = self.after(
                42,
                self._finalize_window_restore,
            )
        except tk.TclError:
            self._window_restore_finalize_job = None
            self._window_restoring = False
            self._window_resizing = False

    def _finalize_window_restore(self) -> None:
        self._window_restore_finalize_job = None
        if not self._window_is_renderable():
            if self._window_mapped and self._window_restore_attempts < 3:
                self._window_restore_attempts += 1
                try:
                    self._window_restore_finalize_job = self.after(
                        28,
                        self._finalize_window_restore,
                    )
                    return
                except tk.TclError:
                    self._window_restore_finalize_job = None
            self._window_restoring = False
            self._window_resizing = False
            return
        physical_size = (max(1, self.winfo_width()), max(1, self.winfo_height()))
        self._last_window_configure_size = physical_size
        self._pending_window_width = self._logical_window_width(physical_size[0])
        self._pending_window_height = self._logical_window_height(physical_size[1])
        self._flush_window_layout(force=True)
        self._flush_setup_canvas_width(force=True)
        self._window_restoring = False
        self._window_resizing = False
        self._window_restore_attempts = 0
        for widget in self.winfo_children():
            self._redraw_rounded_descendants(widget)
        self._schedule_background_redraw(delay=0, force=True)
        self._schedule_empty_drop_redraw(delay=0, force=True)
        self._schedule_progress_track_redraw(delay=0)
        for key in ("original", "result"):
            self._schedule_image_preview_redraw(key, delay=28, force=True)
        _force_windows_window_redraw(self)

    def _finish_window_resize(self) -> None:
        """Run expensive decoration painting once after live resizing settles."""

        self._window_resize_finish_job = None
        if not self._window_is_renderable():
            self._window_resizing = False
            return
        self._window_restoring = False
        if self._window_layout_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._window_layout_job)
            self._window_layout_job = None
        self._flush_window_layout(force=True)
        self._flush_setup_canvas_width(force=True)
        for widget in self.winfo_children():
            self._redraw_rounded_descendants(widget)
        if self._window_restore_finalize_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._window_restore_finalize_job)
        try:
            self._window_restore_finalize_job = self.after(
                36,
                self._finalize_window_restore,
            )
        except tk.TclError:
            self._window_restore_finalize_job = None
            self._window_restoring = False
            self._window_resizing = False

    def _redraw_rounded_descendants(self, widget: tk.Misc) -> None:
        if isinstance(widget, RoundedCard):
            widget._redraw()
        with contextlib.suppress(tk.TclError):
            for child in widget.winfo_children():
                self._redraw_rounded_descendants(child)

    def _flush_window_layout(self, *, force: bool = False) -> None:
        self._window_layout_job = None
        if (
            self._pending_window_width is None
            or self._pending_window_height is None
            or self._closing
            or (not force and not self._window_mapped)
        ):
            return
        width = self._pending_window_width
        height = self._pending_window_height
        self._pending_window_width = None
        self._pending_window_height = None
        self._apply_responsive_layout(width, window_height=height)

    def _logical_window_width(self, physical_width: int) -> int:
        width = max(1, int(physical_width))
        if self._display_scale <= 1.01:
            return width
        # Per-monitor DPI awareness reports physical Configure pixels on Windows.
        # Keep the historical conversion there, while tests/headless Tk retain
        # the direct logical width at ordinary scaling.
        return max(1, round(width / self._display_scale))

    def _logical_window_height(self, physical_height: int) -> int:
        height = max(1, int(physical_height))
        if self._display_scale <= 1.01:
            return height
        return max(1, round(height / self._display_scale))

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
                68 + normal_font.measure(
                    f"●  {operation_display_name(operation.name)}"
                ),
            )
        for root_name in roots:
            content_width = max(content_width, 26 + root_font.measure(root_name))
        for section_name in sections:
            content_width = max(
                content_width,
                47 + section_font.measure(section_name),
            )
        self._catalog_tree_content_width = min(620, content_width + 18)
        self._catalog_preferred_width = min(
            330,
            max(292, self._catalog_tree_content_width + 20),
        )

    def _current_sidebar_width(self, window_width: int) -> int:
        preferred_width = min(
            self._catalog_preferred_width,
            270 if self._layout_mode == "compact" else 330,
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
            self._schedule_indicator_realign()

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
            window_height=self._logical_window_height(self.winfo_height()),
            force=True,
        )

    def _refresh_office_engine_button(self) -> None:
        if not hasattr(self, "engine_select_button"):
            return
        spec = office_engine_parameter_spec(self.current_operation)
        displayed_value = str(self.office_engine_preference.get() or "auto")
        if spec is not None and spec.key in self.param_vars:
            actual = self._parameter_actual_value(spec.key)
            if actual in OFFICE_ENGINE_MENU_LABELS or actual == "none":
                displayed_value = actual
        if displayed_value == "none":
            text = (
                "引擎：仅结构校验  ▾"
                if self._layout_mode == "narrow"
                else "当前引擎：仅结构校验  ▾"
            )
        else:
            text = office_engine_button_text(
                displayed_value,
                compact=self._layout_mode == "narrow",
                active=spec is not None,
            )
        self.engine_select_button.configure(text=text)

    def _show_office_engine_menu(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        self._popup_menu_below(self.office_engine_menu, self.engine_select_button)

    def _select_office_engine(self, value: str) -> None:
        if value not in OFFICE_ENGINE_MENU_LABELS:
            return
        if self.worker and self.worker.is_alive():
            return
        self.office_engine_preference.set(value)
        applied = self._apply_office_engine_preference_to_current()
        self._refresh_office_engine_button()
        label = OFFICE_ENGINE_MENU_LABELS[value]
        if applied:
            self._log(f"Office 引擎已设为：{label}；已同步当前功能参数。")
        else:
            self._log(
                f"Office 默认引擎已设为：{label}；当前功能不调用 Office 引擎。"
            )

    def _apply_office_engine_preference_to_current(self) -> bool:
        spec = office_engine_parameter_spec(self.current_operation)
        if spec is None or spec.key not in self.param_vars:
            return False
        value = str(self.office_engine_preference.get() or "auto")
        allowed = {str(choice) for choice, _label in spec.choices}
        if value not in allowed:
            return False
        self._syncing_office_engine = True
        try:
            self._set_choice_parameter_value(spec.key, value)
        finally:
            self._syncing_office_engine = False
        return True

    def _apply_responsive_layout(
        self,
        window_width: int,
        *,
        window_height: int | None = None,
        force: bool = False,
    ) -> None:
        # ``Configure`` widths may be physical pixels on per-monitor-DPI
        # Windows builds.  Never let the responsive decision exceed the
        # current screen's logical width, which also keeps drag-resize tests
        # and small displays honest.
        try:
            logical_screen_width = round(
                self.winfo_screenwidth() / max(1.0, self._display_scale)
            )
            logical_screen_height = round(
                self.winfo_screenheight() / max(1.0, self._display_scale)
            )
        except (tk.TclError, TypeError, ValueError):
            logical_screen_width = int(window_width)
            logical_screen_height = int(window_height or self.winfo_height())
        window_width = min(max(1, int(window_width)), max(1, logical_screen_width))
        if window_height is None:
            window_height = self._logical_window_height(self.winfo_height())
        window_height = min(
            max(1, int(window_height)), max(1, logical_screen_height)
        )
        mode = responsive_layout_mode(window_width)
        short = short_window_layout(window_height)
        narrow = mode == "narrow"
        compact = mode == "compact"
        stacked = mode != "wide"
        mode_changed = mode != self._layout_mode
        height_mode_changed = short != self._layout_short
        if mode_changed:
            self._sidebar_expanded = not narrow
        self._layout_mode = mode
        self._layout_short = short
        if mode_changed:
            self._refresh_operation_description()

        subtitle_width = max(220, int(window_width) - (70 if narrow else 440))
        self._set_label_wraplength(self.header_subtitle, subtitle_width)
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

        if narrow and self._sidebar_expanded and hasattr(self, "sidebar_shell"):
            sidebar_height = min(
                320 if short else 360,
                max(250 if short else 260, int(window_height * (0.56 if short else 0.45))),
            )
            if abs(sidebar_height - self._last_narrow_sidebar_height) >= 10:
                self.sidebar_shell.configure(height=sidebar_height)
                self._last_narrow_sidebar_height = sidebar_height

        if not mode_changed and not height_mode_changed and not force:
            return

        for column in range(2):
            self.body.columnconfigure(column, weight=0)
        for row in range(2):
            self.body.rowconfigure(row, weight=0)

        if narrow:
            self.header_card.configure(
                height=max(
                    96 if short else 118,
                    round((100 if short else 122) * self._display_scale),
                )
            )
            self.header_card.pack_configure(
                padx=8 if short else 9,
                pady=(6, 4) if short else (9, 7),
            )
            self.header_subtitle.grid_remove()
            self.title_box.grid_configure(row=0, column=0, sticky="ew")
            self.sidebar_toggle_button.grid(
                row=0,
                column=1,
                sticky="e",
                padx=(6, 5),
                pady=(0, 0),
            )
            self.header_actions.grid_configure(
                row=0,
                column=2,
                sticky="e",
                padx=(8, 0),
                pady=(0, 0),
            )
            self.engine_select_button.pack_forget()
            self.particle_effect_button.pack_forget()
            self.preferences_button.pack_forget()
            if not short:
                self.particle_effect_button.pack(side="left", padx=(0, 7))
            self.preferences_button.pack(side="left")
            self.header_art_canvas.grid_remove()
            self.body.configure(padding=(8, 1, 8, 7) if short else (9, 2, 9, 9))
            self.body.columnconfigure(0, weight=1)

            if self._sidebar_expanded:
                sidebar_height = min(
                    320 if short else 360,
                    max(
                        250 if short else 260,
                        int(window_height * (0.56 if short else 0.45)),
                    ),
                )
                self._last_narrow_sidebar_height = sidebar_height
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
            for column in range(2):
                self.file_toolbar.columnconfigure(column, weight=1)
            for widget in (
                *self.file_toolbar_buttons,
                self.file_more_button,
                self.file_count_label,
            ):
                widget.grid_forget()
            for index, button in enumerate(
                (self.add_file_button, self.file_more_button)
            ):
                button.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 6, 0))
            self.file_count_label.grid(
                row=1,
                column=0,
                columnspan=2,
                sticky="w",
                pady=(8, 1),
            )
            self.file_tree.configure(height=3 if short else 5)
            self.empty_drop_canvas.configure(height=126 if short else 176)

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
            self.output_entry.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(5, 0))
            self.output_browse_button.grid(row=2, column=1, padx=(8, 0), pady=(7, 0), sticky="e")
            self.output_open_button.grid(row=2, column=2, padx=(6, 0), pady=(7, 0), sticky="e")

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

            if short:
                self.progress_box.grid_configure(
                    row=0,
                    column=0,
                    columnspan=1,
                    sticky="ew",
                    padx=(0, 8),
                    pady=(0, 0),
                )
                self.cancel_button.grid_configure(
                    row=0,
                    column=1,
                    sticky="",
                    padx=(0, 6),
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
                self.footer.rowconfigure(1, minsize=0)
                self.footer_card.configure(
                    height=max(82, round(88 * self._display_scale))
                )
            else:
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
                    sticky="nsew",
                    padx=(0, 6),
                    pady=(0, 0),
                )
                self.run_button.grid_configure(
                    row=1,
                    column=1,
                    sticky="nsew",
                    padx=(6, 0),
                    pady=(0, 0),
                )
                self.footer.columnconfigure(0, weight=1)
                self.footer.columnconfigure(1, weight=1)
                self.footer.columnconfigure(2, weight=0)
                self.footer.rowconfigure(
                    1,
                    minsize=max(44, round(46 * self._display_scale)),
                )
                self.footer_card.configure(
                    height=max(116, round(120 * self._display_scale))
                )
        else:
            self.sidebar_toggle_button.grid_remove()
            self.footer.rowconfigure(1, minsize=0)
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
                self.header_card.configure(
                    height=max(
                        76 if short else 82,
                        round((80 if short else 86) * self._display_scale),
                    )
                )
                self.header_card.pack_configure(
                    padx=12 if short else 14,
                    pady=(8, 5) if short else (12, 8),
                )
                self.header_subtitle.grid_remove()
                self.body.configure(
                    padding=(12, 1, 12, 10) if short else (14, 2, 14, 14)
                )
            else:
                self.header_card.configure(
                    height=max(
                        80 if short else 88,
                        round((84 if short else 92) * self._display_scale),
                    )
                )
                self.header_card.pack_configure(
                    padx=16 if short else 20,
                    pady=(9, 6) if short else (16, 10),
                )
                if short:
                    self.header_subtitle.grid_remove()
                else:
                    self.header_subtitle.grid()
                self.body.configure(
                    padding=(16, 2, 16, 12) if short else (20, 4, 20, 20)
                )

            self.title_box.grid_configure(row=0, column=0, sticky="ew")
            self.header_actions.grid_configure(
                row=0,
                column=2,
                sticky="e",
                padx=(16, 0),
                pady=(0, 0),
            )
            for widget in (
                self.engine_select_button,
                self.particle_effect_button,
                self.preferences_button,
            ):
                widget.pack_forget()
            if compact:
                self.header_art_canvas.grid_remove()
                self.particle_effect_button.pack(side="left", padx=(0, 7))
                self.preferences_button.pack(side="left")
            else:
                self.header_art_canvas.grid(
                    row=0,
                    column=1,
                    sticky="e",
                    padx=(12, 4),
                )
                self.header_art_canvas.after_idle(self._render_header_art)
                self.engine_select_button.pack(side="left", padx=(0, 7))
                self.particle_effect_button.pack(side="left", padx=(0, 7))
                self.preferences_button.pack(side="left")
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
            for column in range(3):
                self.file_toolbar.columnconfigure(column, weight=0)
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
                    sticky="ew" if compact else "w",
                    padx=(0 if index == 0 else 6, 0),
                )
            self.file_count_label.grid(
                row=1 if compact else 0,
                column=0 if compact else 3,
                columnspan=3 if compact else 1,
                sticky="w" if compact else "e",
                padx=(0 if compact else 12, 0),
                pady=(7, 0) if compact else (0, 0),
            )
            self.file_tree.configure(
                height=(4 if compact else 5) if short else (6 if compact else 7)
            )
            self.empty_drop_canvas.configure(
                height=(142 if compact else 156) if short else 210
            )

            for column in range(4):
                self.output_frame.columnconfigure(column, weight=0)
            for widget in (
                self.output_label,
                self.output_entry,
                self.output_browse_button,
                self.output_open_button,
            ):
                widget.grid_forget()
            if compact:
                self.output_frame.columnconfigure(0, weight=1)
                self.output_label.grid(row=0, column=0, columnspan=3, sticky="w")
                self.output_entry.grid(row=1, column=0, sticky="ew", pady=(5, 0))
                self.output_browse_button.grid(
                    row=1, column=1, padx=(8, 0), pady=(5, 0)
                )
                self.output_open_button.grid(
                    row=1, column=2, padx=(6, 0), pady=(5, 0)
                )
            else:
                self.output_frame.columnconfigure(1, weight=1)
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
            self.footer_card.configure(
                height=max(
                    68 if short else 72,
                    round((70 if short else 76) * self._display_scale),
                )
            )

        self.setup_tab.configure(
            padding=(4, 8, 8, 3) if short else (4, 14, 8, 4)
        )
        self.operation_description.grid_configure(
            pady=(4, 8) if short else (6, 14)
        )
        self.workspace_tabs.grid_configure(pady=(0, 2 if short else 4))
        self.footer_card.grid_configure(pady=(8, 0) if short else (14, 0))
        self.file_toolbar.grid_configure(pady=(0, 5 if short else 8))
        self.output_card.grid_configure(
            pady=(8, 5) if short else (12, 8)
        )
        self._refresh_office_engine_button()
        self._refresh_particle_effect_button()
        self._layout_parameter_rows(stacked)
        self._layout_image_preview_panes(stacked)
        self._schedule_setup_scroll_refresh()
        self._schedule_indicator_realign()

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
        short = bool(self._layout_short)
        layout = f"{'stacked' if stacked else 'side_by_side'}:{'short' if short else 'regular'}"
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
            height = max(
                150 if short else 180,
                round((165 if short else 205) * self._display_scale),
            )
            card_height = height + max(54, round(58 * self._display_scale))
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
            height = max(
                184 if short else 210,
                round((200 if short else 250) * self._display_scale),
            )
            card_height = height + max(54, round(58 * self._display_scale))
        self.image_preview_original_box.configure(height=card_height)
        self.image_preview_result_box.configure(height=card_height)
        self.image_preview_original_canvas.configure(height=height)
        self.image_preview_result_canvas.configure(height=height)
        self.image_preview_card._schedule_fit_requested_height()
        self._schedule_image_preview_redraw("original", delay=30, force=True)
        self._schedule_image_preview_redraw("result", delay=30, force=True)

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
            self.image_preview_card.grid_remove()
            self._image_preview_generation += 1
            self._clear_image_preview_images()
            self._schedule_setup_scroll_refresh()
            return
        self.image_preview_card.grid()
        self.image_preview_card._schedule_fit_requested_height()
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

    def _on_parameter_value_changed(self, key: str | None = None) -> None:
        if key and not self._syncing_office_engine:
            spec = office_engine_parameter_spec(self.current_operation)
            if spec is not None and spec.key == key:
                actual = self._parameter_actual_value(key)
                if actual in OFFICE_ENGINE_MENU_LABELS:
                    self.office_engine_preference.set(actual)
                    self._refresh_office_engine_button()
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
                self._schedule_image_preview_redraw(key, delay=0, force=True)
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
        self._schedule_image_preview_redraw("original", delay=0, force=True)
        self._schedule_image_preview_redraw("result", delay=0, force=True)

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
        self._schedule_image_preview_redraw("original", delay=0, force=True)
        self._schedule_image_preview_redraw("result", delay=0, force=True)
        self._schedule_setup_scroll_refresh()

    def _clear_image_preview_images(self) -> None:
        for image in (self._image_preview_original, self._image_preview_result):
            if image is not None:
                with contextlib.suppress(Exception):
                    image.close()
        self._image_preview_original = None
        self._image_preview_result = None
        self._image_preview_photos.clear()

    def _schedule_image_preview_redraw(
        self,
        key: str,
        *,
        delay: int = 72,
        force: bool = False,
    ) -> None:
        """Resize expensive PIL previews only after Configure events settle."""

        if self._closing or key not in {"original", "result"}:
            return
        existing = self._preview_resize_jobs.get(key)
        if existing is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(existing)
        if force:
            self._preview_last_sizes.pop(key, None)

        def render() -> None:
            self._preview_resize_jobs.pop(key, None)
            self._render_image_preview_canvas(key)

        with contextlib.suppress(tk.TclError):
            self._preview_resize_jobs[key] = self.after(max(0, delay), render)

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
        direct_selection: tuple[float, ...] | None = None
        if key == "original" and self._direct_image_edit_canvas_start is not None:
            with contextlib.suppress(tk.TclError):
                coords = canvas.coords("direct-image-edit")
                if len(coords) == 4:
                    direct_selection = tuple(float(value) for value in coords)
        try:
            width = max(1, canvas.winfo_width())
            height = max(1, canvas.winfo_height())
            signature = (width, height)
            if signature == self._preview_last_sizes.get(key):
                return
            self._preview_last_sizes[key] = signature
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
            if direct_selection is not None:
                canvas.create_rectangle(
                    *direct_selection,
                    outline="#F0445E",
                    width=max(2, round(2 * self._display_scale)),
                    dash=(6, 3),
                    tags=("direct-image-edit",),
                )
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
        width = max(1, int(event.width))
        self._pending_setup_canvas_width = width
        if self._setup_canvas_resize_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._setup_canvas_resize_job)
            self._setup_canvas_resize_job = None
        if self._window_resizing or not self._window_mapped:
            # Keep the existing embedded Frame intact during a live root-window
            # drag.  Resizing every native child for every mouse pixel is what
            # exposed unpainted black rectangles on Windows.
            return
        if abs(width - self._setup_canvas_width) < 2:
            self._pending_setup_canvas_width = None
            return
        with contextlib.suppress(tk.TclError):
            self._setup_canvas_resize_job = self.after(
                34,
                self._flush_setup_canvas_width,
            )

    def _flush_setup_canvas_width(self, *, force: bool = False) -> None:
        self._setup_canvas_resize_job = None
        if self._closing or not hasattr(self, "setup_canvas"):
            return
        if (self._window_resizing or not self._window_mapped) and not force:
            return
        width = self._pending_setup_canvas_width
        if force:
            with contextlib.suppress(tk.TclError):
                width = max(1, int(self.setup_canvas.winfo_width()))
        if width is None:
            return
        if not force and abs(width - self._setup_canvas_width) < 2:
            self._pending_setup_canvas_width = None
            return
        try:
            self.setup_canvas.itemconfigure(self.setup_window, width=width)
        except tk.TclError:
            return
        self._setup_canvas_width = width
        self._pending_setup_canvas_width = None
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
            self._setup_scroll_refresh_job = self.after(
                42,
                self._refresh_setup_scrollregion,
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
        selected_root = self.catalog_root_var.get() or "文档工具"
        catalog: dict[str, list[Operation]] = {}
        for operation in self.operations:
            capability = operation.capability()
            if not self.show_unavailable_var.get() and not capability.runnable:
                continue
            root_name, section_name = operation_catalog_path(operation)
            if not query and root_name != selected_root:
                continue
            haystack = (
                f"{root_name} {section_name} {operation.group} "
                f"{operation.name} {operation.description}".lower()
            )
            if query and query not in haystack:
                continue
            catalog.setdefault(f"{root_name}\0{section_name}", []).append(operation)
        ordered_sections = sorted(
            catalog.items(),
            key=lambda item: catalog_order_key(*item[0].split("\0", 1)),
        )
        for section_key, operations in ordered_sections:
            root_name, section_name = section_key.split("\0", 1)
            section = self.operation_tree.insert(
                "",
                "end",
                text=(
                    f"{root_name} · {section_name}"
                    if query and root_name != selected_root
                    else section_name
                ),
                open=True,
                tags=("catalog_section",),
            )
            for operation in operations:
                self.operation_tree.insert(
                    section,
                    "end",
                    iid=operation.id,
                    text=operation_display_name(operation.name),
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
        if self.current_operation is operation:
            return
        collapse_sidebar = bool(
            _event is not None
            and self._layout_mode == "narrow"
            and self._sidebar_expanded
        )
        if self.current_operation is None:
            self._apply_operation_selection(operation, collapse_sidebar)
            return
        self._run_content_transition(
            lambda: self._apply_operation_selection(operation, collapse_sidebar)
        )

    def _apply_operation_selection(
        self,
        operation: Operation,
        collapse_sidebar: bool = False,
    ) -> None:
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
            "ready": ("已就绪", SUCCESS_SURFACE, SUCCESS),
            "external": ("已就绪", CREAM_BLUE, ACCENT),
            "unavailable": ("需要安装", DANGER_SURFACE, DANGER),
        }[capability.status]
        self.capability_badge.configure(
            text=badge_config[0], bg=badge_config[1], fg=badge_config[2]
        )
        self._build_parameters(operation.parameters)
        self._refresh_office_engine_button()
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
        if collapse_sidebar and self._layout_mode == "narrow" and self._sidebar_expanded:
            self._sidebar_expanded = False
            self.after(
                40,
                lambda: self._apply_responsive_layout(
                    self._logical_window_width(self.winfo_width()),
                    window_height=self._logical_window_height(self.winfo_height()),
                    force=True,
                ),
            )

    def _animate_operation_title(self) -> None:
        if self._title_animation_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._title_animation_job)
            self._title_animation_job = None
        mode = self._motion_mode()
        timing = motion_effect_timing(
            mode,
            "title",
            busy=bool(self.worker and self.worker.is_alive()),
            minimized=self.state() in {"iconic", "withdrawn"},
        )
        if not timing.enabled:
            self.operation_title.configure(foreground=TEXT)
            return
        frames = timing.frames

        def step(index: int) -> None:
            linear = index / max(1, frames)
            progress = ease_out_cubic(linear)
            # Keep colour interpolation monotonic; the tiny spring lives in a
            # local glow pulse instead of invalid opacity/colour overshoot.
            glow = math.sin(math.pi * min(1.0, max(0.0, linear)))
            start_colour = interpolate_hex_colour(DUSTY_BLUE, CYAN, glow * 0.18)
            try:
                self.operation_title.configure(
                    foreground=interpolate_hex_colour(start_colour, TEXT, progress)
                )
            except tk.TclError:
                self._title_animation_job = None
                return
            if index < frames:
                self._title_animation_job = self.after(
                    timing.step_ms,
                    step,
                    index + 1,
                )
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
        self.simple_parameters_card = None
        self.advanced_parameters_button = None
        self.advanced_parameters_expanded = False
        if not specs:
            ttk.Label(
                self.parameters_frame,
                text="此任务无需额外参数。",
                style="Subtle.TLabel",
            ).grid(row=0, column=0, sticky="w")
            self._attach_setup_scroll_bindtag(self.parameters_frame)
            self._refresh_office_engine_button()
            self._schedule_setup_scroll_refresh()
            return
        use_sections = any(spec.section for spec in specs)
        advanced_specs = any(spec.advanced for spec in specs)
        simple_parent: tk.Misc | None = None
        if not use_sections:
            self.simple_parameters_card = RoundedCard(
                self.parameters_frame,
                fill=PANEL,
                background=PANEL,
                outline=PANEL,
                shadow=PANEL,
                radius=24,
                inset=6,
                auto_height=True,
                height=92,
            )
            self.simple_parameters_card.grid(row=0, column=0, columnspan=2, sticky="ew")
            simple_parent = self.simple_parameters_card.inner
            simple_parent.columnconfigure(1, weight=1)
        if use_sections:
            self.parameters_frame.columnconfigure(0, weight=1)
            section_colours = (PANEL, PANEL, PANEL)
            for section_index, spec in enumerate(specs):
                section = spec.section or "其他设置"
                if section in self.parameter_section_frames:
                    continue
                card = RoundedCard(
                    self.parameters_frame,
                    fill=section_colours[
                        len(self.parameter_section_frames) % len(section_colours)
                    ],
                    background=PANEL,
                    outline=PANEL,
                    shadow=PANEL,
                    radius=24,
                    inset=7,
                    auto_height=True,
                    height=96,
                )
                frame = card.inner
                frame.columnconfigure(1, weight=1)
                title = tk.Label(
                    frame,
                    text=f"●  {section}",
                    bg=card.card_fill,
                    fg=(ACCENT_DARK if len(self.parameter_section_frames) % 2 else SUCCESS),
                    font=("Microsoft YaHei UI", 10, "bold"),
                    anchor="w",
                )
                title.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 5))
                self.parameter_section_frames[section] = card
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
                    card = self.parameter_section_frames[section]
                    card.grid(row=grid_row, column=0, sticky="ew", pady=(0, 10))
                    card.grid_remove()
                    if self.advanced_parameters_frame is None:
                        self.advanced_parameters_frame = card
                    grid_row += 1

        section_counts: dict[str, int] = {}
        for row, spec in enumerate(specs):
            section = spec.section or "其他设置"
            parent: tk.Misc = (
                self.parameter_section_frames[section].inner
                if use_sections
                else simple_parent or self.parameters_frame
            )
            local_index = section_counts.get(section, 0)
            section_counts[section] = local_index + 1
            control_row = local_index * 2 + (1 if use_sections else 0)
            label_widget = ttk.Label(
                parent,
                text=spec.label,
                foreground=TEXT,
                style="CardField.TLabel" if use_sections else "CanvasField.TLabel",
            )
            if use_sections:
                label_widget.configure(background=parent.cget("bg"))
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
                    style="CardSubtle.TLabel" if use_sections else "CanvasSubtle.TLabel",
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
                control.configure(
                    postcommand=lambda combo=control: _style_combobox_popdown(combo)
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
                    style="CardSubtle.TLabel" if use_sections else "CanvasSubtle.TLabel",
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
                if use_sections:
                    hint_label.configure(background=parent.cget("bg"))
            self.parameter_rows.append(
                (spec, parent, label_widget, control, action_button, hint_label)
            )
            variable.trace_add(
                "write",
                lambda *_args, key=spec.key: self._on_parameter_value_changed(key),
            )
        self._apply_office_engine_preference_to_current()
        self._refresh_office_engine_button()
        self._layout_parameter_rows(self._layout_mode != "wide")
        self._update_parameter_visibility()
        self._attach_setup_scroll_bindtag(self.parameters_frame)
        self._schedule_setup_scroll_refresh()
        if self.simple_parameters_card is not None:
            self.simple_parameters_card._schedule_fit_requested_height()

    def _parameter_card_for_parent(self, parent: tk.Misc) -> RoundedCard | None:
        if self.simple_parameters_card is not None and self.simple_parameters_card.inner is parent:
            return self.simple_parameters_card
        return next(
            (
                card
                for card in self.parameter_section_frames.values()
                if card.inner is parent
            ),
            None,
        )

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
            title_offset = 1 if self._parameter_card_for_parent(parent) is not None else 0
            if compact:
                base_row = index * 3 + title_offset
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
                base_row = index * 2 + title_offset
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
        for parent in parents:
            card = self._parameter_card_for_parent(parent)
            if card is not None:
                card._schedule_fit_requested_height()

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
            card
            for spec, parent, *_widgets in self.parameter_rows
            if spec.advanced
            if (card := self._parameter_card_for_parent(parent)) is not None
        }
        self.advanced_parameters_expanded = not self.advanced_parameters_expanded
        for frame in advanced_frames:
            if self.advanced_parameters_expanded:
                frame.grid()
                frame._schedule_fit_requested_height()
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
                self.empty_drop_canvas,
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

    def _schedule_empty_drop_redraw(
        self,
        _event: tk.Event | None = None,
        *,
        delay: int | None = None,
        force: bool = False,
    ) -> None:
        """Keep the empty-state illustration stable during live resizing."""

        if self._closing or not hasattr(self, "empty_drop_canvas"):
            return
        if force:
            self._empty_drop_last_signature = None
        if self._empty_drop_redraw_job is not None:
            with contextlib.suppress(tk.TclError):
                self.after_cancel(self._empty_drop_redraw_job)
        wait_ms = (
            150
            if self._window_resizing
            else (110 if self.worker and self.worker.is_alive() else 68)
        ) if delay is None else max(0, int(delay))
        with contextlib.suppress(tk.TclError):
            self._empty_drop_redraw_job = self.after(
                wait_ms,
                self._render_empty_drop_canvas,
            )

    def _render_empty_drop_canvas(self, _event: tk.Event | None = None) -> None:
        self._empty_drop_redraw_job = None
        if not hasattr(self, "empty_drop_canvas"):
            return
        try:
            canvas = self.empty_drop_canvas
            width = max(240, canvas.winfo_width())
            height = max(170, canvas.winfo_height())
            signature = (width, height)
            if signature == self._empty_drop_last_signature:
                return
            previous = self._empty_drop_last_signature
            if (
                previous is not None
                and abs(previous[0] - width) < 4
                and abs(previous[1] - height) < 4
            ):
                return
            self._empty_drop_last_signature = signature
            canvas.delete("all")
            centre_x = width / 2
            portal_y = max(62, height * 0.35)
            portal_radius = min(54, max(42, round(min(width, height) * 0.16)))
            outline = DROP_PORTAL
            doodle_line = _canvas_doodle_line_colour(
                DROP_SURFACE,
                CYAN,
                strength=0.13,
            )
            doodle_accent = _canvas_doodle_line_colour(
                DROP_SURFACE,
                ACCENT,
                strength=0.21,
            )

            # Decorative sketches are confined to edge strips.  The centre
            # 22%..78% / 20%..90% remains an inviolable portal-and-text zone.
            if width >= 620 and height >= 190:
                draw_laptop_doodle(
                    canvas,
                    (width * 0.055, height * 0.26, width * 0.135, height * 0.55),
                    colour=doodle_line,
                    accent=doodle_accent,
                    tags=("doodle",),
                )
                draw_rocket_doodle(
                    canvas,
                    (width * 0.885, height * 0.18, width * 0.94, height * 0.51),
                    colour=doodle_line,
                    accent=doodle_accent,
                    tags=("doodle",),
                )
                for x, y, radius in (
                    (width * 0.08, height * 0.77, 2),
                    (width * 0.15, height * 0.86, 3),
                    (width * 0.90, height * 0.76, 3),
                ):
                    canvas.create_oval(
                        x - radius,
                        y - radius,
                        x + radius,
                        y + radius,
                        fill=doodle_accent,
                        outline="",
                        tags=("doodle",),
                    )
                edge_marks = (
                    ("file", (width * 0.16, height * 0.23, width * 0.19, height * 0.37)),
                    ("satellite", (width * 0.04, height * 0.68, width * 0.09, height * 0.88)),
                    ("file", (width * 0.82, height * 0.70, width * 0.85, height * 0.84)),
                    ("star", (width * 0.95, height * 0.70, width * 0.97, height * 0.79)),
                    ("star", (width * 0.20, height * 0.76, width * 0.22, height * 0.85)),
                )
                for kind, box in edge_marks:
                    if kind == "file":
                        draw_file_doodle(
                            canvas,
                            box,
                            colour=doodle_line,
                            tags=("doodle",),
                        )
                    elif kind == "satellite":
                        draw_satellite_doodle(
                            canvas,
                            box,
                            colour=doodle_line,
                            tags=("doodle",),
                        )
                    else:
                        draw_starlet(
                            canvas,
                            box,
                            colour=doodle_accent,
                            tags=("doodle",),
                        )
            elif width >= 400 and height >= 180:
                draw_laptop_doodle(
                    canvas,
                    (width * 0.04, height * 0.27, width * 0.17, height * 0.61),
                    colour=doodle_line,
                    accent=doodle_accent,
                    tags=("doodle",),
                )

            # Sparse scan lines keep the empty state alive while remaining calm.
            for offset in range(-2, 3):
                y = portal_y + offset * 28
                canvas.create_line(
                    max(22, centre_x - portal_radius * 2.1),
                    y,
                    min(width - 22, centre_x + portal_radius * 2.1),
                    y,
                    fill=DROP_SCAN,
                    width=1,
                    dash=(4, 8),
                    tags=("foreground",),
                )
            canvas.create_oval(
                centre_x - portal_radius - 13,
                portal_y - portal_radius - 13,
                centre_x + portal_radius + 13,
                portal_y + portal_radius + 13,
                fill="",
                outline=DROP_RING,
                width=2,
                dash=(3, 7),
                tags=("foreground",),
            )
            canvas.create_oval(
                centre_x - portal_radius,
                portal_y - portal_radius,
                centre_x + portal_radius,
                portal_y + portal_radius,
                fill=ACCENT_SOFT,
                outline=outline,
                width=2,
                tags=("foreground",),
            )
            canvas.create_oval(
                centre_x - 24,
                portal_y - 24,
                centre_x + 24,
                portal_y + 24,
                fill=DROP_SURFACE,
                outline=CYAN,
                width=2,
                tags=("foreground",),
            )
            canvas.create_line(
                centre_x,
                portal_y + 12,
                centre_x,
                portal_y - 13,
                fill=TEXT,
                width=3,
                arrow="last",
                arrowshape=(9, 10, 4),
                tags=("foreground",),
            )
            canvas.create_line(
                centre_x - 12,
                portal_y + 15,
                centre_x + 12,
                portal_y + 15,
                fill=TEXT,
                width=3,
                tags=("foreground",),
            )
            for angle, colour in ((28, CYAN), (156, PINK), (258, ACCENT)):
                radians = math.radians(angle)
                x = centre_x + math.cos(radians) * (portal_radius + 13)
                y = portal_y + math.sin(radians) * (portal_radius + 13)
                canvas.create_oval(
                    x - 3,
                    y - 3,
                    x + 3,
                    y + 3,
                    fill=colour,
                    outline="",
                    tags=("foreground",),
                )
            canvas.create_text(
                centre_x,
                portal_y + portal_radius + 42,
                text="将文件接入处理舱",
                fill=TEXT,
                font=("Microsoft YaHei UI", 12, "bold"),
                tags=("foreground",),
            )
            canvas.create_text(
                centre_x,
                portal_y + portal_radius + 68,
                text="拖入文件或文件夹，也可点击选择",
                fill=MUTED,
                font=("Microsoft YaHei UI", 9),
                tags=("foreground",),
            )
            canvas.tag_lower("doodle")
            canvas.tag_raise("foreground")
        except tk.TclError:
            return

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
        self._popup_menu_below(self.file_more_menu, self.file_more_button)

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
        if self.input_paths:
            self.empty_drop_canvas.grid_remove()
            self.file_tree.grid()
            self.file_scroll.grid()
            self.file_horizontal_scroll.grid()
            self.drop_hint_label.configure(
                text="继续拖入可追加文件 · 已自动过滤重复与不支持类型",
                style="DropHint.TLabel",
            )
        else:
            self.file_tree.grid_remove()
            self.file_scroll.grid_remove()
            self.file_horizontal_scroll.grid_remove()
            self.empty_drop_canvas.grid()
            self._schedule_empty_drop_redraw(delay=0, force=True)
            self.drop_hint_label.configure(
                text="文件接入区 · 拖放、点击或批量添加",
                style="DropHint.TLabel" if DND_FILES is not None else "DropError.TLabel",
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
        self.engine_select_button.configure(state="disabled")
        self._render_progress_track()
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
        try:
            displayed_percent = float(self.progress_var.get())
        except (tk.TclError, TypeError, ValueError):
            displayed_percent = percent
        self.progress_percent_label.configure(text=f"{displayed_percent:.0f}%")
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
            self._progress_animation_last_at = None
            return
        with contextlib.suppress(tk.TclError):
            self.after_cancel(self._progress_animation_job)
        self._progress_animation_job = None
        self._progress_animation_last_at = None

    def _animate_progress_to(self, target: float) -> None:
        try:
            current = float(self.progress_var.get())
        except (tk.TclError, TypeError, ValueError):
            current = float(target)
        target = min(100.0, max(0.0, float(target)))
        self._progress_target = target
        mode = self._motion_mode()
        timing = motion_effect_timing(
            mode,
            "progress",
            busy=bool(self.worker and self.worker.is_alive()),
            minimized=self.state() in {"iconic", "withdrawn"},
        )
        if (
            target <= current
            or abs(target - current) < 0.08
            or not self.winfo_viewable()
            or not timing.enabled
        ):
            self._cancel_progress_animation()
            self.progress_var.set(target)
            return

        if self._progress_animation_job is not None:
            return
        self._progress_animation_last_at = time.monotonic()

        def step() -> None:
            now = time.monotonic()
            last_at = self._progress_animation_last_at or now
            self._progress_animation_last_at = now
            try:
                displayed = float(self.progress_var.get())
                latest_target = min(100.0, max(0.0, self._progress_target))
                value = smooth_progress_step(
                    displayed,
                    latest_target,
                    (now - last_at) * 1000.0,
                    time_constant_ms=120.0 if latest_target >= 100.0 else 145.0,
                )
                self.progress_var.set(value)
                self.progress_percent_label.configure(text=f"{value:.0f}%")
            except (tk.TclError, TypeError, ValueError):
                self._progress_animation_job = None
                self._progress_animation_last_at = None
                return
            if abs(latest_target - value) > 0.06:
                self._progress_animation_job = self.after(
                    max(16, timing.step_ms),
                    step,
                )
            else:
                self.progress_var.set(latest_target)
                self.progress_percent_label.configure(text=f"{latest_target:.0f}%")
                self._progress_animation_job = None
                self._progress_animation_last_at = None

        self._progress_animation_job = self.after(max(16, timing.step_ms), step)

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
        self._progress_indeterminate = False

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
                            motion_mode=self._motion_mode(),
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
                                motion_mode=self._motion_mode(),
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
                                motion_mode=self._motion_mode(),
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
        self._render_progress_track()
        self.cancel_button.configure(state="disabled")
        if self._closing:
            return
        self.engine_select_button.configure(state="normal")
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
        for job in tuple(self._preview_resize_jobs.values()):
            with contextlib.suppress(tk.TclError):
                self.after_cancel(job)
        self._preview_resize_jobs.clear()
        self._preview_last_sizes.clear()
        self._clear_image_preview_images()

    def _on_close(self) -> None:
        if not self.worker or not self.worker.is_alive():
            self._closing = True
            self._dispose_image_preview_resources()
            self._dispose_motion_resources()
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
        self._dispose_motion_resources()
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
        self._dispose_motion_resources()
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
