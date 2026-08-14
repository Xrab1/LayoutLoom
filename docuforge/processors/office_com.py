"""Optional high-fidelity Microsoft Office COM processors.

This module intentionally contains only operations that need the native Office
layout engines.  Importing it is safe on every platform; an operation raises a
clear :class:`~docuforge.models.MissingEngineError` when Windows, pywin32, or
the required desktop Office application is unavailable.

All public processors validate their arguments before starting Office, keep
the source file untouched, stage generated data away from the final name, and
publish complete files atomically.  They return ``list[Path]`` to match the
rest of LayoutLoom's processor API.
"""

from __future__ import annotations

import importlib
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterator, Sequence
from xml.etree import ElementTree as ET

from docuforge.models import MissingEngineError, ValidationError
from docuforge.utils import atomic_output, ensure_output_dir, safe_filename

PathLike = str | os.PathLike[str]

_POWERPOINT_EXTENSIONS = {
    ".ppt",
    ".pptx",
    ".pptm",
    ".pot",
    ".potx",
    ".potm",
    ".pps",
    ".ppsx",
    ".ppsm",
}
_EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb"}
_WORD_OOXML_EXTENSIONS = {".docx", ".docm"}

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

_A1_RANGE_RE = re.compile(
    r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}):"
    r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$"
)
_A1_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$")
_EXCEL_INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")

_PIVOT_FUNCTIONS = {
    "sum": -4157,  # xlSum
    "求和": -4157,
    "count": -4112,  # xlCount
    "计数": -4112,
    "average": -4106,  # xlAverage
    "avg": -4106,
    "平均值": -4106,
    "min": -4139,  # xlMin
    "最小值": -4139,
    "max": -4136,  # xlMax
    "最大值": -4136,
    "product": -4149,  # xlProduct
    "乘积": -4149,
    "count_numbers": -4113,  # xlCountNums
    "countnums": -4113,
    "数值计数": -4113,
    "stddev": -4155,  # xlStDev
    "stdev": -4155,
    "stddevp": -4156,  # xlStDevP
    "stdevp": -4156,
    "var": -4164,  # xlVar
    "variance": -4164,
    "varp": -4165,  # xlVarP
}

_NAMED_RGB = {
    "black": (0, 0, 0),
    "white": (255, 255, 255),
    "red": (255, 0, 0),
    "green": (0, 128, 0),
    "blue": (0, 0, 255),
    "yellow": (255, 255, 0),
    "gray": (128, 128, 128),
    "grey": (128, 128, 128),
    "transparent": (255, 255, 255),
}


@dataclass(frozen=True)
class _ComRuntime:
    pythoncom: Any
    client: Any


@dataclass(frozen=True)
class WordBlankPageSafety:
    """Pure OOXML preflight result for conservative blank-page removal."""

    explicit_page_breaks: int
    high_risk_reasons: tuple[str, ...] = ()

    @property
    def safe(self) -> bool:
        return not self.high_risk_reasons


def _require_source(source: PathLike, extensions: set[str], family: str) -> Path:
    path = Path(source).expanduser().resolve()
    if not path.is_file():
        raise ValidationError(f"文件不存在：{path}")
    if path.suffix.lower() not in extensions:
        allowed = "、".join(sorted(extensions))
        raise ValidationError(
            f"不支持的 {family} 文件格式：{path.suffix or '无扩展名'}；支持 {allowed}"
        )
    return path


def _require_int(
    value: Any,
    name: str,
    *,
    minimum: int = 1,
    maximum: int | None = None,
) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValidationError(f"{name} 必须是整数")
    if value < minimum:
        raise ValidationError(f"{name} 不能小于 {minimum}")
    if maximum is not None and value > maximum:
        raise ValidationError(f"{name} 不能大于 {maximum}")
    return value


def _require_positive_number(value: Any, name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValidationError(f"{name} 必须是数字")
    result = float(value)
    if result <= 0:
        raise ValidationError(f"{name} 必须大于 0")
    return result


def _normalize_image_format(value: Any) -> str:
    name = str(value).strip().lower().lstrip(".")
    if name == "jpeg":
        name = "jpg"
    if name not in {"png", "jpg"}:
        raise ValidationError("format 仅支持 png、jpg/jpeg")
    return name


def _scaled_dimensions(
    native_width: Any,
    native_height: Any,
    width: int,
    height: int | None,
) -> tuple[int, int]:
    """Return validated export dimensions while preserving aspect if needed."""

    export_width = _require_int(width, "width")
    if height is not None:
        return export_width, _require_int(height, "height")
    if isinstance(native_width, bool) or isinstance(native_height, bool):
        raise ValidationError("PowerPoint 页面尺寸无效")
    try:
        source_width = float(native_width)
        source_height = float(native_height)
    except (TypeError, ValueError) as exc:
        raise ValidationError("PowerPoint 页面尺寸无效") from exc
    if source_width <= 0 or source_height <= 0:
        raise ValidationError("PowerPoint 页面尺寸无效")
    return export_width, max(1, round(export_width * source_height / source_width))


def _parse_rgb(value: Any, name: str = "color") -> tuple[int, int, int]:
    if isinstance(value, str):
        text = value.strip().lower()
        if text in _NAMED_RGB:
            return _NAMED_RGB[text]
        if text.startswith("#"):
            text = text[1:]
        if re.fullmatch(r"[0-9a-f]{3}", text):
            text = "".join(character * 2 for character in text)
        if re.fullmatch(r"[0-9a-f]{6}", text):
            return tuple(int(text[index : index + 2], 16) for index in (0, 2, 4))  # type: ignore[return-value]
        match = re.fullmatch(r"\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*", value)
        if match:
            channels = tuple(int(channel) for channel in match.groups())
            if all(0 <= channel <= 255 for channel in channels):
                return channels  # type: ignore[return-value]
    elif isinstance(value, Sequence) and not isinstance(value, (bytes, bytearray)):
        channels = tuple(value)
        if len(channels) == 3 and all(
            isinstance(channel, int)
            and not isinstance(channel, bool)
            and 0 <= channel <= 255
            for channel in channels
        ):
            return channels  # type: ignore[return-value]
    raise ValidationError(f"{name} 必须是颜色名称、#RRGGBB 或 RGB 三元组")


def _rgb_to_office_bgr(rgb: Sequence[int]) -> int:
    """Convert an RGB triple to the OLE_COLOR integer Office calls ``RGB``."""

    red, green, blue = _parse_rgb(rgb)
    return red | (green << 8) | (blue << 16)


def _column_number(label: str) -> int:
    value = 0
    for character in label.upper():
        value = value * 26 + ord(character) - ord("A") + 1
    return value


def _normalize_a1_range(value: Any) -> str:
    text = str(value).strip()
    match = _A1_RANGE_RE.fullmatch(text)
    if match is None:
        raise ValidationError("source_range 必须是矩形 A1 区域，例如 A1:D100")
    first_col, first_row, last_col, last_row = match.groups()
    first_col_number = _column_number(first_col)
    last_col_number = _column_number(last_col)
    first_row_number = int(first_row)
    last_row_number = int(last_row)
    if first_col_number > 16384 or last_col_number > 16384:
        raise ValidationError("source_range 超出 Excel 最大列 XFD")
    if first_row_number > 1_048_576 or last_row_number > 1_048_576:
        raise ValidationError("source_range 超出 Excel 最大行数")
    if first_col_number > last_col_number or first_row_number > last_row_number:
        raise ValidationError("source_range 的起始单元格必须位于结束单元格左上方")
    if first_row_number == last_row_number:
        raise ValidationError("source_range 至少需要标题行和一行数据")
    return f"{first_col.upper()}{first_row_number}:{last_col.upper()}{last_row_number}"


def _normalize_a1_cell(value: Any, name: str = "target_cell") -> str:
    text = str(value).strip()
    match = _A1_CELL_RE.fullmatch(text)
    if match is None:
        raise ValidationError(f"{name} 必须是单个 A1 单元格，例如 A1")
    column, row = match.groups()
    if _column_number(column) > 16384 or int(row) > 1_048_576:
        raise ValidationError(f"{name} 超出 Excel 工作表范围")
    return f"{column.upper()}{int(row)}"


def _a1_range_bounds(value: str) -> tuple[int, int, int, int]:
    match = _A1_RANGE_RE.fullmatch(value)
    if match is None:  # normalized callers guarantee this invariant
        raise ValidationError("无效的 A1 区域")
    first_col, first_row, last_col, last_row = match.groups()
    return (
        _column_number(first_col),
        int(first_row),
        _column_number(last_col),
        int(last_row),
    )


def _a1_cell_position(value: str) -> tuple[int, int]:
    match = _A1_CELL_RE.fullmatch(value)
    if match is None:  # normalized callers guarantee this invariant
        raise ValidationError("无效的 A1 单元格")
    column, row = match.groups()
    return _column_number(column), int(row)


def _normalize_field_names(value: Any, name: str) -> tuple[str, ...]:
    if value is None:
        return ()
    if isinstance(value, str):
        normalized = value.replace("，", ",")
        items = normalized.split(",") if "," in normalized else (normalized,)
    else:
        try:
            items = tuple(value)
        except TypeError as exc:
            raise ValidationError(f"{name} 必须是字段名序列") from exc
    result = tuple(str(item).strip() for item in items)
    if any(not item for item in result):
        raise ValidationError(f"{name} 不能包含空字段名")
    folded = [item.casefold() for item in result]
    if len(folded) != len(set(folded)):
        raise ValidationError(f"{name} 不能包含重复字段")
    return result


def _normalize_pivot_function(value: Any) -> int:
    key = str(value).strip().lower().replace("-", "_").replace(" ", "_")
    try:
        return _PIVOT_FUNCTIONS[key]
    except KeyError as exc:
        choices = "sum、count、average、min、max、product、count_numbers、stddev、stddevp、var、varp"
        raise ValidationError(f"function 不受支持；可选 {choices}") from exc


def _validate_sheet_name(value: Any, name: str) -> str:
    text = str(value).strip()
    if not text:
        raise ValidationError(f"{name} 不能为空")
    if len(text) > 31 or _EXCEL_INVALID_SHEET_CHARS.search(text):
        raise ValidationError(f"{name} 不是有效的 Excel 工作表名称")
    if text.startswith("'") or text.endswith("'"):
        raise ValidationError(f"{name} 不能以单引号开头或结尾")
    return text


def _analyze_word_document_xml(xml: bytes | str) -> WordBlankPageSafety:
    """Conservatively identify OOXML constructs unsafe for page deletion."""

    try:
        root = ET.fromstring(xml)
    except (ET.ParseError, ValueError, TypeError) as exc:
        raise ValidationError("Word document.xml 无法解析") from exc

    reasons: list[str] = []
    parent = {child: node for node in root.iter() for child in node}
    page_breaks = root.findall(f".//{{{_W}}}br[@{{{_W}}}type='page']")

    if root.findall(f".//{{{_W}}}pPr/{{{_W}}}sectPr"):
        reasons.append("文档包含中间分节符")
    body = root.find(f".//{{{_W}}}body")
    if body is None:
        reasons.append("文档缺少正文结构")
    elif sum(1 for child in body if child.tag == f"{{{_W}}}sectPr") > 1:
        reasons.append("文档包含多个正文级节属性")
    if root.findall(f".//{{{_W}}}pageBreakBefore"):
        reasons.append("文档使用段前分页")
    if root.findall(f".//{{{_W}}}br[@{{{_W}}}type='column']"):
        reasons.append("文档包含分栏符")
    if root.findall(f".//{{{_W}}}altChunk"):
        reasons.append("文档包含外部嵌入内容")
    if root.findall(f".//{{{_WP}}}anchor"):
        reasons.append("文档包含浮动图形")
    if root.findall(f".//{{{_W}}}txbxContent"):
        reasons.append("文档包含文本框正文")
    revision_tags = ("ins", "del", "moveFrom", "moveTo")
    if any(root.findall(f".//{{{_W}}}{tag}") for tag in revision_tags):
        reasons.append("文档包含未处理的修订")
    for columns in root.findall(f".//{{{_W}}}sectPr/{{{_W}}}cols"):
        count = columns.get(f"{{{_W}}}num")
        if (count is not None and count != "1") or columns.findall(f"{{{_W}}}col"):
            reasons.append("文档使用多栏排版")
            break

    for page_break in page_breaks:
        ancestor = parent.get(page_break)
        while ancestor is not None:
            if ancestor.tag == f"{{{_W}}}tc":
                reasons.append("表格单元格中包含手动分页符")
                ancestor = None
                break
            ancestor = parent.get(ancestor)

    return WordBlankPageSafety(
        explicit_page_breaks=len(page_breaks),
        high_risk_reasons=tuple(dict.fromkeys(reasons)),
    )


def _word_settings_risks(xml: bytes | str) -> tuple[str, ...]:
    try:
        root = ET.fromstring(xml)
    except (ET.ParseError, ValueError, TypeError) as exc:
        raise ValidationError("Word settings.xml 无法解析") from exc

    reasons: list[str] = []
    track_revisions = root.find(f".//{{{_W}}}trackRevisions")
    if track_revisions is not None:
        value = str(track_revisions.get(f"{{{_W}}}val", "true")).strip().lower()
        if value not in {"0", "false", "off", "no"}:
            reasons.append("文档已启用修订跟踪")
    protection = root.find(f".//{{{_W}}}documentProtection")
    if protection is not None:
        enforcement = (
            str(protection.get(f"{{{_W}}}enforcement", "true")).strip().lower()
        )
        if enforcement not in {"0", "false", "off", "no"}:
            reasons.append("文档启用了编辑保护")
    return tuple(reasons)


def _is_explicit_blank_page_text(value: Any) -> bool:
    """True only for an otherwise blank Word page containing a manual break."""

    text = "" if value is None else str(value)
    if "\x0c" not in text:
        return False
    return _is_blank_word_page_text(text)


def _is_blank_word_page_text(value: Any) -> bool:
    text = "" if value is None else str(value)
    ignored = " \t\r\n\x07\x0b\x0c\u00a0\u200b\ufeff"
    return not text.strip(ignored)


def _load_com_runtime() -> _ComRuntime:
    if sys.platform != "win32":
        raise MissingEngineError("Microsoft Office COM 仅支持 Windows 桌面环境")
    try:
        pythoncom = importlib.import_module("pythoncom")
        client = importlib.import_module("win32com.client")
    except (ImportError, ModuleNotFoundError) as exc:
        raise MissingEngineError(
            "缺少 pywin32，无法调用 Microsoft Office COM；请安装 pywin32"
        ) from exc
    return _ComRuntime(pythoncom=pythoncom, client=client)


def _com_error_text(exc: BaseException) -> str:
    text = " ".join(str(exc).split())
    return text[:500] or exc.__class__.__name__


@contextmanager
def _com_application(
    runtime: _ComRuntime,
    prog_id: str,
    display_name: str,
) -> Iterator[Any]:
    initialized = False
    application = None
    try:
        runtime.pythoncom.CoInitialize()
        initialized = True
        try:
            application = runtime.client.DispatchEx(prog_id)
        except Exception as exc:
            raise MissingEngineError(
                f"无法启动 Microsoft {display_name}；请确认已安装可自动化的桌面版 Office："
                f"{_com_error_text(exc)}"
            ) from exc
        from .office import _validate_microsoft_com_application

        _validate_microsoft_com_application(application, prog_id, display_name)
        yield application
    except (ValidationError, MissingEngineError, FileExistsError):
        raise
    except Exception as exc:
        raise MissingEngineError(
            f"Microsoft {display_name} COM 操作失败：{_com_error_text(exc)}"
        ) from exc
    finally:
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass
        if initialized:
            try:
                runtime.pythoncom.CoUninitialize()
            except Exception:
                pass


@contextmanager
def _powerpoint_presentation(
    source: Path,
    *,
    runtime: _ComRuntime,
    read_only: bool,
) -> Iterator[tuple[Any, Any]]:
    with _com_application(
        runtime, "PowerPoint.Application", "PowerPoint"
    ) as application:
        presentation = None
        try:
            try:
                application.Visible = False
                application.DisplayAlerts = 1  # ppAlertsNone
                application.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
            presentation = application.Presentations.Open(
                str(source), ReadOnly=read_only, Untitled=False, WithWindow=False
            )
            yield application, presentation
        finally:
            if presentation is not None:
                try:
                    presentation.Close()
                except Exception:
                    pass


@contextmanager
def _excel_workbook(
    source: Path,
    *,
    runtime: _ComRuntime,
) -> Iterator[tuple[Any, Any]]:
    with _com_application(runtime, "Excel.Application", "Excel") as application:
        workbook = None
        try:
            application.Visible = False
            application.DisplayAlerts = False
            application.ScreenUpdating = False
            application.EnableEvents = False
            application.AskToUpdateLinks = False
            try:
                application.AutomationSecurity = 3  # disable workbook macros
            except Exception:
                pass
            workbook = application.Workbooks.Open(
                str(source),
                UpdateLinks=0,
                ReadOnly=False,
                AddToMru=False,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
            )
            yield application, workbook
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass


@contextmanager
def _word_document(
    source: Path,
    *,
    runtime: _ComRuntime,
) -> Iterator[tuple[Any, Any]]:
    with _com_application(runtime, "Word.Application", "Word") as application:
        document = None
        try:
            application.Visible = False
            application.DisplayAlerts = 0
            application.ScreenUpdating = False
            try:
                application.Options.ConfirmConversions = False
                application.Options.SaveNormalPrompt = False
            except Exception:
                pass
            try:
                application.AutomationSecurity = 3  # disable document macros
            except Exception:
                pass
            document = application.Documents.Open(
                str(source),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
                Revert=False,
                Visible=False,
                OpenAndRepair=False,
                NoEncodingDialog=True,
            )
            yield application, document
        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=False)
                except Exception:
                    pass


def _numbered_candidate(path: Path, index: int) -> Path:
    return path if index == 0 else path.with_name(f"{path.stem}_{index}{path.suffix}")


def _reserve_path(path: Path, overwrite: bool) -> tuple[Path, bool]:
    path.parent.mkdir(parents=True, exist_ok=True)
    if overwrite:
        return path, False
    for index in range(10_000):
        candidate = _numbered_candidate(path, index)
        try:
            descriptor = os.open(candidate, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
        except FileExistsError:
            continue
        else:
            os.close(descriptor)
            return candidate, True
    raise ValidationError(f"无法为输出生成唯一文件名：{path.name}")


@contextmanager
def _reserve_paths(paths: Sequence[Path], overwrite: bool) -> Iterator[list[Path]]:
    reservations: list[tuple[Path, bool]] = []
    completed = False
    try:
        for path in paths:
            reservations.append(_reserve_path(path, overwrite))
        yield [path for path, _reserved in reservations]
        completed = True
    finally:
        if not completed:
            for path, reserved in reservations:
                if reserved:
                    path.unlink(missing_ok=True)


def _publish_file(staged: Path, target: Path) -> None:
    if not staged.is_file() or staged.stat().st_size <= 0:
        raise MissingEngineError(f"Office 未生成有效输出文件：{target.name}")
    with atomic_output(target) as temporary:
        shutil.copy2(staged, temporary)


def _directory_target(
    source: Path,
    output_dir: PathLike,
    *,
    tag: str,
    suffix: str | None = None,
) -> Path:
    directory = ensure_output_dir(Path(output_dir))
    extension = source.suffix if suffix is None else suffix
    if not extension.startswith("."):
        extension = f".{extension}"
    target = directory / f"{source.stem}_{safe_filename(tag)}{extension}"
    if target.resolve() == source.resolve():
        target = directory / f"{source.stem}_output{extension}"
    return target


def _explicit_file_target(
    source: Path,
    output_path: PathLike,
    *,
    default_suffix: str,
    allowed_suffixes: set[str],
) -> Path:
    target = Path(output_path).expanduser()
    if not target.suffix:
        target = target.with_suffix(default_suffix)
    if target.suffix.lower() not in allowed_suffixes:
        allowed = "、".join(sorted(allowed_suffixes))
        raise ValidationError(f"输出格式不受支持；请使用 {allowed}")
    directory = ensure_output_dir(target.parent)
    target = (directory / target.name).resolve()
    if target == source.resolve():
        target = target.with_name(f"{target.stem}_output{target.suffix}")
    return target


def _video_target(source: Path, output_path: PathLike) -> Path:
    requested = Path(output_path).expanduser()
    if (requested.exists() and requested.is_dir()) or not requested.suffix:
        directory = ensure_output_dir(requested)
        return (directory / f"{source.stem}.mp4").resolve()
    return _explicit_file_target(
        source, requested, default_suffix=".mp4", allowed_suffixes={".mp4"}
    )


def _temporary_directory(parent: Path, prefix: str) -> tempfile.TemporaryDirectory[str]:
    parent.mkdir(parents=True, exist_ok=True)
    return tempfile.TemporaryDirectory(prefix=prefix, dir=parent)


def _iter_com_collection(collection: Any) -> Iterator[Any]:
    for index in range(1, int(collection.Count) + 1):
        yield collection.Item(index)


def ppt_to_images(
    source: PathLike,
    output_dir: PathLike,
    format: str = "png",
    width: int = 1920,
    height: int | None = None,
    overwrite: bool = False,
) -> list[Path]:
    """Render every PowerPoint slide with the native PowerPoint engine."""

    source_path = _require_source(source, _POWERPOINT_EXTENSIONS, "PowerPoint")
    image_format = _normalize_image_format(format)
    _require_int(width, "width")
    if height is not None:
        _require_int(height, "height")
    runtime = _load_com_runtime()
    directory = ensure_output_dir(Path(output_dir))

    with _powerpoint_presentation(source_path, runtime=runtime, read_only=True) as (
        _,
        presentation,
    ):
        slide_count = int(presentation.Slides.Count)
        if slide_count < 1:
            raise ValidationError("PowerPoint 演示文稿不包含幻灯片")
        export_width, export_height = _scaled_dimensions(
            presentation.PageSetup.SlideWidth,
            presentation.PageSetup.SlideHeight,
            width,
            height,
        )
        requested = [
            directory / f"{source_path.stem}_{index:03d}.{image_format}"
            for index in range(1, slide_count + 1)
        ]
        with _reserve_paths(requested, overwrite) as targets:
            with _temporary_directory(
                directory, ".docuforge-ppt-images-"
            ) as temporary_name:
                staging = Path(temporary_name)
                generated: list[Path] = []
                filter_name = "PNG" if image_format == "png" else "JPG"
                for index in range(1, slide_count + 1):
                    staged = staging / f"slide_{index:06d}.{image_format}"
                    presentation.Slides.Item(index).Export(
                        str(staged), filter_name, export_width, export_height
                    )
                    if not staged.is_file() or staged.stat().st_size <= 0:
                        raise MissingEngineError(
                            f"PowerPoint 未能导出第 {index} 张幻灯片"
                        )
                    generated.append(staged)
                for staged, target in zip(generated, targets):
                    _publish_file(staged, target)
            return targets


def _wait_for_powerpoint_video(
    presentation: Any,
    output: Path,
    timeout: float,
    *,
    clock: Callable[[], float] = time.monotonic,
    sleeper: Callable[[float], None] = time.sleep,
) -> None:
    """Wait for asynchronous ``CreateVideo`` completion.

    ``PpMediaTaskStatus`` values are: none=0, in-progress=1, queued=2,
    done=3, failed=4.
    """

    deadline = clock() + timeout
    while True:
        status = int(presentation.CreateVideoStatus)
        if status == 3:
            if not output.is_file() or output.stat().st_size <= 0:
                raise MissingEngineError("PowerPoint 报告视频完成，但输出文件无效")
            return
        if status == 4:
            raise MissingEngineError("PowerPoint CreateVideo 渲染失败")
        if status not in {0, 1, 2}:
            raise MissingEngineError(f"PowerPoint 返回未知视频任务状态：{status}")
        remaining = deadline - clock()
        if remaining <= 0:
            raise MissingEngineError(f"PowerPoint 视频导出超时（{timeout:g} 秒）")
        sleeper(min(0.5, remaining))


def ppt_to_video(
    source: PathLike,
    output_path: PathLike,
    use_timings: bool = True,
    slide_duration: int = 5,
    resolution: int = 1080,
    fps: int = 30,
    quality: int = 85,
    timeout: float = 1800,
    overwrite: bool = False,
) -> list[Path]:
    """Create an MP4 through PowerPoint, preserving animations and timings."""

    source_path = _require_source(source, _POWERPOINT_EXTENSIONS, "PowerPoint")
    if not isinstance(use_timings, bool):
        raise ValidationError("use_timings 必须是布尔值")
    duration = _require_int(slide_duration, "slide_duration")
    vertical_resolution = _require_int(
        resolution, "resolution", minimum=30, maximum=3072
    )
    frames_per_second = _require_int(fps, "fps", maximum=100)
    video_quality = _require_int(quality, "quality", maximum=100)
    timeout_seconds = _require_positive_number(timeout, "timeout")
    target_request = _video_target(source_path, output_path)
    runtime = _load_com_runtime()

    with _powerpoint_presentation(source_path, runtime=runtime, read_only=True) as (
        _,
        presentation,
    ):
        with _reserve_paths([target_request], overwrite) as targets:
            target = targets[0]
            with _temporary_directory(
                target.parent, ".docuforge-ppt-video-"
            ) as temporary_name:
                staged = Path(temporary_name) / f"{source_path.stem}.mp4"
                try:
                    presentation.CreateVideo(
                        str(staged),
                        use_timings,
                        duration,
                        vertical_resolution,
                        frames_per_second,
                        video_quality,
                    )
                except AttributeError as exc:
                    raise MissingEngineError(
                        "当前 PowerPoint 版本不支持 CreateVideo"
                    ) from exc
                _wait_for_powerpoint_video(presentation, staged, timeout_seconds)
                _publish_file(staged, target)
            return [target]


def _load_pillow() -> Any:
    try:
        return importlib.import_module("PIL.Image")
    except (ImportError, ModuleNotFoundError) as exc:
        raise MissingEngineError("PPT 长图拼接需要安装 Pillow") from exc


def ppt_to_long_image(
    source: PathLike,
    output_path: PathLike,
    direction: str = "vertical",
    spacing: int = 0,
    background: Any = "white",
    width: int = 1920,
    overwrite: bool = False,
) -> list[Path]:
    """Render slides natively and join them into one vertical/horizontal image."""

    source_path = _require_source(source, _POWERPOINT_EXTENSIONS, "PowerPoint")
    direction_name = str(direction).strip().lower()
    if direction_name not in {"vertical", "horizontal"}:
        raise ValidationError("direction 必须是 vertical 或 horizontal")
    gap = _require_int(spacing, "spacing", minimum=0)
    render_width = _require_int(width, "width")
    background_rgb = _parse_rgb(background, "background")
    target_request = _explicit_file_target(
        source_path,
        output_path,
        default_suffix=".png",
        allowed_suffixes={".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"},
    )
    image_module = _load_pillow()
    # Fail before reserving output when COM cannot possibly run.
    _load_com_runtime()

    with _reserve_paths([target_request], overwrite) as targets:
        target = targets[0]
        with _temporary_directory(
            target.parent, ".docuforge-ppt-long-"
        ) as temporary_name:
            render_dir = Path(temporary_name) / "slides"
            render_dir.mkdir()
            slides = ppt_to_images(
                source_path,
                render_dir,
                format="png",
                width=render_width,
                height=None,
                overwrite=True,
            )
            opened = [image_module.open(path).convert("RGBA") for path in slides]
            try:
                if direction_name == "vertical":
                    canvas_width = max(image.width for image in opened)
                    canvas_height = sum(image.height for image in opened) + gap * (
                        len(opened) - 1
                    )
                else:
                    canvas_width = sum(image.width for image in opened) + gap * (
                        len(opened) - 1
                    )
                    canvas_height = max(image.height for image in opened)
                canvas = image_module.new(
                    "RGB", (canvas_width, canvas_height), background_rgb
                )
                try:
                    x = y = 0
                    for slide in opened:
                        canvas.paste(slide, (x, y), slide)
                        if direction_name == "vertical":
                            y += slide.height + gap
                        else:
                            x += slide.width + gap
                    save_format = {
                        ".png": "PNG",
                        ".jpg": "JPEG",
                        ".jpeg": "JPEG",
                        ".webp": "WEBP",
                        ".tif": "TIFF",
                        ".tiff": "TIFF",
                    }[target.suffix.lower()]
                    with atomic_output(target) as temporary:
                        save_options = (
                            {"quality": 95} if save_format in {"JPEG", "WEBP"} else {}
                        )
                        canvas.save(temporary, format=save_format, **save_options)
                finally:
                    canvas.close()
            finally:
                for slide in opened:
                    slide.close()
        return [target]


def _set_shape_font(shape: Any, font_name: str) -> int:
    changed = 0
    try:
        if bool(shape.HasTextFrame) and bool(shape.TextFrame.HasText):
            shape.TextFrame.TextRange.Font.Name = font_name
            changed += 1
    except Exception:
        pass
    try:
        if bool(shape.HasTable):
            table = shape.Table
            for row in range(1, int(table.Rows.Count) + 1):
                for column in range(1, int(table.Columns.Count) + 1):
                    changed += _set_shape_font(table.Cell(row, column).Shape, font_name)
    except Exception:
        pass
    try:
        if int(shape.Type) == 6:  # msoGroup
            for item in _iter_com_collection(shape.GroupItems):
                changed += _set_shape_font(item, font_name)
    except Exception:
        pass
    return changed


def _presentation_masters(presentation: Any) -> list[Any]:
    masters: list[Any] = []
    try:
        for design in _iter_com_collection(presentation.Designs):
            masters.append(design.SlideMaster)
    except Exception:
        pass
    if not masters:
        masters.append(presentation.SlideMaster)
    return masters


def _set_theme_fonts(master: Any, font_name: str) -> int:
    changed = 0
    try:
        scheme = master.Theme.ThemeFontScheme
        for family in (scheme.MajorFont, scheme.MinorFont):
            for script in (1, 2, 3):  # Latin, East Asian, complex script
                try:
                    family(script).Name = font_name
                    changed += 1
                except Exception:
                    pass
    except Exception:
        pass
    return changed


def ppt_modify_master(
    source: PathLike,
    output_dir: PathLike,
    background_color: Any = None,
    font_name: str | None = None,
    footer_text: str | None = None,
    overwrite: bool = False,
) -> list[Path]:
    """Modify PowerPoint slide masters through the native object model."""

    source_path = _require_source(source, _POWERPOINT_EXTENSIONS, "PowerPoint")
    if background_color is None and font_name is None and footer_text is None:
        raise ValidationError(
            "至少需要设置 background_color、font_name 或 footer_text 之一"
        )
    office_color = None
    if background_color is not None:
        office_color = _rgb_to_office_bgr(
            _parse_rgb(background_color, "background_color")
        )
    normalized_font = None
    if font_name is not None:
        normalized_font = str(font_name).strip()
        if not normalized_font:
            raise ValidationError("font_name 不能为空")
    normalized_footer = None if footer_text is None else str(footer_text)
    target_request = _directory_target(source_path, output_dir, tag="母版")
    runtime = _load_com_runtime()

    with _reserve_paths([target_request], overwrite) as targets:
        target = targets[0]
        with _temporary_directory(
            target.parent, ".docuforge-ppt-master-"
        ) as temporary_name:
            staged = Path(temporary_name) / f"source{source_path.suffix}"
            shutil.copyfile(source_path, staged)
            with _powerpoint_presentation(staged, runtime=runtime, read_only=False) as (
                _,
                presentation,
            ):
                masters = _presentation_masters(presentation)
                for master in masters:
                    if office_color is not None:
                        master.Background.Fill.Solid()
                        master.Background.Fill.ForeColor.RGB = office_color
                    if normalized_font is not None:
                        _set_theme_fonts(master, normalized_font)
                        for shape in _iter_com_collection(master.Shapes):
                            _set_shape_font(shape, normalized_font)
                        for layout in _iter_com_collection(master.CustomLayouts):
                            for shape in _iter_com_collection(layout.Shapes):
                                _set_shape_font(shape, normalized_font)
                if normalized_footer is not None:
                    for master in masters:
                        try:
                            master.HeadersFooters.Footer.Visible = True
                            master.HeadersFooters.Footer.Text = normalized_footer
                        except Exception:
                            # Some older PowerPoint versions expose footer
                            # placeholders only through individual slides.
                            pass
                    for slide in _iter_com_collection(presentation.Slides):
                        slide.HeadersFooters.Footer.Visible = True
                        slide.HeadersFooters.Footer.Text = normalized_footer
                presentation.Save()
            _publish_file(staged, target)
        return [target]


def _worksheet_by_name(workbook: Any, name: str) -> Any | None:
    for worksheet in _iter_com_collection(workbook.Worksheets):
        if str(worksheet.Name).casefold() == name.casefold():
            return worksheet
    return None


def _range_headers(source_range: Any) -> tuple[str, ...]:
    headers = []
    for column in range(1, int(source_range.Columns.Count) + 1):
        value = source_range.Cells.Item(1, column).Value
        header = "" if value is None else str(value).strip()
        if not header:
            raise ValidationError("source_range 的标题行不能包含空单元格")
        headers.append(header)
    folded = [header.casefold() for header in headers]
    if len(folded) != len(set(folded)):
        raise ValidationError("source_range 的标题行不能包含重复字段名")
    return tuple(headers)


def _validate_requested_fields(
    requested: Sequence[str], headers: Sequence[str], name: str
) -> None:
    available = {header.casefold() for header in headers}
    missing = [field for field in requested if field.casefold() not in available]
    if missing:
        raise ValidationError(f"{name} 字段不存在：{missing[0]}")


def excel_create_pivot(
    source: PathLike,
    output_dir: PathLike,
    source_sheet: str,
    source_range: str,
    target_sheet: str = "数据透视表",
    target_cell: str = "A1",
    row_fields: Sequence[str] = (),
    column_fields: Sequence[str] = (),
    data_field: str | None = None,
    function: str = "sum",
    overwrite: bool = False,
) -> list[Path]:
    """Create and configure a native Excel PivotTable."""

    source_path = _require_source(source, _EXCEL_EXTENSIONS, "Excel")
    source_sheet_name = _validate_sheet_name(source_sheet, "source_sheet")
    source_range_name = _normalize_a1_range(source_range)
    target_sheet_name = _validate_sheet_name(target_sheet, "target_sheet")
    target_cell_name = _normalize_a1_cell(target_cell)
    rows = _normalize_field_names(row_fields, "row_fields")
    columns = _normalize_field_names(column_fields, "column_fields")
    normalized_data_field = None
    if data_field is not None:
        normalized_data_field = str(data_field).strip()
        if not normalized_data_field:
            raise ValidationError("data_field 不能为空")
    function_code = _normalize_pivot_function(function)
    if not rows and not columns and normalized_data_field is None:
        raise ValidationError("至少需要一个行字段、列字段或数据字段")
    overlap = {field.casefold() for field in rows} & {
        field.casefold() for field in columns
    }
    if overlap:
        raise ValidationError("同一字段不能同时作为行字段和列字段")
    if source_sheet_name.casefold() == target_sheet_name.casefold():
        first_col, first_row, last_col, last_row = _a1_range_bounds(source_range_name)
        target_col, target_row = _a1_cell_position(target_cell_name)
        if first_col <= target_col <= last_col and first_row <= target_row <= last_row:
            raise ValidationError("数据透视表目标单元格不能位于源数据区域内")

    target_request = _directory_target(source_path, output_dir, tag="数据透视表")
    runtime = _load_com_runtime()
    with _reserve_paths([target_request], overwrite) as targets:
        target = targets[0]
        with _temporary_directory(
            target.parent, ".docuforge-excel-pivot-"
        ) as temporary_name:
            staged = Path(temporary_name) / f"source{source_path.suffix}"
            shutil.copyfile(source_path, staged)
            with _excel_workbook(staged, runtime=runtime) as (_, workbook):
                source_worksheet = _worksheet_by_name(workbook, source_sheet_name)
                if source_worksheet is None:
                    raise ValidationError(f"source_sheet 不存在：{source_sheet_name}")
                data_range = source_worksheet.Range(source_range_name)
                headers = _range_headers(data_range)
                _validate_requested_fields(rows, headers, "row_fields")
                _validate_requested_fields(columns, headers, "column_fields")
                if normalized_data_field is not None:
                    _validate_requested_fields(
                        (normalized_data_field,), headers, "data_field"
                    )

                target_worksheet = _worksheet_by_name(workbook, target_sheet_name)
                if target_worksheet is None:
                    target_worksheet = workbook.Worksheets.Add(
                        After=workbook.Worksheets.Item(workbook.Worksheets.Count)
                    )
                    target_worksheet.Name = target_sheet_name
                destination = target_worksheet.Range(target_cell_name)
                if destination.Value not in (None, "") or destination.Formula not in (
                    None,
                    "",
                ):
                    raise ValidationError(
                        f"目标单元格已有内容：{target_sheet_name}!{target_cell_name}"
                    )

                source_address = data_range.Address(
                    RowAbsolute=True,
                    ColumnAbsolute=True,
                    ReferenceStyle=1,  # xlA1
                    External=True,
                )
                cache = workbook.PivotCaches().Create(
                    SourceType=1, SourceData=source_address  # xlDatabase
                )
                table_name = f"LayoutLoomPivot_{int(time.time() * 1000)}"
                pivot = cache.CreatePivotTable(
                    TableDestination=destination,
                    TableName=table_name,
                )
                pivot.ManualUpdate = True
                for position, field_name in enumerate(rows, 1):
                    field = pivot.PivotFields(field_name)
                    field.Orientation = 1  # xlRowField
                    field.Position = position
                for position, field_name in enumerate(columns, 1):
                    field = pivot.PivotFields(field_name)
                    field.Orientation = 2  # xlColumnField
                    field.Position = position
                if normalized_data_field is not None:
                    pivot.AddDataField(
                        pivot.PivotFields(normalized_data_field),
                        f"{function}_{normalized_data_field}",
                        function_code,
                    )
                pivot.ManualUpdate = False
                pivot.RefreshTable()
                workbook.Save()
            _publish_file(staged, target)
        return [target]


def _word_ooxml_preflight(source: Path) -> WordBlankPageSafety:
    try:
        with zipfile.ZipFile(source) as archive:
            xml = archive.read("word/document.xml")
            try:
                settings_xml = archive.read("word/settings.xml")
            except KeyError:
                settings_xml = None
    except (zipfile.BadZipFile, KeyError, OSError) as exc:
        raise ValidationError("Word 文件不是可安全检查的有效 OOXML 文档") from exc
    report = _analyze_word_document_xml(xml)
    settings_reasons = (
        () if settings_xml is None else _word_settings_risks(settings_xml)
    )
    return WordBlankPageSafety(
        explicit_page_breaks=report.explicit_page_breaks,
        high_risk_reasons=tuple(
            dict.fromkeys((*report.high_risk_reasons, *settings_reasons))
        ),
    )


def _range_has_objects(document: Any, page_range: Any) -> bool:
    for collection_name in (
        "Tables",
        "InlineShapes",
        "Fields",
        "FormFields",
        "ContentControls",
    ):
        try:
            if int(getattr(page_range, collection_name).Count) > 0:
                return True
        except Exception:
            # Failure to inspect a collection is not proof that the page is
            # empty, so keep the page rather than making a risky deletion.
            return True
    try:
        start, end = int(page_range.Start), int(page_range.End)
        for shape in _iter_com_collection(document.Shapes):
            anchor = int(shape.Anchor.Start)
            if start <= anchor < end:
                return True
    except Exception:
        pass
    return False


def _word_page_range(document: Any, page_number: int, page_count: int) -> Any:
    start = int(document.GoTo(What=1, Which=1, Count=page_number).Start)
    if page_number < page_count:
        end = int(document.GoTo(What=1, Which=1, Count=page_number + 1).Start)
    else:
        end = int(document.Content.End)
    return document.Range(Start=start, End=end)


def _delete_one_manual_break(document: Any, page_range: Any) -> bool:
    start, end = int(page_range.Start), int(page_range.End)
    for position in range(end - 1, start - 1, -1):
        character = document.Range(Start=position, End=position + 1)
        if str(character.Text) == "\x0c":
            character.Delete()
            return True
    # A trailing manual break belongs to the preceding page in some Word
    # pagination results, while the empty page starts immediately after it.
    if start > 0:
        character = document.Range(Start=start - 1, End=start)
        if str(character.Text) == "\x0c":
            character.Delete()
            return True
    return False


def _word_page_is_explicit_blank(document: Any, page_range: Any) -> bool:
    if not _is_blank_word_page_text(page_range.Text):
        return False
    text = "" if page_range.Text is None else str(page_range.Text)
    if "\x0c" in text:
        return True
    start = int(page_range.Start)
    if start <= 0:
        return False
    preceding = document.Range(Start=start - 1, End=start)
    return str(preceding.Text) == "\x0c"


def word_remove_blank_pages(
    source: PathLike,
    output_dir: PathLike,
    overwrite: bool = False,
) -> list[Path]:
    """Remove only blank pages caused by explicit manual page breaks.

    Word documents with section breaks, paragraph-level pagination, floating
    objects, revisions, or similarly ambiguous structures are rejected instead
    of risking destructive layout edits.
    """

    source_path = _require_source(source, _WORD_OOXML_EXTENSIONS, "Word OOXML")
    runtime = _load_com_runtime()
    safety = _word_ooxml_preflight(source_path)
    if not safety.safe:
        raise ValidationError(
            "为避免误删，拒绝处理此 Word 文档：" + "；".join(safety.high_risk_reasons)
        )
    target_request = _directory_target(source_path, output_dir, tag="无空白页")

    with _reserve_paths([target_request], overwrite) as targets:
        target = targets[0]
        with _temporary_directory(
            target.parent, ".docuforge-word-pages-"
        ) as temporary_name:
            staged = Path(temporary_name) / f"source{source_path.suffix}"
            shutil.copyfile(source_path, staged)
            with _word_document(staged, runtime=runtime) as (_, document):
                # wdStatisticPages=2; repaginate before inspecting native page ranges.
                document.Repaginate()
                page_count = int(document.ComputeStatistics(2))
                candidates: list[int] = []
                for page_number in range(1, page_count + 1):
                    page_range = _word_page_range(document, page_number, page_count)
                    if _word_page_is_explicit_blank(
                        document, page_range
                    ) and not _range_has_objects(document, page_range):
                        candidates.append(page_number)

                for page_number in reversed(candidates):
                    current_count = int(document.ComputeStatistics(2))
                    if page_number > current_count:
                        continue
                    page_range = _word_page_range(document, page_number, current_count)
                    if not _word_page_is_explicit_blank(document, page_range):
                        continue
                    if _range_has_objects(document, page_range):
                        raise ValidationError(
                            f"第 {page_number} 页包含非文本对象，拒绝自动删除"
                        )
                    if not _delete_one_manual_break(document, page_range):
                        raise ValidationError(
                            f"第 {page_number} 页无法安全定位手动分页符"
                        )
                    document.Repaginate()
                    if int(document.ComputeStatistics(2)) >= current_count:
                        raise ValidationError(
                            f"删除第 {page_number} 页的分页符后页数未减少，拒绝发布结果"
                        )
                document.Save()
            _publish_file(staged, target)
        return [target]


__all__ = [
    "WordBlankPageSafety",
    "excel_create_pivot",
    "ppt_modify_master",
    "ppt_to_images",
    "ppt_to_long_image",
    "ppt_to_video",
    "word_remove_blank_pages",
]
