from __future__ import annotations

import os
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from ..models import MissingEngineError, ValidationError
from ..utils import atomic_output, unique_path

PathLike = str | os.PathLike[str]
WpsKind = Literal["writer", "spreadsheets", "presentation"]


@dataclass(frozen=True)
class WpsEngineStatus:
    available: bool
    kind: WpsKind
    prog_id: str
    executable: Path | None = None
    reason: str = ""


_PROG_IDS: dict[WpsKind, tuple[str, ...]] = {
    "writer": ("KWPS.Application", "WPS.Application"),
    "spreadsheets": ("KET.Application",),
    "presentation": ("KWPP.Application",),
}

_EXECUTABLES: dict[WpsKind, str] = {
    "writer": "wps.exe",
    "spreadsheets": "et.exe",
    "presentation": "wpp.exe",
}


def _registered_prog_id(kind: WpsKind) -> str | None:
    if sys.platform != "win32":
        return None
    try:
        import winreg
    except ImportError:
        return None
    for prog_id in _PROG_IDS[kind]:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\CLSID"):
                return prog_id
        except OSError:
            continue
    return None


def _known_wps_roots() -> list[Path]:
    program_files = Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
    program_files_x86 = Path(
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    )
    local_app_data = Path(os.environ.get("LOCALAPPDATA", ""))
    roots = [
        program_files / "Kingsoft/WPS Office",
        program_files_x86 / "Kingsoft/WPS Office",
        local_app_data / "Kingsoft/WPS Office",
        program_files / "WPS Office",
        program_files_x86 / "WPS Office",
        local_app_data / "WPS Office",
    ]
    return [root for root in roots if str(root)]


def _find_wps_executable(kind: WpsKind) -> Path | None:
    executable_name = _EXECUTABLES[kind]
    found = shutil.which(executable_name)
    if found:
        return Path(found).resolve()
    candidates: list[Path] = []
    for root in _known_wps_roots():
        if not root.is_dir():
            continue
        candidates.extend(root.glob(f"office6/{executable_name}"))
        candidates.extend(root.glob(f"*/office6/{executable_name}"))
        candidates.extend(root.glob(f"*/*/office6/{executable_name}"))
    existing = [path.resolve() for path in candidates if path.is_file()]
    return (
        sorted(existing, key=lambda path: path.stat().st_mtime, reverse=True)[0]
        if existing
        else None
    )


def _pywin32_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def detect_wps_engines() -> dict[WpsKind, WpsEngineStatus]:
    statuses: dict[WpsKind, WpsEngineStatus] = {}
    if sys.platform != "win32":
        return {
            kind: WpsEngineStatus(
                False, kind, prog_ids[0], reason="WPS COM 仅支持 Windows"
            )
            for kind, prog_ids in _PROG_IDS.items()
        }
    if not _pywin32_available():
        has_wps = any(_find_wps_executable(kind) for kind in _PROG_IDS)
        reason = (
            "检测到 WPS，但缺少 pywin32" if has_wps else "未检测到 WPS，且缺少 pywin32"
        )
        return {
            kind: WpsEngineStatus(
                False,
                kind,
                prog_ids[0],
                executable=_find_wps_executable(kind),
                reason=reason,
            )
            for kind, prog_ids in _PROG_IDS.items()
        }
    for kind, prog_ids in _PROG_IDS.items():
        prog_id = _registered_prog_id(kind)
        executable = _find_wps_executable(kind)
        if prog_id:
            statuses[kind] = WpsEngineStatus(
                True,
                kind,
                prog_id,
                executable=executable,
                reason=f"已检测到 WPS {kind} 自动化接口",
            )
        else:
            statuses[kind] = WpsEngineStatus(
                False,
                kind,
                prog_ids[0],
                executable=executable,
                reason=(
                    "检测到 WPS 程序，但未注册 COM 自动化接口"
                    if executable
                    else "未检测到 WPS Office"
                ),
            )
    return statuses


def _kind_for_source(source: Path) -> WpsKind:
    suffix = source.suffix.lower()
    if suffix in {".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf", ".txt"}:
        return "writer"
    if suffix in {".xls", ".xlsx", ".xlsm", ".xltx", ".xltm", ".csv"}:
        return "spreadsheets"
    if suffix in {".ppt", ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"}:
        return "presentation"
    raise ValidationError(f"WPS 不支持该输入格式：{source.suffix}")


def _output_target(
    source: Path, output_dir: PathLike, target_format: str, overwrite: bool
) -> Path:
    directory = Path(output_dir).expanduser().resolve()
    directory.mkdir(parents=True, exist_ok=True)
    return unique_path(directory / f"{source.stem}.{target_format}", overwrite)


def _writer_convert(document: object, target: Path, target_format: str) -> None:
    if target_format == "pdf":
        try:
            document.ExportAsFixedFormat(str(target), 17)  # type: ignore[attr-defined]
        except Exception:
            document.SaveAs(str(target), 17)  # type: ignore[attr-defined]
        return
    formats = {"docx": (12, 16), "doc": (0,), "txt": (2,), "html": (10,)}
    if target_format not in formats:
        raise ValidationError(f"WPS Writer 不支持目标格式：{target_format}")
    last_error: Exception | None = None
    for format_code in formats[target_format]:
        try:
            document.SaveAs(str(target), format_code)  # type: ignore[attr-defined]
            return
        except Exception as exc:
            last_error = exc
    assert last_error is not None
    raise last_error


def _sheet_convert(
    workbook: object,
    target: Path,
    target_format: str,
    *,
    application: object | None = None,
    excel_pdf_layout: str = "smart",
    excel_pdf_paper: str = "auto",
    excel_pdf_orientation: str = "auto",
    excel_pdf_margin: str = "auto",
) -> None:
    if target_format == "pdf":
        from .excel_pdf_layout import prepare_excel_workbook_for_pdf

        prepare_excel_workbook_for_pdf(
            workbook,
            application,
            layout=excel_pdf_layout,
            paper=excel_pdf_paper,
            orientation=excel_pdf_orientation,
            margin=excel_pdf_margin,
        )
        workbook.ExportAsFixedFormat(0, str(target))  # type: ignore[attr-defined]
        return
    formats = {"xlsx": 51, "xls": 56, "csv": 6, "xml": 46, "txt": 42}
    if target_format not in formats:
        raise ValidationError(f"WPS Spreadsheets 不支持目标格式：{target_format}")
    workbook.SaveAs(str(target), formats[target_format])  # type: ignore[attr-defined]


def _presentation_convert(
    presentation: object, target: Path, target_format: str
) -> None:
    formats = {"pdf": 32, "pptx": 24, "ppt": 1}
    if target_format not in formats:
        raise ValidationError(f"WPS Presentation 不支持目标格式：{target_format}")
    presentation.SaveAs(str(target), formats[target_format])  # type: ignore[attr-defined]


def _load_pywin32() -> tuple[object, object]:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise MissingEngineError("WPS 自动化需要 pywin32") from exc
    return pythoncom, win32com.client


def _disable_automation_macros(application: object, kind: WpsKind) -> None:
    try:
        application.AutomationSecurity = 3  # type: ignore[attr-defined]
    except Exception as exc:
        raise MissingEngineError(
            f"无法为 WPS {kind} 强制禁用宏，已拒绝打开文件：{exc}"
        ) from exc


def convert_with_wps(
    source: PathLike,
    output_dir: PathLike,
    target_format: str = "pdf",
    *,
    overwrite: bool = False,
    excel_pdf_layout: str = "smart",
    excel_pdf_paper: str = "auto",
    excel_pdf_orientation: str = "auto",
    excel_pdf_margin: str = "auto",
) -> list[Path]:
    """Convert through WPS COM. WPS must expose its registered automation API."""

    from docuforge.runner import cancellation_callback, check_cancelled

    check_cancelled("任务已取消；已完成的文件会保留")
    source_path = Path(source).expanduser().resolve()
    if not source_path.is_file():
        raise ValidationError(f"文件不存在：{source_path}")
    format_name = target_format.lower().lstrip(".")
    if not format_name.isalnum():
        raise ValidationError(f"无效目标格式：{target_format}")
    kind = _kind_for_source(source_path)
    if kind == "spreadsheets" and format_name == "pdf":
        from .excel_pdf_layout import normalize_excel_pdf_options

        (
            excel_pdf_layout,
            excel_pdf_paper,
            excel_pdf_orientation,
            excel_pdf_margin,
        ) = normalize_excel_pdf_options(
            excel_pdf_layout,
            excel_pdf_paper,
            excel_pdf_orientation,
            excel_pdf_margin,
        )
    status = detect_wps_engines()[kind]
    if not status.available:
        raise MissingEngineError(status.reason)

    pythoncom, win32_client = _load_pywin32()

    target = _output_target(source_path, output_dir, format_name, overwrite)
    application = None
    document = None
    initialized = False
    owned_process = None
    process_helpers = None
    expected_process = {
        "writer": "wps.exe",
        "spreadsheets": "et.exe",
        "presentation": "wpp.exe",
    }[kind]
    before_processes: dict[int, object] = {}
    cancel_process_guard = None
    try:
        from .office import (
            _office_application_pid,
            _wait_for_owned_office_exit,
            _windows_process_identity,
            _windows_process_snapshot,
        )

        process_helpers = (
            _office_application_pid,
            _wait_for_owned_office_exit,
            _windows_process_identity,
        )
        before_processes = dict(_windows_process_snapshot(expected_process))
    except Exception:
        process_helpers = None
    try:
        pythoncom.CoInitialize()  # type: ignore[attr-defined]
        initialized = True
        application = win32_client.DispatchEx(status.prog_id)  # type: ignore[attr-defined]
        if process_helpers is not None:
            application_pid = process_helpers[0](application)
            identity = (
                process_helpers[2](application_pid)
                if application_pid is not None
                else None
            )
            if (
                identity is not None
                and identity.pid not in before_processes
                and identity.executable.name.casefold() == expected_process.casefold()
            ):
                owned_process = identity
                try:
                    from .office import _terminate_owned_office_process

                    cancel_process_guard = cancellation_callback(
                        lambda identity=identity: _terminate_owned_office_process(
                            identity
                        )
                    )
                    cancel_process_guard.__enter__()
                except Exception:
                    cancel_process_guard = None
        _disable_automation_macros(application, kind)
        try:
            application.Visible = False
        except Exception:
            pass
        try:
            application.DisplayAlerts = 0
        except Exception:
            pass
        for property_name, value in (
            ("ScreenUpdating", False),
            ("EnableEvents", False),
            ("AskToUpdateLinks", False),
        ):
            try:
                setattr(application, property_name, value)
            except Exception:
                pass
        with atomic_output(target) as temporary:
            temporary.unlink(missing_ok=True)
            if kind == "writer":
                document = application.Documents.Open(str(source_path), ReadOnly=True)
                _writer_convert(document, temporary, format_name)
            elif kind == "spreadsheets":
                document = application.Workbooks.Open(
                    str(source_path),
                    UpdateLinks=0,
                    ReadOnly=True,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                    Notify=False,
                )
                _sheet_convert(
                    document,
                    temporary,
                    format_name,
                    application=application,
                    excel_pdf_layout=excel_pdf_layout,
                    excel_pdf_paper=excel_pdf_paper,
                    excel_pdf_orientation=excel_pdf_orientation,
                    excel_pdf_margin=excel_pdf_margin,
                )
            else:
                document = application.Presentations.Open(
                    str(source_path), WithWindow=False, ReadOnly=True
                )
                _presentation_convert(document, temporary, format_name)
            # SaveAs changes the live WPS document to the temporary output.
            # Close the document and application before atomic_output attempts
            # to rename it, otherwise Windows can report sharing violation 32.
            if document is not None:
                if kind == "presentation":
                    document.Close()
                else:
                    document.Close(False)
                document = None
            if application is not None:
                application.Quit()
                application = None
            if initialized:
                pythoncom.CoUninitialize()  # type: ignore[attr-defined]
                initialized = False
            if owned_process is not None and process_helpers is not None:
                process_helpers[1](owned_process, timeout=1.0)
                owned_process = None
            if not temporary.is_file() or temporary.stat().st_size == 0:
                raise MissingEngineError("WPS 未生成有效输出文件")
    except (ValidationError, MissingEngineError):
        raise
    except Exception as exc:
        raise MissingEngineError(f"WPS 转换失败：{exc}") from exc
    finally:
        if cancel_process_guard is not None:
            try:
                cancel_process_guard.__exit__(None, None, None)
            except Exception:
                pass
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass
        if initialized:
            pythoncom.CoUninitialize()  # type: ignore[attr-defined]
        if owned_process is not None and process_helpers is not None:
            process_helpers[1](owned_process, timeout=1.0)
    return [target]
