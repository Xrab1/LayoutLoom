from __future__ import annotations

import os
import shutil
import sys
import math
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Literal, Mapping, Sequence

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


def _windows_wps_process_command_line(process_id: int) -> str | None:
    """Return one WPS process command line for strict automation ownership."""

    if sys.platform != "win32" or process_id <= 0:
        return None
    initialized = False
    service = None
    rows = None
    row = None
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        initialized = True
        service = win32com.client.GetObject("winmgmts:")
        rows = list(
            service.ExecQuery(
                "SELECT CommandLine FROM Win32_Process "
                f"WHERE ProcessId={int(process_id)}"
            )
        )
        row = rows[0] if rows else None
        value = getattr(row, "CommandLine", None) if row is not None else None
        return str(value) if value else None
    except Exception:
        return None
    finally:
        row = None
        rows = None
        service = None
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _new_owned_wps_processes(
    before: Mapping[int, object],
    *,
    expected_executable: Path,
    reported_pid: int | None = None,
    reported_pids: Sequence[int] | None = None,
    reported_identities: Sequence[object] | None = None,
    require_automation: bool = False,
    snapshot: Callable[[str], Mapping[int, object]] | None = None,
    command_line: Callable[[int], str | None] | None = None,
) -> tuple[object, ...]:
    """Return the compact WPS process group created by one DispatchEx call.

    Current WPS Writer builds commonly launch a broker plus a worker while the
    hidden COM application exposes no HWND.  The generic Microsoft Office
    helper deliberately accepts only one new process, which made LayoutLoom
    misclassify this fresh pair as a reused user instance and skip ``Quit``.
    Accept at most four exact-path processes created in one five-second burst;
    otherwise refuse ownership instead of risking a user's existing WPS.
    """

    if snapshot is None:
        from .office import _windows_process_snapshot

        snapshot = _windows_process_snapshot
    expected_path = expected_executable.expanduser().resolve()
    after = snapshot(expected_path.name)
    candidates = [
        identity
        for process_id, identity in after.items()
        if identity != before.get(process_id)
        and getattr(identity, "executable", None) == expected_path
    ]
    if require_automation:
        lookup = command_line or _windows_wps_process_command_line
        candidates = [
            identity
            for identity in candidates
            if (
                (line := lookup(int(getattr(identity, "pid", 0)))) is not None
                and "/automation" in line.casefold()
                and "-embedding" in line.casefold()
            )
        ]
    if not candidates or len(candidates) > 4:
        return ()
    if reported_pid is not None and not any(
        int(getattr(identity, "pid", 0)) == int(reported_pid) for identity in candidates
    ):
        return ()
    if reported_pids is not None:
        try:
            expected_pids = {int(process_id) for process_id in reported_pids}
        except (TypeError, ValueError):
            return ()
        candidate_pids = {int(getattr(identity, "pid", 0)) for identity in candidates}
        if not expected_pids or candidate_pids != expected_pids:
            return ()
    if reported_identities is not None:
        reported_by_pid: dict[int, tuple[Path, str]] = {}
        for identity in reported_identities:
            try:
                process_id = int(getattr(identity, "pid"))
                executable = (
                    Path(getattr(identity, "executable")).expanduser().resolve()
                )
                created = str(getattr(identity, "created"))
            except (AttributeError, OSError, TypeError, ValueError):
                return ()
            if (
                process_id <= 0
                or not created
                or executable != expected_path
                or process_id in reported_by_pid
            ):
                return ()
            reported_by_pid[process_id] = (executable, created)
        candidates_by_pid = {
            int(getattr(identity, "pid", 0)): identity for identity in candidates
        }
        if not reported_by_pid or set(candidates_by_pid) != set(reported_by_pid):
            return ()
        for process_id, (
            reported_executable,
            reported_created,
        ) in reported_by_pid.items():
            identity = candidates_by_pid[process_id]
            try:
                current_executable = (
                    Path(getattr(identity, "executable")).expanduser().resolve()
                )
                current_created = str(getattr(identity, "created"))
            except (AttributeError, OSError, TypeError, ValueError):
                return ()
            if (
                current_executable != reported_executable
                or current_created != reported_created
            ):
                return ()
    created_values: list[float] = []
    for identity in candidates:
        try:
            created_values.append(
                datetime.fromisoformat(str(getattr(identity, "created"))).timestamp()
            )
        except (TypeError, ValueError):
            if len(candidates) > 1:
                return ()
    if created_values and max(created_values) - min(created_values) > 5.0:
        return ()
    return tuple(sorted(candidates, key=lambda item: int(getattr(item, "pid", 0))))


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
    _ownership_reporter: Callable[[Sequence[object]], bool] | None = None,
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
    application_owned = True
    original_application_settings: dict[str, object] = {}
    owned_processes: list[object] = []
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
            _terminate_owned_office_process,
            _wait_for_owned_office_exit,
            _windows_process_identity,
            _windows_process_snapshot,
        )

        process_helpers = (
            _office_application_pid,
            _wait_for_owned_office_exit,
            _windows_process_identity,
            _terminate_owned_office_process,
        )
        before_processes = dict(_windows_process_snapshot(expected_process))
    except Exception:
        process_helpers = None

    def release_application() -> None:
        nonlocal application
        if application is None:
            return
        if application_owned:
            try:
                application.Quit()
            except Exception:
                pass
        else:
            for property_name, original_value in original_application_settings.items():
                try:
                    setattr(application, property_name, original_value)
                except Exception:
                    pass
        application = None

    def terminate_owned_processes() -> None:
        if process_helpers is None:
            return
        terminate = process_helpers[3]
        for identity in reversed(tuple(owned_processes)):
            try:
                terminate(identity)
            except Exception:
                pass

    try:
        pythoncom.CoInitialize()  # type: ignore[attr-defined]
        initialized = True
        application = win32_client.DispatchEx(status.prog_id)  # type: ignore[attr-defined]
        status_executable = getattr(status, "executable", None)
        if status_executable is not None:
            try:
                expected_executable = Path(status_executable).expanduser().resolve()
            except OSError:
                expected_executable = Path(status_executable).expanduser().absolute()
        else:
            try:
                expected_executable = (
                    Path(str(application.Path)).expanduser().resolve()
                    / expected_process
                ).resolve()
            except Exception:
                expected_executable = None
        ownership_required = process_helpers is not None
        application_owned = not ownership_required
        if process_helpers is not None:
            application_pid = process_helpers[0](application)
            if ownership_required and expected_executable is not None:
                ownership_deadline = time.monotonic() + 1.0
                identities: tuple[object, ...] = ()
                while not identities:
                    identities = _new_owned_wps_processes(
                        before_processes,
                        expected_executable=expected_executable,
                        reported_pid=application_pid,
                    )
                    if identities or time.monotonic() >= ownership_deadline:
                        break
                    time.sleep(0.05)
            else:
                identities = ()
            if ownership_required and not identities:
                existing_identity = (
                    process_helpers[2](application_pid)
                    if application_pid is not None
                    else None
                )
                expected_path = expected_executable
                path_confirmed = bool(
                    existing_identity is not None
                    and expected_path is not None
                    and existing_identity.executable == expected_path
                )
                if not path_confirmed and expected_path is not None:
                    try:
                        application_executable = (
                            Path(str(application.Path)).expanduser().resolve()
                            / expected_process
                        ).resolve()
                    except Exception:
                        application_executable = None
                    path_confirmed = application_executable == expected_path
                collection_name = {
                    "writer": "Documents",
                    "spreadsheets": "Workbooks",
                    "presentation": "Presentations",
                }[kind]
                try:
                    open_documents = int(getattr(application, collection_name).Count)
                except Exception:
                    open_documents = -1
                if not path_confirmed or open_documents != 0:
                    raise MissingEngineError(
                        "WPS 复用了已有实例，且无法确认该实例处于无文档空闲状态；"
                        "为保护用户已打开的文档，已停止自动化"
                    )
                application_owned = False
            if identities:
                approved = True
                if _ownership_reporter is not None:
                    try:
                        approved = bool(_ownership_reporter(tuple(identities)))
                    except Exception:
                        approved = False
                if not approved:
                    application_owned = False
                    raise MissingEngineError(
                        "WPS 新实例的进程所有权未获监督进程确认；"
                        "为保护用户正在使用的 WPS，已停止自动化"
                    )
                owned_processes.extend(identities)
                application_owned = True
                try:
                    cancel_process_guard = cancellation_callback(
                        terminate_owned_processes
                    )
                    cancel_process_guard.__enter__()
                except Exception:
                    cancel_process_guard = None
        if not application_owned:
            for property_name in (
                "AutomationSecurity",
                "Visible",
                "DisplayAlerts",
                "ScreenUpdating",
                "EnableEvents",
                "AskToUpdateLinks",
            ):
                try:
                    original_application_settings[property_name] = getattr(
                        application, property_name
                    )
                except Exception:
                    pass
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
                document = application.Documents.Open(
                    str(source_path),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Revert=False,
                    Visible=False,
                    NoEncodingDialog=True,
                )
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
            release_application()
            if initialized:
                pythoncom.CoUninitialize()  # type: ignore[attr-defined]
                initialized = False
            if owned_processes and process_helpers is not None:
                for identity in tuple(owned_processes):
                    process_helpers[1](identity, timeout=1.0)
                owned_processes.clear()
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
        release_application()
        if initialized:
            pythoncom.CoUninitialize()  # type: ignore[attr-defined]
        if owned_processes and process_helpers is not None:
            for identity in tuple(owned_processes):
                process_helpers[1](identity, timeout=1.0)
            owned_processes.clear()
    return [target]


def _send_wps_worker_message(connection: object, message: Mapping[str, Any]) -> None:
    try:
        connection.send(dict(message))  # type: ignore[attr-defined]
    except (BrokenPipeError, EOFError, OSError):
        pass


def _convert_wps_worker_entry(
    connection: object,
    source: str,
    output_dir: str,
    target_format: str,
    options: Mapping[str, Any],
) -> None:
    """Spawn-safe worker for one isolated WPS COM conversion."""

    def report_processes(identities: Sequence[object]) -> bool:
        process_ids = [int(getattr(identity, "pid", 0)) for identity in identities]
        identity_payloads: list[dict[str, object]] = []
        for identity in identities:
            try:
                executable = (
                    Path(getattr(identity, "executable")).expanduser().resolve()
                )
                created = str(getattr(identity, "created"))
                process_id = int(getattr(identity, "pid"))
            except (AttributeError, OSError, TypeError, ValueError):
                return False
            if process_id <= 0 or not created:
                return False
            identity_payloads.append(
                {
                    "pid": process_id,
                    "executable": str(executable),
                    "created": created,
                }
            )
        _send_wps_worker_message(
            connection,
            {
                "type": "wps_processes",
                "pids": process_ids,
                "identities": identity_payloads,
            },
        )
        try:
            if not connection.poll(15.0):  # type: ignore[attr-defined]
                return False
            response = connection.recv()  # type: ignore[attr-defined]
        except (EOFError, OSError):
            return False
        if not (
            isinstance(response, Mapping)
            and response.get("type") == "ownership"
            and response.get("approved") is True
        ):
            return False
        raw_pids = response.get("pids")
        if not isinstance(raw_pids, Sequence):
            return False
        try:
            approved_pids = {int(process_id) for process_id in raw_pids}
        except (TypeError, ValueError):
            return False
        return approved_pids == set(process_ids)

    try:
        outputs = convert_with_wps(
            source,
            output_dir,
            target_format,
            overwrite=bool(options.get("overwrite", False)),
            excel_pdf_layout=str(options.get("excel_pdf_layout", "smart")),
            excel_pdf_paper=str(options.get("excel_pdf_paper", "auto")),
            excel_pdf_orientation=str(options.get("excel_pdf_orientation", "auto")),
            excel_pdf_margin=str(options.get("excel_pdf_margin", "auto")),
            _ownership_reporter=report_processes,
        )
    except ValidationError as exc:
        _send_wps_worker_message(
            connection,
            {"type": "result", "ok": False, "kind": "validation", "error": str(exc)},
        )
    except FileExistsError as exc:
        _send_wps_worker_message(
            connection,
            {"type": "result", "ok": False, "kind": "exists", "error": str(exc)},
        )
    except MissingEngineError as exc:
        _send_wps_worker_message(
            connection,
            {"type": "result", "ok": False, "kind": "engine", "error": str(exc)},
        )
    except Exception as exc:  # pragma: no cover - defensive worker boundary
        _send_wps_worker_message(
            connection,
            {
                "type": "result",
                "ok": False,
                "kind": "engine",
                "error": f"WPS 隔离转换子进程异常：{exc}",
            },
        )
    else:
        _send_wps_worker_message(
            connection,
            {"type": "result", "ok": True, "outputs": [str(path) for path in outputs]},
        )
    finally:
        try:
            connection.close()  # type: ignore[attr-defined]
        except OSError:
            pass


def _stop_wps_worker(process: object, identities: Sequence[object]) -> None:
    """Stop only this worker and the exact WPS process group it created."""

    from .office import _terminate_owned_office_process, _wait_for_owned_office_exit

    for identity in reversed(tuple(identities)):
        try:
            _terminate_owned_office_process(identity)
        except Exception:
            pass
    try:
        process.join(1.0)  # type: ignore[attr-defined]
    except Exception:
        pass
    try:
        alive = bool(process.is_alive())  # type: ignore[attr-defined]
    except Exception:
        alive = False
    if alive:
        try:
            process.terminate()  # type: ignore[attr-defined]
        except Exception:
            pass
        try:
            process.join(1.0)  # type: ignore[attr-defined]
        except Exception:
            pass
    try:
        alive = bool(process.is_alive())  # type: ignore[attr-defined]
    except Exception:
        alive = False
    if alive and hasattr(process, "kill"):
        try:
            process.kill()  # type: ignore[attr-defined]
            process.join(1.0)  # type: ignore[attr-defined]
        except Exception:
            pass
    for identity in tuple(identities):
        try:
            _wait_for_owned_office_exit(identity, timeout=0.5)
        except Exception:
            pass


def _cleanup_wps_temporary_outputs(
    output_dir: Path,
    existing: set[Path],
    *,
    source_stem: str,
) -> None:
    """Remove only new atomic temp files created for this source."""

    try:
        candidates = tuple(path for path in output_dir.iterdir() if path.is_file())
    except OSError:
        return
    exact_prefix = f".{source_stem}."
    numbered_prefix = f".{source_stem}_"
    for path in candidates:
        try:
            resolved = path.resolve()
            name = path.name
            if resolved in existing or ".tmp" not in name.casefold():
                continue
            if not (name.startswith(exact_prefix) or name.startswith(numbered_prefix)):
                continue
            for attempt in range(11):
                try:
                    path.unlink(missing_ok=True)
                    break
                except PermissionError:
                    if attempt >= 10:
                        break
                    time.sleep(0.05 + attempt * 0.03)
        except OSError:
            continue


def convert_with_wps_supervised(
    source: PathLike,
    output_dir: PathLike,
    target_format: str = "pdf",
    *,
    overwrite: bool = False,
    timeout: float = 90,
    excel_pdf_layout: str = "smart",
    excel_pdf_paper: str = "auto",
    excel_pdf_orientation: str = "auto",
    excel_pdf_margin: str = "auto",
) -> list[Path]:
    """Run WPS COM in a spawned worker with enforceable timeout and cleanup."""

    import multiprocessing

    from docuforge.runner import check_cancelled
    from .office import (
        _OfficeProcessIdentity,
        _wait_for_owned_office_exit,
        _windows_process_snapshot,
    )

    check_cancelled("任务已取消；已完成的文件会保留")
    source_path = Path(source).expanduser().resolve()
    if not source_path.is_file():
        raise ValidationError(f"文件不存在：{source_path}")
    format_name = target_format.lower().lstrip(".")
    if not format_name.isalnum():
        raise ValidationError(f"无效目标格式：{target_format}")
    try:
        timeout_value = float(timeout)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValidationError("timeout 必须是有限且大于 0 的秒数") from exc
    if not math.isfinite(timeout_value) or timeout_value <= 0:
        raise ValidationError("timeout 必须是有限且大于 0 的秒数")

    kind = _kind_for_source(source_path)
    status = detect_wps_engines()[kind]
    if not status.available:
        raise MissingEngineError(status.reason)
    directory = Path(output_dir).expanduser().resolve()
    directory.mkdir(parents=True, exist_ok=True)
    try:
        existing_files = {
            path.resolve() for path in directory.iterdir() if path.is_file()
        }
    except OSError:
        existing_files = set()

    status_executable = getattr(status, "executable", None)
    expected_executable = (
        Path(status_executable).expanduser().resolve()
        if status_executable is not None
        else None
    )
    before = _windows_process_snapshot(_EXECUTABLES[kind])
    context = multiprocessing.get_context("spawn")
    parent_connection, child_connection = context.Pipe(duplex=True)
    process = context.Process(
        target=_convert_wps_worker_entry,
        args=(
            child_connection,
            str(source_path),
            str(directory),
            format_name,
            {
                "overwrite": overwrite,
                "excel_pdf_layout": excel_pdf_layout,
                "excel_pdf_paper": excel_pdf_paper,
                "excel_pdf_orientation": excel_pdf_orientation,
                "excel_pdf_margin": excel_pdf_margin,
            },
        ),
        name=f"docuforge-wps-{kind}",
        daemon=False,
    )
    try:
        process.start()
    except Exception as exc:
        parent_connection.close()
        child_connection.close()
        raise MissingEngineError(f"无法启动 WPS 隔离转换进程：{exc}") from exc
    child_connection.close()

    deadline = time.monotonic() + timeout_value
    owned_processes: tuple[object, ...] = ()
    owned_executable = expected_executable
    reported_process_ids: tuple[int, ...] = ()
    reported_process_identities: tuple[object, ...] = ()
    result: Mapping[str, Any] | None = None

    def refresh_reported_ownership() -> tuple[object, ...]:
        nonlocal owned_processes
        if owned_executable is None or not reported_process_identities:
            return owned_processes
        candidates = _new_owned_wps_processes(
            before,
            expected_executable=owned_executable,
            reported_pids=reported_process_ids,
            reported_identities=reported_process_identities,
            require_automation=True,
        )
        if candidates:
            owned_processes = candidates
        return owned_processes

    try:
        while result is None:
            check_cancelled("任务已取消；正在终止 WPS 转换")
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise MissingEngineError(f"WPS 转换超时（{timeout_value:g} 秒）")
            try:
                has_message = parent_connection.poll(min(0.1, remaining))
            except (EOFError, OSError):
                has_message = False
            if has_message:
                try:
                    message = parent_connection.recv()
                except (EOFError, OSError):
                    message = None
                if isinstance(message, Mapping):
                    message_type = message.get("type")
                    if message_type == "wps_processes":
                        raw_identities = message.get("identities")
                        parsed_identities: list[object] = []
                        if isinstance(raw_identities, Sequence) and not isinstance(
                            raw_identities, (str, bytes, bytearray)
                        ):
                            for raw_identity in raw_identities:
                                if not isinstance(raw_identity, Mapping):
                                    parsed_identities = []
                                    break
                                try:
                                    process_id = int(raw_identity.get("pid") or 0)
                                    reported_path = str(
                                        raw_identity.get("executable") or ""
                                    ).strip()
                                    created = str(
                                        raw_identity.get("created") or ""
                                    ).strip()
                                    reported_executable = (
                                        Path(reported_path).expanduser().resolve()
                                    )
                                except (OSError, TypeError, ValueError):
                                    parsed_identities = []
                                    break
                                if process_id <= 0 or not reported_path or not created:
                                    parsed_identities = []
                                    break
                                parsed_identities.append(
                                    _OfficeProcessIdentity(
                                        process_id, reported_executable, created
                                    )
                                )
                        identity_paths = {
                            getattr(identity, "executable", None)
                            for identity in parsed_identities
                        }
                        parsed_pids = tuple(
                            sorted(
                                int(getattr(identity, "pid", 0))
                                for identity in parsed_identities
                            )
                        )
                        if (
                            parsed_identities
                            and len(set(parsed_pids)) == len(parsed_pids)
                            and len(identity_paths) == 1
                            and next(iter(identity_paths)).name.casefold()
                            == _EXECUTABLES[kind].casefold()
                            and (
                                expected_executable is None
                                or next(iter(identity_paths)) == expected_executable
                            )
                        ):
                            owned_executable = next(iter(identity_paths))
                            reported_process_ids = parsed_pids
                            reported_process_identities = tuple(
                                sorted(
                                    parsed_identities,
                                    key=lambda identity: int(
                                        getattr(identity, "pid", 0)
                                    ),
                                )
                            )
                        else:
                            reported_process_ids = ()
                            reported_process_identities = ()
                        approved = bool(refresh_reported_ownership())
                        try:
                            parent_connection.send(
                                {
                                    "type": "ownership",
                                    "approved": approved,
                                    "pids": (
                                        list(reported_process_ids) if approved else []
                                    ),
                                }
                            )
                        except (BrokenPipeError, EOFError, OSError):
                            pass
                    elif message_type == "result":
                        result = message
            if result is None and not process.is_alive():
                try:
                    if parent_connection.poll(0):
                        final_message = parent_connection.recv()
                        if (
                            isinstance(final_message, Mapping)
                            and final_message.get("type") == "result"
                        ):
                            result = final_message
                except (EOFError, OSError):
                    pass
                if result is None:
                    break
    except BaseException:
        # Only a worker-reported, fully revalidated Automation/Embedding PID
        # set may be reclaimed.  A same-path WPS opened manually by the user
        # during this window is never promoted to task ownership by scanning.
        refresh_reported_ownership()
        if process.is_alive() or owned_processes:
            _stop_wps_worker(process, owned_processes)
        _cleanup_wps_temporary_outputs(
            directory, existing_files, source_stem=source_path.stem
        )
        raise
    finally:
        try:
            parent_connection.close()
        except OSError:
            pass

    try:
        process.join(2.0)
    except Exception:
        pass
    if process.is_alive():
        refresh_reported_ownership()
        _stop_wps_worker(process, owned_processes)
        _cleanup_wps_temporary_outputs(
            directory, existing_files, source_stem=source_path.stem
        )
        raise MissingEngineError("WPS 隔离转换子进程未正常退出")
    for identity in owned_processes:
        _wait_for_owned_office_exit(identity, timeout=1.0)
    if result is None:
        refresh_reported_ownership()
        if owned_processes:
            _stop_wps_worker(process, owned_processes)
        _cleanup_wps_temporary_outputs(
            directory, existing_files, source_stem=source_path.stem
        )
        raise MissingEngineError("WPS 隔离转换进程意外退出")
    if bool(result.get("ok")):
        outputs = [
            Path(str(path)).expanduser().resolve() for path in result.get("outputs", [])
        ]
        invalid = next(
            (
                path
                for path in outputs
                if not path.is_file() or path.stat().st_size == 0
            ),
            None,
        )
        if not outputs or invalid is not None:
            _cleanup_wps_temporary_outputs(
                directory, existing_files, source_stem=source_path.stem
            )
            raise MissingEngineError(
                f"WPS 未生成有效输出文件{f'：{invalid}' if invalid is not None else ''}"
            )
        return outputs
    error = str(result.get("error") or "WPS 隔离转换失败")
    kind_name = str(result.get("kind") or "engine")
    _cleanup_wps_temporary_outputs(
        directory, existing_files, source_stem=source_path.stem
    )
    if kind_name == "validation":
        raise ValidationError(error)
    if kind_name == "exists":
        raise FileExistsError(error)
    raise MissingEngineError(error)
