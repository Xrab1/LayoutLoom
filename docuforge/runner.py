from __future__ import annotations

import threading
import time
import warnings
from contextlib import contextmanager
from contextvars import ContextVar
from pathlib import Path
from typing import Any, Callable, Iterator, Mapping, Sequence

from .models import (
    CancelledError,
    DocuForgeError,
    MissingEngineError,
    Operation,
    TaskFailure,
    TaskResult,
    coerce_result,
)
from .utils import ensure_output_dir, output_directory_lock, validate_inputs

ProgressCallback = Callable[[float, str], None]
HandlerProgressCallback = Callable[[float, str, int | None, int | None], None]
_CURRENT_CANCEL_EVENT: ContextVar[threading.Event | None] = ContextVar(
    "docuforge_cancel_event", default=None
)
_CURRENT_PROGRESS_REPORTER: ContextVar[HandlerProgressCallback | None] = ContextVar(
    "docuforge_progress_reporter", default=None
)
_CURRENT_PROGRESS_SCOPE: ContextVar[
    tuple[float, float, int | None, int | None] | None
] = ContextVar("docuforge_progress_scope", default=None)
_CURRENT_RUNNER: ContextVar["TaskRunner | None"] = ContextVar(
    "docuforge_task_runner", default=None
)


def progress_message(
    stage: str,
    *,
    current_file: int | None = None,
    total_files: int | None = None,
) -> str:
    """Build a compact, consistent status line for UI and CLI consumers."""

    parts = [f"阶段：{str(stage).strip() or '处理中'}"]
    if total_files is not None and total_files > 0:
        if current_file is None:
            parts.append(f"共 {total_files} 个文件")
        else:
            current = min(max(1, int(current_file)), int(total_files))
            parts.append(f"文件 {current}/{int(total_files)}")
    return " · ".join(parts)


def report_progress(
    fraction: float,
    stage: str,
    *,
    current_file: int | None = None,
    total_files: int | None = None,
) -> None:
    """Report handler-relative progress without changing handler signatures.

    The runner installs the reporter in the same context used for cancellation.
    Calls made outside a running task intentionally become no-ops, so processors
    can safely opt in without needing an additional callback parameter.
    """

    scope = _CURRENT_PROGRESS_SCOPE.get()
    mapped_fraction = fraction
    if scope is not None:
        start, span, scope_file, scope_total = scope
        try:
            relative = float(fraction)
        except (TypeError, ValueError, OverflowError):
            relative = 0.0
        if relative != relative:  # NaN
            relative = 0.0
        relative = min(1.0, max(0.0, relative))
        mapped_fraction = start + span * relative
        if current_file is None:
            current_file = scope_file
        if total_files is None:
            total_files = scope_total

    reporter = _CURRENT_PROGRESS_REPORTER.get()
    if reporter is not None:
        reporter(mapped_fraction, stage, current_file, total_files)


@contextmanager
def progress_scope(
    start: float,
    span: float,
    *,
    current_file: int | None = None,
    total_files: int | None = None,
) -> Iterator[None]:
    """Map nested processor progress into a bounded portion of the parent task.

    Batch handlers use one scope per input file.  A processor may then report
    detailed page/stage progress from 0 to 1 without making file two appear to
    jump backwards or remain pinned at the first file's completion value.
    """

    normalized_start = min(1.0, max(0.0, float(start)))
    normalized_span = min(1.0 - normalized_start, max(0.0, float(span)))
    parent = _CURRENT_PROGRESS_SCOPE.get()
    if parent is not None:
        parent_start, parent_span, parent_file, parent_total = parent
        normalized_start = parent_start + parent_span * normalized_start
        normalized_span *= parent_span
        if current_file is None:
            current_file = parent_file
        if total_files is None:
            total_files = parent_total
    token = _CURRENT_PROGRESS_SCOPE.set(
        (normalized_start, normalized_span, current_file, total_files)
    )
    try:
        yield
    finally:
        _CURRENT_PROGRESS_SCOPE.reset(token)


def check_cancelled(message: str = "任务已取消") -> None:
    event = _CURRENT_CANCEL_EVENT.get()
    if event is not None and event.is_set():
        raise CancelledError(message)


def task_runner_active() -> bool:
    """Return whether code is executing inside a supervised TaskRunner task."""

    return _CURRENT_RUNNER.get() is not None


@contextmanager
def task_runner_context(runner: "TaskRunner") -> Iterator[None]:
    """Install an existing runner inside an isolated processor process."""

    cancel_token = _CURRENT_CANCEL_EVENT.set(runner.cancel_event)
    runner_token = _CURRENT_RUNNER.set(runner)
    try:
        yield
    finally:
        _CURRENT_RUNNER.reset(runner_token)
        _CURRENT_CANCEL_EVENT.reset(cancel_token)


@contextmanager
def cancellation_callback(callback: Callable[[], None]) -> Iterator[None]:
    """Register an external-engine stop callback for the active task.

    The callback is invoked from a small background cancellation thread so the
    GUI's stop button remains responsive even when an external process needs a
    moment to terminate. Calls made outside TaskRunner are harmless no-ops.
    """

    runner = _CURRENT_RUNNER.get()
    if runner is None:
        yield
        return
    key = runner._register_cancel_callback(callback)
    try:
        yield
    finally:
        runner._unregister_cancel_callback(key)


def _friendly_failure(exc: BaseException) -> str:
    text = str(exc).strip()
    return text or f"{type(exc).__name__}（未提供错误说明）"


def _snapshot_output_files(output_dir: Path) -> set[Path]:
    try:
        return {path.resolve() for path in output_dir.rglob("*") if path.is_file()}
    except OSError:
        return set()


def _looks_unfinished(path: Path) -> bool:
    name = path.name.casefold()
    return (
        ".tmp" in name
        or name.endswith((".part", ".partial", ".download", ".crdownload"))
        or name.startswith("~$")
    )


def _cleanup_unfinished_outputs(
    output_dir: Path,
    existing: set[Path],
    *,
    preserve: Sequence[Path] = (),
) -> list[Path]:
    """Remove only newly-created temporary/empty artifacts, preserving outputs."""

    removed: list[Path] = []
    preserved = {Path(path).expanduser().resolve() for path in preserve}
    try:
        candidates = [path for path in output_dir.rglob("*") if path.is_file()]
    except OSError:
        return removed
    for path in candidates:
        try:
            resolved = path.resolve()
            if resolved in existing or resolved in preserved:
                continue
            if not _looks_unfinished(path) and path.stat().st_size > 0:
                continue
            path.unlink(missing_ok=True)
            removed.append(resolved)
        except OSError:
            continue
    try:
        directories = sorted(
            (path for path in output_dir.rglob("*") if path.is_dir()),
            key=lambda item: len(item.parts),
            reverse=True,
        )
    except OSError:
        directories = []
    for directory in directories:
        try:
            directory.rmdir()
        except OSError:
            pass
    return removed


def _normalize_and_validate_outputs(
    outputs: Sequence[Path],
    target_dir: Path,
    *,
    allow_external: bool,
) -> list[Path]:
    normalized = [Path(path).expanduser().resolve() for path in outputs]
    missing = [path for path in normalized if not path.is_file()]
    if missing:
        raise DocuForgeError(f"处理器未生成声明的输出文件：{missing[0]}")
    if not allow_external:
        outside = [path for path in normalized if not path.is_relative_to(target_dir)]
        if outside:
            raise DocuForgeError(f"处理器输出超出所选文件夹：{outside[0]}")
    return normalized


class TaskRunner:
    def __init__(self, cancel_event: threading.Event | None = None) -> None:
        self.cancel_event = cancel_event or threading.Event()
        self._cancel_callbacks: dict[int, Callable[[], None]] = {}
        self._cancel_callbacks_lock = threading.Lock()
        self._next_cancel_callback = 0
        self._cancel_dispatched = False

    def cancel(self) -> None:
        self.cancel_event.set()
        with self._cancel_callbacks_lock:
            if self._cancel_dispatched:
                return
            self._cancel_dispatched = True
            callbacks = list(self._cancel_callbacks.values())
        if callbacks:
            threading.Thread(
                target=self._invoke_cancel_callbacks,
                args=(callbacks,),
                name="docuforge-cancel-external-engines",
                daemon=True,
            ).start()

    @staticmethod
    def _invoke_cancel_callbacks(callbacks: Sequence[Callable[[], None]]) -> None:
        for callback in callbacks:
            try:
                callback()
            except Exception:
                pass

    def _register_cancel_callback(self, callback: Callable[[], None]) -> int:
        with self._cancel_callbacks_lock:
            self._next_cancel_callback += 1
            key = self._next_cancel_callback
            self._cancel_callbacks[key] = callback
            already_cancelled = self.cancel_event.is_set()
        if already_cancelled:
            threading.Thread(
                target=self._invoke_cancel_callbacks,
                args=([callback],),
                name="docuforge-cancel-late-engine",
                daemon=True,
            ).start()
        return key

    def _unregister_cancel_callback(self, key: int) -> None:
        with self._cancel_callbacks_lock:
            self._cancel_callbacks.pop(key, None)

    def run(
        self,
        operation: Operation,
        inputs: Sequence[str | Path],
        output_dir: str | Path,
        raw_parameters: Mapping[str, Any] | None = None,
        progress: ProgressCallback | None = None,
    ) -> TaskResult:
        callback = progress or (lambda _value, _message: None)
        if self.cancel_event.is_set():
            raise CancelledError("任务已取消")
        input_count = len(inputs)
        last_progress = 0.0

        def emit(value: float, message: str) -> None:
            nonlocal last_progress
            try:
                normalized = float(value)
            except (TypeError, ValueError, OverflowError):
                normalized = last_progress
            if normalized != normalized:  # NaN
                normalized = last_progress
            normalized = min(1.0, max(last_progress, normalized))
            last_progress = normalized
            callback(normalized, message)

        capability = operation.capability()
        if not capability.runnable:
            raise MissingEngineError(capability.reason)
        emit(
            0.02,
            progress_message("准备任务与处理引擎", total_files=input_count),
        )
        emit(0.04, progress_message("检查输入文件", total_files=input_count))
        paths = validate_inputs(
            inputs, operation.extensions, operation.min_inputs, operation.max_inputs
        )
        if operation.reject_encrypted_pdf_inputs or operation.reject_signed_pdf_inputs:
            try:
                from pypdf import PdfReader
            except ImportError as exc:
                raise MissingEngineError("PDF 安全检查需要 pypdf") from exc
            encrypted: list[Path] = []
            signed: list[Path] = []
            pdf_paths = [path for path in paths if path.suffix.lower() == ".pdf"]
            for pdf_index, path in enumerate(pdf_paths, start=1):
                emit(
                    0.04 + 0.04 * (pdf_index / max(1, len(pdf_paths))),
                    progress_message(
                        f"检查 PDF 安全状态：{path.name}",
                        current_file=pdf_index,
                        total_files=len(pdf_paths),
                    ),
                )
                try:
                    reader = PdfReader(path)
                except Exception as exc:
                    raise DocuForgeError(
                        f"无法检查 PDF 安全状态，文件可能已损坏或格式无效：{path.name}"
                    ) from exc
                try:
                    if reader.is_encrypted:
                        if operation.reject_encrypted_pdf_inputs:
                            encrypted.append(path)
                        continue
                    if operation.reject_signed_pdf_inputs:
                        fields = reader.get_fields() or {}
                        has_signature_field = False
                        for field in fields.values():
                            if str(field.get("/FT", "")) != "/Sig":
                                continue
                            value = field.get("/V")
                            value_object = (
                                value.get_object()
                                if hasattr(value, "get_object")
                                else value
                            )
                            if (
                                value_object
                                and hasattr(value_object, "get")
                                and (
                                    value_object.get("/ByteRange") is not None
                                    or value_object.get("/Contents") is not None
                                )
                            ):
                                has_signature_field = True
                                break
                        root = reader.trailer.get("/Root")
                        root_object = (
                            root.get_object() if hasattr(root, "get_object") else root
                        )
                        has_doc_mdp = bool(
                            root_object
                            and hasattr(root_object, "get")
                            and root_object.get("/Perms") is not None
                        )
                        if has_signature_field or has_doc_mdp:
                            signed.append(path)
                finally:
                    reader.close()
            if encrypted:
                raise DocuForgeError(
                    "为避免静默移除 PDF 加密和权限设置，本任务拒绝直接处理加密 PDF。"
                    "请先用“PDF 解密”另存，处理完成后再用“PDF 加密”设置新密码。"
                )
            if signed:
                raise DocuForgeError(
                    "该 PDF 包含数字签名；当前操作会使现有签名失效，因此已停止。"
                    "如需制作衍生版本，请先明确保存未签名副本，并保留原始已签文件用于验签。"
                )
        target_dir = ensure_output_dir(output_dir)
        params = operation.normalize_parameters(raw_parameters or {})
        emit(
            0.10,
            progress_message(f"准备使用 {capability.engine}", total_files=len(paths)),
        )
        started = time.perf_counter()
        token = _CURRENT_CANCEL_EVENT.set(self.cancel_event)
        runner_token = _CURRENT_RUNNER.set(self)
        handler_start = 0.12
        handler_span = 0.80
        active_file_index: int | None = None

        def handler_progress(
            fraction: float,
            stage: str,
            current_file: int | None,
            total_files: int | None,
        ) -> None:
            if active_file_index is not None:
                current_file = active_file_index
                total_files = len(paths)
            try:
                relative = float(fraction)
            except (TypeError, ValueError, OverflowError):
                relative = 0.0
            if relative != relative:  # NaN
                relative = 0.0
            relative = min(1.0, max(0.0, relative))
            emit(
                handler_start + handler_span * relative,
                progress_message(
                    stage,
                    current_file=current_file,
                    total_files=total_files or len(paths),
                ),
            )

        progress_token = _CURRENT_PROGRESS_REPORTER.set(handler_progress)
        captured_warnings: list[warnings.WarningMessage] = []
        output_snapshot = _snapshot_output_files(target_dir)
        result = TaskResult()
        try:
            emit(
                handler_start,
                progress_message(
                    f"{capability.engine} 正在执行核心处理",
                    total_files=len(paths),
                ),
            )
            with output_directory_lock(target_dir):
                with warnings.catch_warnings(record=True) as warning_records:
                    warnings.simplefilter("always")
                    if operation.independent_inputs and len(paths) > 1:
                        total = len(paths)
                        for index, source in enumerate(paths, start=1):
                            if self.cancel_event.is_set():
                                result.cancelled = True
                                result.cancelled_inputs.extend(paths[index - 1 :])
                                break
                            active_file_index = index
                            emit(
                                handler_start
                                + handler_span * ((index - 1) / max(1, total)),
                                progress_message(
                                    f"处理 {source.name}",
                                    current_file=index,
                                    total_files=total,
                                ),
                            )
                            item_snapshot = _snapshot_output_files(target_dir)
                            try:
                                with progress_scope(
                                    (index - 1) / max(1, total),
                                    1.0 / max(1, total),
                                    current_file=index,
                                    total_files=total,
                                ):
                                    item = coerce_result(
                                        operation.handler([source], target_dir, params)
                                    )
                                if not item.outputs and not operation.allow_empty_outputs:
                                    raise DocuForgeError(
                                        "处理器没有为该文件生成任何输出文件"
                                    )
                                item.outputs = _normalize_and_validate_outputs(
                                    item.outputs,
                                    target_dir,
                                    allow_external=operation.allow_external_outputs,
                                )
                                result.outputs.extend(item.outputs)
                                result.warnings.extend(
                                    warning
                                    for warning in item.warnings
                                    if warning not in result.warnings
                                )
                                if item.details:
                                    result.details.setdefault("per_file", {})[
                                        str(source)
                                    ] = dict(item.details)
                                result.completed_inputs.append(source)
                            except CancelledError:
                                _cleanup_unfinished_outputs(target_dir, item_snapshot)
                                result.cancelled = True
                                result.cancelled_inputs.extend(paths[index - 1 :])
                                break
                            except Exception as exc:
                                _cleanup_unfinished_outputs(target_dir, item_snapshot)
                                result.failed_inputs.append(
                                    TaskFailure(
                                        input_path=source,
                                        error_type=type(exc).__name__,
                                        message=_friendly_failure(exc),
                                    )
                                )
                                report_progress(
                                    index / max(1, total),
                                    f"{source.name} 处理失败，继续下一个文件",
                                    current_file=index,
                                    total_files=total,
                                )
                        active_file_index = None
                    else:
                        result = coerce_result(
                            operation.handler(paths, target_dir, params)
                        )
                        if (
                            not result.completed_inputs
                            and not result.failed_inputs
                            and not result.cancelled_inputs
                        ):
                            result.completed_inputs.extend(paths)
                    captured_warnings = list(warning_records)
        except CancelledError as exc:
            result.cancelled = True
            known_inputs = {
                *result.completed_inputs,
                *(failure.input_path for failure in result.failed_inputs),
                *result.cancelled_inputs,
            }
            result.cancelled_inputs.extend(
                path for path in paths if path not in known_inputs
            )
            _cleanup_unfinished_outputs(
                target_dir,
                output_snapshot,
                preserve=result.outputs,
            )
            result.details.setdefault("elapsed_seconds", round(time.perf_counter() - started, 3))
            raise CancelledError(str(exc), result=result) from None
        finally:
            active_file_index = None
            _CURRENT_PROGRESS_REPORTER.reset(progress_token)
            _CURRENT_RUNNER.reset(runner_token)
            _CURRENT_CANCEL_EVENT.reset(token)
        if self.cancel_event.is_set():
            result.cancelled = True
            known_inputs = {
                *result.completed_inputs,
                *(failure.input_path for failure in result.failed_inputs),
                *result.cancelled_inputs,
            }
            result.cancelled_inputs.extend(
                path for path in paths if path not in known_inputs
            )
            _cleanup_unfinished_outputs(
                target_dir,
                output_snapshot,
                preserve=result.outputs,
            )
            result.details.setdefault(
                "elapsed_seconds", round(time.perf_counter() - started, 3)
            )
            raise CancelledError(
                "任务已取消；已完成的输出文件会保留",
                result=result,
            )
        emit(0.93, progress_message("核对处理结果", total_files=len(paths)))
        for warning in captured_warnings:
            message = str(warning.message)
            if message and message not in result.warnings:
                result.warnings.append(message)
        if (
            not result.outputs
            and not operation.allow_empty_outputs
            and not result.failed_inputs
        ):
            raise DocuForgeError("处理器没有生成任何输出文件")
        normalized_outputs = _normalize_and_validate_outputs(
            result.outputs,
            target_dir,
            allow_external=operation.allow_external_outputs,
        )
        emit(
            0.96,
            progress_message(
                f"验证 {len(normalized_outputs)} 个输出文件",
                total_files=len(paths),
            ),
        )
        for path in normalized_outputs:
            if path.stat().st_size == 0:
                result.warnings.append(f"输出文件为空：{path.name}")
        result.outputs = normalized_outputs
        result.details.setdefault(
            "elapsed_seconds", round(time.perf_counter() - started, 3)
        )
        result.details.setdefault("engine", capability.engine)
        result.details.setdefault("input_count", len(paths))
        result.details.setdefault("completed_count", len(result.completed_inputs))
        result.details.setdefault("failed_count", len(result.failed_inputs))
        result.details.setdefault("cancelled_count", len(result.cancelled_inputs))
        emit(
            1.0,
            progress_message(
                f"完成，共生成 {len(result.outputs)} 个文件",
                current_file=len(paths) if paths else None,
                total_files=len(paths),
            ),
        )
        return result
