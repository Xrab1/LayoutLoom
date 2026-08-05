from __future__ import annotations

import errno
import os
import re
import shutil
import tempfile
import threading
import time
import unicodedata
from contextlib import contextmanager
from hashlib import sha256
from numbers import Integral
from pathlib import Path
from typing import BinaryIO, Iterable, Iterator

from .models import DocuForgeError, ValidationError


INVALID_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    "CONIN$",
    "CONOUT$",
    *(f"COM{index}" for index in range(1, 10)),
    *(f"LPT{index}" for index in range(1, 10)),
}
_ACTIVE_OUTPUT_LOCKS: set[str] = set()
_ACTIVE_OUTPUT_LOCKS_GUARD = threading.Lock()


def _normalized_output_directory(path: str | Path) -> tuple[Path, str]:
    resolved = Path(path).expanduser().resolve()
    normalized = os.path.normcase(os.path.normpath(str(resolved)))
    return resolved, normalized


def _output_lock_path(normalized_directory: str) -> Path:
    digest = sha256(normalized_directory.encode("utf-8")).hexdigest()
    return Path(tempfile.gettempdir()) / f"docuforge-output-{digest}.lock"


def _output_lock_conflict(directory: Path) -> DocuForgeError:
    return DocuForgeError(f"输出文件夹正被其他任务使用，请稍后重试：{directory}")


def _acquire_os_file_lock(handle: BinaryIO, directory: Path) -> None:
    try:
        if os.name == "nt":
            import msvcrt

            handle.seek(0, os.SEEK_END)
            if handle.tell() == 0:
                handle.write(b"\0")
                handle.flush()
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl

            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError as exc:
        busy_errnos = {errno.EACCES, errno.EAGAIN}
        if hasattr(errno, "EDEADLK"):
            busy_errnos.add(errno.EDEADLK)
        if exc.errno in busy_errnos or getattr(exc, "winerror", None) in {33, 36}:
            raise _output_lock_conflict(directory) from None
        raise DocuForgeError(f"无法为输出文件夹建立任务锁：{directory}（{exc}）") from exc


def _release_os_file_lock(handle: BinaryIO) -> None:
    if os.name == "nt":
        import msvcrt

        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
    else:
        import fcntl

        fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


@contextmanager
def output_directory_lock(path: str | Path) -> Iterator[Path]:
    """Hold a non-blocking process-wide and machine-local lock for an output folder."""

    directory, normalized = _normalized_output_directory(path)
    with _ACTIVE_OUTPUT_LOCKS_GUARD:
        if normalized in _ACTIVE_OUTPUT_LOCKS:
            raise _output_lock_conflict(directory)
        _ACTIVE_OUTPUT_LOCKS.add(normalized)

    handle = None
    locked = False
    try:
        lock_path = _output_lock_path(normalized)
        try:
            handle = lock_path.open("a+b", buffering=0)
        except OSError as exc:
            raise DocuForgeError(f"无法创建输出文件夹任务锁：{directory}（{exc}）") from exc
        _acquire_os_file_lock(handle, directory)
        locked = True
        yield directory
    finally:
        try:
            if handle is not None:
                if locked:
                    try:
                        _release_os_file_lock(handle)
                    except (OSError, ValueError):
                        # Closing the descriptor also releases the operating-system lock.
                        pass
                try:
                    handle.close()
                except OSError:
                    pass
        finally:
            with _ACTIVE_OUTPUT_LOCKS_GUARD:
                _ACTIVE_OUTPUT_LOCKS.discard(normalized)


def safe_filename(
    value: str, fallback: str = "output", max_utf16_units: int = 160
) -> str:
    cleaned = unicodedata.normalize("NFC", str(value))
    cleaned = INVALID_FILENAME.sub("_", cleaned).strip(" .")
    cleaned = cleaned or fallback
    if cleaned.split(".", 1)[0].upper() in WINDOWS_RESERVED_NAMES:
        cleaned = f"_{cleaned}"
    if len(cleaned.encode("utf-16-le")) // 2 > max_utf16_units:
        digest = sha256(cleaned.encode("utf-8")).hexdigest()[:10]
        suffix = f"_{digest}"
        budget = max_utf16_units - len(suffix)
        prefix: list[str] = []
        used = 0
        for character in cleaned:
            units = len(character.encode("utf-16-le")) // 2
            if used + units > budget:
                break
            prefix.append(character)
            used += units
        cleaned = "".join(prefix).rstrip(" .") + suffix
    return cleaned


def ensure_output_dir(path: str | Path) -> Path:
    output = Path(path).expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    if not output.is_dir():
        raise ValidationError(f"输出位置不是文件夹：{output}")
    return output


def optimal_worker_count(item_count: int, *, cap: int = 4) -> int:
    """Choose a deterministic, memory-conscious worker count for local batches."""

    if (
        isinstance(item_count, bool)
        or not isinstance(item_count, Integral)
        or item_count < 0
    ):
        raise ValidationError("item_count 必须是非负整数")
    if isinstance(cap, bool) or not isinstance(cap, Integral) or cap < 1:
        raise ValidationError("cap 必须是正整数")
    cpu_count = os.cpu_count() or 1
    return max(1, min(item_count or 1, cap, cpu_count))


def unique_path(path: str | Path, overwrite: bool = False) -> Path:
    target = Path(path)
    if overwrite or not target.exists():
        return target
    for index in range(1, 10000):
        candidate = target.with_name(f"{target.stem}_{index}{target.suffix}")
        if not candidate.exists():
            return candidate
    raise ValidationError(f"无法为输出生成唯一文件名：{target.name}")


def unique_directory(path: str | Path) -> Path:
    target = Path(path)
    if not target.exists():
        return target
    for index in range(1, 10000):
        candidate = target.with_name(f"{target.name}_{index}")
        if not candidate.exists():
            return candidate
    raise ValidationError(f"无法为输出生成唯一文件夹名：{target.name}")


@contextmanager
def atomic_output(target: str | Path) -> Iterator[Path]:
    """Yield a same-directory temporary path and atomically replace the target on success."""

    destination = Path(target)
    destination.parent.mkdir(parents=True, exist_ok=True)
    handle, temp_name = tempfile.mkstemp(
        prefix=f".{destination.stem}.",
        suffix=f".tmp{destination.suffix}",
        dir=destination.parent,
    )
    os.close(handle)
    temp_path = Path(temp_name)
    try:
        yield temp_path
        # WPS/Office may release a just-saved file handle a fraction of a
        # second after Close/Quit returns.  Retry only Windows sharing/lock
        # violations; all other filesystem errors remain immediate.
        for attempt in range(31):
            try:
                os.replace(temp_path, destination)
                break
            except PermissionError as exc:
                if getattr(exc, "winerror", None) not in {32, 33} or attempt >= 30:
                    raise
                time.sleep(min(0.5, 0.04 + attempt * 0.02))
    except Exception:
        for attempt in range(11):
            try:
                temp_path.unlink(missing_ok=True)
                break
            except PermissionError as exc:
                if getattr(exc, "winerror", None) not in {32, 33} or attempt >= 10:
                    break
                time.sleep(0.05 + attempt * 0.03)
        raise


def copy_atomic(source: str | Path, target: str | Path) -> Path:
    destination = Path(target)
    with atomic_output(destination) as temporary:
        shutil.copy2(source, temporary)
    return destination


def validate_inputs(
    paths: Iterable[str | Path],
    extensions: Iterable[str] = (),
    min_inputs: int = 1,
    max_inputs: int | None = None,
) -> list[Path]:
    normalized = [Path(item).expanduser().resolve() for item in paths]
    if len(normalized) < min_inputs:
        raise ValidationError(f"至少需要选择 {min_inputs} 个输入文件")
    if max_inputs is not None and len(normalized) > max_inputs:
        raise ValidationError(f"最多只能选择 {max_inputs} 个输入文件")
    missing = [str(item) for item in normalized if not item.is_file()]
    if missing:
        raise ValidationError(f"文件不存在：{missing[0]}")
    allowed = {
        item.lower() if item.startswith(".") else f".{item.lower()}"
        for item in extensions
    }
    if allowed:
        invalid = [
            item.name for item in normalized if item.suffix.lower() not in allowed
        ]
        if invalid:
            raise ValidationError(f"文件格式不符合当前任务：{invalid[0]}")
    return normalized


def parse_page_spec(spec: str, page_count: int) -> list[int]:
    """Parse one-based page specs such as ``1-3,5,8-`` into zero-based indexes."""

    if page_count < 1:
        return []
    text = str(spec).strip().replace("，", ",").replace("－", "-")
    if not text or text.lower() in {"all", "全部"}:
        return list(range(page_count))
    pages: set[int] = set()
    for part in text.split(","):
        token = part.strip()
        if not token:
            continue
        if "-" in token:
            left, right = token.split("-", 1)
            try:
                start = int(left) if left else 1
                end = int(right) if right else page_count
            except ValueError as exc:
                raise ValidationError(f"无效页码范围：{token}") from exc
            if start > end:
                start, end = end, start
            pages.update(range(start - 1, end))
        else:
            try:
                pages.add(int(token) - 1)
            except ValueError as exc:
                raise ValidationError(f"无效页码：{token}") from exc
    invalid = sorted(page + 1 for page in pages if page < 0 or page >= page_count)
    if invalid:
        raise ValidationError(f"页码超出范围：{invalid[0]}（文档共 {page_count} 页）")
    return sorted(pages)
