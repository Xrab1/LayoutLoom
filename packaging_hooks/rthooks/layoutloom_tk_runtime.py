"""Provide a resilient Tcl/Tk runtime for the frozen Windows application.

PyInstaller normally points Tkinter at ``_tcl_data`` and ``_tk_data`` inside
the one-folder bundle.  Some file-transfer and extraction tools have been
observed to omit one of those script-heavy directories.  LayoutLoom therefore
ships a second, single-file ZIP copy and restores it to a per-user cache when
the primary copy is incomplete.
"""

from __future__ import annotations

import errno
import hashlib
import os
import shutil
import sys
import tempfile
import time
import traceback
import zipfile
from pathlib import Path, PurePosixPath


_SELF_TEST_ERROR_FILE_ENV = "LAYOUTLOOM_SELF_TEST_ERROR_FILE"
_TRANSIENT_RENAME_WINERRORS = {5, 32, 33}
_TRANSIENT_RENAME_ERRNOS = {errno.EACCES, errno.EBUSY, errno.EPERM}


def _runtime_complete(root: Path) -> bool:
    return (root / "_tcl_data" / "init.tcl").is_file() and (
        root / "_tk_data" / "tk.tcl"
    ).is_file()


def _archive_fingerprint(archive: Path) -> str:
    digest = hashlib.sha256()
    with archive.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()[:20]


def _extract_runtime_archive(archive: Path, destination: Path) -> None:
    allowed_roots = {"_tcl_data", "_tk_data"}
    destination.mkdir(parents=True, exist_ok=True)
    resolved_destination = destination.resolve()
    with zipfile.ZipFile(archive) as package:
        for member in package.infolist():
            normalized = member.filename.replace("\\", "/")
            parts = PurePosixPath(normalized).parts
            if not parts or parts[0] not in allowed_roots:
                raise RuntimeError(
                    f"Unexpected entry in bundled Tcl/Tk recovery archive: {normalized}"
                )
            if any(part in {"", ".", ".."} for part in parts):
                raise RuntimeError(
                    f"Unsafe entry in bundled Tcl/Tk recovery archive: {normalized}"
                )
            target = destination.joinpath(*parts)
            resolved_target = target.resolve()
            try:
                resolved_target.relative_to(resolved_destination)
            except ValueError as exc:
                raise RuntimeError(
                    f"Unsafe Tcl/Tk recovery target: {normalized}"
                ) from exc
            if member.is_dir():
                resolved_target.mkdir(parents=True, exist_ok=True)
                continue
            resolved_target.parent.mkdir(parents=True, exist_ok=True)
            with package.open(member) as source, resolved_target.open("wb") as output:
                shutil.copyfileobj(source, output, length=1024 * 1024)


def _publish_runtime_cache(staging: Path, cache_root: Path) -> None:
    """Atomically publish an extracted runtime despite short Windows file locks."""

    deadline = time.monotonic() + 15.0
    delay = 0.10
    while True:
        if _runtime_complete(cache_root):
            shutil.rmtree(staging, ignore_errors=True)
            return
        try:
            staging.replace(cache_root)
            return
        except OSError as exc:
            transient = (
                getattr(exc, "winerror", None) in _TRANSIENT_RENAME_WINERRORS
                or exc.errno in _TRANSIENT_RENAME_ERRNOS
            )
            if not transient or time.monotonic() >= deadline:
                raise
            time.sleep(delay)
            delay = min(delay * 1.7, 1.0)


def _cached_runtime(bundle_root: Path) -> Path:
    archive = bundle_root / "tk_runtime_backup.zip"
    if not archive.is_file():
        raise FileNotFoundError(
            "LayoutLoom portable package is incomplete: both the primary Tcl/Tk "
            f"runtime and its recovery archive are missing under {bundle_root}. "
            "Please extract the complete portable ZIP again."
        )

    cache_parent = Path(
        os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    ) / "LayoutLoom" / "runtime"
    cache_root = cache_parent / f"tk-{_archive_fingerprint(archive)}"
    if _runtime_complete(cache_root):
        return cache_root

    cache_parent.mkdir(parents=True, exist_ok=True)
    lock_path = cache_parent / f"{cache_root.name}.lock"
    lock_handle: int | None = None
    deadline = time.monotonic() + 45.0
    while lock_handle is None:
        try:
            lock_handle = os.open(
                lock_path,
                os.O_CREAT | os.O_EXCL | os.O_WRONLY,
            )
        except FileExistsError:
            if _runtime_complete(cache_root):
                return cache_root
            try:
                age = time.time() - lock_path.stat().st_mtime
            except OSError:
                age = 0.0
            if age > 180.0:
                try:
                    lock_path.unlink()
                except OSError:
                    pass
                continue
            if time.monotonic() >= deadline:
                raise TimeoutError(
                    "Timed out waiting for another LayoutLoom process to "
                    "prepare the Tcl/Tk recovery runtime."
                )
            time.sleep(0.15)
    staging = Path(
        tempfile.mkdtemp(prefix="tk-stage-", dir=str(cache_parent))
    )
    try:
        _extract_runtime_archive(archive, staging)
        if not _runtime_complete(staging):
            raise FileNotFoundError(
                "The bundled Tcl/Tk recovery archive does not contain init.tcl "
                "and tk.tcl. Please download the portable package again."
            )
        if cache_root.exists() and not _runtime_complete(cache_root):
            shutil.rmtree(cache_root, ignore_errors=True)
        _publish_runtime_cache(staging, cache_root)
        return cache_root
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        raise
    finally:
        if lock_handle is not None:
            os.close(lock_handle)
        try:
            lock_path.unlink()
        except OSError:
            pass


def _configure_tk_runtime() -> None:
    if not getattr(sys, "frozen", False) or not hasattr(sys, "_MEIPASS"):
        return
    bundle_root = Path(sys._MEIPASS)
    runtime_root = bundle_root if _runtime_complete(bundle_root) else _cached_runtime(
        bundle_root
    )
    os.environ["TCL_LIBRARY"] = str(runtime_root / "_tcl_data")
    os.environ["TK_LIBRARY"] = str(runtime_root / "_tk_data")


def _configure_with_self_test_diagnostics() -> None:
    try:
        _configure_tk_runtime()
    except Exception:
        error_file = os.environ.get(_SELF_TEST_ERROR_FILE_ENV)
        if not error_file:
            raise
        try:
            error_path = Path(error_file)
            error_path.parent.mkdir(parents=True, exist_ok=True)
            error_path.write_text(traceback.format_exc(), encoding="utf-8")
        except OSError:
            pass
        # A windowed PyInstaller executable otherwise opens a modal bootstrap
        # error dialog, causing automated build validation to wait forever.
        os._exit(86)


_configure_with_self_test_diagnostics()
