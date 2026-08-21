from __future__ import annotations

import json
import os
import shutil
import signal
import subprocess
import sys
import time
from pathlib import Path
from typing import Sequence


def _assign_kill_on_close_job(process: subprocess.Popen[object]) -> object | None:
    """Keep the CLI process tree tied to this wrapper on Windows."""

    if os.name != "nt":
        return None
    job = None
    process_handle = None
    try:
        import win32api
        import win32job

        job = win32job.CreateJobObject(None, "")
        limits = win32job.QueryInformationJobObject(
            job, win32job.JobObjectExtendedLimitInformation
        )
        limits["BasicLimitInformation"][
            "LimitFlags"
        ] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(
            job, win32job.JobObjectExtendedLimitInformation, limits
        )
        process_handle = win32api.OpenProcess(0x0001 | 0x0100, False, process.pid)
        win32job.AssignProcessToJobObject(job, process_handle)
        return job
    except Exception:
        if job is not None:
            try:
                job.Close()
            except Exception:
                pass
        return None
    finally:
        if process_handle is not None:
            try:
                process_handle.Close()
            except Exception:
                pass


def _configured_command(skill_root: Path) -> list[str] | None:
    config_path = skill_root / "layoutloom-cli.json"
    if not config_path.is_file():
        return None
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return None
    command = payload.get("command") if isinstance(payload, dict) else None
    if not isinstance(command, list) or not command:
        return None
    normalized = [str(part) for part in command if str(part)]
    return normalized or None


def _source_checkout_command(skill_root: Path) -> list[str] | None:
    for parent in skill_root.parents:
        if not (parent / "docuforge" / "cli.py").is_file():
            continue
        python = parent / ".venv" / "Scripts" / "python.exe"
        if python.is_file():
            return [str(python), "-m", "docuforge.cli"]
    return None


def locate_command(skill_root: Path) -> list[str]:
    configured = _configured_command(skill_root)
    if configured:
        return configured

    override = os.environ.get("LAYOUTLOOM_CLI", "").strip().strip('"')
    if override:
        path = Path(override).expanduser()
        if path.is_file():
            return [str(path.resolve())]

    portable = skill_root.parent.parent / "LayoutLoom-CLI.exe"
    if portable.is_file():
        return [str(portable.resolve())]

    source_command = _source_checkout_command(skill_root)
    if source_command:
        return source_command

    for name in ("LayoutLoom-CLI.exe", "layoutloom.exe", "layoutloom"):
        executable = shutil.which(name)
        if executable:
            return [executable]

    raise FileNotFoundError(
        "LayoutLoom CLI was not found. Install the portable Skill with "
        "`LayoutLoom-CLI.exe agent install-skill`, set LAYOUTLOOM_CLI, "
        "or install LayoutLoom in the active Python environment."
    )


def main(argv: Sequence[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    skill_root = Path(__file__).resolve().parents[1]
    try:
        command = locate_command(skill_root)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 127
    child_arguments = [*command, "agent", *arguments]
    try:
        process = subprocess.Popen(
            child_arguments,
            creationflags=(
                subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
            ),
        )
    except OSError as exc:
        print(f"Unable to start LayoutLoom CLI: {exc}", file=sys.stderr)
        return 126
    job = _assign_kill_on_close_job(process)

    cancel_started: float | None = None
    previous_handlers: dict[int, object] = {}

    def forward_cancel(signum: int, _frame: object) -> None:
        nonlocal cancel_started
        if process.poll() is not None:
            return
        if cancel_started is not None:
            return
        cancel_started = time.monotonic()
        try:
            if os.name == "nt":
                process.send_signal(signal.CTRL_BREAK_EVENT)
            else:
                process.send_signal(signum)
        except (OSError, ValueError):
            pass

    for signal_name in ("SIGINT", "SIGTERM", "SIGBREAK"):
        signal_value = getattr(signal, signal_name, None)
        if signal_value is None:
            continue
        try:
            previous_handlers[int(signal_value)] = signal.getsignal(signal_value)
            signal.signal(signal_value, forward_cancel)
        except (OSError, RuntimeError, ValueError):
            pass

    try:
        while True:
            try:
                return int(process.wait(timeout=0.2))
            except subprocess.TimeoutExpired:
                if cancel_started is None or time.monotonic() - cancel_started < 15.0:
                    continue
                try:
                    process.terminate()
                    return int(process.wait(timeout=2.0))
                except (OSError, subprocess.TimeoutExpired):
                    try:
                        process.kill()
                        return int(process.wait(timeout=2.0))
                    except (OSError, subprocess.TimeoutExpired):
                        return 130
    finally:
        for signal_number, handler in previous_handlers.items():
            try:
                signal.signal(signal_number, handler)
            except (OSError, RuntimeError, ValueError):
                pass
        if job is not None:
            try:
                job.Close()  # type: ignore[attr-defined]
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
