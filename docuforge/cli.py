from __future__ import annotations

import argparse
import contextlib
import glob
import json
import signal
import sys
import time
from pathlib import Path
from typing import Any, Mapping, Sequence, TextIO

from . import __version__
from .agent_api import (
    AGENT_PROTOCOL_VERSION,
    EXIT_CANCELLED,
    EXIT_SUCCESS,
    catalog_payload,
    dumps_json,
    error_payload,
    exit_code_for_error,
    exit_code_for_result,
    install_codex_skill,
    operation_by_id,
    operation_matches_query,
    operation_payload,
    prepare_agent_request,
    protocol_payload,
    result_payload,
    utc_timestamp,
    validation_payload,
)
from .models import DocuForgeError, ValidationError
from .registry import get_operations
from .runner import TaskRunner


MAX_AGENT_REQUEST_BYTES = 1024 * 1024


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except (OSError, ValueError):
                pass


def _parse_parameter(value: str) -> tuple[str, Any]:
    if "=" not in value:
        raise argparse.ArgumentTypeError("参数格式必须是 key=value")
    key, raw = value.split("=", 1)
    normalized = key.strip()
    if not normalized:
        raise argparse.ArgumentTypeError("参数名不能为空")
    return normalized, raw


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            raise ValidationError(f"JSON 包含重复字段：{key}")
        value[key] = item
    return value


def _parse_json_object(text: str, *, source: str) -> Mapping[str, Any]:
    try:
        value = json.loads(text, object_pairs_hook=_reject_duplicate_keys)
    except UnicodeError as exc:
        raise ValidationError(f"JSON 不是有效 UTF-8：{source}") from exc
    if not isinstance(value, Mapping):
        raise ValidationError(f"JSON 必须包含对象：{source}")
    return value


def _load_json_object(path: Path) -> Mapping[str, Any]:
    try:
        raw = path.read_bytes()
    except OSError as exc:
        raise ValidationError(f"无法读取 JSON 文件：{path}（{exc}）") from exc
    if len(raw) > MAX_AGENT_REQUEST_BYTES:
        raise ValidationError(
            f"JSON 文件超过 {MAX_AGENT_REQUEST_BYTES // 1024} KiB 限制：{path}"
        )
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValidationError(f"JSON 文件必须使用 UTF-8：{path}") from exc
    return _parse_json_object(text, source=str(path))


def _load_agent_request(path_value: str) -> tuple[Mapping[str, Any], Path]:
    if path_value == "-":
        try:
            text = sys.stdin.read(MAX_AGENT_REQUEST_BYTES + 1)
        except OSError as exc:
            raise ValidationError(f"无法从标准输入读取请求：{exc}") from exc
        if len(text.encode("utf-8")) > MAX_AGENT_REQUEST_BYTES:
            raise ValidationError(
                f"标准输入请求超过 {MAX_AGENT_REQUEST_BYTES // 1024} KiB 限制"
            )
        payload = _parse_json_object(text, source="stdin")
        base_dir = Path.cwd().resolve()
    else:
        path = Path(path_value).expanduser().resolve()
        payload = _load_json_object(path)
        base_dir = path.parent
    if not isinstance(payload, Mapping):
        raise ValidationError("Agent 请求必须是 JSON 对象")
    return payload, base_dir


def _write_json(payload: Any, *, pretty: bool = False) -> None:
    if sys.stdout is None:
        raise RuntimeError("当前程序没有可写的标准输出；请使用 LayoutLoom-CLI.exe")
    sys.stdout.write(dumps_json(payload, pretty=pretty) + "\n")
    sys.stdout.flush()


class _JsonEventStream:
    def __init__(self, stream: TextIO | None = None) -> None:
        self.stream = stream or sys.stdout
        if self.stream is None:
            raise RuntimeError(
                "当前程序没有可写的标准输出；请使用 LayoutLoom-CLI.exe"
            )
        self._last_progress: tuple[float, str] | None = None
        self._sequence = 0

    def emit(self, event: str, **payload: Any) -> None:
        self.emit_payload(
            {
            "event": event,
            "protocol_version": AGENT_PROTOCOL_VERSION,
            "timestamp": utc_timestamp(),
            **payload,
            }
        )

    def emit_payload(self, payload: Mapping[str, Any]) -> None:
        self._sequence += 1
        envelope = dict(payload)
        envelope.setdefault("protocol_version", AGENT_PROTOCOL_VERSION)
        envelope.setdefault("timestamp", utc_timestamp())
        envelope["seq"] = self._sequence
        self.stream.write(dumps_json(envelope) + "\n")
        self.stream.flush()

    def progress(
        self,
        value: float,
        message: str,
        *,
        request_id: str,
        operation_id: str,
    ) -> None:
        normalized = min(1.0, max(0.0, float(value)))
        key = (round(normalized, 6), message)
        if self._last_progress == key:
            return
        self._last_progress = key
        self.emit(
            "progress",
            ok=True,
            request_id=request_id,
            operation=operation_id,
            fraction=normalized,
            percent=round(normalized * 100, 2),
            message=message,
        )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="layoutloom", description="页织工坊（LayoutLoom）命令行"
    )
    parser.add_argument(
        "--version", action="version", version=f"LayoutLoom {__version__}"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    list_parser = sub.add_parser("list", help="列出全部处理任务")
    list_parser.add_argument("--json", action="store_true", help="输出 JSON")
    list_parser.add_argument("--query", default="", help="按名称或说明搜索")
    list_parser.add_argument("--group", default="", help="只显示指定分组")

    describe = sub.add_parser("describe", help="查看一个任务的完整参数")
    describe.add_argument("operation", help="任务 ID")
    describe.add_argument("--json", action="store_true", help="输出 JSON")

    run = sub.add_parser("run", help="运行一个处理任务")
    run.add_argument("operation", help="任务 ID，可通过 list 查看")
    run.add_argument("inputs", nargs="+", type=Path)
    run.add_argument("-o", "--output", type=Path, required=True)
    run.add_argument(
        "-p", "--param", action="append", default=[], type=_parse_parameter
    )
    run.add_argument(
        "--params-file", type=Path, help="从 UTF-8 JSON 对象读取参数"
    )
    output_format = run.add_mutually_exclusive_group()
    output_format.add_argument("--json", action="store_true", help="输出最终 JSON")
    output_format.add_argument(
        "--jsonl", action="store_true", help="输出实时 JSON Lines 事件"
    )

    agent = sub.add_parser("agent", help="面向 AI Agent 的稳定 JSON 协议")
    agent_sub = agent.add_subparsers(dest="agent_command", required=True)

    protocol = agent_sub.add_parser("protocol", help="查看 Agent 协议与退出码")
    protocol.add_argument("--pretty", action="store_true")

    catalog = agent_sub.add_parser("catalog", help="以 JSON 查询处理能力")
    catalog.add_argument("--query", default="")
    catalog.add_argument("--group", default="")
    catalog.add_argument("--runnable-only", action="store_true")
    catalog.add_argument("--full", action="store_true")
    catalog.add_argument(
        "--probe", action="store_true", help="实际探测全部候选处理引擎"
    )
    catalog.add_argument("--pretty", action="store_true")

    agent_describe = agent_sub.add_parser("describe", help="查看任务 JSON Schema")
    agent_describe.add_argument("operation")
    agent_describe.add_argument("--pretty", action="store_true")

    validate = agent_sub.add_parser("validate", help="只校验请求，不创建输出")
    validate.add_argument(
        "--request",
        required=True,
        help="UTF-8 JSON 请求文件；使用 - 从 stdin 读取",
    )
    validate.add_argument("--pretty", action="store_true")
    validate.add_argument(
        "--allow-root",
        action="append",
        default=[],
        type=Path,
        help="限制输入、输出和 path 参数只能位于该目录；可重复",
    )
    validate.add_argument(
        "--allow-source-mutation",
        action="store_true",
        help="允许会移动原文件的参数；仅在用户明确授权时使用",
    )

    agent_run = agent_sub.add_parser("run", help="执行请求并返回机器事件")
    agent_run.add_argument(
        "--request",
        required=True,
        help="UTF-8 JSON 请求文件；使用 - 从 stdin 读取",
    )
    agent_run.add_argument(
        "--format",
        choices=("jsonl", "json"),
        default="jsonl",
        help="jsonl 实时输出进度；json 只输出最终结果",
    )
    agent_run.add_argument("--pretty", action="store_true")
    agent_run.add_argument(
        "--allow-root",
        action="append",
        default=[],
        type=Path,
        help="限制输入、输出和 path 参数只能位于该目录；可重复",
    )
    agent_run.add_argument(
        "--allow-source-mutation",
        action="store_true",
        help="允许会移动原文件的参数；仅在用户明确授权时使用",
    )

    quick_run = agent_sub.add_parser(
        "quick-run",
        help="一次调用执行已知任务；内部仍执行完整校验",
    )
    quick_run.add_argument("operation", help="已知任务 ID")
    quick_run.add_argument("inputs", nargs="+", type=Path, help="一个或多个输入文件")
    quick_run.add_argument(
        "-o",
        "--output-dir",
        required=True,
        type=Path,
        help="输出文件夹",
    )
    quick_run.add_argument(
        "-p",
        "--param",
        action="append",
        default=[],
        type=_parse_parameter,
        help="非敏感参数 key=value；可重复",
    )
    quick_run.add_argument(
        "--params-file",
        type=Path,
        help="UTF-8 JSON 参数对象；密码等敏感参数必须使用此方式",
    )
    quick_run.add_argument("--request-id", default="")
    quick_run.add_argument(
        "--expand-globs",
        action="store_true",
        help="显式展开输入路径中的通配符",
    )
    quick_run.add_argument(
        "--format",
        choices=("jsonl", "json"),
        default="jsonl",
        help="jsonl 实时输出进度；json 只输出最终结果",
    )
    quick_run.add_argument("--pretty", action="store_true")
    quick_run.add_argument(
        "--allow-root",
        action="append",
        default=[],
        type=Path,
        help="限制输入、输出和 path 参数只能位于该目录；可重复",
    )
    quick_run.add_argument(
        "--allow-source-mutation",
        action="store_true",
        help="允许会移动原文件的参数；仅在用户明确授权时使用",
    )

    install_skill = agent_sub.add_parser(
        "install-skill", help="安装或升级 Codex Skill"
    )
    install_skill.add_argument(
        "--skills-dir", type=Path, help="Codex skills 根目录，默认自动检测"
    )
    install_skill.add_argument("--force", action="store_true")
    install_skill.add_argument("--pretty", action="store_true")
    return parser


def _human_list(args: argparse.Namespace) -> int:
    operations = get_operations()
    if args.json:
        _write_json(
            catalog_payload(
                operations,
                query=args.query,
                group=args.group,
                probe_capabilities=True,
            ),
            pretty=True,
        )
        return EXIT_SUCCESS
    group_key = args.group.strip().casefold()
    groups: dict[str, list[Any]] = {}
    for operation in operations:
        if group_key and operation.group.casefold() != group_key:
            continue
        if not operation_matches_query(operation, args.query):
            continue
        groups.setdefault(operation.group, []).append(operation)
    for group, items in groups.items():
        print(f"\n[{group}]")
        for operation in items:
            capability = operation.capability()
            print(f"{operation.id:30} {capability.status:11} {operation.name}")
    return EXIT_SUCCESS


def _human_describe(args: argparse.Namespace) -> int:
    try:
        operation = operation_by_id(get_operations(), args.operation)
    except DocuForgeError as exc:
        if args.json:
            _write_json(error_payload(exc), pretty=True)
        else:
            print(str(exc), file=sys.stderr)
        return exit_code_for_error(exc)
    payload = operation_payload(operation, detailed=True)
    if args.json:
        _write_json(
            {
                "ok": True,
                "protocol_version": AGENT_PROTOCOL_VERSION,
                "operation": payload,
            },
            pretty=True,
        )
        return EXIT_SUCCESS
    capability = payload["capability"]
    print(f"{operation.name}（{operation.id}）")
    print(operation.description)
    print(f"分组：{operation.group}")
    print(
        f"引擎：{capability['engine']} · {capability['status']} · {capability['reason']}"
    )
    print(f"输入格式：{', '.join(operation.extensions) or '不限'}")
    if operation.parameters:
        print("参数：")
        for spec in operation.parameters:
            required = "必填" if spec.required else "可选"
            print(f"  {spec.key:24} {spec.kind:8} {required} · {spec.label}")
    if operation.notes:
        print(f"说明：{operation.notes}")
    return EXIT_SUCCESS


def _human_run(args: argparse.Namespace) -> int:
    operations = get_operations()
    try:
        operation = operation_by_id(operations, args.operation)
        parameters: dict[str, Any] = {}
        if args.params_file is not None:
            parameters.update(_load_json_object(args.params_file.expanduser().resolve()))
        parameters.update(dict(args.param))
        expanded_inputs: list[Path] = []
        for input_path in args.inputs:
            raw = str(input_path)
            if any(character in raw for character in "*?["):
                expanded_inputs.extend(
                    Path(item) for item in glob.glob(raw, recursive=True)
                )
            else:
                expanded_inputs.append(input_path)
    except Exception as exc:
        if args.json or args.jsonl:
            _write_json(error_payload(exc), pretty=args.json)
        else:
            print(f"处理失败：{exc}", file=sys.stderr)
        return exit_code_for_error(exc)

    runner = TaskRunner()
    request_id = "interactive-cli"
    events = _JsonEventStream() if args.jsonl else None

    def progress(value: float, message: str) -> None:
        if events is not None:
            events.progress(
                value,
                message,
                request_id=request_id,
                operation_id=operation.id,
            )
        elif not args.json:
            print(f"[{value * 100:5.1f}%] {message}")

    try:
        result = runner.run(
            operation, expanded_inputs, args.output, parameters, progress
        )
    except Exception as exc:
        payload = error_payload(
            exc, request_id=request_id, operation_id=operation.id
        )
        if args.jsonl and events is not None:
            events.emit_payload(payload)
        elif args.json:
            _write_json(payload, pretty=True)
        else:
            print(f"处理失败：{exc}", file=sys.stderr)
        return exit_code_for_error(exc)
    payload = result_payload(
        result, request_id=request_id, operation_id=operation.id
    )
    if args.jsonl and events is not None:
        events.emit_payload(payload)
    elif args.json:
        _write_json(payload, pretty=True)
    else:
        for path in result.outputs:
            print(f"输出：{path}")
        for warning in result.warnings:
            print(f"警告：{warning}")
        for failure in result.failed_inputs:
            print(
                f"失败：{failure.input_path} · {failure.message}", file=sys.stderr
            )
    return exit_code_for_result(result)


def _agent_run(
    args: argparse.Namespace,
    *,
    raw_request: Mapping[str, Any] | None = None,
    base_dir: Path | None = None,
    operations: Sequence[Any] | None = None,
    started_at: float | None = None,
) -> int:
    started_at = time.perf_counter() if started_at is None else started_at
    events = _JsonEventStream() if args.format == "jsonl" else None
    request_id = ""
    operation_id = ""
    runner = TaskRunner()
    previous_handlers: dict[int, Any] = {}
    cancel_announced = False

    def emit(payload: Any) -> None:
        if events is not None:
            events.emit_payload(payload)
        else:
            _write_json(payload, pretty=args.pretty)

    try:
        if raw_request is None:
            raw, request_base = _load_agent_request(args.request)
        else:
            raw = raw_request
            request_base = (base_dir or Path.cwd()).expanduser().resolve()
        request_id = str(raw.get("request_id", ""))
        operation_id = str(raw.get("operation", ""))
        request = prepare_agent_request(
            raw,
            list(operations) if operations is not None else get_operations(),
            base_dir=request_base,
            allow_source_mutation=args.allow_source_mutation,
            allowed_roots=args.allow_root,
        )
        preflight_seconds = round(time.perf_counter() - started_at, 3)
        request_id = request.request_id
        operation_id = request.operation.id

        if events is not None:
            events.emit(
                "accepted",
                ok=True,
                request_id=request_id,
                operation=operation_id,
                input_count=len(request.inputs),
                output_dir=str(request.output_dir),
            )

        def request_cancel(signum: int, _frame: Any) -> None:
            nonlocal cancel_announced
            runner.cancel()
            if events is not None and not cancel_announced:
                cancel_announced = True
                events.emit(
                    "cancel_requested",
                    ok=True,
                    request_id=request_id,
                    operation=operation_id,
                    signal=signum,
                    message="已请求停止；正在保留已完成输出并清理未完成文件。",
                )

        for signal_name in ("SIGINT", "SIGTERM", "SIGBREAK"):
            signal_value = getattr(signal, signal_name, None)
            if signal_value is None:
                continue
            try:
                previous_handlers[int(signal_value)] = signal.getsignal(signal_value)
                signal.signal(signal_value, request_cancel)
            except (OSError, RuntimeError, ValueError):
                pass

        def progress(value: float, message: str) -> None:
            if events is not None:
                events.progress(
                    value,
                    message,
                    request_id=request_id,
                    operation_id=operation_id,
                )

        redirect_target = sys.stderr
        redirect_context = (
            contextlib.redirect_stdout(redirect_target)
            if redirect_target is not None
            else contextlib.nullcontext()
        )
        with redirect_context:
            result = runner.run(
                request.operation,
                request.inputs,
                request.output_dir,
                request.parameters,
                progress,
                capability_override=request.capability,
            )
        result.details.setdefault("agent_preflight_seconds", preflight_seconds)
        result.details.setdefault(
            "agent_total_seconds", round(time.perf_counter() - started_at, 3)
        )
        payload = result_payload(
            result, request_id=request_id, operation_id=operation_id
        )
        emit(payload)
        return exit_code_for_result(result)
    except Exception as exc:
        payload = error_payload(
            exc, request_id=request_id, operation_id=operation_id
        )
        emit(payload)
        return exit_code_for_error(exc)
    finally:
        for signal_number, handler in previous_handlers.items():
            try:
                signal.signal(signal_number, handler)
            except (OSError, RuntimeError, ValueError):
                pass


def _emit_quick_run_error(args: argparse.Namespace, exc: Exception) -> int:
    payload = error_payload(
        exc,
        request_id=str(getattr(args, "request_id", "") or ""),
        operation_id=str(getattr(args, "operation", "") or ""),
    )
    if args.format == "jsonl":
        _JsonEventStream().emit_payload(payload)
    else:
        _write_json(payload, pretty=args.pretty)
    return exit_code_for_error(exc)


def _agent_quick_run(args: argparse.Namespace) -> int:
    """Build an in-memory Agent request and execute it in the same process."""

    started_at = time.perf_counter()
    try:
        operations = get_operations()
        operation = operation_by_id(operations, args.operation)
        parameters: dict[str, Any] = {}
        if args.params_file is not None:
            loaded = _load_json_object(args.params_file.expanduser().resolve())
            parameters.update(loaded)

        sensitive_keys = {
            spec.key for spec in operation.parameters if spec.kind == "password"
        }
        command_keys: set[str] = set()
        for key, value in args.param:
            if key in command_keys or key in parameters:
                raise ValidationError(f"参数重复：{key}")
            if key in sensitive_keys:
                raise ValidationError(
                    f"敏感参数 {key} 不能出现在命令行；请改用 --params-file"
                )
            command_keys.add(key)
            parameters[key] = value

        raw_request: dict[str, Any] = {
            "schema_version": AGENT_PROTOCOL_VERSION,
            "operation": operation.id,
            "inputs": [str(path) for path in args.inputs],
            "output_dir": str(args.output_dir),
            "parameters": parameters,
            "options": {"expand_globs": bool(args.expand_globs)},
        }
        if args.request_id:
            raw_request["request_id"] = args.request_id
    except Exception as exc:
        return _emit_quick_run_error(args, exc)

    return _agent_run(
        args,
        raw_request=raw_request,
        base_dir=Path.cwd(),
        operations=operations,
        started_at=started_at,
    )


def _agent_command(args: argparse.Namespace) -> int:
    if args.agent_command == "protocol":
        _write_json(protocol_payload(), pretty=args.pretty)
        return EXIT_SUCCESS
    if args.agent_command == "install-skill":
        try:
            payload = install_codex_skill(
                skills_directory=args.skills_dir,
                force=args.force,
            )
        except Exception as exc:
            _write_json(error_payload(exc), pretty=args.pretty)
            return exit_code_for_error(exc)
        _write_json(payload, pretty=args.pretty)
        return EXIT_SUCCESS
    if args.agent_command == "run":
        return _agent_run(args)
    if args.agent_command == "quick-run":
        return _agent_quick_run(args)

    try:
        operations = get_operations()
        if args.agent_command == "catalog":
            payload = catalog_payload(
                operations,
                query=args.query,
                group=args.group,
                runnable_only=args.runnable_only,
                detailed=args.full,
                probe_capabilities=args.probe,
            )
        elif args.agent_command == "describe":
            operation = operation_by_id(operations, args.operation)
            payload = {
                "ok": True,
                "protocol_version": AGENT_PROTOCOL_VERSION,
                "operation": operation_payload(operation, detailed=True),
            }
        elif args.agent_command == "validate":
            raw, base_dir = _load_agent_request(args.request)
            request = prepare_agent_request(
                raw,
                operations,
                base_dir=base_dir,
                allow_source_mutation=args.allow_source_mutation,
                allowed_roots=args.allow_root,
            )
            payload = validation_payload(request)
        else:
            raise ValidationError(f"未知 Agent 命令：{args.agent_command}")
    except Exception as exc:
        _write_json(error_payload(exc), pretty=getattr(args, "pretty", False))
        return exit_code_for_error(exc)
    _write_json(payload, pretty=getattr(args, "pretty", False))
    return EXIT_SUCCESS


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    try:
        args = build_parser().parse_args(argv)
        if args.command == "list":
            return _human_list(args)
        if args.command == "describe":
            return _human_describe(args)
        if args.command == "run":
            return _human_run(args)
        if args.command == "agent":
            return _agent_command(args)
        raise ValidationError(f"未知命令：{args.command}")
    except BrokenPipeError:
        return EXIT_SUCCESS
    except KeyboardInterrupt:
        return EXIT_CANCELLED


def agent_main(argv: list[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    return main(["agent", *arguments])


if __name__ == "__main__":
    raise SystemExit(main())
