from __future__ import annotations

import argparse
import glob
import json
import sys
from pathlib import Path
from typing import Any

from .models import DocuForgeError
from .registry import get_operations
from .runner import TaskRunner


def _parse_parameter(value: str) -> tuple[str, Any]:
    if "=" not in value:
        raise argparse.ArgumentTypeError("参数格式必须是 key=value")
    key, raw = value.split("=", 1)
    return key.strip(), raw


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="layoutloom", description="页织工坊（LayoutLoom）命令行"
    )
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("list", help="列出全部处理任务")
    run = sub.add_parser("run", help="运行一个处理任务")
    run.add_argument("operation", help="任务 ID，可通过 list 查看")
    run.add_argument("inputs", nargs="+", type=Path)
    run.add_argument("-o", "--output", type=Path, required=True)
    run.add_argument(
        "-p", "--param", action="append", default=[], type=_parse_parameter
    )
    run.add_argument("--json", action="store_true", help="以 JSON 输出结果")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    operations = get_operations()
    by_id = {item.id: item for item in operations}
    if args.command == "list":
        groups: dict[str, list[Any]] = {}
        for operation in operations:
            groups.setdefault(operation.group, []).append(operation)
        for group, items in groups.items():
            print(f"\n[{group}]")
            for operation in items:
                capability = operation.capability()
                print(f"{operation.id:30} {capability.status:11} {operation.name}")
        return 0

    operation = by_id.get(args.operation)
    if operation is None:
        if args.json:
            print(
                json.dumps(
                    {
                        "ok": False,
                        "error": {
                            "code": "unknown_operation",
                            "message": f"未知任务：{args.operation}",
                        },
                    },
                    ensure_ascii=False,
                )
            )
        else:
            print(f"未知任务：{args.operation}", file=sys.stderr)
        return 2
    parameters = dict(args.param)
    expanded_inputs: list[Path] = []
    for input_path in args.inputs:
        raw = str(input_path)
        if any(character in raw for character in "*?["):
            expanded_inputs.extend(
                Path(item) for item in glob.glob(raw, recursive=True)
            )
        else:
            expanded_inputs.append(input_path)

    def progress(value: float, message: str) -> None:
        if not args.json:
            print(f"[{value * 100:5.1f}%] {message}")

    try:
        result = TaskRunner().run(
            operation, expanded_inputs, args.output, parameters, progress
        )
    except DocuForgeError as exc:
        if args.json:
            print(
                json.dumps(
                    {
                        "ok": False,
                        "error": {
                            "code": type(exc).__name__,
                            "message": str(exc),
                        },
                    },
                    ensure_ascii=False,
                )
            )
        else:
            print(f"处理失败：{exc}", file=sys.stderr)
        return 1
    except Exception:
        if args.json:
            print(
                json.dumps(
                    {
                        "ok": False,
                        "error": {
                            "code": "internal_error",
                            "message": "处理时发生未预期错误，请检查输入文件和运行日志。",
                        },
                    },
                    ensure_ascii=False,
                )
            )
        else:
            print("处理失败：发生未预期错误。", file=sys.stderr)
        return 1
    if args.json:
        print(
            json.dumps(
                {
                    "ok": True,
                    "outputs": [str(path) for path in result.outputs],
                    "warnings": result.warnings,
                    "details": result.details,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    else:
        for path in result.outputs:
            print(f"输出：{path}")
        for warning in result.warnings:
            print(f"警告：{warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
