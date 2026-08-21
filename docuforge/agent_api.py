from __future__ import annotations

import glob
import json
import math
import os
import re
import shutil
import sys
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence

from . import __version__
from .models import (
    CancelledError,
    Capability,
    DocuForgeError,
    MissingEngineError,
    Operation,
    ParameterSpec,
    TaskFailure,
    TaskResult,
    ValidationError,
)
from .utils import validate_inputs


AGENT_PROTOCOL_VERSION = "1.0"
EXIT_SUCCESS = 0
EXIT_USAGE_ERROR = 2
EXIT_ENGINE_UNAVAILABLE = 3
EXIT_RUNTIME_ERROR = 4
EXIT_PARTIAL = 5
EXIT_INTERNAL_ERROR = 70
EXIT_CANCELLED = 130


@dataclass(frozen=True)
class PreparedAgentRequest:
    request_id: str
    operation: Operation
    inputs: tuple[Path, ...]
    output_dir: Path
    parameters: dict[str, Any]
    capability: Capability
    base_dir: Path


def utc_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace(
        "+00:00", "Z"
    )


def json_safe(value: Any) -> Any:
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, TaskFailure):
        return {
            "input": str(value.input_path),
            "error_type": value.error_type,
            "message": value.message,
        }
    if isinstance(value, Mapping):
        return {str(key): json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set, frozenset)):
        return [json_safe(item) for item in value]
    return str(value)


def dumps_json(value: Any, *, pretty: bool = False) -> str:
    return json.dumps(
        json_safe(value),
        ensure_ascii=False,
        allow_nan=False,
        indent=2 if pretty else None,
        separators=None if pretty else (",", ":"),
    )


def protocol_payload() -> dict[str, Any]:
    return {
        "ok": True,
        "protocol": {
            "name": "layoutloom-agent",
            "version": AGENT_PROTOCOL_VERSION,
            "application_version": __version__,
            "transport": "stdio",
            "request_encoding": "UTF-8 JSON",
            "event_encoding": "UTF-8 JSON Lines",
        },
        "commands": [
            "catalog",
            "describe",
            "validate",
            "run",
            "quick-run",
            "install-skill",
        ],
        "exit_codes": {
            str(EXIT_SUCCESS): "success",
            str(EXIT_RUNTIME_ERROR): "runtime_error",
            str(EXIT_USAGE_ERROR): "invalid_request_or_usage",
            str(EXIT_PARTIAL): "partial_success",
            str(EXIT_ENGINE_UNAVAILABLE): "engine_unavailable",
            str(EXIT_INTERNAL_ERROR): "internal_error",
            str(EXIT_CANCELLED): "cancelled",
        },
    }


def _parameter_payload(spec: ParameterSpec) -> dict[str, Any]:
    default = None if spec.kind == "password" else json_safe(spec.default)
    visible_when = None
    if spec.visible_when is not None:
        key, values = spec.visible_when
        visible_when = {"parameter": key, "values": list(values)}
    return {
        "key": spec.key,
        "label": spec.label,
        "kind": spec.kind,
        "required": spec.required,
        "default": default,
        "choices": [
            {"value": value, "label": label} for value, label in spec.choices
        ],
        "help": spec.help_text,
        "minimum": spec.minimum,
        "maximum": spec.maximum,
        "section": spec.section,
        "advanced": spec.advanced,
        "visible_when": visible_when,
        "sensitive": spec.kind == "password",
    }


def _capability_payload(capability: Capability) -> dict[str, Any]:
    return {
        "status": capability.status,
        "runnable": capability.runnable,
        "reason": capability.reason,
        "engine": capability.engine,
    }


def operation_payload(
    operation: Operation,
    *,
    detailed: bool,
    capability: Capability | None = None,
    probe_capability: bool = True,
) -> dict[str, Any]:
    capability_payload: dict[str, Any]
    if probe_capability:
        capability_payload = _capability_payload(capability or operation.capability())
    else:
        capability_payload = {
            "status": "not_probed",
            "runnable": None,
            "reason": "尚未探测；使用 describe 或 validate 检查本机引擎",
            "engine": "",
        }
    payload: dict[str, Any] = {
        "id": operation.id,
        "group": operation.group,
        "name": operation.name,
        "description": operation.description,
        "fidelity": operation.fidelity,
        "capability": capability_payload,
        "inputs": {
            "extensions": list(operation.extensions),
            "minimum": operation.min_inputs,
            "maximum": operation.max_inputs,
            "independent": operation.independent_inputs,
        },
        "interaction": (
            {
                "mode": "prebuilt_plan_required",
                "reason": "该任务需要先在 GUI 补修工作台生成页面修复方案",
            }
            if operation.id == "video.repair_slides_ppt"
            else {"mode": "automatic"}
        ),
    }
    if detailed:
        payload.update(
            {
                "notes": operation.notes,
                "parameters": [
                    _parameter_payload(spec) for spec in operation.parameters
                ],
                "output_policy": {
                    "allows_empty": operation.allow_empty_outputs,
                    "allows_external_paths": operation.allow_external_outputs,
                },
                "input_safety": {
                    "rejects_encrypted_pdf": operation.reject_encrypted_pdf_inputs,
                    "rejects_signed_pdf": operation.reject_signed_pdf_inputs,
                },
            }
        )
    return payload


def operation_matches_query(operation: Operation, query: str) -> bool:
    """Match human/agent search terms without requiring punctuation to agree."""

    query_key = str(query).strip().casefold()
    if not query_key:
        return True
    haystack = "\n".join(
        (
            operation.id,
            operation.group,
            operation.name,
            operation.description,
            operation.notes,
        )
    ).casefold()
    if query_key in haystack:
        return True
    terms = tuple(
        term
        for term in re.split(r"[^0-9a-z\u3400-\u9fff]+", query_key)
        if term
    )
    return bool(terms) and all(term in haystack for term in terms)


def catalog_payload(
    operations: Sequence[Operation],
    *,
    query: str = "",
    group: str = "",
    runnable_only: bool = False,
    detailed: bool = False,
    probe_capabilities: bool = False,
) -> dict[str, Any]:
    group_key = group.strip().casefold()
    selected: list[dict[str, Any]] = []
    for operation in operations:
        if group_key and operation.group.casefold() != group_key:
            continue
        if not operation_matches_query(operation, query):
            continue
        should_probe = probe_capabilities or runnable_only
        capability = operation.capability() if should_probe else None
        if runnable_only and capability is not None and not capability.runnable:
            continue
        selected.append(
            operation_payload(
                operation,
                detailed=detailed,
                capability=capability,
                probe_capability=should_probe,
            )
        )
    return {
        "ok": True,
        "protocol_version": AGENT_PROTOCOL_VERSION,
        "application_version": __version__,
        "count": len(selected),
        "operations": selected,
    }


def operation_by_id(
    operations: Sequence[Operation], operation_id: str
) -> Operation:
    normalized = str(operation_id).strip()
    for operation in operations:
        if operation.id == normalized:
            return operation
    raise ValidationError(f"未知任务：{normalized or '（空）'}")


def _resolved_path(raw: str, base_dir: Path) -> Path:
    path = Path(raw).expanduser()
    if not path.is_absolute():
        path = base_dir / path
    return path.resolve()


def _resolve_inputs(
    raw_inputs: Sequence[Any], base_dir: Path, *, expand_globs: bool
) -> list[Path]:
    resolved: list[Path] = []
    for raw_value in raw_inputs:
        if not isinstance(raw_value, (str, os.PathLike)):
            raise ValidationError("inputs 中的每一项都必须是文件路径字符串")
        raw = os.fspath(raw_value).strip()
        if not raw:
            raise ValidationError("inputs 不能包含空路径")
        candidate = _resolved_path(raw, base_dir)
        if expand_globs and any(character in raw for character in "*?["):
            matches = sorted(
                Path(item).resolve()
                for item in glob.glob(str(candidate), recursive=True)
            )
            if not matches:
                raise ValidationError(f"通配符没有匹配任何文件：{raw}")
            resolved.extend(matches)
        else:
            resolved.append(candidate)
    return resolved


_OFFICE_ENGINE_PARAMETERS = {
    "word.to_pdf": "engine",
    "word.full_compatibility": "verification_engine",
    "excel.to_pdf": "engine",
    "ppt.to_pdf": "engine",
    "legacy.doc_to_docx": "engine",
    "legacy.xls_to_xlsx": "engine",
    "ppt.to_images": "renderer",
}


def request_capability(
    operation: Operation,
    parameters: Mapping[str, Any],
    inputs: Sequence[Path],
) -> Capability:
    """Return the capability for the selected mode/engine, not only its default."""

    if operation.id == "pdf.to_word" and parameters.get("mode") == "office_native":
        from .engines import microsoft_word_capability

        return microsoft_word_capability()

    engine_parameter = _OFFICE_ENGINE_PARAMETERS.get(operation.id)
    if engine_parameter is not None:
        requested = str(parameters.get(engine_parameter, "auto")).strip().casefold()
        if requested == "none":
            return operation.capability()
        from .processors.office import _select_conversion_engine

        selected = _select_conversion_engine(inputs[0], requested)
        labels = {
            "wps": "WPS Office COM",
            "microsoft_office": "Microsoft Office COM",
            "libreoffice": "LibreOffice",
        }
        label = labels.get(selected, selected)
        return Capability("external", f"已按请求选择 {label}", label)

    return operation.capability()


def prepare_agent_request(
    payload: Mapping[str, Any],
    operations: Sequence[Operation],
    *,
    base_dir: str | Path | None = None,
    allow_source_mutation: bool = False,
    allowed_roots: Sequence[str | Path] = (),
) -> PreparedAgentRequest:
    if not isinstance(payload, Mapping):
        raise ValidationError("请求必须是一个 JSON 对象")
    allowed = {
        "schema_version",
        "request_id",
        "operation",
        "inputs",
        "output_dir",
        "parameters",
        "options",
    }
    unknown = sorted(str(key) for key in payload if key not in allowed)
    if unknown:
        raise ValidationError(f"请求包含未知字段：{unknown[0]}")
    raw_schema_version = payload.get("schema_version", AGENT_PROTOCOL_VERSION)
    if isinstance(raw_schema_version, bool) or not isinstance(
        raw_schema_version, (str, int, float)
    ):
        raise ValidationError("schema_version 必须是字符串或数字")
    schema_version = str(raw_schema_version)
    if schema_version not in {"1", AGENT_PROTOCOL_VERSION}:
        raise ValidationError(
            f"不支持的 Agent 协议版本：{schema_version}；当前为 {AGENT_PROTOCOL_VERSION}"
        )

    raw_operation = payload.get("operation")
    if not isinstance(raw_operation, str) or not raw_operation.strip():
        raise ValidationError("operation 必须是非空任务 ID 字符串")
    operation = operation_by_id(operations, raw_operation)
    raw_inputs = payload.get("inputs")
    if not isinstance(raw_inputs, list):
        raise ValidationError("inputs 必须是文件路径数组")
    options = payload.get("options", {})
    if not isinstance(options, Mapping):
        raise ValidationError("options 必须是 JSON 对象")
    unknown_options = sorted(
        str(key) for key in options if key not in {"expand_globs"}
    )
    if unknown_options:
        raise ValidationError(f"options 包含未知字段：{unknown_options[0]}")
    expand_globs = options.get("expand_globs", False)
    if not isinstance(expand_globs, bool):
        raise ValidationError("options.expand_globs 必须是 JSON 布尔值")

    request_base = Path(base_dir or Path.cwd()).expanduser().resolve()
    inputs = validate_inputs(
        _resolve_inputs(raw_inputs, request_base, expand_globs=expand_globs),
        operation.extensions,
        operation.min_inputs,
        operation.max_inputs,
    )

    raw_output = payload.get("output_dir")
    if not isinstance(raw_output, (str, os.PathLike)) or not os.fspath(
        raw_output
    ).strip():
        raise ValidationError("output_dir 必须是非空文件夹路径")
    output_dir = _resolved_path(os.fspath(raw_output), request_base)
    if output_dir.exists() and not output_dir.is_dir():
        raise ValidationError(f"输出位置不是文件夹：{output_dir}")

    raw_parameters = payload.get("parameters", {})
    if not isinstance(raw_parameters, Mapping):
        raise ValidationError("parameters 必须是 JSON 对象")
    parameters = operation.normalize_parameters(raw_parameters)
    for spec in operation.parameters:
        if spec.kind != "path":
            continue
        value = parameters.get(spec.key)
        if value is not None:
            parameters[spec.key] = _resolved_path(str(value), request_base)
    if operation.id == "image.rename" and parameters.get("move"):
        if not allow_source_mutation:
            raise ValidationError(
                "Agent 安全模式默认禁止 image.rename 的 move=true，"
                "因为它会移动原文件；仅在用户明确授权后添加 --allow-source-mutation"
            )

    capability = request_capability(operation, parameters, inputs)
    if not capability.runnable:
        raise MissingEngineError(capability.reason)

    roots = [Path(root).expanduser().resolve() for root in allowed_roots]
    if roots:
        guarded_paths: list[Path] = [*inputs, output_dir]
        path_parameter_keys = {
            spec.key for spec in operation.parameters if spec.kind == "path"
        }
        guarded_paths.extend(
            value.resolve()
            for key, value in parameters.items()
            if key in path_parameter_keys and isinstance(value, Path)
        )
        for guarded in guarded_paths:
            if not any(guarded == root or guarded.is_relative_to(root) for root in roots):
                raise ValidationError(f"路径超出 Agent 获准范围：{guarded}")

    raw_request_id = payload.get("request_id", "")
    if raw_request_id is not None and not isinstance(raw_request_id, str):
        raise ValidationError("request_id 必须是字符串")
    request_id = str(raw_request_id or "").strip()
    if not request_id:
        request_id = f"layoutloom-{uuid.uuid4().hex[:16]}"
    if len(request_id) > 128 or any(ord(character) < 32 for character in request_id):
        raise ValidationError("request_id 必须是 1 至 128 个可显示字符")

    return PreparedAgentRequest(
        request_id=request_id,
        operation=operation,
        inputs=tuple(inputs),
        output_dir=output_dir,
        parameters=parameters,
        capability=capability,
        base_dir=request_base,
    )


def validation_payload(request: PreparedAgentRequest) -> dict[str, Any]:
    sensitive_keys = {
        spec.key for spec in request.operation.parameters if spec.kind == "password"
    }
    return {
        "ok": True,
        "protocol_version": AGENT_PROTOCOL_VERSION,
        "request_id": request.request_id,
        "operation": request.operation.id,
        "inputs": [str(path) for path in request.inputs],
        "output_dir": str(request.output_dir),
        "parameter_keys": sorted(request.parameters),
        "sensitive_parameter_keys": sorted(sensitive_keys & request.parameters.keys()),
        "capability": _capability_payload(request.capability),
        "message": (
            "请求结构、文件、参数和本机处理引擎检查通过；尚未创建输出。"
            "运行时仍会执行 PDF 签名/加密、输出锁和最终文件完整性检查。"
        ),
    }


def result_payload(
    result: TaskResult,
    *,
    request_id: str,
    operation_id: str,
) -> dict[str, Any]:
    outcome = result.outcome
    return {
        "ok": outcome == "success",
        "event": "result",
        "protocol_version": AGENT_PROTOCOL_VERSION,
        "timestamp": utc_timestamp(),
        "request_id": request_id,
        "operation": operation_id,
        "outcome": outcome,
        "outputs": [str(path) for path in result.outputs],
        "warnings": list(result.warnings),
        "details": json_safe(result.details),
        "completed_inputs": [str(path) for path in result.completed_inputs],
        "failed_inputs": [json_safe(item) for item in result.failed_inputs],
        "cancelled_inputs": [str(path) for path in result.cancelled_inputs],
        "cancelled": result.cancelled,
    }


def exit_code_for_result(result: TaskResult) -> int:
    if result.outcome == "success":
        return EXIT_SUCCESS
    if result.outcome == "partial":
        return EXIT_PARTIAL
    if result.outcome == "cancelled":
        return EXIT_CANCELLED
    return EXIT_RUNTIME_ERROR


def error_code(exc: BaseException) -> str:
    if isinstance(exc, CancelledError):
        return "cancelled"
    if isinstance(exc, MissingEngineError):
        return "engine_unavailable"
    if isinstance(exc, ValidationError):
        message = str(exc)
        if message.startswith("未知任务："):
            return "unknown_operation"
        return "validation_error"
    if isinstance(exc, json.JSONDecodeError):
        return "invalid_json"
    if isinstance(exc, DocuForgeError):
        return "processing_error"
    return "internal_error"


def exit_code_for_error(exc: BaseException) -> int:
    if isinstance(exc, CancelledError):
        return EXIT_CANCELLED
    if isinstance(exc, MissingEngineError):
        return EXIT_ENGINE_UNAVAILABLE
    if isinstance(exc, (ValidationError, json.JSONDecodeError, ValueError)):
        return EXIT_USAGE_ERROR
    if isinstance(exc, DocuForgeError):
        return EXIT_RUNTIME_ERROR
    return EXIT_INTERNAL_ERROR


def error_payload(
    exc: BaseException,
    *,
    request_id: str = "",
    operation_id: str = "",
) -> dict[str, Any]:
    message = str(exc).strip()
    if not message:
        message = "处理时发生未预期错误，请检查请求、输入文件和本机环境。"
    payload: dict[str, Any] = {
        "ok": False,
        "event": "error",
        "protocol_version": AGENT_PROTOCOL_VERSION,
        "timestamp": utc_timestamp(),
        "request_id": request_id,
        "operation": operation_id,
        "error": {
            "code": error_code(exc),
            "type": type(exc).__name__,
            "message": message,
        },
    }
    if isinstance(exc, CancelledError) and exc.result is not None:
        payload["partial_result"] = result_payload(
            exc.result,
            request_id=request_id,
            operation_id=operation_id,
        )
    return payload


def _bundled_skill_directory() -> Path:
    candidates: list[Path] = []
    if getattr(sys, "frozen", False):
        candidates.append(
            Path(sys.executable).resolve().parent
            / "agent_skill"
            / "layoutloom-agent"
        )
    project_root = Path(__file__).resolve().parents[1]
    candidates.append(project_root / "integrations" / "codex" / "layoutloom-agent")
    for candidate in candidates:
        if (candidate / "SKILL.md").is_file():
            return candidate
    raise DocuForgeError("当前安装中缺少 LayoutLoom Agent Skill 文件")


def _default_codex_skills_directory() -> Path:
    codex_home = os.environ.get("CODEX_HOME")
    root = Path(codex_home).expanduser() if codex_home else Path.home() / ".codex"
    return root.resolve() / "skills"


def _current_cli_command() -> list[str]:
    executable = Path(sys.executable).resolve()
    if getattr(sys, "frozen", False):
        return [str(executable)]
    return [str(executable), "-m", "docuforge.cli"]


def install_codex_skill(
    *,
    skills_directory: str | Path | None = None,
    force: bool = False,
) -> dict[str, Any]:
    source = _bundled_skill_directory().resolve()
    root = Path(skills_directory or _default_codex_skills_directory()).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    target = (root / "layoutloom-agent").resolve()
    if target.parent != root:
        raise ValidationError("Skill 安装目标无效")
    if target == source:
        raise ValidationError("Skill 已位于目标目录，无需重复安装")
    if target.exists() and not force:
        raise ValidationError(
            f"Skill 已存在：{target}；确认覆盖时请添加 --force"
        )

    temporary = root / f".layoutloom-agent.install-{uuid.uuid4().hex}"
    backup: Path | None = None
    try:
        shutil.copytree(source, temporary)
        for cache in temporary.rglob("__pycache__"):
            shutil.rmtree(cache, ignore_errors=True)
        config = {
            "schema_version": AGENT_PROTOCOL_VERSION,
            "application_version": __version__,
            "command": _current_cli_command(),
        }
        (temporary / "layoutloom-cli.json").write_text(
            dumps_json(config, pretty=True) + "\n", encoding="utf-8"
        )
        if target.exists():
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            backup = root / f"layoutloom-agent.backup-{stamp}"
            suffix = 1
            while backup.exists():
                backup = root / f"layoutloom-agent.backup-{stamp}-{suffix}"
                suffix += 1
            target.replace(backup)
        temporary.replace(target)
    except Exception:
        shutil.rmtree(temporary, ignore_errors=True)
        if backup is not None and backup.exists() and not target.exists():
            backup.replace(target)
        raise
    return {
        "ok": True,
        "protocol_version": AGENT_PROTOCOL_VERSION,
        "skill": "layoutloom-agent",
        "installed_to": str(target),
        "backup": str(backup) if backup is not None else None,
        "command": _current_cli_command(),
        "restart_required": True,
        "message": "Skill 已安装；重新打开 Codex 后即可自动发现。",
    }
