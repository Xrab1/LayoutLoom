from __future__ import annotations

import math
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Literal, Mapping, Sequence


class DocuForgeError(Exception):
    """Base exception shown to end users without a Python traceback."""


class ValidationError(DocuForgeError):
    pass


class MissingEngineError(DocuForgeError):
    pass


class CancelledError(DocuForgeError):
    def __init__(self, message: str = "任务已取消", result: "TaskResult | None" = None):
        super().__init__(message)
        self.result = result


ParameterKind = Literal[
    "text",
    "password",
    "integer",
    "number",
    "boolean",
    "choice",
    "path",
    "color",
    "colors",
    "region",
]


@dataclass(frozen=True)
class ParameterSpec:
    key: str
    label: str
    kind: ParameterKind = "text"
    default: Any = ""
    required: bool = False
    choices: tuple[tuple[str, str], ...] = ()
    help_text: str = ""
    minimum: float | None = None
    maximum: float | None = None
    section: str = ""
    advanced: bool = False
    visible_when: tuple[str, tuple[str, ...]] | None = None

    def normalize(self, raw: Any) -> Any:
        if raw is None or raw == "":
            if self.required and self.default in (None, ""):
                raise ValidationError(f"请填写“{self.label}”")
            raw = self.default

        if self.kind == "boolean":
            if isinstance(raw, bool):
                return raw
            text = str(raw).strip().lower()
            if text in {"1", "true", "yes", "on", "是"}:
                return True
            if text in {"0", "false", "no", "off", "否"}:
                return False
            raise ValidationError(f"“{self.label}”必须是 true/false、yes/no、on/off 或 1/0")
        if self.kind == "integer":
            if isinstance(raw, bool):
                raise ValidationError(f"“{self.label}”必须是整数")
            try:
                value = int(raw)
            except (TypeError, ValueError, OverflowError) as exc:
                raise ValidationError(f"“{self.label}”必须是整数") from exc
            self._validate_bounds(value)
            return value
        if self.kind == "number":
            if isinstance(raw, bool):
                raise ValidationError(f"“{self.label}”必须是有限数字")
            try:
                value = float(raw)
            except (TypeError, ValueError) as exc:
                raise ValidationError(f"“{self.label}”必须是数字") from exc
            if not math.isfinite(value):
                raise ValidationError(f"“{self.label}”必须是有限数字")
            self._validate_bounds(value)
            return value
        if self.kind == "choice":
            value = str(raw)
            allowed = {item[0] for item in self.choices}
            if allowed and value not in allowed:
                raise ValidationError(f"“{self.label}”选项无效")
            return value
        if self.kind == "path":
            return Path(str(raw)).expanduser() if raw else None
        return str(raw)

    def _validate_bounds(self, value: float) -> None:
        if self.minimum is not None and value < self.minimum:
            raise ValidationError(f"“{self.label}”不能小于 {self.minimum:g}")
        if self.maximum is not None and value > self.maximum:
            raise ValidationError(f"“{self.label}”不能大于 {self.maximum:g}")


@dataclass(frozen=True)
class Capability:
    status: Literal["ready", "external", "unavailable"] = "ready"
    reason: str = "本机可直接处理"
    engine: str = "内置引擎"

    @property
    def runnable(self) -> bool:
        return self.status != "unavailable"


@dataclass
class TaskFailure:
    input_path: Path
    error_type: str
    message: str


@dataclass
class TaskResult:
    outputs: list[Path] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    details: dict[str, Any] = field(default_factory=dict)
    completed_inputs: list[Path] = field(default_factory=list)
    failed_inputs: list[TaskFailure] = field(default_factory=list)
    cancelled_inputs: list[Path] = field(default_factory=list)
    cancelled: bool = False

    @property
    def outcome(self) -> Literal["success", "partial", "failure", "cancelled"]:
        """Return the user-facing aggregate outcome for this task."""

        if self.cancelled and not self.outputs and not self.completed_inputs:
            return "cancelled"
        if self.failed_inputs:
            return "partial" if self.outputs or self.completed_inputs else "failure"
        if self.cancelled:
            return "partial"
        return "success"


Handler = Callable[
    [Sequence[Path], Path, Mapping[str, Any]], TaskResult | Sequence[Path] | Path
]
CapabilityProbe = Callable[[], Capability]


@dataclass(frozen=True)
class Operation:
    id: str
    group: str
    name: str
    description: str
    handler: Handler
    extensions: tuple[str, ...] = ()
    parameters: tuple[ParameterSpec, ...] = ()
    min_inputs: int = 1
    max_inputs: int | None = None
    fidelity: Literal[
        "lossless", "visual", "editable", "extract", "transform"
    ] = "transform"
    capability_probe: CapabilityProbe | None = None
    notes: str = ""
    allow_empty_outputs: bool = False
    allow_external_outputs: bool = False
    reject_encrypted_pdf_inputs: bool = False
    reject_signed_pdf_inputs: bool = False
    independent_inputs: bool = False

    def capability(self) -> Capability:
        if self.capability_probe is None:
            return Capability()
        try:
            return self.capability_probe()
        except Exception as exc:  # capability checks must never crash the UI
            return Capability("unavailable", f"引擎检测失败：{exc}", "检测器")

    def normalize_parameters(self, raw: Mapping[str, Any]) -> dict[str, Any]:
        allowed = {spec.key for spec in self.parameters}
        unknown = sorted(str(key) for key in raw if key not in allowed)
        if unknown:
            raise ValidationError(f"未知参数：{unknown[0]}")
        return {spec.key: spec.normalize(raw.get(spec.key)) for spec in self.parameters}


def coerce_result(value: TaskResult | Sequence[Path] | Path) -> TaskResult:
    if isinstance(value, TaskResult):
        return value
    if isinstance(value, Path):
        return TaskResult(outputs=[value])
    return TaskResult(outputs=[Path(item) for item in value])
