from __future__ import annotations

import json
from pathlib import Path

import pytest

from docuforge import agent_api
from docuforge.models import (
    Capability,
    MissingEngineError,
    Operation,
    ParameterSpec,
    TaskFailure,
    TaskResult,
    ValidationError,
)
from docuforge.registry import CORE_OPERATION_IDS, get_operations


def _operation(
    operation_id: str = "test.copy",
    *,
    parameters: tuple[ParameterSpec, ...] = (),
    probe=None,
) -> Operation:
    return Operation(
        operation_id,
        "测试",
        "测试任务",
        "用于 Agent API 测试",
        lambda paths, output, params: [],
        (".txt",),
        parameters,
        capability_probe=probe,
        allow_empty_outputs=True,
    )


def test_complete_catalog_is_json_serializable_without_probing() -> None:
    payload = agent_api.catalog_payload(get_operations(), detailed=True)
    assert payload["count"] == len(CORE_OPERATION_IDS) == 73
    assert all(
        operation["capability"]["status"] == "not_probed"
        for operation in payload["operations"]
    )
    json.dumps(payload, ensure_ascii=False, allow_nan=False)


def test_catalog_only_probes_when_requested() -> None:
    calls = 0

    def probe() -> Capability:
        nonlocal calls
        calls += 1
        return Capability("ready", "可运行", "测试引擎")

    operation = _operation(probe=probe)
    unprobed = agent_api.catalog_payload([operation])
    assert calls == 0
    assert unprobed["operations"][0]["capability"]["runnable"] is None

    probed = agent_api.catalog_payload([operation], probe_capabilities=True)
    assert calls == 1
    assert probed["operations"][0]["capability"]["runnable"] is True


def test_catalog_query_matches_multiple_terms_across_punctuation() -> None:
    operation = _operation("pdf.to_word")
    payload = agent_api.catalog_payload([operation], query="PDF Word")
    assert payload["count"] == 1


def test_operation_schema_redacts_password_and_marks_gui_plan() -> None:
    operation = _operation(
        "video.repair_slides_ppt",
        parameters=(ParameterSpec("password", "密码", "password", "secret"),),
    )
    payload = agent_api.operation_payload(operation, detailed=True)
    assert payload["parameters"][0]["default"] is None
    assert payload["parameters"][0]["sensitive"] is True
    assert payload["interaction"]["mode"] == "prebuilt_plan_required"


def test_prepare_request_resolves_relative_paths_and_normalizes_parameters(
    tmp_path: Path,
) -> None:
    source = tmp_path / "input.txt"
    source.write_text("hello", encoding="utf-8")
    operation = _operation(
        parameters=(ParameterSpec("count", "数量", "integer", 1, minimum=1),)
    )
    request = agent_api.prepare_agent_request(
        {
            "operation": operation.id,
            "inputs": [source.name],
            "output_dir": "output",
            "parameters": {"count": "3"},
        },
        [operation],
        base_dir=tmp_path,
    )
    assert request.inputs == (source.resolve(),)
    assert request.output_dir == (tmp_path / "output").resolve()
    assert request.parameters["count"] == 3
    assert not request.output_dir.exists()


def test_prepare_request_resolves_relative_path_parameters(tmp_path: Path) -> None:
    source = tmp_path / "input.txt"
    source.write_text("hello", encoding="utf-8")
    asset = tmp_path / "assets" / "watermark.png"
    asset.parent.mkdir()
    asset.write_bytes(b"asset")
    operation = _operation(
        parameters=(ParameterSpec("asset", "资源", "path", None),)
    )
    request = agent_api.prepare_agent_request(
        {
            "operation": operation.id,
            "inputs": [source.name],
            "output_dir": "output",
            "parameters": {"asset": "assets/watermark.png"},
        },
        [operation],
        base_dir=tmp_path,
        allowed_roots=[tmp_path],
    )
    assert request.parameters["asset"] == asset.resolve()


def test_prepare_request_rejects_non_boolean_glob_option(tmp_path: Path) -> None:
    source = tmp_path / "input.txt"
    source.write_text("hello", encoding="utf-8")
    with pytest.raises(ValidationError, match="JSON 布尔值"):
        agent_api.prepare_agent_request(
            {
                "operation": "test.copy",
                "inputs": [str(source)],
                "output_dir": str(tmp_path / "output"),
                "options": {"expand_globs": "false"},
            },
            [_operation()],
        )


def test_agent_mode_blocks_source_mutation_until_host_authorizes_it(
    tmp_path: Path,
) -> None:
    source = tmp_path / "input.txt"
    source.write_text("hello", encoding="utf-8")
    operation = _operation(
        "image.rename",
        parameters=(ParameterSpec("move", "移动原文件", "boolean", False),),
    )
    payload = {
        "operation": operation.id,
        "inputs": [str(source)],
        "output_dir": str(tmp_path / "output"),
        "parameters": {"move": True},
    }
    with pytest.raises(ValidationError, match="allow-source-mutation"):
        agent_api.prepare_agent_request(payload, [operation])
    prepared = agent_api.prepare_agent_request(
        payload, [operation], allow_source_mutation=True
    )
    assert prepared.parameters["move"] is True


def test_allowed_roots_guard_inputs_outputs_and_path_parameters(tmp_path: Path) -> None:
    allowed = tmp_path / "allowed"
    outside = tmp_path / "outside"
    allowed.mkdir()
    outside.mkdir()
    source = allowed / "input.txt"
    source.write_text("hello", encoding="utf-8")
    asset = outside / "asset.txt"
    asset.write_text("asset", encoding="utf-8")
    operation = _operation(
        parameters=(ParameterSpec("asset", "资源", "path", None),)
    )
    with pytest.raises(ValidationError, match="获准范围"):
        agent_api.prepare_agent_request(
            {
                "operation": operation.id,
                "inputs": [str(source)],
                "output_dir": str(allowed / "output"),
                "parameters": {"asset": str(asset)},
            },
            [operation],
            allowed_roots=[allowed],
        )


def test_office_native_pdf_mode_uses_parameter_aware_word_probe(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "input.txt"
    source.write_text("hello", encoding="utf-8")
    operation = _operation(
        "pdf.to_word",
        parameters=(
            ParameterSpec(
                "mode",
                "模式",
                "choice",
                "hybrid",
                choices=(("hybrid", "混合"), ("office_native", "Word 原生")),
            ),
        ),
    )
    from docuforge import engines

    monkeypatch.setattr(
        engines,
        "microsoft_word_capability",
        lambda: Capability("unavailable", "未安装 Word", "Microsoft Word"),
    )
    with pytest.raises(MissingEngineError, match="未安装 Word"):
        agent_api.prepare_agent_request(
            {
                "operation": operation.id,
                "inputs": [str(source)],
                "output_dir": str(tmp_path / "output"),
                "parameters": {"mode": "office_native"},
            },
            [operation],
        )


def test_partial_result_has_stable_payload_and_exit_code(tmp_path: Path) -> None:
    output = tmp_path / "done.txt"
    output.write_text("done", encoding="utf-8")
    result = TaskResult(
        outputs=[output],
        completed_inputs=[tmp_path / "good.txt"],
        failed_inputs=[
            TaskFailure(tmp_path / "bad.txt", "ValidationError", "bad input")
        ],
    )
    payload = agent_api.result_payload(
        result, request_id="request-1", operation_id="test.copy"
    )
    assert payload["outcome"] == "partial"
    assert payload["ok"] is False
    assert payload["failed_inputs"][0]["message"] == "bad input"
    assert agent_api.exit_code_for_result(result) == agent_api.EXIT_PARTIAL == 5


def test_skill_installer_writes_cli_config_and_backs_up_existing_copy(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "source-skill"
    source.mkdir()
    (source / "SKILL.md").write_text("skill", encoding="utf-8")
    (source / "scripts").mkdir()
    (source / "scripts" / "wrapper.py").write_text("pass", encoding="utf-8")
    skills = tmp_path / "skills"
    monkeypatch.setattr(agent_api, "_bundled_skill_directory", lambda: source)
    monkeypatch.setattr(
        agent_api, "_current_cli_command", lambda: ["C:\\LayoutLoom-CLI.exe"]
    )

    installed = agent_api.install_codex_skill(skills_directory=skills)
    target = Path(installed["installed_to"])
    config = json.loads((target / "layoutloom-cli.json").read_text(encoding="utf-8"))
    assert config["command"] == ["C:\\LayoutLoom-CLI.exe"]

    (target / "user-file.txt").write_text("old", encoding="utf-8")
    upgraded = agent_api.install_codex_skill(
        skills_directory=skills, force=True
    )
    backup = Path(upgraded["backup"])
    assert backup.is_dir()
    assert (backup / "user-file.txt").read_text(encoding="utf-8") == "old"
