from __future__ import annotations

import json
from pathlib import Path

import pytest

from docuforge import cli
from docuforge.models import Capability, Operation, ParameterSpec, ValidationError
from docuforge.runner import report_progress


def _write_request(path: Path, payload: dict[str, object]) -> Path:
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return path


def _operation(handler, *, independent: bool = False) -> Operation:
    return Operation(
        "test.copy",
        "测试",
        "测试复制",
        "Agent CLI 测试任务",
        handler,
        (".txt",),
        (ParameterSpec("suffix", "后缀", "text", "done"),),
        independent_inputs=independent,
    )


def test_protocol_command_returns_machine_readable_json(capsys) -> None:
    code = cli.main(["agent", "protocol"])
    payload = json.loads(capsys.readouterr().out)
    assert code == 0
    assert payload["protocol"]["name"] == "layoutloom-agent"
    assert "quick-run" in payload["commands"]
    assert payload["exit_codes"]["70"] == "internal_error"


def test_duplicate_request_keys_are_rejected(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    request = tmp_path / "request.json"
    request.write_text(
        '{"operation":"test.copy","operation":"other"}', encoding="utf-8"
    )
    monkeypatch.setattr(cli, "get_operations", lambda: [])
    code = cli.main(["agent", "validate", "--request", str(request)])
    payload = json.loads(capsys.readouterr().out)
    assert code == 2
    assert payload["error"]["code"] == "validation_error"
    assert "重复字段" in payload["error"]["message"]


def test_validate_does_not_create_output_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    source = tmp_path / "source.txt"
    source.write_text("hello", encoding="utf-8")
    output = tmp_path / "output"
    operation = _operation(lambda paths, target, params: [])
    request = _write_request(
        tmp_path / "request.json",
        {
            "operation": operation.id,
            "inputs": [str(source)],
            "output_dir": str(output),
            "parameters": {"suffix": "ok"},
        },
    )
    monkeypatch.setattr(cli, "get_operations", lambda: [operation])
    code = cli.main(
        [
            "agent",
            "validate",
            "--request",
            str(request),
            "--allow-root",
            str(tmp_path),
        ]
    )
    payload = json.loads(capsys.readouterr().out)
    assert code == 0
    assert payload["ok"] is True
    assert payload["output_dir"] == str(output.resolve())
    assert not output.exists()


def test_jsonl_run_is_not_polluted_by_handler_stdout(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    source = tmp_path / "source.txt"
    source.write_text("hello", encoding="utf-8")

    def handler(paths, output, parameters):
        print("third-party diagnostic")
        report_progress(0.5, "正在复制")
        target = output / f"result-{parameters['suffix']}.txt"
        target.write_text(paths[0].read_text(encoding="utf-8"), encoding="utf-8")
        return target

    operation = _operation(handler)
    request = _write_request(
        tmp_path / "request.json",
        {
            "request_id": "jsonl-test",
            "operation": operation.id,
            "inputs": [str(source)],
            "output_dir": str(tmp_path / "output"),
            "parameters": {"suffix": "ok"},
        },
    )
    monkeypatch.setattr(cli, "get_operations", lambda: [operation])
    code = cli.main(
        ["agent", "run", "--request", str(request), "--format", "jsonl"]
    )
    captured = capsys.readouterr()
    events = [json.loads(line) for line in captured.out.splitlines()]
    assert code == 0
    assert "third-party diagnostic" not in captured.out
    assert "third-party diagnostic" in captured.err
    assert [event["seq"] for event in events] == list(range(1, len(events) + 1))
    assert events[0]["event"] == "accepted"
    assert events[-1]["event"] == "result"
    assert events[-1]["outcome"] == "success"
    assert Path(events[-1]["outputs"][0]).is_file()


def test_independent_batch_returns_partial_exit_and_keeps_successful_output(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    good = tmp_path / "good.txt"
    bad = tmp_path / "bad.txt"
    good.write_text("good", encoding="utf-8")
    bad.write_text("bad", encoding="utf-8")

    def handler(paths, output, parameters):
        source = paths[0]
        if source.name == "bad.txt":
            raise ValidationError("deliberate failure")
        target = output / f"{source.stem}-{parameters['suffix']}.txt"
        target.write_text("done", encoding="utf-8")
        return target

    operation = _operation(handler, independent=True)
    request = _write_request(
        tmp_path / "request.json",
        {
            "operation": operation.id,
            "inputs": [str(good), str(bad)],
            "output_dir": str(tmp_path / "output"),
            "parameters": {},
        },
    )
    monkeypatch.setattr(cli, "get_operations", lambda: [operation])
    code = cli.main(
        ["agent", "run", "--request", str(request), "--format", "json"]
    )
    payload = json.loads(capsys.readouterr().out)
    assert code == 5
    assert payload["outcome"] == "partial"
    assert len(payload["outputs"]) == 1
    assert len(payload["failed_inputs"]) == 1
    assert Path(payload["outputs"][0]).is_file()


def test_unknown_operation_uses_stable_error_code(
    monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    monkeypatch.setattr(cli, "get_operations", lambda: [])
    code = cli.main(["agent", "describe", "missing.operation"])
    payload = json.loads(capsys.readouterr().out)
    assert code == 2
    assert payload["error"]["code"] == "unknown_operation"


def test_agent_run_uses_parameter_aware_capability_override(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    source = tmp_path / "source.docx"
    source.write_text("placeholder", encoding="utf-8")

    def handler(paths, output, parameters):
        target = output / "converted.pdf"
        target.write_bytes(b"%PDF-test")
        return target

    operation = Operation(
        "word.to_pdf",
        "测试",
        "Word 转 PDF",
        "参数感知能力测试",
        handler,
        (".docx",),
        (
            ParameterSpec(
                "engine",
                "引擎",
                "choice",
                "auto",
                choices=(("auto", "自动"), ("microsoft_office", "Office")),
            ),
        ),
        capability_probe=lambda: Capability("unavailable", "默认探测不可用", "默认"),
    )
    from docuforge.processors import office

    monkeypatch.setattr(
        office,
        "_select_conversion_engine",
        lambda _source, _requested: "microsoft_office",
    )
    request = _write_request(
        tmp_path / "request.json",
        {
            "operation": operation.id,
            "inputs": [str(source)],
            "output_dir": str(tmp_path / "output"),
            "parameters": {"engine": "microsoft_office"},
        },
    )
    monkeypatch.setattr(cli, "get_operations", lambda: [operation])
    code = cli.main(
        ["agent", "run", "--request", str(request), "--format", "json"]
    )
    payload = json.loads(capsys.readouterr().out)
    assert code == 0
    assert payload["outcome"] == "success"
    assert payload["details"]["engine"] == "Microsoft Office COM"


def test_quick_run_executes_known_operation_without_request_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    source = tmp_path / "source.txt"
    source.write_text("hello", encoding="utf-8")

    def handler(paths, output, parameters):
        target = output / f"result-{parameters['suffix']}.txt"
        target.write_text(paths[0].read_text(encoding="utf-8"), encoding="utf-8")
        return target

    operation = _operation(handler)
    monkeypatch.setattr(cli, "get_operations", lambda: [operation])
    code = cli.main(
        [
            "agent",
            "quick-run",
            operation.id,
            str(source),
            "--output-dir",
            str(tmp_path / "output"),
            "--param",
            "suffix=fast",
            "--format",
            "json",
        ]
    )
    payload = json.loads(capsys.readouterr().out)
    assert code == 0
    assert payload["outcome"] == "success"
    assert Path(payload["outputs"][0]).read_text(encoding="utf-8") == "hello"
    assert payload["details"]["agent_preflight_seconds"] >= 0
    assert payload["details"]["agent_total_seconds"] >= 0


def test_quick_run_rejects_sensitive_command_line_parameter(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    source = tmp_path / "source.txt"
    source.write_text("hello", encoding="utf-8")
    operation = Operation(
        "test.secret",
        "测试",
        "敏感参数测试",
        "敏感参数测试",
        lambda paths, output, parameters: [],
        (".txt",),
        (ParameterSpec("password", "密码", "password", ""),),
        allow_empty_outputs=True,
    )
    monkeypatch.setattr(cli, "get_operations", lambda: [operation])
    code = cli.main(
        [
            "agent",
            "quick-run",
            operation.id,
            str(source),
            "--output-dir",
            str(tmp_path / "output"),
            "--param",
            "password=do-not-print-this",
            "--format",
            "json",
        ]
    )
    captured = capsys.readouterr().out
    payload = json.loads(captured)
    assert code == 2
    assert payload["error"]["code"] == "validation_error"
    assert "--params-file" in payload["error"]["message"]
    assert "do-not-print-this" not in captured


def test_quick_run_accepts_sensitive_parameter_from_json_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    source = tmp_path / "source.txt"
    source.write_text("hello", encoding="utf-8")

    def handler(paths, output, parameters):
        assert parameters["password"] == "s3ns1t1ve-value"
        target = output / "done.txt"
        target.write_text("done", encoding="utf-8")
        return target

    operation = Operation(
        "test.secret",
        "测试",
        "敏感参数测试",
        "敏感参数测试",
        handler,
        (".txt",),
        (ParameterSpec("password", "密码", "password", ""),),
    )
    parameters = tmp_path / "parameters.json"
    parameters.write_text('{"password":"s3ns1t1ve-value"}', encoding="utf-8")
    monkeypatch.setattr(cli, "get_operations", lambda: [operation])
    code = cli.main(
        [
            "agent",
            "quick-run",
            operation.id,
            str(source),
            "--output-dir",
            str(tmp_path / "output"),
            "--params-file",
            str(parameters),
            "--format",
            "json",
        ]
    )
    captured = capsys.readouterr().out
    payload = json.loads(captured)
    assert code == 0
    assert payload["outcome"] == "success"
    assert "s3ns1t1ve-value" not in captured
