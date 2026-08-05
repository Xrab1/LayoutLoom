from __future__ import annotations

import subprocess
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest import mock

import docuforge.runner as runner_module
from docuforge.models import (
    CancelledError,
    Capability,
    DocuForgeError,
    Operation,
    ParameterSpec,
    TaskResult,
    ValidationError,
)
from docuforge.runner import (
    TaskRunner,
    cancellation_callback,
    check_cancelled,
    progress_message,
    progress_scope,
    report_progress,
)
from docuforge.utils import (
    optimal_worker_count,
    parse_page_spec,
    safe_filename,
    unique_path,
)


class CoreTests(unittest.TestCase):
    def test_numeric_parameters_reject_booleans_and_non_finite_values(self) -> None:
        integer = ParameterSpec("count", "Count", "integer", 0)
        number = ParameterSpec("ratio", "Ratio", "number", 0.0)

        for spec in (integer, number):
            for value in (True, False):
                with self.subTest(kind=spec.kind, value=value):
                    with self.assertRaises(ValidationError):
                        spec.normalize(value)

        for spec in (integer, number):
            for value in (
                float("nan"),
                float("inf"),
                float("-inf"),
                "nan",
                "inf",
                "-inf",
            ):
                with self.subTest(kind=spec.kind, value=repr(value)):
                    with self.assertRaises(ValidationError):
                        spec.normalize(value)

    def test_parse_page_spec(self) -> None:
        self.assertEqual(parse_page_spec("1-3, 5, 8-", 10), [0, 1, 2, 4, 7, 8, 9])
        self.assertEqual(parse_page_spec("全部", 3), [0, 1, 2])
        with self.assertRaises(ValidationError):
            parse_page_spec("9", 3)

    def test_safe_filename_and_unique_path(self) -> None:
        self.assertEqual(safe_filename("报表:/2026*?"), "报表__2026__")
        with tempfile.TemporaryDirectory() as folder:
            first = Path(folder) / "结果.txt"
            first.write_text("x", encoding="utf-8")
            self.assertEqual(unique_path(first).name, "结果_1.txt")

    def test_optimal_worker_count_is_bounded_and_rejects_non_integers(self) -> None:
        with mock.patch("docuforge.utils.os.cpu_count", return_value=8):
            self.assertEqual(optimal_worker_count(0), 1)
            self.assertEqual(optimal_worker_count(2), 2)
            self.assertEqual(optimal_worker_count(20), 4)
            self.assertEqual(optimal_worker_count(20, cap=3), 3)
        for value in (-1, 1.5, "2", True):
            with self.subTest(value=value), self.assertRaises(ValidationError):
                optimal_worker_count(value)  # type: ignore[arg-type]

    def test_runner_normalizes_and_reports(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "输入.txt"
            source.write_text("hello", encoding="utf-8")

            def handler(
                inputs: list[Path], output: Path, params: dict[str, object]
            ) -> TaskResult:
                target = output / "结果.txt"
                target.write_text(str(params["count"]), encoding="utf-8")
                return TaskResult([target])

            operation = Operation(
                id="test.copy",
                group="测试",
                name="测试任务",
                description="",
                handler=handler,
                extensions=(".txt",),
                parameters=(ParameterSpec("count", "数量", "integer", 2, minimum=1),),
                capability_probe=lambda: Capability("ready", "ok", "test"),
            )
            result = TaskRunner().run(operation, [source], root / "out", {"count": "3"})
            self.assertEqual(
                (root / "out" / "结果.txt").read_text(encoding="utf-8"), "3"
            )
            self.assertEqual(len(result.outputs), 1)

    def test_runner_reports_monotonic_stage_and_file_progress(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            sources = [root / "first.txt", root / "second.txt"]
            for source in sources:
                source.write_text(source.stem, encoding="utf-8")

            def handler(
                inputs: list[Path], output: Path, params: dict[str, object]
            ) -> TaskResult:
                for index, source in enumerate(inputs, start=1):
                    report_progress(
                        (index - 1) / len(inputs),
                        f"复制 {source.name}",
                        current_file=index,
                        total_files=len(inputs),
                    )
                target = output / "result.txt"
                target.write_text("done", encoding="utf-8")
                report_progress(
                    1.0,
                    "核对批量输出",
                    current_file=len(inputs),
                    total_files=len(inputs),
                )
                return TaskResult([target])

            operation = Operation(
                id="test.progress",
                group="test",
                name="progress",
                description="",
                handler=handler,
                extensions=(".txt",),
                capability_probe=lambda: Capability("ready", "ok", "test"),
            )
            events: list[tuple[float, str]] = []

            TaskRunner().run(
                operation,
                sources,
                root / "out",
                progress=lambda value, message: events.append((value, message)),
            )

            values = [value for value, _message in events]
            messages = [message for _value, message in events]
            self.assertEqual(values, sorted(values))
            self.assertEqual(values[-1], 1.0)
            self.assertTrue(any("文件 1/2" in message for message in messages))
            self.assertTrue(any("文件 2/2" in message for message in messages))
            self.assertTrue(
                any("阶段：核对处理结果" in message for message in messages)
            )

    def test_progress_message_clamps_file_index_and_supports_aggregate_tasks(
        self,
    ) -> None:
        self.assertEqual(
            progress_message("转换", current_file=9, total_files=3),
            "阶段：转换 · 文件 3/3",
        )
        self.assertEqual(
            progress_message("合并", total_files=4),
            "阶段：合并 · 共 4 个文件",
        )

    def test_nested_progress_scopes_map_each_file_without_pinning_later_files(
        self,
    ) -> None:
        captured: list[tuple[float, str, int | None, int | None]] = []
        token = runner_module._CURRENT_PROGRESS_REPORTER.set(
            lambda value, stage, current, total: captured.append(
                (value, stage, current, total)
            )
        )
        try:
            with progress_scope(0.0, 0.5, current_file=1, total_files=2):
                report_progress(0.5, "分析第一个文件")
                report_progress(1.0, "完成第一个文件")
            with progress_scope(0.5, 0.5, current_file=2, total_files=2):
                report_progress(0.5, "分析第二个文件")
        finally:
            runner_module._CURRENT_PROGRESS_REPORTER.reset(token)

        self.assertEqual(
            captured,
            [
                (0.25, "分析第一个文件", 1, 2),
                (0.5, "完成第一个文件", 1, 2),
                (0.75, "分析第二个文件", 2, 2),
            ],
        )

    def test_runner_rejects_parallel_threads_and_releases_output_lock(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "input.txt"
            source.write_text("hello", encoding="utf-8")
            output = root / "out"
            entered = threading.Event()
            release = threading.Event()
            failures: list[BaseException] = []

            def handler(
                inputs: list[Path], target_dir: Path, params: dict[str, object]
            ) -> TaskResult:
                entered.set()
                if not release.wait(5):
                    raise RuntimeError("test handler timed out")
                target = target_dir / "result.txt"
                target.write_text(
                    inputs[0].read_text(encoding="utf-8"), encoding="utf-8"
                )
                return TaskResult([target])

            operation = Operation(
                id="test.lock.threads",
                group="test",
                name="thread lock",
                description="",
                handler=handler,
                extensions=(".txt",),
                capability_probe=lambda: Capability("ready", "ok", "test"),
            )

            def run_first() -> None:
                try:
                    TaskRunner().run(operation, [source], output)
                except BaseException as exc:
                    failures.append(exc)

            worker = threading.Thread(target=run_first)
            worker.start()
            self.assertTrue(entered.wait(5), "first task did not enter its handler")
            try:
                with self.assertRaisesRegex(DocuForgeError, "正被其他任务使用"):
                    TaskRunner().run(operation, [source], output)
            finally:
                release.set()
                worker.join(5)

            self.assertFalse(worker.is_alive(), "first task did not finish")
            self.assertEqual(failures, [])
            retry = TaskRunner().run(operation, [source], output)
            self.assertEqual(retry.outputs, [output.resolve() / "result.txt"])

    def test_runner_output_lock_conflicts_across_processes_then_retries(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "input.txt"
            source.write_text("hello", encoding="utf-8")
            output = root / "out"
            output.mkdir()
            ready = root / "ready"
            release = root / "release"
            project_root = Path(__file__).resolve().parents[1]
            script = """
import sys
import time
from pathlib import Path

sys.path.insert(0, sys.argv[1])
from docuforge.utils import output_directory_lock

output = Path(sys.argv[2])
ready = Path(sys.argv[3])
release = Path(sys.argv[4])
with output_directory_lock(output):
    ready.write_text("ready", encoding="utf-8")
    deadline = time.monotonic() + 10
    while not release.exists():
        if time.monotonic() >= deadline:
            raise TimeoutError("parent did not release child lock")
        time.sleep(0.02)
"""
            process = subprocess.Popen(
                [
                    sys.executable,
                    "-c",
                    script,
                    str(project_root),
                    str(output),
                    str(ready),
                    str(release),
                ],
                cwd=project_root,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )

            def handler(
                inputs: list[Path], target_dir: Path, params: dict[str, object]
            ) -> TaskResult:
                target = target_dir / "result.txt"
                target.write_text(
                    inputs[0].read_text(encoding="utf-8"), encoding="utf-8"
                )
                return TaskResult([target])

            operation = Operation(
                id="test.lock.processes",
                group="test",
                name="process lock",
                description="",
                handler=handler,
                extensions=(".txt",),
                capability_probe=lambda: Capability("ready", "ok", "test"),
            )

            try:
                deadline = time.monotonic() + 5
                while not ready.exists():
                    if process.poll() is not None:
                        stdout, stderr = process.communicate()
                        self.fail(
                            f"lock holder exited early ({process.returncode}): "
                            f"{stdout}\n{stderr}"
                        )
                    if time.monotonic() >= deadline:
                        self.fail("lock holder did not become ready")
                    time.sleep(0.02)

                with self.assertRaisesRegex(DocuForgeError, "正被其他任务使用"):
                    TaskRunner().run(operation, [source], output)
            finally:
                release.write_text("release", encoding="utf-8")
                try:
                    stdout, stderr = process.communicate(timeout=5)
                except subprocess.TimeoutExpired:
                    process.terminate()
                    stdout, stderr = process.communicate(timeout=5)
                    self.fail(f"lock holder did not exit: {stdout}\n{stderr}")

            self.assertEqual(process.returncode, 0, f"{stdout}\n{stderr}")
            retry = TaskRunner().run(operation, [source], output)
            self.assertEqual(retry.outputs, [output.resolve() / "result.txt"])

    def test_runner_releases_output_lock_after_handler_error(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "input.txt"
            source.write_text("hello", encoding="utf-8")
            output = root / "out"

            def failing_handler(
                inputs: list[Path], target_dir: Path, params: dict[str, object]
            ) -> TaskResult:
                raise RuntimeError("expected failure")

            failing_operation = Operation(
                id="test.lock.failure",
                group="test",
                name="failure lock",
                description="",
                handler=failing_handler,
                extensions=(".txt",),
                capability_probe=lambda: Capability("ready", "ok", "test"),
            )
            with self.assertRaisesRegex(RuntimeError, "expected failure"):
                TaskRunner().run(failing_operation, [source], output)

            def successful_handler(
                inputs: list[Path], target_dir: Path, params: dict[str, object]
            ) -> TaskResult:
                target = target_dir / "result.txt"
                target.write_text("done", encoding="utf-8")
                return TaskResult([target])

            successful_operation = Operation(
                id="test.lock.retry",
                group="test",
                name="retry lock",
                description="",
                handler=successful_handler,
                extensions=(".txt",),
                capability_probe=lambda: Capability("ready", "ok", "test"),
            )
            result = TaskRunner().run(successful_operation, [source], output)
            self.assertEqual(result.outputs, [output.resolve() / "result.txt"])

    def test_independent_batch_continues_after_one_file_fails(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            sources = [root / name for name in ("first.txt", "broken.txt", "last.txt")]
            for source in sources:
                source.write_text(source.stem, encoding="utf-8")
            calls: list[str] = []

            def handler(
                inputs: list[Path], target_dir: Path, params: dict[str, object]
            ) -> TaskResult:
                source = inputs[0]
                calls.append(source.name)
                if source.name == "broken.txt":
                    raise RuntimeError("模拟单文件损坏")
                target = target_dir / f"{source.stem}.out.txt"
                target.write_text(source.read_text(encoding="utf-8"), encoding="utf-8")
                return TaskResult([target])

            operation = Operation(
                id="test.batch.continue",
                group="test",
                name="batch continue",
                description="",
                handler=handler,
                extensions=(".txt",),
                capability_probe=lambda: Capability("ready", "ok", "test"),
                independent_inputs=True,
            )
            result = TaskRunner().run(operation, sources, root / "out")

            self.assertEqual(calls, [source.name for source in sources])
            self.assertEqual([path.name for path in result.outputs], ["first.out.txt", "last.out.txt"])
            self.assertEqual(result.completed_inputs, [sources[0].resolve(), sources[2].resolve()])
            self.assertEqual(len(result.failed_inputs), 1)
            self.assertEqual(result.failed_inputs[0].input_path, sources[1].resolve())
            self.assertIn("模拟单文件损坏", result.failed_inputs[0].message)
            self.assertEqual(result.outcome, "partial")

    def test_combined_operation_is_never_split_into_single_inputs(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            sources = [root / f"{index}.txt" for index in range(3)]
            for source in sources:
                source.write_text(source.stem, encoding="utf-8")
            received: list[list[Path]] = []

            def handler(
                inputs: list[Path], target_dir: Path, params: dict[str, object]
            ) -> TaskResult:
                received.append(list(inputs))
                target = target_dir / "combined.txt"
                target.write_text("|".join(path.stem for path in inputs), encoding="utf-8")
                return TaskResult([target])

            operation = Operation(
                id="test.batch.combined",
                group="test",
                name="combined",
                description="",
                handler=handler,
                extensions=(".txt",),
                capability_probe=lambda: Capability("ready", "ok", "test"),
            )
            result = TaskRunner().run(operation, sources, root / "out")

            self.assertEqual(received, [[source.resolve() for source in sources]])
            self.assertEqual(result.completed_inputs, [source.resolve() for source in sources])

    def test_cancel_cleans_current_partial_file_and_preserves_completed_output(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            sources = [root / "first.txt", root / "second.txt", root / "third.txt"]
            for source in sources:
                source.write_text(source.stem, encoding="utf-8")
            entered_second = threading.Event()
            failures: list[BaseException] = []

            def handler(
                inputs: list[Path], target_dir: Path, params: dict[str, object]
            ) -> TaskResult:
                source = inputs[0]
                if source.name == "second.txt":
                    (target_dir / "second.partial").write_bytes(b"unfinished")
                    entered_second.set()
                    while True:
                        check_cancelled("cancelled by test")
                        time.sleep(0.01)
                target = target_dir / f"{source.stem}.done.txt"
                target.write_text("done", encoding="utf-8")
                return TaskResult([target])

            operation = Operation(
                id="test.batch.cancel",
                group="test",
                name="cancel batch",
                description="",
                handler=handler,
                extensions=(".txt",),
                capability_probe=lambda: Capability("ready", "ok", "test"),
                independent_inputs=True,
            )
            runner = TaskRunner()

            def work() -> None:
                try:
                    runner.run(operation, sources, root / "out")
                except BaseException as exc:
                    failures.append(exc)

            worker = threading.Thread(target=work)
            worker.start()
            self.assertTrue(entered_second.wait(5), "second file did not start")
            runner.cancel()
            worker.join(5)

            self.assertFalse(worker.is_alive())
            self.assertEqual(len(failures), 1)
            self.assertIsInstance(failures[0], CancelledError)
            cancelled = failures[0]
            assert isinstance(cancelled, CancelledError)
            self.assertIsNotNone(cancelled.result)
            self.assertTrue((root / "out" / "first.done.txt").is_file())
            self.assertFalse((root / "out" / "second.partial").exists())
            self.assertFalse((root / "out" / "third.done.txt").exists())
            self.assertEqual(
                [path.name for path in (cancelled.result or TaskResult()).outputs],
                ["first.done.txt"],
            )
            self.assertEqual(
                [path.name for path in (cancelled.result or TaskResult()).cancelled_inputs],
                ["second.txt", "third.txt"],
            )

    def test_cancel_invokes_registered_external_engine_stop_callback(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "input.txt"
            source.write_text("input", encoding="utf-8")
            entered = threading.Event()
            stopped = threading.Event()
            failures: list[BaseException] = []

            def handler(
                inputs: list[Path], target_dir: Path, params: dict[str, object]
            ) -> TaskResult:
                with cancellation_callback(stopped.set):
                    entered.set()
                    while True:
                        check_cancelled("cancelled")
                        time.sleep(0.01)

            operation = Operation(
                id="test.cancel.callback",
                group="test",
                name="external callback",
                description="",
                handler=handler,
                extensions=(".txt",),
                capability_probe=lambda: Capability("ready", "ok", "test"),
            )
            runner = TaskRunner()

            def work() -> None:
                try:
                    runner.run(operation, [source], root / "out")
                except BaseException as exc:
                    failures.append(exc)

            worker = threading.Thread(target=work)
            worker.start()
            self.assertTrue(entered.wait(5))
            runner.cancel()
            self.assertTrue(stopped.wait(2), "external stop callback was not invoked")
            worker.join(5)

            self.assertFalse(worker.is_alive())
            self.assertEqual(len(failures), 1)
            self.assertIsInstance(failures[0], CancelledError)


if __name__ == "__main__":
    unittest.main()
