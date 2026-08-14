from __future__ import annotations

import os
import runpy
import sys
import tempfile
import threading
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch


HOOK = (
    Path(__file__).resolve().parents[1]
    / "packaging_hooks"
    / "rthooks"
    / "layoutloom_tk_runtime.py"
)


class TkRuntimeHookTests(unittest.TestCase):
    @staticmethod
    def _write_primary_runtime(root: Path) -> None:
        (root / "_tcl_data").mkdir(parents=True)
        (root / "_tk_data").mkdir(parents=True)
        (root / "_tcl_data" / "init.tcl").write_text("# tcl", encoding="utf-8")
        (root / "_tk_data" / "tk.tcl").write_text("# tk", encoding="utf-8")

    @staticmethod
    def _write_backup(root: Path) -> None:
        with zipfile.ZipFile(root / "tk_runtime_backup.zip", "w") as package:
            package.writestr("_tcl_data/init.tcl", "# tcl")
            package.writestr("_tk_data/tk.tcl", "# tk")

    def test_hook_uses_primary_frozen_runtime(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            self._write_primary_runtime(root)
            with patch.object(sys, "frozen", True, create=True), patch.object(
                sys, "_MEIPASS", str(root), create=True
            ), patch.dict(os.environ, {}, clear=True):
                runpy.run_path(str(HOOK))
                self.assertEqual(
                    Path(os.environ["TCL_LIBRARY"]), root / "_tcl_data"
                )
                self.assertEqual(Path(os.environ["TK_LIBRARY"]), root / "_tk_data")

    def test_hook_recovers_missing_runtime_from_single_file_backup(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder) / "bundle"
            cache = Path(folder) / "cache"
            root.mkdir()
            self._write_backup(root)
            with patch.object(sys, "frozen", True, create=True), patch.object(
                sys, "_MEIPASS", str(root), create=True
            ), patch.dict(
                os.environ, {"LOCALAPPDATA": str(cache)}, clear=True
            ):
                runpy.run_path(str(HOOK))
                tcl_library = Path(os.environ["TCL_LIBRARY"])
                tk_library = Path(os.environ["TK_LIBRARY"])
                self.assertTrue((tcl_library / "init.tcl").is_file())
                self.assertTrue((tk_library / "tk.tcl").is_file())
                self.assertTrue(tcl_library.is_relative_to(cache))
                self.assertTrue(tk_library.is_relative_to(cache))

    def test_recovery_cache_is_safe_when_two_frozen_processes_start_together(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder) / "bundle"
            cache = Path(folder) / "cache"
            root.mkdir()
            self._write_backup(root)
            failures: list[BaseException] = []

            def recover() -> None:
                try:
                    namespace = runpy.run_path(str(HOOK), run_name="layoutloom_hook_test")
                    namespace["_cached_runtime"](root)
                except BaseException as exc:  # pragma: no cover - only records failure
                    failures.append(exc)

            with patch.dict(os.environ, {"LOCALAPPDATA": str(cache)}, clear=True):
                workers = [threading.Thread(target=recover) for _ in range(2)]
                for worker in workers:
                    worker.start()
                for worker in workers:
                    worker.join(10)

            self.assertTrue(all(not worker.is_alive() for worker in workers))
            self.assertEqual(failures, [])
            runtime_root = cache / "LayoutLoom" / "runtime"
            completed = [
                path
                for path in runtime_root.iterdir()
                if path.is_dir()
                and (path / "_tcl_data" / "init.tcl").is_file()
                and (path / "_tk_data" / "tk.tcl").is_file()
            ]
            self.assertEqual(len(completed), 1)
            self.assertEqual(list(runtime_root.glob("*.lock")), [])

    def test_recovery_publication_retries_transient_windows_directory_lock(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder) / "bundle"
            cache = Path(folder) / "cache"
            root.mkdir()
            self._write_backup(root)
            namespace = runpy.run_path(str(HOOK), run_name="layoutloom_hook_retry_test")
            real_replace = Path.replace
            attempts = 0

            def flaky_replace(path: Path, target: Path) -> Path:
                nonlocal attempts
                if path.name.startswith("tk-stage-"):
                    attempts += 1
                    if attempts < 3:
                        error = PermissionError(13, "temporarily locked")
                        error.winerror = 5
                        raise error
                return real_replace(path, target)

            with patch.dict(
                os.environ, {"LOCALAPPDATA": str(cache)}, clear=True
            ), patch.object(Path, "replace", new=flaky_replace):
                runtime = namespace["_cached_runtime"](root)

            self.assertEqual(attempts, 3)
            self.assertTrue((runtime / "_tcl_data" / "init.tcl").is_file())
            self.assertTrue((runtime / "_tk_data" / "tk.tcl").is_file())
            runtime_parent = cache / "LayoutLoom" / "runtime"
            self.assertEqual(list(runtime_parent.glob("*.lock")), [])
            self.assertEqual(list(runtime_parent.glob("tk-stage-*")), [])

    def test_hook_explains_an_incomplete_portable_package(self) -> None:
        with tempfile.TemporaryDirectory() as folder, patch.object(
            sys, "frozen", True, create=True
        ), patch.object(sys, "_MEIPASS", folder, create=True), patch.dict(
            os.environ, {}, clear=True
        ):
            with self.assertRaisesRegex(
                FileNotFoundError, "portable package is incomplete"
            ):
                runpy.run_path(str(HOOK))

    def test_hook_reports_bootstrap_failure_without_opening_a_modal_dialog(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            error_file = Path(folder) / "bootstrap-error.log"
            namespace = runpy.run_path(
                str(HOOK), run_name="layoutloom_hook_diagnostic_test"
            )

            def fail_configuration() -> None:
                raise PermissionError("runtime unavailable")

            diagnostic = namespace["_configure_with_self_test_diagnostics"]
            with patch.dict(
                os.environ,
                {namespace["_SELF_TEST_ERROR_FILE_ENV"]: str(error_file)},
                clear=True,
            ), patch.dict(
                diagnostic.__globals__,
                {"_configure_tk_runtime": fail_configuration},
            ), patch.object(os, "_exit", side_effect=SystemExit(86)):
                with self.assertRaisesRegex(SystemExit, "86"):
                    diagnostic()

            self.assertIn(
                "runtime unavailable", error_file.read_text(encoding="utf-8")
            )


if __name__ == "__main__":
    unittest.main()
