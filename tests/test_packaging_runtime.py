from __future__ import annotations

import os
import runpy
import shutil
import subprocess
import sys
import tempfile
import threading
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
HOOK = (
    PROJECT_ROOT
    / "packaging_hooks"
    / "rthooks"
    / "layoutloom_tk_runtime.py"
)
BUILD_SCRIPT = PROJECT_ROOT / "build.ps1"
RELEASE_SCRIPT = PROJECT_ROOT / "release.ps1"
AGENT_SKILL = PROJECT_ROOT / "integrations" / "codex" / "layoutloom-agent"


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


class AgentPackagingContractTests(unittest.TestCase):
    @staticmethod
    def _script(path: Path) -> str:
        return path.read_text(encoding="utf-8-sig")

    def test_powershell_packaging_scripts_have_valid_syntax(self) -> None:
        powershell = shutil.which("pwsh") or shutil.which("powershell")
        if powershell is None:
            self.skipTest("PowerShell is not available")
        parser = (
            "& { param([string]$Path) "
            "$tokens=$null; $errors=$null; "
            "[System.Management.Automation.Language.Parser]::ParseFile("
            "$Path,[ref]$tokens,[ref]$errors) | Out-Null; "
            "if($errors.Count){ $errors | ForEach-Object { "
            "[Console]::Error.WriteLine($_.ToString()) }; exit 1 } }"
        )
        for script in (BUILD_SCRIPT, RELEASE_SCRIPT):
            with self.subTest(script=script.name):
                completed = subprocess.run(
                    [
                        powershell,
                        "-NoProfile",
                        "-NonInteractive",
                        "-Command",
                        parser,
                        str(script),
                    ],
                    check=False,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                )
                self.assertEqual(
                    completed.returncode,
                    0,
                    completed.stderr or completed.stdout,
                )

    def test_build_contract_creates_console_cli_with_a_safe_shared_runtime(self) -> None:
        script = self._script(BUILD_SCRIPT)
        self.assertIn('"--name", "LayoutLoom-CLI"', script)
        self.assertIn('"--console"', script)
        self.assertIn('$protocol.commands -notcontains "quick-run"', script)
        self.assertIn(
            '$cliEntryPoint = Join-Path $projectRoot "agent_launcher.py"', script
        )
        self.assertIn("function Merge-CompatiblePyInstallerRuntime", script)
        self.assertIn(
            "Merge-CompatiblePyInstallerRuntime -SourceDirectory "
            "$cliSourceInternal -DestinationDirectory $bundleInternal",
            script,
        )
        self.assertIn(
            "The GUI and Agent CLI PyInstaller runtimes conflict", script
        )
        self.assertNotIn(
            "Copy-Item -LiteralPath $item.FullName -Destination "
            "$bundleInternal -Recurse -Force",
            script,
        )

    def test_shared_runtime_merge_preserves_identical_files_and_rejects_conflicts(
        self,
    ) -> None:
        powershell = shutil.which("pwsh") or shutil.which("powershell")
        if powershell is None:
            self.skipTest("PowerShell is not available")
        command = (
            "& { param([string]$BuildScript,[string]$SourceDirectory,"
            "[string]$DestinationDirectory) "
            "$tokens=$null; $errors=$null; "
            "$ast=[System.Management.Automation.Language.Parser]::ParseFile("
            "$BuildScript,[ref]$tokens,[ref]$errors); "
            "if($errors.Count){ throw ($errors -join [Environment]::NewLine) }; "
            "$function=$ast.Find({ param($node) "
            "$node -is [System.Management.Automation.Language.FunctionDefinitionAst] "
            "-and $node.Name -eq 'Merge-CompatiblePyInstallerRuntime' },$true); "
            "if($null -eq $function){ throw 'Runtime merge function is missing.' }; "
            "Invoke-Expression $function.Extent.Text; "
            "Merge-CompatiblePyInstallerRuntime -SourceDirectory $SourceDirectory "
            "-DestinationDirectory $DestinationDirectory }"
        )
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "cli-internal"
            destination = root / "gui-internal"
            (source / "package").mkdir(parents=True)
            destination.mkdir()
            (source / "shared.bin").write_bytes(b"same-runtime")
            (destination / "shared.bin").write_bytes(b"same-runtime")
            (source / "package" / "cli-only.dat").write_bytes(b"agent")
            source_zip_info = zipfile.ZipInfo(
                "encodings/__init__.pyc", date_time=(2020, 1, 1, 0, 0, 0)
            )
            source_zip_info.compress_type = zipfile.ZIP_STORED
            destination_zip_info = zipfile.ZipInfo(
                "encodings/__init__.pyc", date_time=(2024, 1, 1, 0, 0, 0)
            )
            destination_zip_info.compress_type = zipfile.ZIP_DEFLATED
            with zipfile.ZipFile(source / "base_library.zip", "w") as package:
                package.writestr(source_zip_info, b"same-stdlib-module")
            with zipfile.ZipFile(destination / "base_library.zip", "w") as package:
                package.writestr(destination_zip_info, b"same-stdlib-module")
            invocation = [
                powershell,
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                command,
                str(BUILD_SCRIPT),
                str(source),
                str(destination),
            ]
            compatible = subprocess.run(
                invocation,
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            self.assertEqual(
                compatible.returncode,
                0,
                compatible.stderr or compatible.stdout,
            )
            self.assertEqual(
                (destination / "package" / "cli-only.dat").read_bytes(), b"agent"
            )
            self.assertEqual(
                (destination / "shared.bin").read_bytes(), b"same-runtime"
            )
            with zipfile.ZipFile(destination / "base_library.zip") as package:
                self.assertEqual(
                    package.read("encodings/__init__.pyc"), b"same-stdlib-module"
                )

            (source / "shared.bin").write_bytes(b"DIFF-runtime")
            conflict = subprocess.run(
                invocation,
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            self.assertNotEqual(conflict.returncode, 0)
            self.assertIn(
                "PyInstaller runtimes conflict",
                conflict.stderr + conflict.stdout,
            )

            (source / "shared.bin").write_bytes(b"same-runtime")
            with zipfile.ZipFile(source / "base_library.zip", "w") as package:
                package.writestr(source_zip_info, b"DIFF-stdlib-module")
            zip_conflict = subprocess.run(
                invocation,
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            self.assertNotEqual(zip_conflict.returncode, 0)
            self.assertIn(
                "base_library.zip",
                zip_conflict.stderr + zip_conflict.stdout,
            )

    def test_portable_bundle_contract_includes_cli_and_complete_skill(self) -> None:
        build_script = self._script(BUILD_SCRIPT).replace("\\", "/")
        release_script = self._script(RELEASE_SCRIPT).replace("\\", "/")
        self.assertIn(
            '$agentSkillTarget = Join-Path $bundleDirectory '
            '"agent_skill/layoutloom-agent"',
            build_script,
        )
        for relative in (
            "SKILL.md",
            "agents/openai.yaml",
            "scripts/layoutloom_agent.py",
            "references/protocol.md",
        ):
            self.assertTrue((AGENT_SKILL / relative).is_file(), relative)
            self.assertIn(
                f'"LayoutLoom/agent_skill/layoutloom-agent/{relative}"',
                release_script,
            )
        self.assertIn('"LayoutLoom/LayoutLoom-CLI.exe"', release_script)
        self.assertIn('"LayoutLoom/AGENT_INTEGRATION.md"', release_script)
        self.assertIn('"AGENT_INTEGRATION.md"', build_script)

    def test_portable_skill_wrapper_finds_sibling_cli(self) -> None:
        namespace = runpy.run_path(
            str(AGENT_SKILL / "scripts" / "layoutloom_agent.py"),
            run_name="layoutloom_portable_skill_test",
        )
        with tempfile.TemporaryDirectory() as folder:
            bundle = Path(folder) / "LayoutLoom"
            skill_root = bundle / "agent_skill" / "layoutloom-agent"
            skill_root.mkdir(parents=True)
            cli = bundle / "LayoutLoom-CLI.exe"
            cli.touch()
            with patch.dict(os.environ, {}, clear=True):
                command = namespace["locate_command"](skill_root)
            self.assertEqual(command, [str(cli.resolve())])

    def test_source_archive_contract_includes_agent_bridge(self) -> None:
        script = self._script(RELEASE_SCRIPT).replace("\\", "/")
        self.assertIn('$protocol.commands -notcontains "quick-run"', script)
        self.assertIn(
            '$sourceDirectories = @("docuforge", "tests", '
            '"packaging_hooks", "integrations")',
            script,
        )
        self.assertIn('"agent_launcher.py"', script)
        self.assertIn("function Assert-SourceArchiveContents", script)
        self.assertIn(
            "Assert-SourceArchiveContents -ArchivePath $sourceZip "
            "-Version $version",
            script,
        )
        for relative in (
            "AGENT_INTEGRATION.md",
            "agent_launcher.py",
            "docuforge/agent_api.py",
            "docuforge/cli.py",
            "integrations/codex/layoutloom-agent/SKILL.md",
            "integrations/codex/layoutloom-agent/agents/openai.yaml",
            "integrations/codex/layoutloom-agent/scripts/layoutloom_agent.py",
            "integrations/codex/layoutloom-agent/references/protocol.md",
            "tests/test_agent_api.py",
            "tests/test_agent_cli.py",
        ):
            self.assertIn(f'"$root/{relative}"', script)


if __name__ == "__main__":
    unittest.main()
