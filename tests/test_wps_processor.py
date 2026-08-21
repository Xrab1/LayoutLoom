from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docuforge.models import CancelledError, MissingEngineError, ValidationError
from docuforge.processors import office, wps
from docuforge.processors.wps import convert_with_wps, detect_wps_engines


class WpsProcessorTests(unittest.TestCase):
    def test_probe_has_all_application_kinds(self) -> None:
        statuses = detect_wps_engines()
        self.assertEqual(set(statuses), {"writer", "spreadsheets", "presentation"})
        self.assertTrue(all(status.prog_id for status in statuses.values()))

    def test_validation_and_missing_engine_are_clear(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            with self.assertRaises(ValidationError):
                convert_with_wps(root / "missing.docx", root / "out")
            source = root / "sample.docx"
            source.write_bytes(b"placeholder")
            statuses = detect_wps_engines()
            if not statuses["writer"].available:
                with self.assertRaises(MissingEngineError):
                    convert_with_wps(source, root / "out")

    def test_probe_does_not_infer_missing_components_from_writer(self) -> None:
        def registered(kind: wps.WpsKind) -> str | None:
            return "KWPS.Application" if kind == "writer" else None

        with patch.object(wps.sys, "platform", "win32"), patch.object(
            wps, "_pywin32_available", return_value=True
        ), patch.object(
            wps, "_registered_prog_id", side_effect=registered
        ), patch.object(
            wps, "_find_wps_executable", return_value=None
        ):
            statuses = detect_wps_engines()
        self.assertTrue(statuses["writer"].available)
        self.assertFalse(statuses["spreadsheets"].available)
        self.assertFalse(statuses["presentation"].available)

    def test_finds_wps_installation_without_kingsoft_parent_folder(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            executable = root / "WPS Office" / "12.1.0" / "office6" / "wps.exe"
            executable.parent.mkdir(parents=True)
            executable.touch()
            with patch.dict(
                wps.os.environ,
                {
                    "ProgramFiles": str(root / "Program Files"),
                    "ProgramFiles(x86)": str(root),
                    "LOCALAPPDATA": str(root / "Local"),
                },
            ), patch.object(wps.shutil, "which", return_value=None):
                detected = wps._find_wps_executable("writer")
            self.assertEqual(detected, executable.resolve())

    def test_new_owned_wps_processes_accepts_broker_worker_group_only(self) -> None:
        executable = Path("C:/Program Files/Kingsoft/WPS Office/office6/wps.exe")
        existing = office._OfficeProcessIdentity(100, executable, "2026-08-16T12:00:00")
        broker = office._OfficeProcessIdentity(201, executable, "2026-08-16T12:00:10")
        worker = office._OfficeProcessIdentity(202, executable, "2026-08-16T12:00:11")

        owned = wps._new_owned_wps_processes(
            {existing.pid: existing},
            expected_executable=executable,
            snapshot=lambda _name: {
                existing.pid: existing,
                broker.pid: broker,
                worker.pid: worker,
            },
        )

        self.assertEqual(owned, (broker, worker))
        self.assertNotIn(existing, owned)

    def test_owned_wps_requires_automation_and_exact_reported_pid_set(self) -> None:
        executable = Path("C:/Program Files/Kingsoft/WPS Office/office6/wps.exe")
        existing = office._OfficeProcessIdentity(100, executable, "2026-08-16T12:00:00")
        broker = office._OfficeProcessIdentity(201, executable, "2026-08-16T12:00:10")
        worker = office._OfficeProcessIdentity(202, executable, "2026-08-16T12:00:11")
        manual = office._OfficeProcessIdentity(203, executable, "2026-08-16T12:00:11")
        after = {
            existing.pid: existing,
            broker.pid: broker,
            worker.pid: worker,
            manual.pid: manual,
        }
        command_lines = {
            broker.pid: '"wps.exe" /prometheus /Automation -Embedding',
            worker.pid: '"wps.exe" /from_prome /Automation -Embedding',
            manual.pid: '"wps.exe" /wps',
        }

        owned = wps._new_owned_wps_processes(
            {existing.pid: existing},
            expected_executable=executable,
            reported_pids=(broker.pid, worker.pid),
            reported_identities=(broker, worker),
            require_automation=True,
            snapshot=lambda _name: after,
            command_line=lambda process_id: command_lines.get(process_id),
        )
        rejected = wps._new_owned_wps_processes(
            {existing.pid: existing},
            expected_executable=executable,
            reported_pids=(broker.pid, worker.pid, manual.pid),
            require_automation=True,
            snapshot=lambda _name: after,
            command_line=lambda process_id: command_lines.get(process_id),
        )

        self.assertEqual(owned, (broker, worker))
        self.assertEqual(rejected, ())

    def test_owned_wps_rejects_reported_creation_time_mismatch(self) -> None:
        executable = Path("C:/Program Files/Kingsoft/WPS Office/office6/wps.exe")
        reported = office._OfficeProcessIdentity(201, executable, "2026-08-16T12:00:10")
        current = office._OfficeProcessIdentity(201, executable, "2026-08-16T12:00:12")

        owned = wps._new_owned_wps_processes(
            {},
            expected_executable=executable,
            reported_identities=(reported,),
            require_automation=True,
            snapshot=lambda _name: {current.pid: current},
            command_line=lambda _process_id: '"wps.exe" /Automation -Embedding',
        )

        self.assertEqual(owned, ())

    def test_owned_wps_rejects_pid_reuse_after_worker_report(self) -> None:
        executable = Path("C:/Program Files/Kingsoft/WPS Office/office6/wps.exe")
        before = office._OfficeProcessIdentity(201, executable, "2026-08-16T12:00:00")
        worker_report = office._OfficeProcessIdentity(
            201, executable, "2026-08-16T12:00:10"
        )
        reused = office._OfficeProcessIdentity(201, executable, "2026-08-16T12:00:20")

        owned = wps._new_owned_wps_processes(
            {before.pid: before},
            expected_executable=executable,
            reported_identities=(worker_report,),
            require_automation=True,
            snapshot=lambda _name: {reused.pid: reused},
            command_line=lambda _process_id: '"wps.exe" /Automation -Embedding',
        )

        self.assertEqual(owned, ())

    def test_wps_claims_broker_worker_and_quits_owned_application(self) -> None:
        events: list[str] = []

        class PythonCom:
            @staticmethod
            def CoInitialize() -> None:
                events.append("initialize")

            @staticmethod
            def CoUninitialize() -> None:
                events.append("uninitialize")

        class Document:
            def SaveAs(self, path: str, _format_code: int) -> None:
                Path(path).write_bytes(b"converted")

            def Close(self, _save: bool) -> None:
                events.append("close")

        class Documents:
            def Open(self, _path: str, **_kwargs: object) -> Document:
                return Document()

        class Application:
            def __init__(self) -> None:
                self.Documents = Documents()

            def Quit(self) -> None:
                events.append("quit")

        class Client:
            @staticmethod
            def DispatchEx(_prog_id: str) -> Application:
                return Application()

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            executable = (root / "wps.exe").resolve()
            executable.touch()
            broker = office._OfficeProcessIdentity(
                201, executable, "2026-08-16T12:00:10"
            )
            worker = office._OfficeProcessIdentity(
                202, executable, "2026-08-16T12:00:11"
            )
            statuses = {
                "writer": wps.WpsEngineStatus(
                    True, "writer", "KWPS.Application", executable=executable
                ),
                "spreadsheets": wps.WpsEngineStatus(
                    False, "spreadsheets", "KET.Application"
                ),
                "presentation": wps.WpsEngineStatus(
                    False, "presentation", "KWPP.Application"
                ),
            }
            source = root / "合同.docx"
            source.write_bytes(b"placeholder")
            with patch.object(
                wps, "detect_wps_engines", return_value=statuses
            ), patch.object(
                wps, "_load_pywin32", return_value=(PythonCom, Client)
            ), patch.object(
                office,
                "_windows_process_snapshot",
                side_effect=[{}, {broker.pid: broker, worker.pid: worker}],
            ), patch.object(
                office, "_office_application_pid", return_value=None
            ), patch.object(
                wps,
                "_windows_wps_process_command_line",
                return_value='"C:/fake/wps.exe" /Automation -Embedding',
            ), patch.object(
                office, "_wait_for_owned_office_exit"
            ) as wait_for_exit:
                output = convert_with_wps(source, root / "out", "docx")
                output_bytes = output[0].read_bytes()

        self.assertEqual(output_bytes, b"converted")
        self.assertIn("quit", events)
        self.assertEqual(
            [entry.args[0] for entry in wait_for_exit.call_args_list],
            [broker, worker],
        )

    def test_supervised_wps_timeout_stops_worker(self) -> None:
        class ParentConnection:
            def poll(self, _timeout: float) -> bool:
                return False

            def close(self) -> None:
                return None

        class ChildConnection:
            def close(self) -> None:
                return None

        class Process:
            def start(self) -> None:
                return None

            def is_alive(self) -> bool:
                return True

        class Context:
            def Pipe(self, *, duplex: bool) -> tuple[ParentConnection, ChildConnection]:
                self.duplex = duplex
                return ParentConnection(), ChildConnection()

            def Process(self, **_kwargs: object) -> Process:
                return Process()

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            executable = (root / "wps.exe").resolve()
            executable.touch()
            source = root / "合同.docx"
            source.write_bytes(b"placeholder")
            statuses = {
                "writer": wps.WpsEngineStatus(
                    True, "writer", "KWPS.Application", executable=executable
                ),
                "spreadsheets": wps.WpsEngineStatus(
                    False, "spreadsheets", "KET.Application"
                ),
                "presentation": wps.WpsEngineStatus(
                    False, "presentation", "KWPP.Application"
                ),
            }
            context = Context()
            with patch.object(wps, "detect_wps_engines", return_value=statuses), patch(
                "multiprocessing.get_context", return_value=context
            ), patch.object(
                office, "_windows_process_snapshot", return_value={}
            ), patch.object(
                wps, "_new_owned_wps_processes", return_value=()
            ), patch.object(
                wps.time, "monotonic", side_effect=[0.0, 1.0]
            ), patch.object(
                wps, "_stop_wps_worker"
            ) as stop_worker:
                with self.assertRaisesRegex(MissingEngineError, "转换超时"):
                    wps.convert_with_wps_supervised(
                        source, root / "out", "pdf", timeout=0.5
                    )

        self.assertTrue(context.duplex)
        stop_worker.assert_called_once()

    def test_supervised_wps_cancel_stops_worker_and_cleans_temps(self) -> None:
        class ParentConnection:
            def poll(self, _timeout: float) -> bool:
                return False

            def close(self) -> None:
                return None

        class ChildConnection:
            def close(self) -> None:
                return None

        class Process:
            def start(self) -> None:
                return None

            def is_alive(self) -> bool:
                return True

        class Context:
            def Pipe(self, *, duplex: bool) -> tuple[ParentConnection, ChildConnection]:
                return ParentConnection(), ChildConnection()

            def Process(self, **_kwargs: object) -> Process:
                return Process()

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            executable = (root / "wps.exe").resolve()
            executable.touch()
            source = root / "合同.docx"
            source.write_bytes(b"placeholder")
            statuses = {
                "writer": wps.WpsEngineStatus(
                    True, "writer", "KWPS.Application", executable=executable
                ),
                "spreadsheets": wps.WpsEngineStatus(
                    False, "spreadsheets", "KET.Application"
                ),
                "presentation": wps.WpsEngineStatus(
                    False, "presentation", "KWPP.Application"
                ),
            }
            with patch.object(wps, "detect_wps_engines", return_value=statuses), patch(
                "multiprocessing.get_context", return_value=Context()
            ), patch.object(
                office, "_windows_process_snapshot", return_value={}
            ), patch.object(
                wps, "_new_owned_wps_processes", return_value=()
            ), patch(
                "docuforge.runner.check_cancelled",
                side_effect=[None, CancelledError("stop")],
            ), patch.object(
                wps, "_stop_wps_worker"
            ) as stop_worker, patch.object(
                wps, "_cleanup_wps_temporary_outputs"
            ) as cleanup:
                with self.assertRaises(CancelledError):
                    wps.convert_with_wps_supervised(source, root / "out", "pdf")

        stop_worker.assert_called_once()
        cleanup.assert_called_once()

    def test_wps_macro_security_is_set_before_open(self) -> None:
        events: list[object] = []

        class PythonCom:
            @staticmethod
            def CoInitialize() -> None:
                events.append("initialize")

            @staticmethod
            def CoUninitialize() -> None:
                events.append("uninitialize")

        class Document:
            def SaveAs(self, path: str, _format_code: int) -> None:
                Path(path).write_bytes(b"converted")

            def Close(self, _save: bool) -> None:
                events.append("close")

        class Documents:
            def Open(self, _path: str, **_kwargs: object) -> Document:
                events.append("open")
                return Document()

        class Application:
            def __init__(self) -> None:
                object.__setattr__(self, "Documents", Documents())

            def __setattr__(self, name: str, value: object) -> None:
                if name == "AutomationSecurity":
                    events.append(("security", value))
                object.__setattr__(self, name, value)

            def Quit(self) -> None:
                events.append("quit")

        class Client:
            @staticmethod
            def DispatchEx(_prog_id: str) -> Application:
                return Application()

        executable = Path("C:/fake/wps.exe")
        identity = office._OfficeProcessIdentity(
            1001, executable, "2026-08-16T12:00:00"
        )
        status = wps.WpsEngineStatus(
            True, "writer", "KWPS.Application", executable=executable
        )
        statuses = {
            "writer": status,
            "spreadsheets": wps.WpsEngineStatus(
                False, "spreadsheets", "KET.Application"
            ),
            "presentation": wps.WpsEngineStatus(
                False, "presentation", "KWPP.Application"
            ),
        }
        with tempfile.TemporaryDirectory() as folder, patch.object(
            wps, "detect_wps_engines", return_value=statuses
        ), patch.object(
            wps, "_load_pywin32", return_value=(PythonCom, Client)
        ), patch.object(
            wps, "_new_owned_wps_processes", return_value=(identity,)
        ), patch.object(
            office, "_wait_for_owned_office_exit"
        ):
            root = Path(folder)
            source = root / "合同.docx"
            source.write_bytes(b"placeholder")
            output = convert_with_wps(source, root / "out", "docx")
            self.assertEqual(output[0].read_bytes(), b"converted")

        self.assertLess(events.index(("security", 3)), events.index("open"))
        self.assertEqual(events[0], "initialize")
        self.assertEqual(events[-1], "uninitialize")

    def test_wps_unknown_executable_never_quits_an_unproven_instance(self) -> None:
        events: list[str] = []

        class PythonCom:
            @staticmethod
            def CoInitialize() -> None:
                events.append("initialize")

            @staticmethod
            def CoUninitialize() -> None:
                events.append("uninitialize")

        class Documents:
            Count = 0

        class Application:
            def __init__(self) -> None:
                self.Documents = Documents()

            def Quit(self) -> None:
                events.append("quit")

        class Client:
            @staticmethod
            def DispatchEx(_prog_id: str) -> Application:
                return Application()

        statuses = {
            "writer": wps.WpsEngineStatus(True, "writer", "KWPS.Application"),
            "spreadsheets": wps.WpsEngineStatus(
                False, "spreadsheets", "KET.Application"
            ),
            "presentation": wps.WpsEngineStatus(
                False, "presentation", "KWPP.Application"
            ),
        }
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "合同.docx"
            source.write_bytes(b"placeholder")
            with patch.object(
                wps, "detect_wps_engines", return_value=statuses
            ), patch.object(
                wps, "_load_pywin32", return_value=(PythonCom, Client)
            ), patch.object(
                office, "_windows_process_snapshot", return_value={}
            ):
                with self.assertRaisesRegex(MissingEngineError, "保护用户"):
                    convert_with_wps(source, root / "out", "pdf")

        self.assertEqual(events, ["initialize", "uninitialize"])

    def test_wps_rejects_unowned_existing_process_without_quitting_it(self) -> None:
        events: list[str] = []

        class PythonCom:
            @staticmethod
            def CoInitialize() -> None:
                events.append("initialize")

            @staticmethod
            def CoUninitialize() -> None:
                events.append("uninitialize")

        class Application:
            def Quit(self) -> None:
                events.append("quit")

        class Client:
            @staticmethod
            def DispatchEx(_prog_id: str) -> Application:
                return Application()

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            executable = root / "wps.exe"
            executable.touch()
            source = root / "合同.docx"
            source.write_bytes(b"placeholder")
            statuses = {
                "writer": wps.WpsEngineStatus(
                    True,
                    "writer",
                    "KWPS.Application",
                    executable=executable,
                ),
                "spreadsheets": wps.WpsEngineStatus(
                    False, "spreadsheets", "KET.Application"
                ),
                "presentation": wps.WpsEngineStatus(
                    False, "presentation", "KWPP.Application"
                ),
            }
            with patch.object(
                wps, "detect_wps_engines", return_value=statuses
            ), patch.object(
                wps, "_load_pywin32", return_value=(PythonCom, Client)
            ), patch(
                "docuforge.processors.office._windows_process_snapshot",
                return_value={},
            ), patch(
                "docuforge.processors.office._office_application_pid",
                return_value=321,
            ), patch(
                "docuforge.processors.office._new_owned_office_process",
                return_value=None,
            ):
                with self.assertRaisesRegex(MissingEngineError, "无文档空闲状态"):
                    convert_with_wps(source, root / "out", "pdf")

        self.assertEqual(events, ["initialize", "uninitialize"])

    def test_wps_reuses_idle_instance_and_restores_global_settings(self) -> None:
        events: list[str] = []

        class PythonCom:
            @staticmethod
            def CoInitialize() -> None:
                events.append("initialize")

            @staticmethod
            def CoUninitialize() -> None:
                events.append("uninitialize")

        class Document:
            def SaveAs(self, path: str, _format_code: int) -> None:
                Path(path).write_bytes(b"converted")

            def Close(self, _save: bool) -> None:
                events.append("close")

        class Documents:
            Count = 0

            def Open(self, _path: str, **_kwargs: object) -> Document:
                return Document()

        class Application:
            def __init__(self) -> None:
                self.Documents = Documents()
                self.AutomationSecurity = 1
                self.Visible = True
                self.DisplayAlerts = 1
                self.ScreenUpdating = True
                self.EnableEvents = True
                self.AskToUpdateLinks = True

            def Quit(self) -> None:
                events.append("quit")

        application = Application()

        class Client:
            @staticmethod
            def DispatchEx(_prog_id: str) -> Application:
                return application

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            executable = (root / "wps.exe").resolve()
            executable.touch()
            source = root / "合同.docx"
            source.write_bytes(b"placeholder")
            identity = type("Identity", (), {"executable": executable})()
            statuses = {
                "writer": wps.WpsEngineStatus(
                    True,
                    "writer",
                    "KWPS.Application",
                    executable=executable,
                ),
                "spreadsheets": wps.WpsEngineStatus(
                    False, "spreadsheets", "KET.Application"
                ),
                "presentation": wps.WpsEngineStatus(
                    False, "presentation", "KWPP.Application"
                ),
            }
            with patch.object(
                wps, "detect_wps_engines", return_value=statuses
            ), patch.object(
                wps, "_load_pywin32", return_value=(PythonCom, Client)
            ), patch(
                "docuforge.processors.office._windows_process_snapshot",
                return_value={},
            ), patch(
                "docuforge.processors.office._office_application_pid",
                return_value=321,
            ), patch(
                "docuforge.processors.office._new_owned_office_process",
                return_value=None,
            ), patch(
                "docuforge.processors.office._windows_process_identity",
                return_value=identity,
            ):
                output = convert_with_wps(source, root / "out", "docx")
                output_bytes = output[0].read_bytes()

        self.assertEqual(output_bytes, b"converted")
        self.assertNotIn("quit", events)
        self.assertEqual(application.AutomationSecurity, 1)
        self.assertTrue(application.Visible)
        self.assertEqual(application.DisplayAlerts, 1)
        self.assertTrue(application.ScreenUpdating)
        self.assertTrue(application.EnableEvents)
        self.assertTrue(application.AskToUpdateLinks)

    def test_spreadsheets_open_safely_and_prepare_layout_before_export(
        self,
    ) -> None:
        events: list[object] = []
        open_calls: list[tuple[str, dict[str, object]]] = []

        class PythonCom:
            @staticmethod
            def CoInitialize() -> None:
                events.append("initialize")

            @staticmethod
            def CoUninitialize() -> None:
                events.append("uninitialize")

        class Workbook:
            def ExportAsFixedFormat(self, output_type: int, path: str) -> None:
                events.append("export")
                self.export_args = (output_type, path)
                Path(path).write_bytes(b"pdf")

            def Close(self, _save: bool) -> None:
                events.append("close")

        workbook = Workbook()

        class Workbooks:
            def Open(self, path: str, **kwargs: object) -> Workbook:
                events.append("open")
                open_calls.append((path, kwargs))
                return workbook

        class Application:
            def __init__(self) -> None:
                object.__setattr__(self, "Workbooks", Workbooks())

            def __setattr__(self, name: str, value: object) -> None:
                if name == "AutomationSecurity":
                    events.append(("security", value))
                object.__setattr__(self, name, value)

            def Quit(self) -> None:
                events.append("quit")

        application = Application()

        class Client:
            @staticmethod
            def DispatchEx(_prog_id: str) -> Application:
                return application

        executable = Path("C:/fake/et.exe")
        identity = office._OfficeProcessIdentity(
            2001, executable, "2026-08-16T12:00:00"
        )
        statuses = {
            "writer": wps.WpsEngineStatus(False, "writer", "KWPS.Application"),
            "spreadsheets": wps.WpsEngineStatus(
                True, "spreadsheets", "KET.Application", executable=executable
            ),
            "presentation": wps.WpsEngineStatus(
                False, "presentation", "KWPP.Application"
            ),
        }

        def record_layout(*_args: object, **_kwargs: object) -> None:
            events.append("layout")

        with tempfile.TemporaryDirectory() as folder, patch.object(
            wps, "detect_wps_engines", return_value=statuses
        ), patch.object(wps, "_load_pywin32", return_value=(PythonCom, Client)), patch(
            "docuforge.processors.wps._new_owned_wps_processes",
            return_value=(identity,),
        ), patch.object(
            office, "_wait_for_owned_office_exit"
        ), patch(
            "docuforge.processors.excel_pdf_layout.prepare_excel_workbook_for_pdf",
            side_effect=record_layout,
        ) as prepare_layout:
            root = Path(folder)
            source = root / "报表.xlsx"
            source.write_bytes(b"placeholder")
            output = convert_with_wps(
                source,
                root / "out",
                "pdf",
                excel_pdf_layout="fit_width",
                excel_pdf_paper="a4",
                excel_pdf_orientation="landscape",
                excel_pdf_margin="narrow",
            )
            self.assertEqual(output[0].read_bytes(), b"pdf")

        self.assertEqual(
            open_calls,
            [
                (
                    str(source.resolve()),
                    {
                        "UpdateLinks": 0,
                        "ReadOnly": True,
                        "AddToMru": False,
                        "IgnoreReadOnlyRecommended": True,
                        "Notify": False,
                    },
                )
            ],
        )
        prepare_layout.assert_called_once_with(
            workbook,
            application,
            layout="fit_width",
            paper="a4",
            orientation="landscape",
            margin="narrow",
        )
        self.assertLess(events.index("layout"), events.index("export"))
        self.assertEqual(workbook.export_args[0], 0)


if __name__ == "__main__":
    unittest.main()
