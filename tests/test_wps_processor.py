from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docuforge.models import MissingEngineError, ValidationError
from docuforge.processors import wps
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

        status = wps.WpsEngineStatus(True, "writer", "KWPS.Application")
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
        ), patch.object(wps, "_load_pywin32", return_value=(PythonCom, Client)):
            root = Path(folder)
            source = root / "合同.docx"
            source.write_bytes(b"placeholder")
            output = convert_with_wps(source, root / "out", "docx")
            self.assertEqual(output[0].read_bytes(), b"converted")

        self.assertLess(events.index(("security", 3)), events.index("open"))
        self.assertEqual(events[0], "initialize")
        self.assertEqual(events[-1], "uninitialize")

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

        statuses = {
            "writer": wps.WpsEngineStatus(False, "writer", "KWPS.Application"),
            "spreadsheets": wps.WpsEngineStatus(
                True, "spreadsheets", "KET.Application"
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
