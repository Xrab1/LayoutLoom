from __future__ import annotations

import inspect
import tempfile
import unittest
from pathlib import Path
from typing import Any, Callable
from unittest.mock import patch

from docuforge.models import MissingEngineError, ValidationError
from docuforge.processors import office_com


class OfficeComProcessorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory(prefix="docuforge-com-")
        self.root = Path(self.temporary.name)
        self.output = self.root / "中文输出"
        self.output.mkdir()
        self.ppt = self.root / "演示文稿.pptx"
        self.excel = self.root / "数据报表.xlsx"
        self.word = self.root / "合同文档.docx"
        # Public functions validate the path and extension before probing COM.
        # They do not parse these placeholders in dependency/parameter tests.
        self.ppt.write_bytes(b"placeholder")
        self.excel.write_bytes(b"placeholder")
        self.word.write_bytes(b"placeholder")

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_public_signatures_are_plain_and_stable(self) -> None:
        expected: dict[Callable[..., Any], list[str]] = {
            office_com.ppt_to_images: [
                "source",
                "output_dir",
                "format",
                "width",
                "height",
                "overwrite",
            ],
            office_com.ppt_to_video: [
                "source",
                "output_path",
                "use_timings",
                "slide_duration",
                "resolution",
                "fps",
                "quality",
                "timeout",
                "overwrite",
            ],
            office_com.ppt_to_long_image: [
                "source",
                "output_path",
                "direction",
                "spacing",
                "background",
                "width",
                "overwrite",
            ],
            office_com.ppt_modify_master: [
                "source",
                "output_dir",
                "background_color",
                "font_name",
                "footer_text",
                "overwrite",
            ],
            office_com.excel_create_pivot: [
                "source",
                "output_dir",
                "source_sheet",
                "source_range",
                "target_sheet",
                "target_cell",
                "row_fields",
                "column_fields",
                "data_field",
                "function",
                "overwrite",
            ],
            office_com.word_remove_blank_pages: ["source", "output_dir", "overwrite"],
        }
        for function, parameter_names in expected.items():
            signature = inspect.signature(function)
            self.assertEqual(list(signature.parameters), parameter_names)
            self.assertFalse(
                any(
                    parameter.kind is inspect.Parameter.KEYWORD_ONLY
                    for parameter in signature.parameters.values()
                )
            )

    def test_non_windows_reports_missing_engine_for_each_office_family(self) -> None:
        calls = (
            lambda: office_com.ppt_to_images(self.ppt, self.output),
            lambda: office_com.excel_create_pivot(
                self.excel,
                self.output,
                "数据",
                "A1:B2",
                row_fields=("类别",),
            ),
            lambda: office_com.word_remove_blank_pages(self.word, self.output),
        )
        with patch.object(office_com.sys, "platform", "linux"):
            for call in calls:
                with self.subTest(call=call), self.assertRaisesRegex(
                    MissingEngineError, "Windows"
                ):
                    call()
        self.assertEqual(list(self.output.iterdir()), [])

    def test_missing_pywin32_is_reported_clearly(self) -> None:
        with patch.object(office_com.sys, "platform", "win32"), patch.object(
            office_com.importlib,
            "import_module",
            side_effect=ModuleNotFoundError("pythoncom"),
        ):
            with self.assertRaisesRegex(MissingEngineError, "pywin32"):
                office_com._load_com_runtime()

    def test_unregistered_office_cleans_up_com(self) -> None:
        class PythonCom:
            initialized = 0
            uninitialized = 0

            @classmethod
            def CoInitialize(cls) -> None:
                cls.initialized += 1

            @classmethod
            def CoUninitialize(cls) -> None:
                cls.uninitialized += 1

        class Client:
            @staticmethod
            def DispatchEx(_prog_id: str):
                raise OSError("class not registered")

        runtime = office_com._ComRuntime(PythonCom, Client)
        with self.assertRaisesRegex(MissingEngineError, "桌面版 Office"):
            with office_com._com_application(
                runtime, "PowerPoint.Application", "PowerPoint"
            ):
                self.fail("unregistered COM application must not yield")
        self.assertEqual(PythonCom.initialized, 1)
        self.assertEqual(PythonCom.uninitialized, 1)

    def test_com_application_reuses_strict_microsoft_identity_validation(self) -> None:
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

        runtime = office_com._ComRuntime(PythonCom, Client)
        with patch(
            "docuforge.processors.office._validate_microsoft_com_application"
        ) as validate:
            with office_com._com_application(
                runtime, "PowerPoint.Application", "PowerPoint"
            ) as application:
                self.assertIsInstance(application, Application)

        validate.assert_called_once_with(
            application, "PowerPoint.Application", "PowerPoint"
        )
        self.assertEqual(events, ["initialize", "quit", "uninitialize"])

    def test_powerpoint_parameter_validation_precedes_engine_probe(self) -> None:
        with patch.object(office_com.sys, "platform", "linux"):
            invalid_calls = (
                lambda: office_com.ppt_to_images(self.ppt, self.output, format="gif"),
                lambda: office_com.ppt_to_images(self.ppt, self.output, width=0),
                lambda: office_com.ppt_to_video(
                    self.ppt, self.output, use_timings="yes"  # type: ignore[arg-type]
                ),
                lambda: office_com.ppt_to_video(self.ppt, self.output, fps=101),
                lambda: office_com.ppt_to_video(self.ppt, self.output, timeout=0),
                lambda: office_com.ppt_to_long_image(
                    self.ppt, self.output / "long.png", direction="diagonal"
                ),
                lambda: office_com.ppt_to_long_image(
                    self.ppt, self.output / "long.png", background="not-a-color"
                ),
                lambda: office_com.ppt_modify_master(self.ppt, self.output),
            )
            for call in invalid_calls:
                with self.subTest(call=call), self.assertRaises(ValidationError):
                    call()

    def test_excel_parameter_validation_precedes_engine_probe(self) -> None:
        with patch.object(office_com.sys, "platform", "linux"):
            invalid_calls = (
                lambda: office_com.excel_create_pivot(
                    self.excel, self.output, "数据", "A1", row_fields=("类别",)
                ),
                lambda: office_com.excel_create_pivot(
                    self.excel, self.output, "数据", "D1:A10", row_fields=("类别",)
                ),
                lambda: office_com.excel_create_pivot(
                    self.excel,
                    self.output,
                    "数据",
                    "A1:D10",
                    target_sheet="bad/name",
                    row_fields=("类别",),
                ),
                lambda: office_com.excel_create_pivot(
                    self.excel,
                    self.output,
                    "数据",
                    "A1:D10",
                    row_fields=("类别",),
                    column_fields=("类别",),
                ),
                lambda: office_com.excel_create_pivot(
                    self.excel,
                    self.output,
                    "数据",
                    "A1:D10",
                    row_fields=("类别",),
                    function="median",
                ),
                lambda: office_com.excel_create_pivot(
                    self.excel,
                    self.output,
                    "数据",
                    "A1:D10",
                    target_sheet="数据",
                    target_cell="B2",
                    row_fields=("类别",),
                ),
            )
            for call in invalid_calls:
                with self.subTest(call=call), self.assertRaises(ValidationError):
                    call()

    def test_image_dimension_color_and_video_target_helpers(self) -> None:
        self.assertEqual(office_com._normalize_image_format(".JPEG"), "jpg")
        self.assertEqual(office_com._scaled_dimensions(16, 9, 1920, None), (1920, 1080))
        self.assertEqual(office_com._scaled_dimensions(16, 9, 800, 600), (800, 600))
        self.assertEqual(office_com._parse_rgb("#1a2B3c"), (26, 43, 60))
        self.assertEqual(office_com._parse_rgb("f00"), (255, 0, 0))
        self.assertEqual(office_com._rgb_to_office_bgr((1, 2, 3)), 0x030201)

        video_directory = self.root / "视频输出"
        target = office_com._video_target(self.ppt, video_directory)
        self.assertEqual(target, (video_directory / "演示文稿.mp4").resolve())
        self.assertTrue(video_directory.is_dir())
        explicit = office_com._video_target(self.ppt, self.root / "成片.mp4")
        self.assertEqual(explicit.name, "成片.mp4")
        with self.assertRaises(ValidationError):
            office_com._video_target(self.ppt, self.root / "成片.avi")

    def test_output_reservation_and_atomic_publication(self) -> None:
        existing = self.output / "页面_001.png"
        existing.write_bytes(b"old")
        staged = self.root / "staged.png"
        staged.write_bytes(b"complete-new-file")

        with office_com._reserve_paths([existing], overwrite=False) as targets:
            target = targets[0]
            self.assertNotEqual(target, existing)
            office_com._publish_file(staged, target)
        self.assertEqual(existing.read_bytes(), b"old")
        self.assertEqual(target.read_bytes(), b"complete-new-file")
        self.assertFalse(
            any(path.name.endswith(".tmp.png") for path in self.output.iterdir())
        )

        reserved: Path | None = None
        with self.assertRaises(RuntimeError):
            with office_com._reserve_paths(
                [self.output / "失败.png"], overwrite=False
            ) as targets:
                reserved = targets[0]
                raise RuntimeError("simulated pre-publication failure")
        assert reserved is not None
        self.assertFalse(reserved.exists())

    def test_excel_range_field_and_function_helpers(self) -> None:
        self.assertEqual(office_com._normalize_a1_range("$a$1:$xfd$20"), "A1:XFD20")
        self.assertEqual(office_com._normalize_a1_cell("$c$7"), "C7")
        self.assertEqual(
            office_com._normalize_field_names("地区, 金额", "fields"), ("地区", "金额")
        )
        self.assertEqual(office_com._normalize_pivot_function("average"), -4106)
        self.assertEqual(office_com._normalize_pivot_function("求和"), -4157)
        with self.assertRaises(ValidationError):
            office_com._normalize_a1_range("XFE1:XFE10")
        with self.assertRaises(ValidationError):
            office_com._normalize_field_names(("地区", "地区"), "fields")

    def test_word_ooxml_safety_analysis_is_conservative(self) -> None:
        safe_xml = f"""
        <w:document xmlns:w="{office_com._W}">
          <w:body>
            <w:p><w:r><w:br w:type="page"/></w:r></w:p>
            <w:sectPr/>
          </w:body>
        </w:document>
        """.encode()
        report = office_com._analyze_word_document_xml(safe_xml)
        self.assertTrue(report.safe)
        self.assertEqual(report.explicit_page_breaks, 1)

        section_xml = f"""
        <w:document xmlns:w="{office_com._W}">
          <w:body>
            <w:p><w:pPr><w:sectPr/></w:pPr></w:p>
            <w:sectPr/>
          </w:body>
        </w:document>
        """.encode()
        section_report = office_com._analyze_word_document_xml(section_xml)
        self.assertFalse(section_report.safe)
        self.assertIn("文档包含中间分节符", section_report.high_risk_reasons)

        table_break_xml = f"""
        <w:document xmlns:w="{office_com._W}">
          <w:body>
            <w:tbl><w:tr><w:tc><w:p><w:r><w:br w:type="page"/></w:r></w:p></w:tc></w:tr></w:tbl>
            <w:sectPr/>
          </w:body>
        </w:document>
        """.encode()
        table_report = office_com._analyze_word_document_xml(table_break_xml)
        self.assertFalse(table_report.safe)
        self.assertIn("表格单元格中包含手动分页符", table_report.high_risk_reasons)

        columns_xml = f"""
        <w:document xmlns:w="{office_com._W}">
          <w:body><w:p/><w:sectPr><w:cols w:num="2"/></w:sectPr></w:body>
        </w:document>
        """.encode()
        columns_report = office_com._analyze_word_document_xml(columns_xml)
        self.assertFalse(columns_report.safe)
        self.assertIn("文档使用多栏排版", columns_report.high_risk_reasons)

        settings_xml = f"""
        <w:settings xmlns:w="{office_com._W}">
          <w:trackRevisions/>
          <w:documentProtection w:enforcement="1"/>
        </w:settings>
        """.encode()
        self.assertEqual(
            office_com._word_settings_risks(settings_xml),
            ("文档已启用修订跟踪", "文档启用了编辑保护"),
        )

    def test_explicit_blank_page_text_needs_a_manual_break(self) -> None:
        self.assertTrue(office_com._is_explicit_blank_page_text("\r\x0c\r"))
        self.assertTrue(office_com._is_blank_word_page_text("\r\r"))
        self.assertFalse(office_com._is_explicit_blank_page_text("\r\r"))
        self.assertFalse(office_com._is_explicit_blank_page_text("正文\x0c\r"))


if __name__ == "__main__":
    unittest.main()
