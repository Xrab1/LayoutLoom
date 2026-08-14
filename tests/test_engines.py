from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from docuforge import engines
from docuforge.models import Capability
from docuforge.processors.office import OfficeEngineStatus
from docuforge.processors.video import VideoEngineStatus


class EngineCapabilityTests(unittest.TestCase):
    def setUp(self) -> None:
        self._clear_caches()

    def tearDown(self) -> None:
        self._clear_caches()

    @staticmethod
    def _clear_caches() -> None:
        for probe in (
            engines.poppler_bin_path,
            engines.pdf_to_word_capability,
            engines.slideshow_video_capability,
            engines.video_transform_capability,
            engines.audio_extraction_capability,
            engines.video_capability,
            engines.video_slide_extraction_capability,
            engines.ppt_video_capability,
            engines.office_render_capability,
            engines.microsoft_powerpoint_capability,
        ):
            probe.cache_clear()

    @staticmethod
    def _status(
        video_encoders: tuple[str, ...] = (),
        audio_encoders: tuple[str, ...] = (),
    ) -> VideoEngineStatus:
        return VideoEngineStatus(
            bool(video_encoders),
            Path("ffmpeg"),
            video_encoders,
            "mock FFmpeg",
            ffprobe_executable=Path("ffprobe"),
            audio_encoders=audio_encoders,
        )

    def test_video_probes_require_their_actual_default_encoders(self) -> None:
        cases = (
            (("libvpx",), (), False, False),
            (("mpeg4",), (), True, False),
            (("libx264",), (), True, False),
            (("libx264",), ("aac",), True, True),
        )
        for encoders, audio_encoders, slideshow_ready, transform_ready in cases:
            with self.subTest(encoders=encoders, audio_encoders=audio_encoders):
                self._clear_caches()
                with patch(
                    "docuforge.processors.video.detect_video_engine",
                    return_value=self._status(encoders, audio_encoders),
                ):
                    self.assertEqual(
                        engines.slideshow_video_capability().runnable,
                        slideshow_ready,
                    )
                    self.assertEqual(
                        engines.video_transform_capability().runnable,
                        transform_ready,
                    )

    def test_audio_probe_is_independent_from_video_encoders(self) -> None:
        status = self._status(audio_encoders=("pcm_s16le",))
        with patch(
            "docuforge.processors.video.detect_video_engine", return_value=status
        ):
            self.assertFalse(engines.slideshow_video_capability().runnable)
            capability = engines.audio_extraction_capability()
        self.assertTrue(capability.runnable)
        self.assertIn("WAV", capability.reason)

        self._clear_caches()
        with patch(
            "docuforge.processors.video.detect_video_engine",
            return_value=self._status(audio_encoders=("aac",)),
        ):
            self.assertFalse(engines.audio_extraction_capability().runnable)

    def test_video_slide_probe_is_decode_only_and_reports_missing_cv_stack(self) -> None:
        with patch.object(
            engines,
            "has_module",
            side_effect=lambda name: name in {"numpy", "pptx"},
        ):
            capability = engines.video_slide_extraction_capability()
        self.assertFalse(capability.runnable)
        self.assertIn("OpenCV", capability.reason)

        engines.video_slide_extraction_capability.cache_clear()
        with patch.object(engines, "has_module", return_value=True):
            capability = engines.video_slide_extraction_capability()
        self.assertTrue(capability.runnable)
        self.assertIn("不依赖云端 OCR", capability.reason)

    def test_native_powerpoint_video_does_not_require_static_pipeline(self) -> None:
        native = Capability("external", "PowerPoint ready", "Microsoft PowerPoint COM")
        with patch.object(
            engines, "microsoft_powerpoint_capability", return_value=native
        ), patch.object(
            engines,
            "office_render_capability",
            side_effect=AssertionError("native probe must not require static renderer"),
        ):
            capability = engines.ppt_video_capability()
        self.assertTrue(capability.runnable)
        self.assertIn("静态模式", capability.reason)

    def test_wps_ppt_video_requires_poppler_and_slideshow_encoder(self) -> None:
        wps = Capability("external", "WPS ready", "WPS Office COM")
        ready = Capability("external", "ready", "mock")
        unavailable = Capability("unavailable", "missing", "mock")

        with patch.object(
            engines, "microsoft_powerpoint_capability", return_value=unavailable
        ), patch.object(
            engines, "office_render_capability", return_value=wps
        ), patch.object(
            engines, "pdf_render_capability", return_value=unavailable
        ), patch.object(
            engines, "slideshow_video_capability", return_value=ready
        ):
            self.assertFalse(engines.ppt_video_capability().runnable)

        engines.ppt_video_capability.cache_clear()
        with patch.object(
            engines, "microsoft_powerpoint_capability", return_value=unavailable
        ), patch.object(
            engines, "office_render_capability", return_value=wps
        ), patch.object(
            engines, "pdf_render_capability", return_value=ready
        ), patch.object(
            engines, "slideshow_video_capability", return_value=unavailable
        ):
            self.assertFalse(engines.ppt_video_capability().runnable)

        engines.ppt_video_capability.cache_clear()
        with patch.object(
            engines, "microsoft_powerpoint_capability", return_value=unavailable
        ), patch.object(
            engines, "office_render_capability", return_value=wps
        ), patch.object(
            engines, "pdf_render_capability", return_value=ready
        ), patch.object(
            engines, "slideshow_video_capability", return_value=ready
        ):
            capability = engines.ppt_video_capability()
        self.assertTrue(capability.runnable)
        self.assertIn("Poppler", capability.reason)
        self.assertIn("FFmpeg", capability.reason)

    def test_office_render_capability_prefers_wps_over_microsoft(self) -> None:
        statuses = {
            "microsoft_office": OfficeEngineStatus(True, reason="Microsoft ready"),
            "microsoft_word": OfficeEngineStatus(True, reason="Word ready"),
            "microsoft_excel": OfficeEngineStatus(True, reason="Excel ready"),
            "microsoft_powerpoint": OfficeEngineStatus(True, reason="PowerPoint ready"),
            "libreoffice": OfficeEngineStatus(True, reason="LibreOffice ready"),
        }
        wps_statuses = {
            "writer": SimpleNamespace(
                available=True, kind="writer", reason="WPS ready"
            ),
            "spreadsheets": SimpleNamespace(
                available=True, kind="spreadsheets", reason="WPS ready"
            ),
            "presentation": SimpleNamespace(
                available=True, kind="presentation", reason="WPS ready"
            ),
        }
        with patch(
            "docuforge.processors.office.detect_office_engines",
            return_value=statuses,
        ), patch(
            "docuforge.processors.wps.detect_wps_engines",
            return_value=wps_statuses,
        ):
            capability = engines.office_render_capability("powerpoint")

        self.assertEqual(capability.engine, "WPS Office COM")

    def test_bundled_poppler_library_bin_layout_is_detected(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            binary_dir = Path(folder) / "poppler" / "Library" / "bin"
            binary_dir.mkdir(parents=True)
            (binary_dir / "pdftoppm.exe").touch()
            (binary_dir / "pdfinfo.exe").touch()
            with patch.object(engines.sys, "_MEIPASS", folder, create=True):
                self.assertEqual(engines.poppler_bin_path(), str(binary_dir))

    def test_poppler_explicit_path_works_without_shell_path(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            binary_dir = Path(folder) / "Library" / "bin"
            binary_dir.mkdir(parents=True)
            (binary_dir / "pdftoppm.exe").touch()
            (binary_dir / "pdfinfo.exe").touch()
            with patch.dict(
                engines.os.environ,
                {"DOCUFORGE_POPPLER_PATH": str(binary_dir)},
                clear=False,
            ):
                self.assertEqual(engines.poppler_bin_path(), str(binary_dir))

    def test_pdf_to_word_requires_high_precision_layout_dependencies(self) -> None:
        with patch.object(
            engines, "has_module", side_effect=lambda name: name == "docx"
        ):
            capability = engines.pdf_to_word_capability()
        self.assertFalse(capability.runnable)
        self.assertIn("pdf2docx", capability.reason)
        self.assertIn("pymupdf", capability.reason)

    def test_pdf_to_word_exposes_builtin_modes_and_optional_word_reflow(self) -> None:
        with patch.object(
            engines,
            "has_module",
            side_effect=lambda name: name in {"docx", "pdf2docx", "pymupdf"},
        ):
            capability = engines.pdf_to_word_capability()
        self.assertTrue(capability.runnable)
        self.assertEqual(capability.status, "ready")
        self.assertIn("版式优先混合", capability.reason)
        self.assertIn("Microsoft Word 原生转换", capability.reason)
        self.assertIn("桌面版 Word", capability.reason)
        self.assertIn("pdf2docx", capability.engine)


if __name__ == "__main__":
    unittest.main()
