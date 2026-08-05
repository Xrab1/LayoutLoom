from __future__ import annotations

import math
import struct
import subprocess
import tempfile
import unittest
import wave
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from docuforge.models import CancelledError, MissingEngineError, ValidationError
from docuforge.processors import video
from docuforge.processors.office import OfficeEngineStatus


class _WaitingProcess:
    def __init__(self) -> None:
        self.returncode: int | None = None
        self.terminated = False

    def communicate(self, timeout: float) -> tuple[str, str]:
        raise subprocess.TimeoutExpired(["fake"], timeout)

    def poll(self) -> int | None:
        return self.returncode

    def terminate(self) -> None:
        self.terminated = True
        self.returncode = -15

    def kill(self) -> None:
        self.returncode = -9

    def wait(self, timeout: float) -> int:
        assert self.returncode is not None
        return self.returncode


class VideoProcessorTests(unittest.TestCase):
    @staticmethod
    def _write_tone(path: Path, duration: float = 0.25) -> None:
        sample_rate = 48_000
        frame_count = round(duration * sample_rate)
        with wave.open(str(path), "wb") as writer:
            writer.setnchannels(1)
            writer.setsampwidth(2)
            writer.setframerate(sample_rate)
            writer.writeframes(
                b"".join(
                    struct.pack(
                        "<h",
                        round(9000 * math.sin(2 * math.pi * 440 * index / sample_rate)),
                    )
                    for index in range(frame_count)
                )
            )

    def test_engine_probe(self) -> None:
        status = video.detect_video_engine()
        self.assertIsInstance(status.encoders, tuple)
        self.assertIsInstance(status.audio_encoders, tuple)
        if status.available:
            self.assertIsNotNone(status.executable)
            self.assertTrue(status.encoders)
            if status.ffprobe_executable is not None:
                self.assertTrue(status.ffprobe_executable.is_file())

    def test_engine_probe_parses_exact_encoder_names_and_ffprobe(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            ffmpeg = root / "ffmpeg.exe"
            ffprobe = root / "ffprobe.exe"
            ffmpeg.touch()
            ffprobe.touch()
            encoder_output = """Encoders:
 V..... libx264 H.264
 V..... h264_nvenc NVIDIA
 V..... notlibx264 Not an exact match
 A..... aac AAC
 A..... libmp3lame MP3
"""
            responses = [
                subprocess.CompletedProcess(
                    [str(ffmpeg), "-hide_banner", "-version"], 0, "ffmpeg 7", ""
                ),
                subprocess.CompletedProcess(
                    [str(ffmpeg), "-encoders"], 0, encoder_output, ""
                ),
                subprocess.CompletedProcess([str(ffprobe), "-version"], 0, "ok", ""),
            ]
            video.detect_video_engine.cache_clear()
            try:
                with patch.object(
                    video.shutil,
                    "which",
                    side_effect=lambda name: str(
                        ffmpeg if name == "ffmpeg" else ffprobe
                    ),
                ), patch.object(
                    video, "_ffmpeg_candidates", return_value=[ffmpeg.resolve()]
                ), patch.object(video.subprocess, "run", side_effect=responses) as run:
                    status = video.detect_video_engine()
                self.assertTrue(status.available)
                self.assertEqual(status.encoders, ("libx264", "h264_nvenc"))
                self.assertEqual(status.audio_encoders, ("aac", "libmp3lame"))
                self.assertEqual(status.ffprobe_executable, ffprobe.resolve())
                self.assertIn("-hide_banner", run.call_args_list[0].args[0])
            finally:
                video.detect_video_engine.cache_clear()

    def test_engine_probe_preserves_audio_and_ffprobe_without_video_encoder(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            ffmpeg = root / "ffmpeg.exe"
            ffprobe = root / "ffprobe.exe"
            ffmpeg.touch()
            ffprobe.touch()
            responses = [
                subprocess.CompletedProcess(
                    [str(ffmpeg), "-hide_banner", "-version"], 0, "ffmpeg 7", ""
                ),
                subprocess.CompletedProcess(
                    [str(ffmpeg), "-encoders"],
                    0,
                    "Encoders:\n A..... aac AAC\n A..... pcm_s16le PCM\n",
                    "",
                ),
                subprocess.CompletedProcess([str(ffprobe), "-version"], 0, "ok", ""),
            ]
            video.detect_video_engine.cache_clear()
            try:
                with patch.object(
                    video.shutil,
                    "which",
                    side_effect=lambda name: str(
                        ffmpeg if name == "ffmpeg" else ffprobe
                    ),
                ), patch.object(
                    video, "_ffmpeg_candidates", return_value=[ffmpeg.resolve()]
                ), patch.object(video.subprocess, "run", side_effect=responses):
                    status = video.detect_video_engine()
                self.assertFalse(status.available)
                self.assertEqual(status.encoders, ())
                self.assertEqual(status.audio_encoders, ("aac", "pcm_s16le"))
                self.assertEqual(status.ffprobe_executable, ffprobe.resolve())
                self.assertIn("未提供可用的视频编码器", status.reason)
                with patch.object(video, "detect_video_engine", return_value=status):
                    self.assertIs(video._require_engine(require_video=False), status)
            finally:
                video.detect_video_engine.cache_clear()

    def test_images_to_mp4_with_fade(self) -> None:
        status = video.detect_video_engine()
        if not status.available:
            self.skipTest(status.reason)
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            first = root / "红色 图片.png"
            second = root / "蓝色 图片.png"
            Image.new("RGB", (160, 100), "red").save(first)
            Image.new("RGB", (100, 160), "blue").save(second)
            output = root / "幻灯片 成片.mp4"
            result = video.images_to_video(
                [first, second],
                output,
                slide_duration=0.8,
                fps=5,
                resolution=(320, 240),
                transition="fade",
                transition_duration=0.2,
            )
            self.assertEqual(result, [output])
            self.assertGreater(output.stat().st_size, 1000)
            info = video._probe_media(output, status, require_video=True)
            if info is not None:
                duration = video._payload_duration(info)
                self.assertIsNotNone(duration)
                self.assertGreaterEqual(duration or 0, 1.35)

    def test_normalized_and_transition_frames_are_lossless_png(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            first_source = root / "small.png"
            second_source = root / "second.png"
            Image.new("RGBA", (20, 10), (20, 180, 40, 255)).save(first_source)
            Image.new("RGB", (20, 10), "blue").save(second_source)
            first_frame = root / "frame_00001.png"
            second_frame = root / "frame_00002.png"
            video._normalize_frame(first_source, first_frame, (40, 40), (0, 0, 0))
            video._normalize_frame(second_source, second_frame, (40, 40), (0, 0, 0))

            with Image.open(first_frame) as frame:
                self.assertEqual(frame.format, "PNG")
                self.assertEqual(frame.size, (40, 40))
                self.assertEqual(frame.getpixel((20, 0)), (0, 0, 0))
                self.assertNotEqual(frame.getpixel((20, 10)), (0, 0, 0))

            manifest = root / "frames.ffconcat"
            generated = video._write_manifest(
                [first_frame, second_frame],
                manifest,
                slide_duration=1,
                transition="fade",
                transition_duration=0.2,
                fps=5,
            )
            self.assertEqual(len(generated), 1)
            self.assertTrue(all(path.suffix == ".png" for path in generated))
            with Image.open(generated[0]) as transition:
                self.assertEqual(transition.format, "PNG")

    def test_short_background_audio_does_not_truncate_video(self) -> None:
        status = video.detect_video_engine()
        if not status.available:
            self.skipTest(status.reason)
        if not any(
            encoder in status.audio_encoders
            for encoder in ("aac", "libfdk_aac", "libvo_aacenc")
        ):
            self.skipTest("FFmpeg 没有 AAC 编码器")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            image = root / "画面.png"
            audio = root / "短 音乐.wav"
            output = root / "带音乐 成片.mp4"
            Image.new("RGB", (160, 90), "green").save(image)
            self._write_tone(audio, duration=0.2)
            video.images_to_video(
                [image],
                output,
                slide_duration=1.2,
                fps=10,
                resolution=(320, 180),
                audio_path=audio,
            )
            info = video._probe_media(
                output, status, require_video=True, require_audio=True
            )
            if info is not None:
                self.assertGreaterEqual(video._payload_duration(info) or 0, 1.0)

    def test_real_transcode_compress_trim_and_extract_audio(self) -> None:
        status = video.detect_video_engine()
        if not status.available or "libx264" not in status.encoders:
            self.skipTest("真实测试需要 libx264")
        if not any(
            encoder in status.audio_encoders
            for encoder in ("aac", "libfdk_aac", "libvo_aacenc")
        ):
            self.skipTest("真实测试需要 AAC 编码器")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            image = root / "源 画面.png"
            tone = root / "源 音频.wav"
            source = root / "视频 源.mp4"
            Image.new("RGB", (160, 90), "purple").save(image)
            self._write_tone(tone, duration=0.35)
            video.images_to_video(
                [image],
                source,
                slide_duration=1.2,
                fps=10,
                resolution=(320, 180),
                audio_path=tone,
            )

            mkv = root / "转码 结果.mkv"
            mov = root / "转码 结果.mov"
            compressed = root / "压缩 结果.mp4"
            precise = root / "精准 裁剪.mp4"
            copied = root / "快速 裁剪.mkv"
            self.assertEqual(video.video_transcode(source, mkv), [mkv])
            self.assertEqual(video.video_transcode(source, mov), [mov])
            self.assertEqual(
                video.video_compress(source, compressed, codec="h264", quality=30),
                [compressed],
            )
            self.assertEqual(
                video.video_trim(
                    source, precise, start="00:00.10", duration=0.4, mode="precise"
                ),
                [precise],
            )
            self.assertEqual(
                video.video_trim(source, copied, start=0, duration=0.4, mode="copy"),
                [copied],
            )
            for output in (mkv, mov, compressed, precise, copied):
                self.assertGreater(output.stat().st_size, 500)
                video._probe_media(output, status, require_video=True)

            audio_outputs = [root / "提取 音频.aac", root / "提取 音频.wav"]
            if any(
                encoder in status.audio_encoders
                for encoder in ("libmp3lame", "libshine")
            ):
                audio_outputs.append(root / "提取 音频.mp3")
            for extracted in audio_outputs:
                self.assertEqual(
                    video.video_extract_audio(source, extracted), [extracted]
                )
                self.assertGreater(extracted.stat().st_size, 500)
                video._probe_media(extracted, status, require_audio=True)

            if any(
                encoder in status.encoders for encoder in ("libvpx-vp9", "libvpx")
            ) and any(
                encoder in status.audio_encoders for encoder in ("libopus", "libvorbis")
            ):
                webm = root / "网页 视频.webm"
                self.assertEqual(video.video_transcode(source, webm), [webm])
                video._probe_media(webm, status, require_video=True, require_audio=True)

    def test_atomic_output_preserves_existing_file_on_encoder_failure(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            image = root / "frame.png"
            output = root / "movie.mp4"
            Image.new("RGB", (50, 50), "white").save(image)
            output.write_bytes(b"existing movie")
            status = video.VideoEngineStatus(
                True,
                Path("ffmpeg"),
                ("libx264",),
                "mock",
                audio_encoders=("aac",),
            )
            with patch.object(
                video, "detect_video_engine", return_value=status
            ), patch.object(
                video,
                "_run_process",
                side_effect=MissingEngineError("simulated failure"),
            ):
                with self.assertRaises(MissingEngineError):
                    video.images_to_video(
                        [image],
                        output,
                        resolution=(160, 120),
                        slide_duration=0.2,
                        fps=10,
                        overwrite=True,
                    )
            self.assertEqual(output.read_bytes(), b"existing movie")
            self.assertFalse(list(root.glob(".movie.*.tmp.mp4")))

    def test_process_cancellation_terminates_child(self) -> None:
        process = _WaitingProcess()
        with patch.object(
            video.subprocess, "Popen", return_value=process
        ), patch.object(
            video,
            "check_cancelled",
            side_effect=[None, CancelledError("cancelled")],
        ):
            with self.assertRaises(CancelledError):
                video._run_process(["ffmpeg", "-version"], timeout=10, label="测试进程")
        self.assertTrue(process.terminated)

    def test_transcode_and_trim_warn_when_extra_streams_are_discarded(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "multi-stream.mkv"
            source.write_bytes(b"input")
            status = video.VideoEngineStatus(
                True,
                Path("ffmpeg"),
                ("libx264",),
                "mock",
                audio_encoders=("aac",),
            )
            source_info = {
                "streams": [
                    {"codec_type": "video", "width": 320, "height": 180},
                    {"codec_type": "audio"},
                    {"codec_type": "audio"},
                    {"codec_type": "subtitle"},
                    {"codec_type": "data"},
                ],
                "format": {"duration": "1.0"},
            }
            output_info = {
                "streams": [
                    {"codec_type": "video", "width": 320, "height": 180},
                    {"codec_type": "audio"},
                ],
                "format": {"duration": "0.5"},
            }

            def fake_run(
                command: list[str], **_: object
            ) -> subprocess.CompletedProcess[str]:
                Path(command[-1]).write_bytes(b"output")
                return subprocess.CompletedProcess(command, 0, "", "")

            operations = (
                lambda target: video.video_transcode(source, target),
                lambda target: video.video_trim(
                    source, target, start=0, duration=0.5, mode="precise"
                ),
            )
            for index, operation in enumerate(operations):
                with self.subTest(index=index), patch.object(
                    video, "detect_video_engine", return_value=status
                ), patch.object(
                    video, "_probe_media", side_effect=[source_info, output_info]
                ), patch.object(
                    video, "_run_process", side_effect=fake_run
                ):
                    with self.assertWarnsRegex(
                        UserWarning,
                        r"2 条音轨.*字幕流.*数据/附件流.*主视频/主音轨",
                    ):
                        operation(root / f"output-{index}.mp4")

    def test_powerpoint_native_parameter_routing_and_quality_mapping(self) -> None:
        self.assertEqual(video._native_powerpoint_quality(0), 100)
        self.assertEqual(video._native_powerpoint_quality(20), 85)
        self.assertEqual(video._native_powerpoint_quality(40), 70)

        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "native.mp4"

            def export_video(*_args: object, **_kwargs: object) -> list[Path]:
                output.write_bytes(b"video")
                return [output]

            with patch.object(
                video,
                "_microsoft_powerpoint_status",
                return_value=OfficeEngineStatus(True, reason="ready"),
            ), patch(
                "docuforge.processors.office_com.ppt_to_video",
                side_effect=export_video,
            ) as native:
                self.assertEqual(
                    video.ppt_to_video(
                        "slides.pptx",
                        output,
                        mode="native",
                        slide_duration=5.0,
                        fps=100,
                        quality=20,
                    ),
                    [output],
                )
        native_kwargs = native.call_args.kwargs
        self.assertEqual(native_kwargs["slide_duration"], 5)
        self.assertIs(type(native_kwargs["slide_duration"]), int)
        self.assertEqual(native_kwargs["fps"], 100)
        self.assertIs(type(native_kwargs["fps"]), int)
        self.assertEqual(native_kwargs["quality"], 85)

        with self.assertRaisesRegex(ValidationError, "整数秒"):
            video.ppt_to_video(
                "slides.pptx",
                output,
                mode="native",
                slide_duration=2.5,
            )
        with self.assertRaisesRegex(ValidationError, "不能超过 100"):
            video.ppt_to_video(
                "slides.pptx",
                output,
                mode="native",
                slide_duration=5,
                fps=101,
            )

        fallback_cases = ((2.5, 30), (5.0, 101))
        for slide_duration, fps in fallback_cases:
            with self.subTest(slide_duration=slide_duration, fps=fps), patch.object(
                video, "ppt_to_static_video", return_value=[output]
            ) as static, patch(
                "docuforge.processors.office_com.ppt_to_video"
            ) as native:
                self.assertEqual(
                    video.ppt_to_video(
                        "slides.pptx",
                        output,
                        mode="auto",
                        slide_duration=slide_duration,
                        fps=fps,
                    ),
                    [output],
                )
                native.assert_not_called()
                static.assert_called_once()

    def test_ppt_static_rendering_auto_delegates_to_wps_first_office_router(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "slides.pptx"
            source.write_bytes(b"presentation")
            rendered_pdf = root / "render" / "slides.pdf"

            def convert(*_args: object, **_kwargs: object) -> list[Path]:
                rendered_pdf.parent.mkdir(parents=True, exist_ok=True)
                rendered_pdf.write_bytes(b"pdf")
                return [rendered_pdf]

            with patch(
                "docuforge.processors.office.convert_with_office",
                side_effect=convert,
            ) as convert_office:
                result = video._render_ppt_to_pdf(source, root / "render", "auto")

        self.assertEqual(result, rendered_pdf)
        self.assertEqual(convert_office.call_args.kwargs["engine"], "auto")

    def test_ppt_images_auto_does_not_call_powerpoint_native_export(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "slides.pptx"
            source.write_bytes(b"presentation")
            rendered_pdf = root / "slides.pdf"
            rendered_pdf.write_bytes(b"pdf")
            rendered_image = root / "slide.png"
            rendered_image.write_bytes(b"png")
            with patch(
                "docuforge.processors.office_com.ppt_to_images",
                side_effect=AssertionError("auto must not call PowerPoint native"),
            ) as native, patch.object(
                video, "_render_ppt_to_pdf", return_value=rendered_pdf
            ) as render_pdf, patch(
                "docuforge.processors.pdf.pdf_to_images",
                return_value=[rendered_image],
            ):
                result = video.ppt_to_images(source, root / "out", renderer="auto")

        self.assertEqual(result, [rendered_image])
        native.assert_not_called()
        self.assertEqual(render_pdf.call_args.args[2], "auto")

    def test_explicit_microsoft_ppt_images_failure_does_not_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "slides.pptx"
            source.write_bytes(b"presentation")
            with patch.object(
                video,
                "_microsoft_powerpoint_status",
                return_value=OfficeEngineStatus(True, reason="ready"),
            ), patch(
                "docuforge.processors.office_com.ppt_to_images",
                side_effect=MissingEngineError("native failed"),
            ), patch.object(
                video, "_render_ppt_to_pdf"
            ) as fallback:
                with self.assertRaisesRegex(MissingEngineError, "native failed"):
                    video.ppt_to_images(
                        source, root / "out", renderer="microsoft_office"
                    )
        fallback.assert_not_called()

    def test_single_slide_ppt_long_image_supports_png_jpg_and_webp(self) -> None:
        expected_formats = {"png": "PNG", "jpg": "JPEG", "webp": "WEBP"}
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "slides.pptx"
            source.write_bytes(b"presentation")
            slide = root / "slide.png"
            Image.new("RGBA", (64, 36), (20, 120, 220, 160)).save(slide)
            for extension, expected_format in expected_formats.items():
                with self.subTest(extension=extension), patch.object(
                    video, "ppt_to_images", return_value=[slide]
                ):
                    target = root / "out" / f"single.{extension}"
                    result = video.ppt_to_long_image(source, target, renderer="auto")
                    self.assertEqual(result, [target])
                    with Image.open(target) as converted:
                        self.assertEqual(converted.format, expected_format)
                        self.assertEqual(converted.size, (64, 36))

    def test_single_slide_ppt_long_image_honors_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "slides.pptx"
            source.write_bytes(b"presentation")
            slide = root / "slide.png"
            Image.new("RGB", (32, 18), "green").save(slide)
            target = root / "single.png"
            target.write_bytes(b"existing")
            with patch.object(video, "ppt_to_images", return_value=[slide]):
                unique_result = video.ppt_to_long_image(
                    source, target, overwrite=False
                )[0]
                self.assertNotEqual(unique_result, target)
                self.assertEqual(target.read_bytes(), b"existing")

                replaced = video.ppt_to_long_image(source, target, overwrite=True)[0]
            self.assertEqual(replaced, target)
            with Image.open(target) as converted:
                self.assertEqual(converted.size, (32, 18))

    def test_hardware_quality_arguments(self) -> None:
        self.assertIn("-cq", video._quality_args("h264_nvenc", 20))
        self.assertIn("-global_quality", video._quality_args("h264_qsv", 20))
        self.assertIn("-qp_i", video._quality_args("h264_amf", 20))
        self.assertIn("-crf", video._quality_args("libx264", 20))

    def test_compress_reports_unavailable_h265_and_av1_encoders(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "source.mp4"
            source.write_bytes(b"not-empty")
            status = video.VideoEngineStatus(
                True,
                Path("ffmpeg"),
                ("libx264",),
                "mock",
                audio_encoders=("aac",),
            )
            with patch.object(video, "detect_video_engine", return_value=status):
                with self.assertRaises(MissingEngineError):
                    video.video_compress(source, root / "h265.mp4", codec="h265")
                with self.assertRaises(MissingEngineError):
                    video.video_compress(source, root / "av1.mkv", codec="av1")

    def test_validation_and_resource_limits(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            image = root / "frame.png"
            Image.new("RGB", (50, 50), "white").save(image)
            with self.assertRaises(ValidationError):
                video.images_to_video([image], root / "bad.avi")
            with self.assertRaises(ValidationError):
                video.images_to_video([image], root / "bad.mp4", fps=0)
            with self.assertRaises(ValidationError):
                video.images_to_video([image], root / "bad.mp4", fps=5.5)
            with self.assertRaises(ValidationError):
                video.images_to_video(
                    [image], root / "bad.mp4", slide_duration=float("nan")
                )
            with self.assertRaises(ValidationError):
                video.images_to_video(
                    [image], root / "bad.mp4", resolution=(7680, 4320)
                )
            with self.assertRaises(ValidationError):
                video.images_to_video(
                    [image] * 2001,
                    root / "too-many.mp4",
                    resolution=(160, 120),
                )
            with self.assertRaises(ValidationError):
                video._time_value("00:61", "开始时间")


if __name__ == "__main__":
    unittest.main()
