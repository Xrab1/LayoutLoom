from __future__ import annotations

import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
import time
import warnings
from concurrent.futures import Future, ThreadPoolExecutor
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Iterable, Sequence

from PIL import Image, ImageColor, ImageOps

from ..models import MissingEngineError, ValidationError
from ..runner import cancellation_callback, check_cancelled
from ..utils import atomic_output, optimal_worker_count, unique_path

PathLike = str | Path

_VIDEO_OUTPUT_SUFFIXES = {".mp4", ".mkv", ".mov", ".webm"}
_AUDIO_OUTPUT_SUFFIXES = {".mp3", ".aac", ".wav"}
_MAX_INPUT_FRAMES = 2000
_MAX_SOURCE_PIXELS = 100_000_000
_MAX_OUTPUT_PIXELS = 3840 * 2160
_MAX_STAGED_PIXEL_WORK = 8_000_000_000
_MAX_TOTAL_DURATION = 24 * 60 * 60
_MAX_TIMEOUT = 24 * 60 * 60
_MAX_STAGED_AUDIO_BYTES = 4 * 1024 * 1024 * 1024

_KNOWN_VIDEO_ENCODERS = (
    "libx264",
    "h264_nvenc",
    "h264_qsv",
    "h264_amf",
    "libx265",
    "hevc_nvenc",
    "hevc_qsv",
    "hevc_amf",
    "libsvtav1",
    "libaom-av1",
    "av1_nvenc",
    "av1_qsv",
    "av1_amf",
    "libvpx-vp9",
    "libvpx",
    "mpeg4",
)
_KNOWN_AUDIO_ENCODERS = (
    "aac",
    "libfdk_aac",
    "libvo_aacenc",
    "libmp3lame",
    "libshine",
    "libopus",
    "libvorbis",
    "pcm_s16le",
)
_ENCODER_FAMILIES = {
    "libx264": "h264",
    "h264_nvenc": "h264",
    "h264_qsv": "h264",
    "h264_amf": "h264",
    "libx265": "h265",
    "hevc_nvenc": "h265",
    "hevc_qsv": "h265",
    "hevc_amf": "h265",
    "libsvtav1": "av1",
    "libaom-av1": "av1",
    "av1_nvenc": "av1",
    "av1_qsv": "av1",
    "av1_amf": "av1",
    "libvpx-vp9": "vp9",
    "libvpx": "vp8",
    "mpeg4": "mpeg4",
}


@dataclass(frozen=True)
class VideoEngineStatus:
    available: bool
    executable: Path | None
    encoders: tuple[str, ...]
    reason: str
    ffprobe_executable: Path | None = None
    audio_encoders: tuple[str, ...] = ()


def _creation_flags() -> int:
    return subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0


def _parse_encoders(output: str) -> tuple[set[str], set[str]]:
    video: set[str] = set()
    audio: set[str] = set()
    for line in output.splitlines():
        parts = line.split()
        if len(parts) < 2 or len(parts[0]) < 2:
            continue
        flags, name = parts[0], parts[1]
        if flags[0] == "V":
            video.add(name)
        elif flags[0] == "A":
            audio.add(name)
    return video, audio


def _path_executable_candidates(name: str) -> list[Path]:
    executable_name = f"{name}.exe" if os.name == "nt" else name
    candidates: list[Path] = []
    located = shutil.which(name)
    if located:
        candidates.append(Path(located))
    for raw_directory in os.environ.get("PATH", "").split(os.pathsep):
        directory = raw_directory.strip().strip('"')
        if directory:
            candidates.append(Path(directory) / executable_name)
    unique: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        try:
            resolved = candidate.expanduser().resolve()
        except OSError:
            continue
        key = os.path.normcase(str(resolved))
        if key in seen or not resolved.is_file():
            continue
        seen.add(key)
        unique.append(resolved)
    return unique


def _ffmpeg_candidates() -> list[Path]:
    candidates: list[Path] = []
    explicit = (
        os.environ.get("LAYOUTLOOM_FFMPEG_PATH", "").strip()
        or os.environ.get("DOCUFORGE_FFMPEG_PATH", "").strip()
    )
    if explicit:
        configured = Path(explicit).expanduser()
        candidates.append(
            configured / ("ffmpeg.exe" if os.name == "nt" else "ffmpeg")
            if configured.is_dir()
            else configured
        )
    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        base = Path(bundle_root)
        candidates.extend(
            (
                base / "ffmpeg" / "bin" / "ffmpeg.exe",
                base / "ffmpeg" / "ffmpeg.exe",
                base / "ffmpeg.exe",
            )
        )
    executable_root = Path(sys.executable).resolve().parent
    candidates.extend(
        (
            executable_root / "ffmpeg" / "bin" / "ffmpeg.exe",
            executable_root / "ffmpeg" / "ffmpeg.exe",
            executable_root / "_internal" / "ffmpeg" / "bin" / "ffmpeg.exe",
            executable_root / "_internal" / "ffmpeg" / "ffmpeg.exe",
        )
    )
    project_root = Path(__file__).resolve().parents[2]
    candidates.extend(
        (
            project_root / "third_party" / "ffmpeg" / "bin" / "ffmpeg.exe",
            project_root / "third_party" / "ffmpeg" / "ffmpeg.exe",
        )
    )
    candidates.extend(_path_executable_candidates("ffmpeg"))
    unique: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            continue
        key = os.path.normcase(str(resolved))
        if key in seen or not resolved.is_file():
            continue
        seen.add(key)
        unique.append(resolved)
    return unique


def _find_ffprobe(ffmpeg: Path) -> Path | None:
    # Prefer the ffprobe shipped beside the selected ffmpeg.  Taking the first
    # PATH entry can accidentally pair a modern ffmpeg with an incompatible
    # decade-old ffprobe from another installation.
    candidates = [ffmpeg.with_name("ffprobe.exe"), ffmpeg.with_name("ffprobe")]
    candidates.extend(_path_executable_candidates("ffprobe"))
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            continue
        if not resolved.is_file():
            continue
        try:
            result = subprocess.run(
                [str(resolved), "-version"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=8,
                check=False,
                creationflags=_creation_flags(),
            )
        except (OSError, subprocess.TimeoutExpired):
            continue
        if result.returncode == 0:
            return resolved
    return None


@lru_cache(maxsize=1)
def detect_video_engine() -> VideoEngineStatus:
    candidates = _ffmpeg_candidates()
    if not candidates:
        return VideoEngineStatus(False, None, (), "未检测到 FFmpeg")
    usable: list[
        tuple[Path, tuple[str, ...], tuple[str, ...], Path | None]
    ] = []
    failures: list[str] = []
    for executable in candidates:
        try:
            version = subprocess.run(
                [str(executable), "-hide_banner", "-version"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=8,
                check=False,
                creationflags=_creation_flags(),
            )
            if version.returncode != 0:
                failures.append(f"{executable}：版本过旧或不兼容")
                continue
            result = subprocess.run(
                [str(executable), "-encoders"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=15,
                check=False,
                creationflags=_creation_flags(),
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            failures.append(f"{executable}：{exc}")
            continue
        if result.returncode != 0:
            diagnostic = (result.stderr or result.stdout or "未知错误").strip()
            failures.append(f"{executable}：{diagnostic[-240:]}")
            continue
        output = f"{result.stdout}\n{result.stderr}"
        detected_video, detected_audio = _parse_encoders(output)
        encoders = tuple(
            name for name in _KNOWN_VIDEO_ENCODERS if name in detected_video
        )
        audio_encoders = tuple(
            name for name in _KNOWN_AUDIO_ENCODERS if name in detected_audio
        )
        usable.append(
            (executable, encoders, audio_encoders, _find_ffprobe(executable))
        )
        if encoders:
            break
    if not usable:
        detail = "；".join(failures[:3]) or "没有候选程序通过兼容性检查"
        return VideoEngineStatus(False, None, (), f"未检测到兼容的 FFmpeg；{detail}")
    selected = next((item for item in usable if item[1]), usable[0])
    executable, encoders, audio_encoders, ffprobe = selected
    if encoders:
        reason = f"已检测到 FFmpeg：{executable}"
    else:
        reason = f"已检测到 FFmpeg，但未提供可用的视频编码器：{executable}"
        if audio_encoders:
            reason += f"；音频编码器：{'、'.join(audio_encoders)}"
    if ffprobe is None:
        reason += "；未检测到 ffprobe，将仅做基础输出校验"
    else:
        reason += f"；ffprobe：{ffprobe}"
    return VideoEngineStatus(
        bool(encoders),
        executable,
        encoders,
        reason,
        ffprobe_executable=ffprobe,
        audio_encoders=audio_encoders,
    )


def _finite_number(
    value: object,
    name: str,
    *,
    minimum: float | None = None,
    maximum: float | None = None,
) -> float:
    if isinstance(value, bool):
        raise ValidationError(f"{name} 必须是数字")
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValidationError(f"{name} 必须是数字") from exc
    if not math.isfinite(number):
        raise ValidationError(f"{name} 必须是有限数字")
    if minimum is not None and number < minimum:
        raise ValidationError(f"{name} 不能小于 {minimum:g}")
    if maximum is not None and number > maximum:
        raise ValidationError(f"{name} 不能大于 {maximum:g}")
    return number


def _integer(
    value: object,
    name: str,
    *,
    minimum: int | None = None,
    maximum: int | None = None,
) -> int:
    number = _finite_number(value, name)
    integer = int(number)
    if number != integer:
        raise ValidationError(f"{name} 必须是整数")
    if minimum is not None and integer < minimum:
        raise ValidationError(f"{name} 不能小于 {minimum}")
    if maximum is not None and integer > maximum:
        raise ValidationError(f"{name} 不能大于 {maximum}")
    return integer


def _timeout_seconds(value: object | None, default: float) -> float:
    return _finite_number(
        default if value is None else value,
        "超时时间",
        minimum=1,
        maximum=_MAX_TIMEOUT,
    )


def _stop_process(process: subprocess.Popen[str]) -> None:
    if process.poll() is not None:
        return
    try:
        process.terminate()
        process.wait(timeout=3)
    except (OSError, subprocess.TimeoutExpired):
        try:
            process.kill()
            process.wait(timeout=3)
        except (OSError, subprocess.TimeoutExpired):
            pass


def _run_process(
    command: Sequence[str],
    *,
    cwd: Path | None = None,
    timeout: float,
    label: str,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    """Run an FFmpeg-family process while remaining responsive to cancellation."""

    check_cancelled()
    try:
        process = subprocess.Popen(
            list(command),
            cwd=str(cwd) if cwd else None,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=_creation_flags(),
        )
    except OSError as exc:
        raise MissingEngineError(f"无法启动 {label}：{exc}") from exc
    deadline = time.monotonic() + timeout
    try:
        with cancellation_callback(lambda: _stop_process(process)):
            while True:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    _stop_process(process)
                    raise MissingEngineError(f"{label} 超时（{timeout:g} 秒）")
                try:
                    stdout, stderr = process.communicate(timeout=min(0.25, remaining))
                    break
                except subprocess.TimeoutExpired:
                    check_cancelled()
    except BaseException:
        _stop_process(process)
        raise
    completed = subprocess.CompletedProcess(
        list(command), process.returncode, stdout, stderr
    )
    if check and completed.returncode != 0:
        diagnostic = (completed.stderr or completed.stdout or "未知错误").strip()
        raise MissingEngineError(f"{label} 失败：{diagnostic[-1600:]}")
    return completed


def _require_engine(*, require_video: bool = True) -> VideoEngineStatus:
    status = detect_video_engine()
    if status.executable is None:
        raise MissingEngineError(status.reason)
    if require_video and not status.available:
        raise MissingEngineError(status.reason)
    if not require_video and not status.audio_encoders:
        raise MissingEngineError(status.reason)
    return status


def _probe_media(
    path: Path,
    status: VideoEngineStatus,
    *,
    require_video: bool = False,
    require_audio: bool = False,
    expected_size: tuple[int, int] | None = None,
    expected_duration: float | None = None,
    input_file: bool = False,
) -> dict[str, object] | None:
    if not path.is_file() or path.stat().st_size == 0:
        raise MissingEngineError(f"媒体输出不存在或为空：{path}")
    ffprobe = status.ffprobe_executable
    if ffprobe is None:
        return None
    result = _run_process(
        [
            str(ffprobe),
            "-v",
            "error",
            "-show_entries",
            "format=duration:stream=codec_type,width,height",
            "-of",
            "json",
            str(path),
        ],
        timeout=30,
        label="ffprobe 媒体校验",
        check=False,
    )
    if result.returncode != 0:
        diagnostic = (result.stderr or result.stdout or "无法读取媒体信息").strip()
        error = ValidationError if input_file else MissingEngineError
        raise error(f"ffprobe 无法校验 {path.name}：{diagnostic[-800:]}")
    try:
        payload = json.loads(result.stdout or "{}")
    except json.JSONDecodeError as exc:
        raise MissingEngineError(f"ffprobe 返回了无效结果：{exc}") from exc
    streams = payload.get("streams", [])
    if not isinstance(streams, list):
        streams = []
    video_streams = [
        item
        for item in streams
        if isinstance(item, dict) and item.get("codec_type") == "video"
    ]
    audio_streams = [
        item
        for item in streams
        if isinstance(item, dict) and item.get("codec_type") == "audio"
    ]
    error = ValidationError if input_file else MissingEngineError
    if require_video and not video_streams:
        raise error(f"媒体文件没有视频流：{path.name}")
    if require_audio and not audio_streams:
        raise error(f"媒体文件没有音频流：{path.name}")
    if expected_size and video_streams:
        first = video_streams[0]
        actual = (int(first.get("width", 0)), int(first.get("height", 0)))
        if actual != expected_size:
            raise MissingEngineError(
                f"视频分辨率校验失败：期望 {expected_size[0]}×{expected_size[1]}，"
                f"实际 {actual[0]}×{actual[1]}"
            )
    if expected_duration is not None:
        raw_duration = payload.get("format", {})
        if isinstance(raw_duration, dict):
            try:
                actual_duration = float(raw_duration.get("duration", "nan"))
            except (TypeError, ValueError):
                actual_duration = math.nan
            tolerance = max(0.25, expected_duration * 0.03)
            if (
                math.isfinite(actual_duration)
                and actual_duration + tolerance < expected_duration
            ):
                raise MissingEngineError(
                    f"视频时长异常：期望约 {expected_duration:g} 秒，"
                    f"实际 {actual_duration:g} 秒"
                )
    return payload


def _resolution(value: str | int | tuple[int, int]) -> tuple[int, int]:
    if isinstance(value, tuple):
        if len(value) != 2:
            raise ValidationError("分辨率必须是宽、高二元组")
        width = _integer(value[0], "分辨率宽度", minimum=160)
        height = _integer(value[1], "分辨率高度", minimum=120)
    elif isinstance(value, int) or str(value).strip().isdigit():
        height = _integer(value, "分辨率高度", minimum=120)
        width = round(height * 16 / 9)
    else:
        aliases = {
            "480p": (854, 480),
            "720p": (1280, 720),
            "1080p": (1920, 1080),
            "1440p": (2560, 1440),
            "2160p": (3840, 2160),
            "4k": (3840, 2160),
        }
        key = str(value).strip().lower()
        if key not in aliases:
            raise ValidationError(f"不支持的分辨率：{value}")
        width, height = aliases[key]
    width += width % 2
    height += height % 2
    if width > 4096 or height > 4096 or width * height > _MAX_OUTPUT_PIXELS:
        raise ValidationError("分辨率不能超过 4K（总像素不超过 3840×2160）")
    return width, height


def _encoder(value: str, status: VideoEngineStatus) -> str:
    requested = str(value).strip().lower()
    if requested in {"", "auto"}:
        for candidate in ("libx264", "h264_nvenc", "h264_qsv", "h264_amf", "mpeg4"):
            if candidate in status.encoders:
                return candidate
        raise MissingEngineError("FFmpeg 没有可用的 H.264/MPEG-4 编码器")
    if requested not in {"libx264", "h264_nvenc", "h264_qsv", "h264_amf", "mpeg4"}:
        raise ValidationError(f"幻灯片不支持视频编码器：{requested}")
    if requested not in status.encoders:
        raise MissingEngineError(
            f"当前 FFmpeg 不支持编码器 {requested}；可用：{', '.join(status.encoders)}"
        )
    return requested


def _quality_args(encoder: str, quality: int) -> list[str]:
    if encoder in {"libx264", "libx265"}:
        return ["-preset", "medium", "-crf", str(quality)]
    if encoder.endswith("_nvenc"):
        return ["-preset", "medium", "-cq", str(quality), "-b:v", "0"]
    if encoder.endswith("_qsv"):
        return ["-global_quality", str(quality)]
    if encoder.endswith("_amf"):
        return [
            "-quality",
            "balanced",
            "-rc",
            "cqp",
            "-qp_i",
            str(quality),
            "-qp_p",
            str(quality),
        ]
    if encoder in {"libsvtav1", "libaom-av1", "libvpx-vp9"}:
        return ["-crf", str(quality), "-b:v", "0"]
    if encoder == "libvpx":
        return ["-crf", str(quality), "-b:v", "1M"]
    if encoder == "mpeg4":
        return ["-q:v", str(max(1, min(31, round(quality / 3))))]
    return []


def _aac_encoder(status: VideoEngineStatus) -> str:
    for candidate in ("aac", "libfdk_aac", "libvo_aacenc"):
        if candidate in status.audio_encoders:
            return candidate
    raise MissingEngineError("FFmpeg 没有可用的 AAC 音频编码器")


def _normalize_frame(
    source: Path,
    target: Path,
    size: tuple[int, int],
    background: tuple[int, int, int],
) -> None:
    check_cancelled()
    normalized: Image.Image | None = None
    fitted: Image.Image | None = None
    canvas: Image.Image | None = None
    try:
        with Image.open(source) as opened:
            opened.seek(0)
            if opened.width * opened.height > _MAX_SOURCE_PIXELS:
                raise ValidationError(
                    f"图片像素过大（上限 {_MAX_SOURCE_PIXELS:,} 像素）：{source.name}"
                )
            transposed = ImageOps.exif_transpose(opened)
            try:
                normalized = transposed.convert("RGBA")
            finally:
                if transposed is not opened:
                    transposed.close()
            fitted = ImageOps.contain(normalized, size, Image.Resampling.LANCZOS)
            canvas = Image.new("RGB", size, background)
            left = (size[0] - fitted.width) // 2
            top = (size[1] - fitted.height) // 2
            canvas.paste(fitted, (left, top), fitted)
            canvas.save(target, "PNG", optimize=False, compress_level=1)
    except Exception as exc:
        raise ValidationError(f"无法读取视频画面图片 {source.name}：{exc}") from exc
    finally:
        if canvas is not None:
            canvas.close()
        if fitted is not None:
            fitted.close()
        if normalized is not None:
            normalized.close()


def _write_manifest(
    frame_paths: Sequence[Path],
    manifest: Path,
    *,
    slide_duration: float,
    transition: str,
    transition_duration: float,
    fps: int,
) -> list[Path]:
    entries: list[tuple[Path, float]] = []
    generated: list[Path] = []
    if transition == "none" or transition_duration == 0:
        entries = [(frame, slide_duration) for frame in frame_paths]
    else:
        transition_frames = max(1, round(transition_duration * fps))
        still_duration = max(1 / fps, slide_duration - transition_duration)
        for index, frame in enumerate(frame_paths):
            check_cancelled()
            entries.append(
                (
                    frame,
                    slide_duration if index == len(frame_paths) - 1 else still_duration,
                )
            )
            if index + 1 >= len(frame_paths):
                continue
            with Image.open(frame) as current, Image.open(
                frame_paths[index + 1]
            ) as following:
                current_rgb = current.convert("RGB")
                following_rgb = following.convert("RGB")
                try:
                    for step in range(1, transition_frames + 1):
                        check_cancelled()
                        alpha = step / (transition_frames + 1)
                        blended = Image.blend(current_rgb, following_rgb, alpha)
                        blended_path = (
                            manifest.parent / f"transition_{index:04d}_{step:04d}.png"
                        )
                        blended.save(
                            blended_path, "PNG", optimize=False, compress_level=1
                        )
                        blended.close()
                        entries.append(
                            (blended_path, transition_duration / transition_frames)
                        )
                        generated.append(blended_path)
                finally:
                    current_rgb.close()
                    following_rgb.close()
    lines = ["ffconcat version 1.0"]
    for frame, duration in entries:
        lines.append(f"file '{frame.name}'")
        lines.append(f"duration {duration:.6f}")
    lines.append(f"file '{entries[-1][0].name}'")
    manifest.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return generated


def images_to_video(
    image_paths: Iterable[PathLike],
    output_path: PathLike,
    *,
    slide_duration: float = 3.0,
    fps: int = 30,
    resolution: str | int | tuple[int, int] = "1080p",
    transition: str = "none",
    transition_duration: float = 0.5,
    background: str = "black",
    audio_path: PathLike | None = None,
    encoder: str = "auto",
    quality: int = 20,
    overwrite: bool = False,
    timeout: float | None = None,
) -> list[Path]:
    """Create an MP4 slideshow with CPU/GPU FFmpeg encoders and optional audio."""

    sources = [Path(item).expanduser().resolve() for item in image_paths]
    if not sources:
        raise ValidationError("至少需要一张图片")
    if len(sources) > _MAX_INPUT_FRAMES:
        raise ValidationError(f"单次最多处理 {_MAX_INPUT_FRAMES} 张图片")
    missing = next((path for path in sources if not path.is_file()), None)
    if missing:
        raise ValidationError(f"图片不存在：{missing}")
    slide_seconds = _finite_number(
        slide_duration, "每页时长", minimum=0.01, maximum=3600
    )
    frame_rate = _integer(fps, "帧率", minimum=1, maximum=120)
    if slide_seconds < 1 / frame_rate:
        raise ValidationError("每页时长不能短于一帧")
    transition_name = str(transition).strip().lower()
    if transition_name not in {"none", "fade"}:
        raise ValidationError("转场仅支持 none 或 fade")
    transition_seconds = _finite_number(
        transition_duration, "转场时长", minimum=0, maximum=30
    )
    if transition_name == "none":
        transition_seconds = 0
    elif transition_seconds:
        if transition_seconds >= slide_seconds:
            raise ValidationError("转场时长必须大于等于 0 且小于每页时长")
        if transition_seconds < 1 / frame_rate:
            raise ValidationError("淡入淡出时长不能短于一帧")
        if slide_seconds - transition_seconds < 1 / frame_rate:
            raise ValidationError("转场后每页至少需要保留一帧静止画面")
    try:
        background_rgb = ImageColor.getcolor(str(background), "RGB")
    except (TypeError, ValueError) as exc:
        raise ValidationError(f"无效背景色：{background}") from exc
    size = _resolution(resolution)
    expected_duration = len(sources) * slide_seconds
    if expected_duration > _MAX_TOTAL_DURATION:
        raise ValidationError("输出视频总时长不能超过 24 小时")
    transition_frames = (
        max(1, round(transition_seconds * frame_rate))
        if transition_name == "fade" and transition_seconds
        else 0
    )
    staged_frames = len(sources) + max(0, len(sources) - 1) * transition_frames
    if staged_frames * size[0] * size[1] > _MAX_STAGED_PIXEL_WORK:
        raise ValidationError("图片数量、分辨率和转场组合过大，请降低其中一项")
    quality_value = _integer(quality, "画质参数", minimum=0, maximum=40)
    status = _require_engine()
    selected_encoder = _encoder(encoder, status)
    target = unique_path(Path(output_path).expanduser().resolve(), overwrite)
    if target.suffix.lower() != ".mp4":
        raise ValidationError("视频输出必须使用 .mp4 扩展名")
    target.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="docuforge-video-") as folder_name:
        folder = Path(folder_name)
        frames = [
            folder / f"frame_{index:05d}.png" for index in range(1, len(sources) + 1)
        ]
        workers = optimal_worker_count(len(sources), cap=4)
        if workers == 1:
            for source, frame in zip(sources, frames):
                _normalize_frame(source, frame, size, background_rgb)
        else:
            futures: list[Future[None]] = []
            with ThreadPoolExecutor(
                max_workers=workers, thread_name_prefix="docuforge-video-frame"
            ) as executor:
                try:
                    futures = [
                        executor.submit(
                            _normalize_frame, source, frame, size, background_rgb
                        )
                        for source, frame in zip(sources, frames)
                    ]
                    for future in futures:
                        future.result()
                except Exception:
                    for future in futures:
                        future.cancel()
                    raise
        manifest = folder / "frames.ffconcat"
        _write_manifest(
            frames,
            manifest,
            slide_duration=slide_seconds,
            transition=transition_name,
            transition_duration=transition_seconds,
            fps=frame_rate,
        )
        command = [
            str(status.executable),
            "-y",
            "-nostdin",
            "-loglevel",
            "error",
            "-f",
            "concat",
            "-i",
            manifest.name,
        ]
        if audio_path:
            audio_source = Path(audio_path).expanduser().resolve()
            if not audio_source.is_file():
                raise ValidationError(f"背景音频不存在：{audio_source}")
            staged_audio = folder / f"audio{audio_source.suffix.lower()}"
            _stage_audio(audio_source, staged_audio)
            command.extend(["-i", staged_audio.name])
        command.extend(["-map", "0:v:0"])
        if audio_path:
            command.extend(["-map", "1:a:0"])
        command.extend(["-vf", f"fps={frame_rate}", "-c:v", selected_encoder])
        command.extend(_quality_args(selected_encoder, quality_value))
        command.extend(["-pix_fmt", "yuv420p"])
        with atomic_output(target) as temporary:
            if audio_path:
                audio_encoder = _aac_encoder(status)
                command.extend(["-c:a", audio_encoder])
                if audio_encoder == "aac":
                    command.extend(["-strict", "-2"])
                command.extend(["-ar", "48000", "-b:a", "192k"])
            else:
                command.append("-an")
            # An explicit duration prevents the concat demuxer's terminal frame
            # from extending the movie and prevents short audio from truncating it.
            command.extend(
                [
                    "-t",
                    f"{expected_duration:.6f}",
                    "-movflags",
                    "+faststart",
                    str(temporary),
                ]
            )
            render_timeout = _timeout_seconds(
                timeout, min(_MAX_TIMEOUT, max(60, expected_duration * 10 + 30))
            )
            _run_process(
                command,
                cwd=folder,
                timeout=render_timeout,
                label="FFmpeg 视频编码",
            )
            _probe_media(
                temporary,
                status,
                require_video=True,
                require_audio=bool(audio_path),
                expected_size=size,
                expected_duration=expected_duration,
            )
    if not target.is_file() or target.stat().st_size == 0:
        raise MissingEngineError("FFmpeg 未生成有效的视频文件")
    return [target]


def _input_media(path: PathLike) -> Path:
    source = Path(path).expanduser().resolve()
    if not source.is_file():
        raise ValidationError(f"媒体文件不存在：{source}")
    if source.stat().st_size == 0:
        raise ValidationError(f"媒体文件为空：{source}")
    return source


def _output_media(
    path: PathLike,
    allowed_suffixes: set[str],
    overwrite: bool,
) -> Path:
    requested = Path(path).expanduser().resolve()
    if requested.suffix.lower() not in allowed_suffixes:
        allowed = "、".join(sorted(allowed_suffixes))
        raise ValidationError(f"输出格式必须是：{allowed}")
    if requested.exists() and requested.is_dir():
        raise ValidationError(f"输出位置不能是文件夹：{requested}")
    requested.parent.mkdir(parents=True, exist_ok=True)
    return unique_path(requested, overwrite)


def _payload_has_stream(payload: dict[str, object] | None, kind: str) -> bool | None:
    if payload is None:
        return None
    streams = payload.get("streams", [])
    if not isinstance(streams, list):
        return False
    return any(
        isinstance(stream, dict) and stream.get("codec_type") == kind
        for stream in streams
    )


def _payload_duration(payload: dict[str, object] | None) -> float | None:
    if payload is None:
        return None
    format_info = payload.get("format", {})
    if not isinstance(format_info, dict):
        return None
    try:
        duration = float(format_info.get("duration", "nan"))
    except (TypeError, ValueError):
        return None
    return duration if math.isfinite(duration) and duration >= 0 else None


def _warn_discarded_streams(
    payload: dict[str, object] | None,
    operation: str,
) -> None:
    if payload is None:
        return
    streams = payload.get("streams", [])
    if not isinstance(streams, list):
        return
    counts: dict[str, int] = {}
    for stream in streams:
        if not isinstance(stream, dict):
            continue
        kind = str(stream.get("codec_type", "")).lower()
        counts[kind] = counts.get(kind, 0) + 1
    details: list[str] = []
    if counts.get("video", 0) > 1:
        details.append(f"{counts['video']} 条视频流")
    if counts.get("audio", 0) > 1:
        details.append(f"{counts['audio']} 条音轨")
    if counts.get("subtitle", 0):
        details.append(f"{counts['subtitle']} 条字幕流")
    auxiliary_count = counts.get("data", 0) + counts.get("attachment", 0)
    if auxiliary_count:
        details.append(f"{auxiliary_count} 条数据/附件流")
    if not details:
        return
    warnings.warn(
        f"{operation}输入包含{'、'.join(details)}；当前只保留第一条视频流和"
        "第一条音轨（主视频/主音轨），其余流将被丢弃。",
        UserWarning,
        stacklevel=2,
    )


def _video_encoder(
    value: object,
    status: VideoEngineStatus,
    output_suffix: str,
) -> str:
    requested = str(value).strip().lower().replace("hevc", "h265")
    auto_requested = requested in {"", "auto"}
    if requested in {"", "auto"}:
        requested = "vp9" if output_suffix == ".webm" else "h264"
    if requested == "copy":
        return requested
    if requested in _ENCODER_FAMILIES:
        selected = requested
        if selected not in status.encoders:
            raise MissingEngineError(f"当前 FFmpeg 不支持编码器 {selected}")
    else:
        aliases = {
            "avc": "h264",
            "h264": "h264",
            "h265": "h265",
            "av1": "av1",
            "vp9": "vp9",
            "vp8": "vp8",
            "mpeg4": "mpeg4",
        }
        family = aliases.get(requested)
        if family is None:
            raise ValidationError(f"不支持的视频编码格式：{value}")
        preferences = {
            "h264": ("libx264", "h264_nvenc", "h264_qsv", "h264_amf"),
            "h265": ("libx265", "hevc_nvenc", "hevc_qsv", "hevc_amf"),
            "av1": (
                "libsvtav1",
                "libaom-av1",
                "av1_nvenc",
                "av1_qsv",
                "av1_amf",
            ),
            "vp9": ("libvpx-vp9",),
            "vp8": ("libvpx",),
            "mpeg4": ("mpeg4",),
        }[family]
        selected = next(
            (candidate for candidate in preferences if candidate in status.encoders),
            "",
        )
        if not selected and auto_requested and output_suffix == ".webm":
            selected = next(
                (
                    candidate
                    for candidate in (
                        "libvpx",
                        "libsvtav1",
                        "libaom-av1",
                        "av1_nvenc",
                        "av1_qsv",
                        "av1_amf",
                    )
                    if candidate in status.encoders
                ),
                "",
            )
        if not selected:
            raise MissingEngineError(f"当前 FFmpeg 没有可用的 {family.upper()} 编码器")
    family = _ENCODER_FAMILIES[selected]
    if output_suffix == ".webm" and family not in {"vp8", "vp9", "av1"}:
        raise ValidationError("WebM 输出仅支持 VP8、VP9 或 AV1 视频")
    if output_suffix in {".mp4", ".mov"} and family in {"vp8", "vp9"}:
        raise ValidationError("MP4/MOV 输出不支持 VP8/VP9，请改用 WebM 或 MKV")
    return selected


def _audio_encoder(
    value: object,
    status: VideoEngineStatus,
    output_suffix: str,
) -> str:
    requested = str(value).strip().lower()
    if requested in {"none", "disable", "无"}:
        return "none"
    if requested == "copy":
        return "copy"
    if requested in {"", "auto"}:
        candidates = (
            ("libopus", "libvorbis")
            if output_suffix == ".webm"
            else ("aac", "libfdk_aac", "libvo_aacenc")
        )
    elif requested in {"aac", "mp3", "opus", "vorbis", "wav", "pcm"}:
        candidates = {
            "aac": ("aac", "libfdk_aac", "libvo_aacenc"),
            "mp3": ("libmp3lame", "libshine"),
            "opus": ("libopus",),
            "vorbis": ("libvorbis",),
            "wav": ("pcm_s16le",),
            "pcm": ("pcm_s16le",),
        }[requested]
    elif requested in _KNOWN_AUDIO_ENCODERS:
        candidates = (requested,)
    else:
        raise ValidationError(f"不支持的音频编码格式：{value}")
    selected = next(
        (candidate for candidate in candidates if candidate in status.audio_encoders),
        "",
    )
    if not selected:
        raise MissingEngineError(
            f"当前 FFmpeg 没有可用的音频编码器：{'、'.join(candidates)}"
        )
    if output_suffix == ".webm" and selected not in {"libopus", "libvorbis"}:
        raise ValidationError("WebM 输出音频仅支持 Opus 或 Vorbis")
    if output_suffix in {".mp4", ".mov"} and selected in {
        "libopus",
        "libvorbis",
        "pcm_s16le",
    }:
        raise ValidationError("所选音频编码与 MP4/MOV 容器不兼容")
    return selected


def _audio_codec_args(encoder: str, bitrate: int) -> list[str]:
    if encoder == "none":
        return ["-an"]
    if encoder == "copy":
        return ["-c:a", "copy"]
    args = ["-c:a", encoder]
    if encoder in {"aac", "libopus"}:
        args.extend(["-strict", "-2"])
    if encoder not in {"pcm_s16le"}:
        # Normalizing lossy output to 48 kHz avoids impossible combinations
        # such as a 192 kb/s AAC stream fed by an 8 kHz voice recording.
        args.extend(["-ar", "48000", "-b:a", f"{bitrate}k"])
    return args


def _transcode_output_args(
    status: VideoEngineStatus,
    *,
    output_suffix: str,
    video_codec: object,
    audio_codec: object,
    quality: object,
    resolution: str | int | tuple[int, int] | None,
    fps: object | None,
    audio_bitrate: object,
    has_audio: bool | None,
) -> tuple[list[str], tuple[int, int] | None, str, str]:
    selected_video = _video_encoder(video_codec, status, output_suffix)
    selected_audio = _audio_encoder(audio_codec, status, output_suffix)
    quality_value = _integer(quality, "画质参数", minimum=0, maximum=51)
    bitrate_value = _integer(audio_bitrate, "音频码率", minimum=16, maximum=1024)
    size = _resolution(resolution) if resolution is not None else None
    frame_rate = (
        _integer(fps, "帧率", minimum=1, maximum=120) if fps is not None else None
    )
    if selected_video == "copy" and (size is not None or frame_rate is not None):
        raise ValidationError("视频流复制模式不能同时修改分辨率或帧率")
    if has_audio is None:
        args = ["-sn", "-dn"]
    else:
        args = ["-map", "0:v:0"]
        if has_audio and selected_audio != "none":
            args.extend(["-map", "0:a:0"])
        args.extend(["-sn", "-dn"])
    if selected_video == "copy":
        args.extend(["-c:v", "copy"])
    else:
        filters: list[str] = []
        if size is not None:
            filters.append(
                f"scale={size[0]}:{size[1]}:force_original_aspect_ratio=decrease,"
                f"pad={size[0]}:{size[1]}:(ow-iw)/2:(oh-ih)/2"
            )
        if frame_rate is not None:
            filters.append(f"fps={frame_rate}")
        if filters:
            args.extend(["-vf", ",".join(filters)])
        args.extend(["-c:v", selected_video])
        args.extend(_quality_args(selected_video, quality_value))
        family = _ENCODER_FAMILIES[selected_video]
        if family in {"h264", "h265", "mpeg4"}:
            args.extend(["-pix_fmt", "yuv420p"])
    args.extend(_audio_codec_args(selected_audio, bitrate_value))
    if output_suffix in {".mp4", ".mov"}:
        args.extend(["-movflags", "+faststart"])
    return args, size, selected_video, selected_audio


def _default_media_timeout(duration: float | None, multiplier: float = 10) -> float:
    if duration is None:
        return 3600
    return min(_MAX_TIMEOUT, max(60, duration * multiplier + 30))


def _stage_audio(source: Path, target: Path) -> None:
    if source.stat().st_size > _MAX_STAGED_AUDIO_BYTES:
        raise ValidationError("背景音频不能超过 4 GB")
    try:
        os.link(source, target)
        return
    except OSError:
        pass
    with source.open("rb") as reader, target.open("wb") as writer:
        while True:
            check_cancelled()
            chunk = reader.read(8 * 1024 * 1024)
            if not chunk:
                break
            writer.write(chunk)
    try:
        shutil.copystat(source, target)
    except OSError:
        pass


def video_transcode(
    input_video: PathLike,
    output_path: PathLike,
    *,
    video_codec: str = "auto",
    audio_codec: str = "auto",
    quality: int = 23,
    resolution: str | int | tuple[int, int] | None = None,
    fps: int | None = None,
    audio_bitrate: int = 192,
    overwrite: bool = False,
    timeout: float | None = None,
) -> list[Path]:
    """Transcode one video to MP4, MKV, MOV or WebM with atomic publishing."""

    source = _input_media(input_video)
    target = _output_media(output_path, _VIDEO_OUTPUT_SUFFIXES, overwrite)
    status = _require_engine()
    source_info = _probe_media(source, status, require_video=True, input_file=True)
    _warn_discarded_streams(source_info, "视频转码：")
    has_audio = _payload_has_stream(source_info, "audio")
    effective_audio_codec = "none" if has_audio is False else audio_codec
    output_args, size, _, selected_audio = _transcode_output_args(
        status,
        output_suffix=target.suffix.lower(),
        video_codec=video_codec,
        audio_codec=effective_audio_codec,
        quality=quality,
        resolution=resolution,
        fps=fps,
        audio_bitrate=audio_bitrate,
        has_audio=has_audio,
    )
    duration = _payload_duration(source_info)
    if duration is not None and duration > _MAX_TOTAL_DURATION:
        raise ValidationError("单次转码的视频时长不能超过 24 小时")
    with atomic_output(target) as temporary:
        command = [
            str(status.executable),
            "-y",
            "-nostdin",
            "-loglevel",
            "error",
            "-i",
            str(source),
            *output_args,
            str(temporary),
        ]
        _run_process(
            command,
            timeout=_timeout_seconds(timeout, _default_media_timeout(duration)),
            label="FFmpeg 视频转码",
        )
        _probe_media(
            temporary,
            status,
            require_video=True,
            require_audio=has_audio is True and selected_audio != "none",
            expected_size=size,
        )
    return [target]


def video_compress(
    input_video: PathLike,
    output_path: PathLike,
    *,
    codec: str = "h264",
    encoder: str = "auto",
    quality: int = 28,
    resolution: str | int | tuple[int, int] | None = None,
    fps: int | None = None,
    audio_bitrate: int = 128,
    overwrite: bool = False,
    timeout: float | None = None,
) -> list[Path]:
    """Compress video with H.264, H.265/HEVC or AV1 when available."""

    family = str(codec).strip().lower().replace("hevc", "h265")
    if family not in {"h264", "h265", "av1"}:
        raise ValidationError("压缩编码必须是 h264、h265/hevc 或 av1")
    requested_encoder = str(encoder).strip().lower()
    if requested_encoder in {"", "auto"}:
        requested_encoder = family
    elif _ENCODER_FAMILIES.get(requested_encoder) != family:
        raise ValidationError(
            f"编码器 {requested_encoder} 不属于 {family.upper()} 系列"
        )
    return video_transcode(
        input_video,
        output_path,
        video_codec=requested_encoder,
        audio_codec="auto",
        quality=quality,
        resolution=resolution,
        fps=fps,
        audio_bitrate=audio_bitrate,
        overwrite=overwrite,
        timeout=timeout,
    )


def _time_value(value: object, name: str) -> float:
    if isinstance(value, str) and ":" in value:
        pieces = value.strip().split(":")
        if len(pieces) not in {2, 3}:
            raise ValidationError(f"{name} 时间格式应为 HH:MM:SS 或 MM:SS")
        try:
            numbers = [float(piece) for piece in pieces]
        except ValueError as exc:
            raise ValidationError(f"{name} 时间格式无效") from exc
        if any(not math.isfinite(number) or number < 0 for number in numbers):
            raise ValidationError(f"{name} 时间格式无效")
        if numbers[-1] >= 60 or numbers[-2] >= 60:
            raise ValidationError(f"{name} 的分钟和秒必须小于 60")
        if len(numbers) == 2:
            return numbers[0] * 60 + numbers[1]
        return numbers[0] * 3600 + numbers[1] * 60 + numbers[2]
    return _finite_number(value, name, minimum=0, maximum=_MAX_TOTAL_DURATION)


def video_trim(
    input_video: PathLike,
    output_path: PathLike,
    *,
    start: float | str = 0,
    end: float | str | None = None,
    duration: float | str | None = None,
    mode: str = "precise",
    video_codec: str = "auto",
    audio_codec: str = "auto",
    quality: int = 23,
    overwrite: bool = False,
    timeout: float | None = None,
) -> list[Path]:
    """Trim video by fast stream copy or frame-accurate re-encoding."""

    source = _input_media(input_video)
    target = _output_media(output_path, _VIDEO_OUTPUT_SUFFIXES, overwrite)
    start_seconds = _time_value(start, "开始时间")
    if end is not None and duration is not None:
        raise ValidationError("结束时间和持续时长只能填写一个")
    end_seconds = _time_value(end, "结束时间") if end is not None else None
    duration_seconds = (
        _time_value(duration, "持续时长") if duration is not None else None
    )
    if end_seconds is not None:
        if end_seconds <= start_seconds:
            raise ValidationError("结束时间必须晚于开始时间")
        duration_seconds = end_seconds - start_seconds
    if duration_seconds is not None and duration_seconds <= 0:
        raise ValidationError("持续时长必须大于 0")
    requested_mode = str(mode).strip().lower()
    aliases = {"fast": "copy", "快速": "copy", "reencode": "precise", "精准": "precise"}
    requested_mode = aliases.get(requested_mode, requested_mode)
    if requested_mode not in {"copy", "precise"}:
        raise ValidationError("裁剪模式必须是 copy/fast 或 precise/reencode")

    status = _require_engine()
    source_info = _probe_media(source, status, require_video=True, input_file=True)
    _warn_discarded_streams(source_info, "视频裁剪：")
    source_duration = _payload_duration(source_info)
    if source_duration is not None:
        tolerance = 0.05
        if start_seconds >= source_duration - tolerance:
            raise ValidationError("开始时间已超出视频时长")
        if (
            duration_seconds is not None
            and start_seconds + duration_seconds > source_duration + tolerance
        ):
            raise ValidationError("裁剪结束时间已超出视频时长")
    has_audio = _payload_has_stream(source_info, "audio")
    require_output_audio = has_audio is True
    with atomic_output(target) as temporary:
        base = [str(status.executable), "-y", "-nostdin", "-loglevel", "error"]
        if requested_mode == "copy":
            command = [*base, "-ss", f"{start_seconds:.6f}", "-i", str(source)]
            if duration_seconds is not None:
                command.extend(["-t", f"{duration_seconds:.6f}"])
            if has_audio is None:
                command.extend(["-sn", "-dn", "-c", "copy"])
            else:
                command.extend(["-map", "0:v:0"])
                if has_audio:
                    command.extend(["-map", "0:a:0"])
                command.extend(["-sn", "-dn", "-c", "copy"])
            if target.suffix.lower() in {".mp4", ".mov"}:
                command.extend(["-movflags", "+faststart"])
            command.append(str(temporary))
        else:
            effective_audio_codec = "none" if has_audio is False else audio_codec
            output_args, _, _, selected_audio = _transcode_output_args(
                status,
                output_suffix=target.suffix.lower(),
                video_codec=video_codec,
                audio_codec=effective_audio_codec,
                quality=quality,
                resolution=None,
                fps=None,
                audio_bitrate=192,
                has_audio=has_audio,
            )
            command = [*base, "-i", str(source), "-ss", f"{start_seconds:.6f}"]
            if duration_seconds is not None:
                command.extend(["-t", f"{duration_seconds:.6f}"])
            command.extend([*output_args, str(temporary)])
            require_output_audio = has_audio is True and selected_audio != "none"
        operation_duration = duration_seconds
        if operation_duration is None and source_duration is not None:
            operation_duration = source_duration - start_seconds
        if operation_duration is not None and operation_duration > _MAX_TOTAL_DURATION:
            raise ValidationError("单次裁剪输出时长不能超过 24 小时")
        _run_process(
            command,
            timeout=_timeout_seconds(
                timeout,
                _default_media_timeout(
                    operation_duration,
                    2 if requested_mode == "copy" else 10,
                ),
            ),
            label="FFmpeg 视频裁剪",
        )
        _probe_media(
            temporary,
            status,
            require_video=True,
            require_audio=require_output_audio,
        )
    return [target]


def video_extract_audio(
    input_video: PathLike,
    output_path: PathLike,
    *,
    audio_codec: str = "auto",
    bitrate: int = 192,
    sample_rate: int | None = None,
    channels: int | None = None,
    overwrite: bool = False,
    timeout: float | None = None,
) -> list[Path]:
    """Extract the first audio stream as MP3, AAC or WAV."""

    source = _input_media(input_video)
    target = _output_media(output_path, _AUDIO_OUTPUT_SUFFIXES, overwrite)
    status = _require_engine(require_video=False)
    source_info = _probe_media(source, status, require_audio=True, input_file=True)
    suffix = target.suffix.lower()
    default_codec = {".mp3": "mp3", ".aac": "aac", ".wav": "wav"}[suffix]
    requested_codec = str(audio_codec).strip().lower()
    if requested_codec in {"", "auto"}:
        requested_codec = default_codec
    selected = _audio_encoder(requested_codec, status, ".mkv")
    expected_families = {
        ".mp3": {"libmp3lame", "libshine"},
        ".aac": {"aac", "libfdk_aac", "libvo_aacenc"},
        ".wav": {"pcm_s16le"},
    }[suffix]
    if selected not in expected_families:
        raise ValidationError(f"音频编码器 {selected} 与 {suffix} 输出格式不匹配")
    bitrate_value = _integer(bitrate, "音频码率", minimum=16, maximum=1024)
    rate_value = (
        _integer(sample_rate, "采样率", minimum=8000, maximum=192000)
        if sample_rate is not None
        else None
    )
    channels_value = (
        _integer(channels, "声道数", minimum=1, maximum=8)
        if channels is not None
        else None
    )
    duration_value = _payload_duration(source_info)
    if duration_value is not None and duration_value > _MAX_TOTAL_DURATION:
        raise ValidationError("单次音频提取的媒体时长不能超过 24 小时")
    with atomic_output(target) as temporary:
        command = [
            str(status.executable),
            "-y",
            "-nostdin",
            "-loglevel",
            "error",
            "-i",
            str(source),
            "-map",
            "0:a:0",
            "-vn",
            *_audio_codec_args(selected, bitrate_value),
        ]
        if rate_value is not None:
            command.extend(["-ar", str(rate_value)])
        if channels_value is not None:
            command.extend(["-ac", str(channels_value)])
        command.append(str(temporary))
        _run_process(
            command,
            timeout=_timeout_seconds(
                timeout, _default_media_timeout(duration_value, multiplier=3)
            ),
            label="FFmpeg 音频提取",
        )
        _probe_media(temporary, status, require_audio=True)
    return [target]


def pdf_to_video(
    input_pdf: PathLike,
    output_path: PathLike,
    *,
    dpi: int = 160,
    password: str | None = None,
    **video_options: object,
) -> list[Path]:
    from ..engines import poppler_bin_path
    from .pdf import pdf_to_images

    dpi_value = _integer(dpi, "PDF 渲染 DPI", minimum=72, maximum=600)
    check_cancelled()
    with tempfile.TemporaryDirectory(prefix="docuforge-pdf-video-") as folder_name:
        folder = Path(folder_name)
        pages = pdf_to_images(
            input_pdf,
            folder,
            image_format="png",
            dpi=dpi_value,
            password=password,
            poppler_path=poppler_bin_path(),
        )
        check_cancelled()
        return images_to_video(pages, output_path, **video_options)


def _require_nonempty_paths(outputs: Iterable[PathLike], label: str) -> list[Path]:
    paths = [Path(path) for path in outputs]
    if not paths:
        raise MissingEngineError(f"{label} 未返回任何输出文件")
    invalid = next(
        (path for path in paths if not path.is_file() or path.stat().st_size == 0),
        None,
    )
    if invalid is not None:
        raise MissingEngineError(f"{label} 输出不存在或为空：{invalid}")
    return paths


def _microsoft_powerpoint_status():
    from .office import OfficeEngineStatus, detect_office_engines

    try:
        statuses = detect_office_engines()
        return statuses.get(
            "microsoft_powerpoint",
            OfficeEngineStatus(False, reason="未检测到 Microsoft PowerPoint COM"),
        )
    except Exception as exc:
        return OfficeEngineStatus(False, reason=f"PowerPoint COM 检测失败：{exc}")


def _require_microsoft_powerpoint() -> None:
    status = _microsoft_powerpoint_status()
    if not status.available:
        raise MissingEngineError(status.reason)


def _render_ppt_to_pdf(source: Path, output_dir: Path, renderer: str) -> Path:
    from .office import convert_with_office

    requested = renderer.strip().lower()
    if requested not in {"auto", "microsoft_office", "wps", "libreoffice"}:
        raise ValidationError(
            "PPT 渲染器必须是 auto、microsoft_office、wps 或 libreoffice"
        )
    outputs = convert_with_office(source, output_dir, "pdf", engine=requested)
    return _require_nonempty_paths(outputs, "PPT 静态渲染")[0]


def ppt_to_images(
    input_ppt: PathLike,
    output_dir: PathLike,
    *,
    renderer: str = "auto",
    image_format: str = "png",
    width: int = 1920,
    dpi: int = 160,
    overwrite: bool = False,
) -> list[Path]:
    """Render slides with WPS-first auto routing or explicit PowerPoint native export."""

    source = Path(input_ppt).expanduser().resolve()
    if not source.is_file():
        raise ValidationError(f"PPT 不存在：{source}")
    requested = renderer.strip().lower()
    if requested == "microsoft_office":
        _require_microsoft_powerpoint()
        try:
            from .office_com import ppt_to_images as native_ppt_to_images

            return _require_nonempty_paths(
                native_ppt_to_images(
                    source,
                    output_dir,
                    format=image_format,
                    width=int(width),
                    overwrite=overwrite,
                ),
                "PowerPoint 原生图片导出",
            )
        except (ImportError, MissingEngineError):
            raise
    with tempfile.TemporaryDirectory(prefix="docuforge-ppt-images-") as folder_name:
        folder = Path(folder_name)
        pdf_path = _render_ppt_to_pdf(source, folder, requested)
        from ..engines import poppler_bin_path
        from .pdf import pdf_to_images

        return _require_nonempty_paths(
            pdf_to_images(
                pdf_path,
                output_dir,
                image_format=image_format,
                dpi=int(dpi),
                prefix=source.stem,
                poppler_path=poppler_bin_path(),
                overwrite=overwrite,
            ),
            "PPT 页面图片渲染",
        )


def _save_single_slide_image(
    source: Path,
    output_path: PathLike,
    *,
    background: str,
    overwrite: bool,
) -> list[Path]:
    requested = Path(output_path).expanduser()
    if not requested.suffix:
        requested = requested.with_suffix(".png")
    extension = requested.suffix.lower()
    formats = {".png": "PNG", ".jpg": "JPEG", ".jpeg": "JPEG", ".webp": "WEBP"}
    if extension not in formats:
        raise ValidationError("单页 PPT 长图仅支持 PNG、JPG 或 WebP")
    target = unique_path(requested, overwrite=overwrite)
    image_format = formats[extension]
    with Image.open(source) as image:
        image.load()
        prepared: Image.Image
        if image_format == "JPEG":
            rgba = image.convert("RGBA")
            prepared = Image.new(
                "RGB", rgba.size, ImageColor.getcolor(str(background), "RGB")
            )
            prepared.paste(rgba, mask=rgba.getchannel("A"))
            rgba.close()
        else:
            prepared = image.copy()
        try:
            options: dict[str, object] = {}
            if image_format == "PNG":
                options.update(optimize=False, compress_level=6)
            elif image_format == "JPEG":
                options.update(quality=95, subsampling=0, optimize=True)
            else:
                options.update(quality=95, method=6)
            with atomic_output(target) as temporary:
                prepared.save(temporary, image_format, **options)
                if not temporary.is_file() or temporary.stat().st_size == 0:
                    raise MissingEngineError("PPT 单页长图未生成有效输出文件")
        finally:
            prepared.close()
    return _require_nonempty_paths([target], "PPT 单页长图")


def ppt_to_long_image(
    input_ppt: PathLike,
    output_path: PathLike,
    *,
    renderer: str = "auto",
    direction: str = "vertical",
    spacing: int = 0,
    background: str = "white",
    width: int = 1920,
    dpi: int = 160,
    overwrite: bool = False,
) -> list[Path]:
    """Render slides and stitch them into a long PNG/JPG image."""

    requested = renderer.strip().lower()
    if requested == "microsoft_office":
        _require_microsoft_powerpoint()
        try:
            from .office_com import ppt_to_long_image as native_ppt_to_long_image

            return _require_nonempty_paths(
                native_ppt_to_long_image(
                    input_ppt,
                    output_path,
                    direction=direction,
                    spacing=int(spacing),
                    background=background,
                    width=int(width),
                    overwrite=overwrite,
                ),
                "PowerPoint 原生长图导出",
            )
        except (ImportError, MissingEngineError):
            raise
    with tempfile.TemporaryDirectory(prefix="docuforge-ppt-long-") as folder_name:
        folder = Path(folder_name)
        slides = ppt_to_images(
            input_ppt,
            folder,
            renderer=requested,
            image_format="png",
            width=width,
            dpi=dpi,
        )
        if len(slides) == 1:
            return _save_single_slide_image(
                slides[0],
                output_path,
                background=background,
                overwrite=overwrite,
            )
        from .image import stitch_images

        return _require_nonempty_paths(
            stitch_images(
                slides,
                output_path,
                direction=direction,
                spacing=int(spacing),
                background=background,
                overwrite=overwrite,
            ),
            "PPT 长图拼接",
        )


def ppt_to_static_video(
    input_ppt: PathLike,
    output_path: PathLike,
    *,
    renderer: str = "auto",
    dpi: int = 160,
    **video_options: object,
) -> list[Path]:
    source = Path(input_ppt).expanduser().resolve()
    if not source.is_file():
        raise ValidationError(f"PPT 不存在：{source}")
    with tempfile.TemporaryDirectory(prefix="docuforge-ppt-video-") as folder_name:
        folder = Path(folder_name)
        pdf_path = _render_ppt_to_pdf(source, folder, renderer)
        return pdf_to_video(pdf_path, output_path, dpi=dpi, **video_options)


def _native_powerpoint_quality(quality: int) -> int:
    """Map FFmpeg-style 0-40 quality to PowerPoint's 1-100 scale."""

    return max(1, min(100, round(100 - quality * 0.75)))


def ppt_to_video(
    input_ppt: PathLike,
    output_path: PathLike,
    *,
    mode: str = "auto",
    renderer: str = "auto",
    use_timings: bool = True,
    slide_duration: float = 5.0,
    resolution: str | int | tuple[int, int] = "1080p",
    fps: int = 30,
    quality: int = 20,
    **video_options: object,
) -> list[Path]:
    """Use native PowerPoint video first, then a WPS-first static fallback."""

    requested = str(mode).strip().lower()
    if requested not in {"auto", "native", "static"}:
        raise ValidationError("视频模式必须是 auto、native 或 static")
    slide_seconds = _finite_number(
        slide_duration, "每页时长", minimum=0.01, maximum=3600
    )
    frame_rate = _integer(fps, "帧率", minimum=1, maximum=120)
    quality_value = _integer(quality, "画质参数", minimum=0, maximum=40)
    ffmpeg_only_options = bool(video_options.get("audio_path")) or str(
        video_options.get("encoder", "auto")
    ).strip().lower() not in {"", "auto"}
    native_parameter_errors: list[str] = []
    if not slide_seconds.is_integer():
        native_parameter_errors.append("每页时长必须是整数秒")
    if frame_rate > 100:
        native_parameter_errors.append("帧率不能超过 100")
    if requested == "native" and ffmpeg_only_options:
        raise ValidationError(
            "原生 PowerPoint 视频模式不支持外部背景音频或 FFmpeg 编码器"
        )
    if requested == "native" and native_parameter_errors:
        raise ValidationError(
            "原生 PowerPoint 视频模式参数无效：" + "；".join(native_parameter_errors)
        )
    if requested == "auto" and (ffmpeg_only_options or native_parameter_errors):
        requested = "static"
    if requested in {"auto", "native"}:
        native_status = _microsoft_powerpoint_status()
        if not native_status.available:
            if requested == "native":
                raise MissingEngineError(native_status.reason)
        else:
            try:
                from .office_com import ppt_to_video as native_ppt_to_video

                height = _resolution(resolution)[1]
                return _require_nonempty_paths(
                    native_ppt_to_video(
                        input_ppt,
                        output_path,
                        use_timings=use_timings,
                        slide_duration=int(slide_seconds),
                        resolution=height,
                        fps=int(frame_rate),
                        quality=_native_powerpoint_quality(quality_value),
                    ),
                    "PowerPoint 原生视频导出",
                )
            except (ImportError, MissingEngineError):
                if requested == "native":
                    raise
    return ppt_to_static_video(
        input_ppt,
        output_path,
        renderer=renderer,
        slide_duration=slide_seconds,
        resolution=resolution,
        fps=frame_rate,
        quality=quality_value,
        **video_options,
    )
