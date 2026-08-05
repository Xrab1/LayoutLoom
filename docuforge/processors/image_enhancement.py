"""High-fidelity local image enhancement with a conservative AI fallback.

The public helpers in this module operate on OpenCV BGR arrays so the same
pipeline can be shared by ordinary image jobs and the video-to-PPT extractor.
Real-ESRGAN NCNN Vulkan is optional at runtime: traditional preprocessing and
multi-frame reconstruction remain usable when the portable executable is not
present, while every AI result is structurally audited before it is accepted.
"""

from __future__ import annotations

import math
import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile
import threading
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Sequence

try:  # Keep the application catalog importable without the optional CV stack.
    import cv2  # type: ignore[import-not-found]
    import numpy as np  # type: ignore[import-not-found]
except Exception:  # pragma: no cover - capability probes cover this case
    cv2 = None
    np = None

from ..models import MissingEngineError, ValidationError
from ..runner import cancellation_callback, check_cancelled


_REAL_ESRGAN_EXECUTABLES = (
    "realesrgan-ncnn-vulkan.exe",
    "realesrgan-ncnn-vulkan",
)
_REAL_ESRGAN_MODEL = "realesrgan-x4plus"
_AI_LOCK = threading.Lock()


def _silence_native_worker_output() -> None:
    """Keep NCNN device diagnostics out of the GUI/CLI console."""

    try:
        null_fd = os.open(os.devnull, os.O_WRONLY)
        os.dup2(null_fd, 1)
        os.dup2(null_fd, 2)
        os.close(null_fd)
    except OSError:
        pass


def _binding_gpu_probe_worker(connection: object) -> None:
    """Probe Vulkan in a disposable process.

    Some releases of ``realesrgan_ncnn_py`` raise a native access violation
    while Python unloads the Vulkan runtime.  ``os._exit`` deliberately skips
    that unsafe native destructor path after the result has been sent.
    """

    _silence_native_worker_output()
    try:
        from realesrgan_ncnn_py import realesrgan_ncnn_vulkan_wrapper as wrapped

        connection.send(("ok", max(0, int(wrapped.get_gpu_count()))))  # type: ignore[attr-defined]
    except BaseException as exc:
        try:
            connection.send(("error", f"{type(exc).__name__}: {exc}"))  # type: ignore[attr-defined]
        except BaseException:
            pass
    finally:
        try:
            connection.close()  # type: ignore[attr-defined]
        except BaseException:
            pass
        os._exit(0)


def _binding_enhancement_worker(
    connection: object,
    input_path: str,
    output_path: str,
    scale: int,
    tile_size: int,
) -> None:
    """Run the native NCNN binding outside the long-lived GUI process."""

    _silence_native_worker_output()
    try:
        import cv2 as worker_cv2  # type: ignore[import-not-found]
        import numpy as worker_np  # type: ignore[import-not-found]
        from realesrgan_ncnn_py import Realesrgan
        from realesrgan_ncnn_py import realesrgan_ncnn_vulkan_wrapper as wrapped

        payload = worker_np.fromfile(input_path, dtype=worker_np.uint8)
        source = worker_cv2.imdecode(payload, worker_cv2.IMREAD_COLOR)
        if source is None or source.size == 0:
            raise RuntimeError("无法读取送入 Real-ESRGAN 的图片")
        explicit_gpu = (
            os.environ.get("LAYOUTLOOM_REALESRGAN_GPU", "").strip()
            or os.environ.get("DOCUFORGE_REALESRGAN_GPU", "").strip()
        )
        if explicit_gpu:
            try:
                gpu_id = max(0, int(explicit_gpu))
            except ValueError:
                gpu_id = 0
        else:
            gpu_count = max(1, int(wrapped.get_gpu_count()))
            gpu_id = 1 if gpu_count > 1 and shutil.which("nvidia-smi") else 0
        model_number = 0 if int(scale) == 2 else 4
        model = Realesrgan(
            gpuid=gpu_id,
            tta_mode=False,
            tilesize=max(32, int(tile_size)),
            model=model_number,
        )
        enhanced = model.process_cv2(source)
        if not isinstance(enhanced, worker_np.ndarray) or not enhanced.size:
            raise RuntimeError("Real-ESRGAN Vulkan 绑定返回空结果")
        encoded, output_payload = worker_cv2.imencode(
            ".png", enhanced, [worker_cv2.IMWRITE_PNG_COMPRESSION, 1]
        )
        if not encoded:
            raise RuntimeError("Real-ESRGAN 输出图片编码失败")
        Path(output_path).write_bytes(output_payload.tobytes())
        connection.send(("ok", ""))  # type: ignore[attr-defined]
    except BaseException as exc:
        try:
            connection.send(("error", f"{type(exc).__name__}: {exc}"))  # type: ignore[attr-defined]
        except BaseException:
            pass
    finally:
        try:
            connection.close()  # type: ignore[attr-defined]
        except BaseException:
            pass
        os._exit(0)


@dataclass(frozen=True)
class EnhancementResult:
    image: "np.ndarray"
    engine: str
    requested_mode: str
    scale: float
    ai_attempted: bool = False
    ai_accepted: bool = False
    fallback_blocks: int = 0
    total_blocks: int = 0
    reason: str = ""
    metrics: dict[str, float] = field(default_factory=dict)


@dataclass(frozen=True)
class FusionResult:
    image: "np.ndarray"
    input_frames: int
    registered_frames: int
    fused_pixels: int
    rejected_frames: int


def _require_cv() -> None:
    if cv2 is None or np is None:
        raise MissingEngineError("高清图像增强需要 OpenCV 与 NumPy")


def _candidate_engine_roots() -> list[Path]:
    roots: list[Path] = []
    explicit = (
        os.environ.get("LAYOUTLOOM_REALESRGAN_PATH", "").strip()
        or os.environ.get("DOCUFORGE_REALESRGAN_PATH", "").strip()
    )
    if explicit:
        path = Path(explicit).expanduser()
        roots.append(path.parent if path.is_file() else path)

    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        roots.extend(
            [
                Path(bundle_root) / "realesrgan-ncnn-vulkan",
                Path(bundle_root) / "realesrgan",
            ]
        )

    executable_root = Path(sys.executable).resolve().parent
    roots.extend(
        [
            executable_root / "realesrgan-ncnn-vulkan",
            executable_root / "realesrgan",
            executable_root / "_internal" / "realesrgan-ncnn-vulkan",
            executable_root / "_internal" / "realesrgan",
        ]
    )

    project_root = Path(__file__).resolve().parents[2]
    roots.extend(
        [
            project_root / "third_party" / "realesrgan-ncnn-vulkan",
            project_root / "standalone" / "realesrgan-ncnn-vulkan",
        ]
    )
    return roots


def _models_present(root: Path) -> bool:
    model_root = root / "models"
    return (model_root / f"{_REAL_ESRGAN_MODEL}.param").is_file() and (
        model_root / f"{_REAL_ESRGAN_MODEL}.bin"
    ).is_file()


@lru_cache(maxsize=1)
def realesrgan_executable() -> Path | None:
    """Locate a complete portable Real-ESRGAN NCNN Vulkan installation."""

    explicit = (
        os.environ.get("LAYOUTLOOM_REALESRGAN_PATH", "").strip()
        or os.environ.get("DOCUFORGE_REALESRGAN_PATH", "").strip()
    )
    if explicit:
        candidate = Path(explicit).expanduser()
        if candidate.is_file() and _models_present(candidate.parent):
            return candidate.resolve()

    for root in _candidate_engine_roots():
        for name in _REAL_ESRGAN_EXECUTABLES:
            candidate = root / name
            if candidate.is_file() and _models_present(root):
                return candidate.resolve()

    for name in _REAL_ESRGAN_EXECUTABLES:
        located = shutil.which(name)
        if located:
            candidate = Path(located).resolve()
            if _models_present(candidate.parent):
                return candidate
    return None


@lru_cache(maxsize=1)
def realesrgan_binding_available() -> bool:
    return importlib.util.find_spec("realesrgan_ncnn_py") is not None


def realesrgan_available() -> bool:
    return realesrgan_binding_available() or realesrgan_executable() is not None


@lru_cache(maxsize=1)
def realesrgan_gpu_available() -> bool:
    if realesrgan_binding_available():
        import multiprocessing

        context = multiprocessing.get_context("spawn")
        parent_connection, child_connection = context.Pipe(duplex=False)
        process = context.Process(
            target=_binding_gpu_probe_worker,
            args=(child_connection,),
            name="docuforge-realesrgan-gpu-probe",
            daemon=False,
        )
        try:
            process.start()
            child_connection.close()
            if not parent_connection.poll(15.0):
                return False
            status, value = parent_connection.recv()
            return status == "ok" and int(value) > 0
        except (EOFError, OSError, RuntimeError, ValueError):
            return False
        finally:
            parent_connection.close()
            process.join(1.5)
            if process.is_alive():
                process.terminate()
                process.join(1.5)
    # The portable executable performs its own Vulkan-device probe. Its
    # presence is enough for automatic mode; launch failures still fall back.
    return realesrgan_executable() is not None


@lru_cache(maxsize=1)
def _preferred_gpu_id() -> int:
    explicit = (
        os.environ.get("LAYOUTLOOM_REALESRGAN_GPU", "").strip()
        or os.environ.get("DOCUFORGE_REALESRGAN_GPU", "").strip()
    )
    if explicit:
        try:
            return max(0, int(explicit))
        except ValueError:
            pass
    if not realesrgan_binding_available():
        return 0
    try:
        from realesrgan_ncnn_py import realesrgan_ncnn_vulkan_wrapper as wrapped

        count = max(1, int(wrapped.get_gpu_count()))
    except Exception:
        return 0
    # Hybrid Windows laptops usually enumerate the integrated adapter first
    # and the discrete NVIDIA adapter second. Prefer that Vulkan device when
    # nvidia-smi confirms a discrete GPU; single-GPU systems remain on index 0.
    if count > 1 and shutil.which("nvidia-smi"):
        return 1
    return 0


def _validate_image(image: "np.ndarray", label: str) -> "np.ndarray":
    _require_cv()
    if not isinstance(image, np.ndarray) or image.ndim != 3 or image.shape[2] != 3:
        raise ValidationError(f"{label}必须是三通道 BGR 图片")
    if image.size == 0 or image.shape[0] < 2 or image.shape[1] < 2:
        raise ValidationError(f"{label}尺寸无效")
    if image.dtype != np.uint8:
        image = np.clip(image, 0, 255).astype(np.uint8)
    return np.ascontiguousarray(image)


def _gray_correlation(first: "np.ndarray", second: "np.ndarray") -> float:
    first_values = first.astype(np.float32).ravel()
    second_values = second.astype(np.float32).ravel()
    first_values -= float(first_values.mean())
    second_values -= float(second_values.mean())
    denominator = math.sqrt(
        float(np.dot(first_values, first_values))
        * float(np.dot(second_values, second_values))
    )
    if denominator <= 1e-6:
        return 1.0 if float(np.mean(np.abs(first_values - second_values))) <= 1.0 else 0.0
    return float(np.dot(first_values, second_values) / denominator)


def high_fidelity_preprocess(image: "np.ndarray") -> "np.ndarray":
    """Apply conservative denoising, local contrast and halo-limited sharpening."""

    source = _validate_image(image, "输入图片")
    gray = cv2.cvtColor(source, cv2.COLOR_BGR2GRAY)
    smooth_gray = cv2.GaussianBlur(gray, (0, 0), 1.05)
    residual = cv2.absdiff(gray, smooth_gray)
    laplacian = cv2.Laplacian(gray, cv2.CV_32F)
    flat = np.abs(laplacian) < 7.0
    flat_noise = float(np.median(residual[flat])) if np.any(flat) else 0.0

    working = source.copy()
    if flat_noise >= 2.0:
        denoised = cv2.bilateralFilter(
            source,
            5,
            sigmaColor=min(30.0, 10.0 + flat_noise * 2.4),
            sigmaSpace=4.0,
        )
        flat_mask = cv2.GaussianBlur(flat.astype(np.float32), (0, 0), 1.2)
        strength = min(0.58, 0.18 + flat_noise * 0.055)
        weight = (flat_mask * strength)[:, :, None]
        working = np.clip(
            working.astype(np.float32) * (1.0 - weight)
            + denoised.astype(np.float32) * weight,
            0,
            255,
        ).astype(np.uint8)

    lab = cv2.cvtColor(working, cv2.COLOR_BGR2LAB)
    lightness, channel_a, channel_b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=1.35, tileGridSize=(8, 8))
    local_lightness = clahe.apply(lightness)
    lightness = cv2.addWeighted(lightness, 0.86, local_lightness, 0.14, 0)
    working = cv2.cvtColor(
        cv2.merge((lightness, channel_a, channel_b)), cv2.COLOR_LAB2BGR
    )

    current_gray = cv2.cvtColor(working, cv2.COLOR_BGR2GRAY)
    sharpness = float(cv2.Laplacian(current_gray, cv2.CV_32F).var())
    amount = 0.28 if sharpness < 100.0 else 0.20 if sharpness < 260.0 else 0.13
    blurred = cv2.GaussianBlur(working, (0, 0), 0.78)
    sharpened = cv2.addWeighted(working, 1.0 + amount, blurred, -amount, 0)

    # Prevent unsharp masking from creating bright/dark halos beyond the local
    # colour envelope already present in the authentic source pixels.
    local_min = cv2.erode(working, np.ones((3, 3), np.uint8)).astype(np.int16) - 3
    local_max = cv2.dilate(working, np.ones((3, 3), np.uint8)).astype(np.int16) + 3
    limited = np.minimum(
        np.maximum(sharpened.astype(np.int16), local_min), local_max
    )
    return np.clip(limited, 0, 255).astype(np.uint8)


def _registration_score(reference: "np.ndarray", candidate: "np.ndarray") -> float:
    difference = cv2.absdiff(reference, candidate)
    return float(np.mean(difference))


def _register_translation(
    reference: "np.ndarray", candidate: "np.ndarray"
) -> tuple["np.ndarray", bool]:
    """Conservatively align a frame and reject transforms that do not help."""

    if candidate.shape != reference.shape:
        return candidate, False
    height, width = reference.shape[:2]
    longest = max(height, width)
    scale = min(1.0, 720.0 / max(1, longest))
    analysis_size = (max(48, round(width * scale)), max(48, round(height * scale)))
    reference_gray = cv2.cvtColor(reference, cv2.COLOR_BGR2GRAY)
    candidate_gray = cv2.cvtColor(candidate, cv2.COLOR_BGR2GRAY)
    if scale < 1.0:
        reference_small = cv2.resize(reference_gray, analysis_size, interpolation=cv2.INTER_AREA)
        candidate_small = cv2.resize(candidate_gray, analysis_size, interpolation=cv2.INTER_AREA)
    else:
        reference_small = reference_gray
        candidate_small = candidate_gray
    reference_float = reference_small.astype(np.float32) / 255.0
    candidate_float = candidate_small.astype(np.float32) / 255.0
    warp = np.eye(2, 3, dtype=np.float32)
    try:
        _coefficient, warp = cv2.findTransformECC(
            reference_float,
            candidate_float,
            warp,
            cv2.MOTION_TRANSLATION,
            (cv2.TERM_CRITERIA_EPS | cv2.TERM_CRITERIA_COUNT, 45, 1e-5),
            None,
            3,
        )
    except cv2.error:
        return candidate, False
    dx = float(warp[0, 2]) / scale
    dy = float(warp[1, 2]) / scale
    if abs(dx) > width * 0.012 or abs(dy) > height * 0.012:
        return candidate, False
    full_warp = np.asarray([[1.0, 0.0, dx], [0.0, 1.0, dy]], dtype=np.float32)
    aligned = cv2.warpAffine(
        candidate,
        full_warp,
        (width, height),
        flags=cv2.INTER_LINEAR | cv2.WARP_INVERSE_MAP,
        borderMode=cv2.BORDER_REFLECT101,
    )
    before = _registration_score(reference_gray, candidate_gray)
    after = _registration_score(reference_gray, cv2.cvtColor(aligned, cv2.COLOR_BGR2GRAY))
    if after > before * 0.985 + 0.05:
        return candidate, False
    return aligned, True


def multiframe_fuse(
    reference: "np.ndarray",
    frames: Sequence["np.ndarray"],
    *,
    repaired_reference: "np.ndarray | None" = None,
    excluded_boxes: Sequence[tuple[int, int, int, int]] = (),
    maximum_frames: int = 12,
) -> FusionResult:
    """Fuse stable real pixels from aligned observations without inventing text."""

    authentic = _validate_image(reference, "参考帧")
    repaired = (
        _validate_image(repaired_reference, "修复参考帧")
        if repaired_reference is not None
        else authentic
    )
    compatible = [
        _validate_image(frame, "候选帧")
        for frame in frames
        if isinstance(frame, np.ndarray) and frame.shape == authentic.shape
    ]
    if not compatible:
        return FusionResult(repaired.copy(), 0, 0, 0, 0)
    if len(compatible) > maximum_frames:
        indices = np.linspace(0, len(compatible) - 1, maximum_frames).round().astype(int)
        compatible = [compatible[int(index)] for index in indices]

    aligned: list[np.ndarray] = [authentic]
    rejected = 0
    reference_gray = cv2.cvtColor(authentic, cv2.COLOR_BGR2GRAY)
    for candidate in compatible:
        check_cancelled("已取消视频幻灯片多帧融合")
        if np.shares_memory(candidate, authentic) or np.array_equal(candidate, authentic):
            continue
        transformed, _changed = _register_translation(authentic, candidate)
        candidate_gray = cv2.cvtColor(transformed, cv2.COLOR_BGR2GRAY)
        correlation = _gray_correlation(reference_gray, candidate_gray)
        mean_difference = float(np.mean(cv2.absdiff(reference_gray, candidate_gray)))
        if correlation < 0.93 or mean_difference > 22.0:
            rejected += 1
            continue
        aligned.append(transformed)

    if len(aligned) < 2:
        return FusionResult(repaired.copy(), len(compatible), len(aligned), 0, rejected)

    stack = np.stack(aligned).astype(np.float32)
    median = np.median(stack, axis=0)
    deviation = np.mean(np.abs(stack - median[None, ...]), axis=3)
    temporal_mad = np.median(deviation, axis=0)
    stable = temporal_mad <= 5.5

    repair_delta = np.max(
        np.abs(repaired.astype(np.int16) - authentic.astype(np.int16)), axis=2
    )
    protected = repair_delta > 3
    if np.any(protected):
        protected = cv2.dilate(protected.astype(np.uint8), np.ones((5, 5), np.uint8)) > 0
    for left, top, right, bottom in excluded_boxes:
        protected[
            max(0, top) : min(protected.shape[0], bottom),
            max(0, left) : min(protected.shape[1], right),
        ] = True
    stable &= ~protected

    weights = 1.0 / (1.0 + np.square(deviation / 4.0))
    weights *= stable[None, ...]
    total_weight = np.sum(weights, axis=0)
    weighted = np.sum(stack * weights[:, :, :, None], axis=0) / np.maximum(
        total_weight[:, :, None], 1e-6
    )

    # On text and table edges, select the sharpest agreeing real observation
    # instead of averaging strokes. Flat areas use the weighted mean to reduce
    # codec noise and blocking.
    sharpness_maps = []
    for frame in aligned:
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        sharpness_maps.append(np.abs(cv2.Laplacian(gray, cv2.CV_32F)))
    sharpness_stack = np.stack(sharpness_maps)
    agreeing = deviation <= np.maximum(7.0, temporal_mad[None, ...] * 2.5 + 2.0)
    sharpness_stack[~agreeing] = -1.0
    best_indices = np.argmax(sharpness_stack, axis=0)
    rows, columns = np.indices(best_indices.shape)
    sharpest = stack[best_indices, rows, columns]
    median_gray = cv2.cvtColor(np.clip(median, 0, 255).astype(np.uint8), cv2.COLOR_BGR2GRAY)
    edge_mask = cv2.dilate(
        cv2.Canny(median_gray, 50, 145), np.ones((3, 3), np.uint8)
    ).astype(np.float32) / 255.0
    edge_weight = cv2.GaussianBlur(edge_mask, (0, 0), 0.65)[:, :, None]
    fused = weighted * (1.0 - edge_weight) + sharpest * edge_weight

    output = repaired.copy()
    fusion_mask = stable & (total_weight >= 1.5)
    output[fusion_mask] = np.clip(fused[fusion_mask], 0, 255).astype(np.uint8)
    return FusionResult(
        output,
        len(compatible),
        len(aligned),
        int(np.count_nonzero(fusion_mask)),
        rejected,
    )


def _read_cv_image(path: Path) -> "np.ndarray | None":
    try:
        payload = np.fromfile(str(path), dtype=np.uint8)
    except OSError:
        return None
    if payload.size == 0:
        return None
    return cv2.imdecode(payload, cv2.IMREAD_COLOR)


def _write_cv_png(path: Path, image: "np.ndarray") -> None:
    encoded, payload = cv2.imencode(".png", image, [cv2.IMWRITE_PNG_COMPRESSION, 1])
    if not encoded:
        raise ValidationError("Real-ESRGAN 输入图片编码失败")
    path.write_bytes(payload.tobytes())


def _run_realesrgan(
    image: "np.ndarray", *, scale: int, tile_size: int = 256
) -> tuple["np.ndarray | None", str]:
    binding_error = ""
    if realesrgan_binding_available():
        import multiprocessing
        import time

        check_cancelled("已取消 Real-ESRGAN 高清增强")
        with tempfile.TemporaryDirectory(prefix="df-esrgan-binding-") as folder:
            root = Path(folder)
            input_path = root / "input.png"
            output_path = root / "output.png"
            _write_cv_png(input_path, image)
            context = multiprocessing.get_context("spawn")
            parent_connection, child_connection = context.Pipe(duplex=False)
            process = context.Process(
                target=_binding_enhancement_worker,
                args=(
                    child_connection,
                    str(input_path),
                    str(output_path),
                    int(scale),
                    max(32, int(tile_size)),
                ),
                name="docuforge-realesrgan-worker",
                daemon=False,
            )
            timeout = max(
                180.0,
                image.shape[0] * image.shape[1] / 1_000_000 * 150.0,
            )
            message: tuple[str, object] | None = None
            started = time.monotonic()
            try:
                with _AI_LOCK:
                    process.start()
                    child_connection.close()
                    with cancellation_callback(
                        lambda: process.terminate() if process.is_alive() else None
                    ):
                        while time.monotonic() - started < timeout:
                            check_cancelled("已取消 Real-ESRGAN 高清增强")
                            if parent_connection.poll(0.12):
                                message = parent_connection.recv()
                                break
                            if not process.is_alive():
                                break
                    if message is None and process.is_alive():
                        process.terminate()
            except (EOFError, OSError, RuntimeError) as exc:
                binding_error = f"Real-ESRGAN Vulkan 隔离进程失败：{exc}"
            finally:
                parent_connection.close()
                process.join(2.0)
                if process.is_alive():
                    process.terminate()
                    process.join(2.0)
            check_cancelled("已取消 Real-ESRGAN 高清增强")
            if message is not None and message[0] == "ok" and output_path.is_file():
                enhanced = _read_cv_image(output_path)
                if enhanced is not None and enhanced.size:
                    return np.ascontiguousarray(enhanced), ""
                binding_error = "Real-ESRGAN 隔离进程输出无法读取"
            elif message is not None:
                binding_error = f"Real-ESRGAN Vulkan 绑定失败：{message[1]}"
            elif not binding_error:
                binding_error = (
                    "Real-ESRGAN Vulkan 绑定超时"
                    if time.monotonic() - started >= timeout
                    else "Real-ESRGAN Vulkan 隔离进程意外退出"
                )

    executable = realesrgan_executable()
    if executable is None:
        return None, binding_error or "未找到完整的 Real-ESRGAN NCNN Vulkan 引擎"
    check_cancelled("已取消 Real-ESRGAN 高清增强")
    with tempfile.TemporaryDirectory(prefix="df-esrgan-") as folder:
        root = Path(folder)
        input_path = root / "input.png"
        output_path = root / "output.png"
        _write_cv_png(input_path, image)
        model_root = executable.parent / "models"
        command = [
            str(executable),
            "-i",
            str(input_path),
            "-o",
            str(output_path),
            "-n",
            _REAL_ESRGAN_MODEL,
            "-m",
            str(model_root),
            "-s",
            str(scale),
            "-t",
            str(max(32, int(tile_size))),
            "-g",
            "auto",
            "-j",
            "1:2:2",
            "-f",
            "png",
        ]
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        timeout = max(180.0, image.shape[0] * image.shape[1] / 1_000_000 * 150.0)
        with _AI_LOCK:
            process = None
            try:
                process = subprocess.Popen(
                    command,
                    cwd=executable.parent,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    creationflags=creation_flags,
                )
                import time

                started = time.monotonic()
                with cancellation_callback(
                    lambda: process.terminate() if process.poll() is None else None
                ):
                    while True:
                        try:
                            stdout, _stderr = process.communicate(timeout=0.15)
                            break
                        except subprocess.TimeoutExpired:
                            check_cancelled("已取消 Real-ESRGAN 高清增强")
                            if time.monotonic() - started >= timeout:
                                process.kill()
                                process.communicate()
                                raise subprocess.TimeoutExpired(command, timeout)
                completed = subprocess.CompletedProcess(
                    command,
                    process.returncode,
                    stdout,
                    None,
                )
            except (OSError, subprocess.TimeoutExpired) as exc:
                detail = f"Real-ESRGAN 便携引擎启动失败：{exc}"
                return None, f"{binding_error}；{detail}" if binding_error else detail
            except BaseException:
                if process is not None and process.poll() is None:
                    try:
                        process.terminate()
                        process.wait(timeout=1.0)
                    except (OSError, subprocess.TimeoutExpired):
                        try:
                            process.kill()
                            process.wait(timeout=1.0)
                        except (OSError, subprocess.TimeoutExpired):
                            pass
                raise
        check_cancelled("已取消 Real-ESRGAN 高清增强")
        if completed.returncode != 0 or not output_path.is_file():
            detail = (completed.stdout or "").strip().splitlines()
            suffix = detail[-1] if detail else f"退出码 {completed.returncode}"
            detail = f"Real-ESRGAN 便携引擎处理失败：{suffix}"
            return None, f"{binding_error}；{detail}" if binding_error else detail
        enhanced = _read_cv_image(output_path)
        if enhanced is None:
            detail = "Real-ESRGAN 便携引擎输出无法解码"
            return None, f"{binding_error}；{detail}" if binding_error else detail
        return enhanced, ""


def _edge_metrics(
    source_gray: "np.ndarray",
    candidate_gray: "np.ndarray",
    low_threshold: int = 48,
    high_threshold: int = 142,
) -> tuple[float, float]:
    source_edges = cv2.Canny(source_gray, low_threshold, high_threshold) > 0
    candidate_edges = cv2.Canny(candidate_gray, low_threshold, high_threshold) > 0
    source_count = int(np.count_nonzero(source_edges))
    candidate_count = int(np.count_nonzero(candidate_edges))
    if source_count == 0:
        return 1.0, 1.0 if candidate_count == 0 else float(candidate_count)
    candidate_dilated = cv2.dilate(candidate_edges.astype(np.uint8), np.ones((3, 3), np.uint8)) > 0
    recall = float(np.count_nonzero(source_edges & candidate_dilated)) / source_count
    return recall, candidate_count / source_count


def _audit_and_fallback(
    source: "np.ndarray",
    enhanced: "np.ndarray",
    *,
    scale: int,
    content_type: str,
    protected_mask: "np.ndarray | None" = None,
) -> tuple["np.ndarray", bool, int, int, dict[str, float], str]:
    expected_height = source.shape[0] * scale
    expected_width = source.shape[1] * scale
    if (
        abs(enhanced.shape[0] - expected_height) > 2
        or abs(enhanced.shape[1] - expected_width) > 2
    ):
        fallback = cv2.resize(source, (expected_width, expected_height), interpolation=cv2.INTER_LANCZOS4)
        return fallback, False, 1, 1, {}, "AI 输出尺寸异常，已整图回退"
    if enhanced.shape[:2] != (expected_height, expected_width):
        enhanced = cv2.resize(enhanced, (expected_width, expected_height), interpolation=cv2.INTER_LANCZOS4)

    downsampled = cv2.resize(
        enhanced, (source.shape[1], source.shape[0]), interpolation=cv2.INTER_AREA
    )
    source_gray = cv2.cvtColor(source, cv2.COLOR_BGR2GRAY)
    candidate_gray = cv2.cvtColor(downsampled, cv2.COLOR_BGR2GRAY)
    correlation = _gray_correlation(source_gray, candidate_gray)
    mae = float(np.mean(cv2.absdiff(source, downsampled)))
    edge_recall, edge_ratio = _edge_metrics(source_gray, candidate_gray)
    strong_edge_recall, strong_edge_ratio = _edge_metrics(
        source_gray,
        candidate_gray,
        90,
        220,
    )
    metrics = {
        "correlation": correlation,
        "mean_absolute_difference": mae,
        "edge_recall": edge_recall,
        "edge_ratio": edge_ratio,
        "strong_edge_recall": strong_edge_recall,
        "strong_edge_ratio": strong_edge_ratio,
    }
    ordinary_edges_ok = (
        edge_recall >= (0.84 if content_type == "photo" else 0.89)
        and 0.68 <= edge_ratio <= (2.7 if content_type == "photo" else 2.15)
    )
    # Photographed paper and phone-camera images often contain millions of
    # low-contrast sensor/paper-texture edges.  A useful restoration removes
    # those weak edges while preserving the strong strokes of text, stamps and
    # table rules.  Auditing only the weak Canny map incorrectly classified
    # legitimate denoising as structural loss.  The strong-edge path remains
    # gated by excellent whole-image similarity, so invented or erased content
    # still causes a full fallback.
    denoised_photo_edges_ok = (
        content_type == "photo"
        and correlation >= 0.985
        and mae <= 8.5
        and strong_edge_recall >= 0.86
        and 0.65 <= strong_edge_ratio <= 1.45
    )
    global_ok = (
        correlation >= 0.955
        and mae <= (14.5 if content_type == "photo" else 11.5)
        and (ordinary_edges_ok or denoised_photo_edges_ok)
    )
    bicubic = cv2.resize(
        source, (expected_width, expected_height), interpolation=cv2.INTER_LANCZOS4
    )
    if not global_ok:
        return bicubic, False, 1, 1, metrics, "AI 结构二检未通过，已整图回退"

    block = 256
    blocks_x = math.ceil(source.shape[1] / block)
    blocks_y = math.ceil(source.shape[0] / block)
    risk_mask = np.zeros(source.shape[:2], dtype=np.uint8)
    fallback_blocks = 0
    for block_y in range(blocks_y):
        for block_x in range(blocks_x):
            top = block_y * block
            bottom = min(source.shape[0], top + block)
            left = block_x * block
            right = min(source.shape[1], left + block)
            original_block = source_gray[top:bottom, left:right]
            candidate_block = candidate_gray[top:bottom, left:right]
            local_mae = float(np.mean(cv2.absdiff(original_block, candidate_block)))
            local_corr = _gray_correlation(original_block, candidate_block)
            local_recall, local_edge_ratio = _edge_metrics(original_block, candidate_block)
            local_strong_recall, local_strong_edge_ratio = _edge_metrics(
                original_block,
                candidate_block,
                90,
                220,
            )
            local_edges_ok = (
                local_recall >= (0.73 if content_type == "photo" else 0.80)
                and local_edge_ratio <= (3.2 if content_type == "photo" else 2.55)
            )
            if content_type == "photo":
                local_edges_ok = local_edges_ok or (
                    local_corr >= 0.92
                    and local_strong_recall >= 0.78
                    and 0.50 <= local_strong_edge_ratio <= 1.75
                )
            risky = (
                (local_mae > 16.0 and local_corr < 0.94)
                or not local_edges_ok
            )
            if risky:
                risk_mask[top:bottom, left:right] = 255
                fallback_blocks += 1

    total_blocks = blocks_x * blocks_y
    if fallback_blocks:
        risk_mask = cv2.GaussianBlur(risk_mask, (0, 0), max(1.0, block * 0.045))
        upscaled_mask = cv2.resize(
            risk_mask,
            (expected_width, expected_height),
            interpolation=cv2.INTER_LINEAR,
        ).astype(np.float32)[:, :, None] / 255.0
        enhanced = np.clip(
            enhanced.astype(np.float32) * (1.0 - upscaled_mask)
            + bicubic.astype(np.float32) * upscaled_mask,
            0,
            255,
        ).astype(np.uint8)

    if content_type == "document":
        # Small text, formulae and thin table lines receive a real-pixel anchor.
        # The AI result still improves anti-aliasing, but cannot freely redraw
        # high-frequency glyph structure.
        edges = cv2.Canny(source_gray, 42, 132)
        structure = cv2.dilate(edges, np.ones((3, 3), np.uint8))
        structure = cv2.GaussianBlur(structure, (0, 0), 0.75)
        structure = cv2.resize(
            structure,
            (expected_width, expected_height),
            interpolation=cv2.INTER_LINEAR,
        ).astype(np.float32)[:, :, None] / 255.0
        anchor = structure * 0.30
        enhanced = np.clip(
            enhanced.astype(np.float32) * (1.0 - anchor)
            + bicubic.astype(np.float32) * anchor,
            0,
            255,
        ).astype(np.uint8)

    if protected_mask is not None:
        mask = np.asarray(protected_mask)
        if mask.shape != source.shape[:2]:
            raise ValidationError("增强保护区域尺寸与输入图片不一致")
        mask = (mask > 0).astype(np.uint8) * 255
        if np.any(mask):
            mask = cv2.dilate(mask, np.ones((5, 5), np.uint8))
            mask = cv2.GaussianBlur(mask, (0, 0), 1.15)
            mask = cv2.resize(
                mask,
                (expected_width, expected_height),
                interpolation=cv2.INTER_LINEAR,
            ).astype(np.float32)[:, :, None] / 255.0
            enhanced = np.clip(
                enhanced.astype(np.float32) * (1.0 - mask)
                + bicubic.astype(np.float32) * mask,
                0,
                255,
            ).astype(np.uint8)
            metrics["protected_pixel_ratio"] = float(
                np.mean(np.asarray(protected_mask) > 0)
            )

    reason = (
        f"AI 二检通过，{fallback_blocks}/{total_blocks} 个高风险分块回退"
        if fallback_blocks
        else "AI 二检通过"
    )
    return enhanced, True, fallback_blocks, total_blocks, metrics, reason


def enhance_bgr(
    image: "np.ndarray",
    *,
    mode: str = "high_fidelity",
    content_type: str = "auto",
    scale: int = 2,
    max_dimension: int = 4096,
    tile_size: int = 256,
    protected_mask: "np.ndarray | None" = None,
) -> EnhancementResult:
    """Enhance one BGR image and automatically fall back on unsafe AI output."""

    source = _validate_image(image, "输入图片")
    requested_mode = str(mode or "auto").strip().lower()
    aliases = {
        "compatible": "preprocess",
        "gpu_ai": "high_fidelity",
    }
    normalized_mode = aliases.get(requested_mode, requested_mode)
    if normalized_mode == "auto":
        normalized_mode = (
            "high_fidelity" if realesrgan_gpu_available() else "preprocess"
        )
    if normalized_mode not in {"off", "preprocess", "high_fidelity"}:
        raise ValidationError("图像增强模式无效")
    normalized_content = str(content_type or "auto").strip().lower()
    if normalized_content not in {"auto", "document", "photo"}:
        raise ValidationError("图像内容类型无效")
    if isinstance(scale, bool) or int(scale) not in {2, 4}:
        raise ValidationError("AI 放大倍率必须是 2 或 4")
    scale = int(scale)
    maximum = max(1024, int(max_dimension))
    if normalized_mode == "off":
        return EnhancementResult(source.copy(), "原图", normalized_mode, 1.0, reason="未启用增强")

    preprocessed = high_fidelity_preprocess(source)
    if normalized_mode == "preprocess":
        return EnhancementResult(
            preprocessed,
            "OpenCV 高保真预处理",
            normalized_mode,
            1.0,
            reason="已完成保守去噪、局部对比度和限幅锐化",
        )

    if normalized_content == "auto":
        gray = cv2.cvtColor(source, cv2.COLOR_BGR2GRAY)
        edges = float(np.count_nonzero(cv2.Canny(gray, 50, 145))) / max(1, gray.size)
        saturation = float(np.mean(cv2.cvtColor(source, cv2.COLOR_BGR2HSV)[:, :, 1]))
        normalized_content = "document" if edges >= 0.075 and saturation <= 82.0 else "photo"

    longest_after = max(source.shape[:2]) * scale
    if longest_after > maximum:
        if max(source.shape[:2]) >= maximum:
            return EnhancementResult(
                preprocessed,
                "OpenCV 高保真预处理",
                normalized_mode,
                1.0,
                reason=f"原图已达到 {max(source.shape[:2])} 像素，跳过无意义的再次放大",
            )

    enhanced, error = _run_realesrgan(preprocessed, scale=scale, tile_size=tile_size)
    if enhanced is None:
        return EnhancementResult(
            preprocessed,
            "OpenCV 高保真预处理（AI 自动降级）",
            normalized_mode,
            1.0,
            ai_attempted=realesrgan_available(),
            reason=error,
        )

    checked, accepted, fallback_blocks, total_blocks, metrics, reason = _audit_and_fallback(
        preprocessed,
        enhanced,
        scale=scale,
        content_type=normalized_content,
        protected_mask=protected_mask,
    )
    target_longest = min(maximum, max(checked.shape[:2]))
    if max(checked.shape[:2]) > target_longest:
        resize_scale = target_longest / max(checked.shape[:2])
        checked = cv2.resize(
            checked,
            (
                max(1, round(checked.shape[1] * resize_scale)),
                max(1, round(checked.shape[0] * resize_scale)),
            ),
            interpolation=cv2.INTER_AREA,
        )
    actual_scale = checked.shape[1] / source.shape[1]
    return EnhancementResult(
        checked,
        "Real-ESRGAN NCNN Vulkan + OpenCV 二检" if accepted else "Lanczos 安全回退",
        normalized_mode,
        actual_scale,
        ai_attempted=True,
        ai_accepted=accepted,
        fallback_blocks=fallback_blocks,
        total_blocks=total_blocks,
        reason=reason,
        metrics=metrics,
    )


__all__ = [
    "EnhancementResult",
    "FusionResult",
    "enhance_bgr",
    "high_fidelity_preprocess",
    "multiframe_fuse",
    "realesrgan_available",
    "realesrgan_binding_available",
    "realesrgan_executable",
    "realesrgan_gpu_available",
]
