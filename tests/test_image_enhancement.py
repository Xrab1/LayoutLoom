from __future__ import annotations

from pathlib import Path

import cv2
import numpy as np

from docuforge.processors import image as image_processor
from docuforge.processors import image_enhancement


def _document_image(height: int = 180, width: int = 320) -> np.ndarray:
    image = np.full((height, width, 3), 244, dtype=np.uint8)
    cv2.putText(
        image,
        "LayoutLoom 123",
        (18, 76),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.78,
        (22, 22, 22),
        2,
        cv2.LINE_AA,
    )
    cv2.rectangle(image, (18, 105), (300, 155), (45, 45, 45), 2)
    cv2.line(image, (18, 130), (300, 130), (80, 80, 80), 1)
    return image


def test_high_fidelity_preprocess_is_conservative_and_shape_safe() -> None:
    source = _document_image()
    noisy = np.clip(
        source.astype(np.int16)
        + np.random.default_rng(7).normal(0, 2.5, source.shape),
        0,
        255,
    ).astype(np.uint8)

    result = image_enhancement.high_fidelity_preprocess(noisy)

    assert result.shape == source.shape
    assert result.dtype == np.uint8
    assert float(np.mean(cv2.absdiff(result, noisy))) < 8.0
    assert float(cv2.Laplacian(cv2.cvtColor(result, cv2.COLOR_BGR2GRAY), cv2.CV_32F).var()) > 10


def test_multiframe_fusion_reduces_codec_noise_without_averaging_text() -> None:
    clean = _document_image()
    rng = np.random.default_rng(12)
    frames: list[np.ndarray] = []
    for index in range(7):
        frame = np.clip(
            clean.astype(np.int16) + rng.normal(0, 4.0, clean.shape),
            0,
            255,
        ).astype(np.uint8)
        if index:
            cv2.rectangle(frame, (15 + index * 28, 8), (78 + index * 28, 25), (250, 250, 250), -1)
        frames.append(frame)

    fused = image_enhancement.multiframe_fuse(frames[0], frames)
    original_error = float(
        np.mean((frames[0].astype(np.float32) - clean.astype(np.float32)) ** 2)
    )
    fused_error = float(
        np.mean((fused.image.astype(np.float32) - clean.astype(np.float32)) ** 2)
    )

    assert fused.registered_frames >= 3
    assert fused.fused_pixels > clean.shape[0] * clean.shape[1] // 2
    assert fused_error < original_error


def test_ai_structural_check_accepts_safe_upscale(monkeypatch) -> None:
    source = _document_image(96, 160)

    def safe_engine(image: np.ndarray, *, scale: int, tile_size: int):
        return (
            cv2.resize(
                image,
                (image.shape[1] * scale, image.shape[0] * scale),
                interpolation=cv2.INTER_LANCZOS4,
            ),
            "",
        )

    monkeypatch.setattr(image_enhancement, "_run_realesrgan", safe_engine)
    result = image_enhancement.enhance_bgr(
        source,
        mode="high_fidelity",
        content_type="document",
        scale=2,
        max_dimension=1024,
    )

    assert result.ai_attempted
    assert result.ai_accepted
    assert result.image.shape[:2] == (192, 320)
    assert result.total_blocks == 1


def test_photo_audit_accepts_denoising_when_strong_real_edges_survive() -> None:
    rng = np.random.default_rng(33)
    clean = np.full((240, 360, 3), 230, dtype=np.uint8)
    cv2.putText(
        clean,
        "FORM 2026",
        (20, 100),
        cv2.FONT_HERSHEY_SIMPLEX,
        1.3,
        (20, 20, 20),
        3,
        cv2.LINE_AA,
    )
    cv2.rectangle(clean, (18, 140), (340, 210), (35, 35, 35), 2)
    noise = rng.normal(0, 10, clean.shape[:2])[:, :, None]
    photographed = np.clip(clean.astype(np.float32) + noise, 0, 255).astype(np.uint8)
    denoised = cv2.bilateralFilter(photographed, 7, 35, 5)
    enhanced = cv2.resize(denoised, (720, 480), interpolation=cv2.INTER_LANCZOS4)

    checked, accepted, _fallback, _total, metrics, _reason = (
        image_enhancement._audit_and_fallback(
            photographed,
            enhanced,
            scale=2,
            content_type="photo",
        )
    )

    assert accepted
    assert checked.shape[:2] == (480, 720)
    assert metrics["edge_recall"] < 0.5
    assert metrics["strong_edge_recall"] > 0.9


def test_auto_mode_without_vulkan_gpu_uses_compatible_preprocess(monkeypatch) -> None:
    source = _document_image(96, 160)
    monkeypatch.setattr(image_enhancement, "realesrgan_gpu_available", lambda: False)

    def must_not_run(*_args: object, **_kwargs: object):
        raise AssertionError("automatic compatibility mode must not start the AI engine")

    monkeypatch.setattr(image_enhancement, "_run_realesrgan", must_not_run)
    result = image_enhancement.enhance_bgr(
        source,
        mode="auto",
        content_type="document",
        scale=2,
        max_dimension=1024,
    )

    assert not result.ai_attempted
    assert not result.ai_accepted
    assert result.image.shape == source.shape
    assert "OpenCV" in result.engine


def test_explicit_no_discrete_gpu_mode_never_starts_ai_engine(monkeypatch) -> None:
    source = _document_image(96, 160)

    def must_not_run(*_args: object, **_kwargs: object):
        raise AssertionError("explicit CPU compatibility mode must never start AI")

    monkeypatch.setattr(image_enhancement, "_run_realesrgan", must_not_run)
    result = image_enhancement.enhance_bgr(
        source,
        mode="compatible",
        content_type="document",
        scale=2,
        max_dimension=1024,
    )

    assert not result.ai_attempted
    assert not result.ai_accepted
    assert result.scale == 1.0
    assert result.image.shape == source.shape
    assert "OpenCV" in result.engine


def test_ai_structural_check_rejects_hallucinated_page(monkeypatch) -> None:
    source = _document_image(96, 160)

    def unsafe_engine(image: np.ndarray, *, scale: int, tile_size: int):
        result = cv2.resize(
            image,
            (image.shape[1] * scale, image.shape[0] * scale),
            interpolation=cv2.INTER_LANCZOS4,
        )
        cv2.rectangle(result, (0, 0), (result.shape[1] - 1, result.shape[0] - 1), (0, 0, 0), -1)
        cv2.putText(
            result,
            "WRONG",
            (20, result.shape[0] // 2),
            cv2.FONT_HERSHEY_SIMPLEX,
            1.2,
            (255, 255, 255),
            3,
            cv2.LINE_AA,
        )
        return result, ""

    monkeypatch.setattr(image_enhancement, "_run_realesrgan", unsafe_engine)
    result = image_enhancement.enhance_bgr(
        source,
        mode="high_fidelity",
        content_type="document",
        scale=2,
        max_dimension=1024,
    )

    assert result.ai_attempted
    assert not result.ai_accepted
    assert "整图回退" in result.reason
    assert result.image.shape[:2] == (192, 320)


def test_public_image_enhancement_writes_lossless_png_with_alpha(
    tmp_path: Path, monkeypatch
) -> None:
    source = tmp_path / "alpha.png"
    rgba = np.zeros((40, 64, 4), dtype=np.uint8)
    rgba[:, :, :3] = (40, 80, 220)
    rgba[:, :, 3] = np.linspace(0, 255, 64, dtype=np.uint8)[None, :]
    cv2.imencode(".png", rgba)[1].tofile(str(source))

    def safe_result(image: np.ndarray, **_kwargs: object):
        doubled = cv2.resize(
            image,
            (image.shape[1] * 2, image.shape[0] * 2),
            interpolation=cv2.INTER_LANCZOS4,
        )
        return image_enhancement.EnhancementResult(
            doubled,
            "test",
            "high_fidelity",
            2.0,
            ai_attempted=True,
            ai_accepted=True,
            reason="AI 二检通过",
        )

    monkeypatch.setattr(image_enhancement, "enhance_bgr", safe_result)
    # image.py imports the callable inside the function, so the module patch is
    # observed without changing the public API.
    outputs = image_processor.enhance_images(
        [source], tmp_path / "out", output_format="png"
    )

    assert len(outputs) == 1
    decoded = cv2.imdecode(np.fromfile(str(outputs[0]), dtype=np.uint8), cv2.IMREAD_UNCHANGED)
    assert decoded is not None
    assert decoded.shape == (80, 128, 4)
    assert decoded[:, :, 3].min() == 0
    assert decoded[:, :, 3].max() == 255
