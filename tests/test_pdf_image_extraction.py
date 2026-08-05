from __future__ import annotations

import io
import json
from pathlib import Path
from typing import Any, Iterable

import pymupdf
import pytest
from PIL import Image
from reportlab.pdfgen.canvas import Canvas

from docuforge.processors import pdf_images as pdf_images_module
from docuforge.models import DocuForgeError, ValidationError
from docuforge.processors.pdf import PDFPasswordError, PDFProcessorError
from docuforge.processors.pdf_images import extract_pdf_images


def _encoded_image(
    image_format: str,
    size: tuple[int, int],
    color: tuple[int, ...],
    *,
    quality: int = 92,
) -> bytes:
    mode = "RGBA" if len(color) == 4 else "RGB"
    image = Image.new(mode, size, color)
    output = io.BytesIO()
    save_options: dict[str, Any] = {}
    if image_format.upper() == "JPEG":
        save_options.update(quality=quality, subsampling=0)
    image.save(output, image_format, **save_options)
    image.close()
    return output.getvalue()


def _transparent_png(size: tuple[int, int] = (48, 32)) -> bytes:
    image = Image.new("RGBA", size)
    width, height = size
    for y in range(height):
        for x in range(width):
            alpha = round(255 * x / max(1, width - 1))
            image.putpixel((x, y), (30, 120, 220, alpha))
    output = io.BytesIO()
    image.save(output, "PNG")
    image.close()
    return output.getvalue()


def _write_pdf(
    target: Path,
    pages: Iterable[
        Iterable[tuple[tuple[float, float, float, float], bytes, dict[str, Any]]]
    ],
    *,
    page_size: tuple[float, float] = (240, 180),
) -> Path:
    document = pymupdf.open()
    try:
        for placements in pages:
            page = document.new_page(width=page_size[0], height=page_size[1])
            for rectangle, payload, options in placements:
                page.insert_image(
                    pymupdf.Rect(rectangle),
                    stream=payload,
                    keep_proportion=False,
                    **options,
                )
        document.save(str(target))
    finally:
        document.close()
    return target


def _write_inline_image_pdf(target: Path) -> Path:
    image = Image.new("RGB", (37, 23), (10, 170, 80))
    try:
        document = Canvas(str(target), pagesize=(200, 150))
        document.drawInlineImage(image, 20, 30, width=111, height=69)
        document.save()
    finally:
        image.close()
    return target


def _encrypt_pdf(source: Path, target: Path, *, password: str) -> Path:
    with pymupdf.open(str(source)) as document:
        document.save(
            str(target),
            encryption=pymupdf.PDF_ENCRYPT_AES_256,
            owner_pw="owner-secret",
            user_pw=password,
        )
    return target


def _image_outputs(outputs: Iterable[Path]) -> list[Path]:
    return [
        Path(output) for output in outputs if Path(output).suffix.lower() != ".json"
    ]


def _json_outputs(outputs: Iterable[Path]) -> list[Path]:
    return [
        Path(output) for output in outputs if Path(output).suffix.lower() == ".json"
    ]


def _contains_json_value(value: object, expected: str) -> bool:
    if isinstance(value, dict):
        return any(
            _contains_json_value(key, expected) or _contains_json_value(item, expected)
            for key, item in value.items()
        )
    if isinstance(value, (list, tuple)):
        return any(_contains_json_value(item, expected) for item in value)
    return expected in str(value)


def test_original_mode_preserves_jpeg_format_and_intrinsic_dimensions(
    tmp_path: Path,
) -> None:
    jpeg = _encoded_image("JPEG", (80, 52), (24, 90, 210))
    source = _write_pdf(
        tmp_path / "photo.pdf",
        [[((24, 30, 184, 134), jpeg, {})]],
    )

    outputs = extract_pdf_images(
        source,
        tmp_path / "original",
        mode="original",
        image_format="auto",
        write_manifest=False,
    )

    images = _image_outputs(outputs)
    assert len(images) == 1
    assert images[0].suffix.lower() in {".jpg", ".jpeg"}
    with Image.open(images[0]) as extracted:
        assert extracted.format == "JPEG"
        assert extracted.size == (80, 52)


def test_original_mode_converts_cmyk_jpeg_to_png(tmp_path: Path) -> None:
    cmyk_path = tmp_path / "cmyk.jpg"
    image = Image.new("CMYK", (36, 22), (20, 80, 140, 10))
    image.save(cmyk_path, "JPEG")
    image.close()
    source = tmp_path / "cmyk.pdf"
    document = Canvas(str(source), pagesize=(200, 150))
    document.drawImage(str(cmyk_path), 20, 30, width=144, height=88)
    document.save()

    outputs = extract_pdf_images(
        source,
        tmp_path / "cmyk-output",
        mode="original",
        image_format="png",
        write_manifest=False,
    )

    images = _image_outputs(outputs)
    assert len(images) == 1
    with Image.open(images[0]) as extracted:
        assert extracted.format == "PNG"
        assert extracted.mode in {"RGB", "RGBA"}
        assert extracted.size == (36, 22)


def test_original_mode_composes_pdf_soft_mask_into_png_alpha(tmp_path: Path) -> None:
    transparent = _transparent_png()
    source = _write_pdf(
        tmp_path / "transparent.pdf",
        [[((20, 24, 164, 120), transparent, {})]],
    )

    outputs = extract_pdf_images(
        source,
        tmp_path / "alpha",
        mode="original",
        image_format="auto",
        write_manifest=False,
    )

    images = _image_outputs(outputs)
    assert len(images) == 1
    assert images[0].suffix.lower() == ".png"
    with Image.open(images[0]) as extracted:
        assert extracted.size == (48, 32)
        assert extracted.mode in {"LA", "RGBA"}
        assert extracted.getchannel("A").getextrema() == (0, 255)


def test_original_mode_deduplicates_identical_images_across_pages(
    tmp_path: Path,
) -> None:
    repeated = _encoded_image("JPEG", (64, 40), (210, 70, 30))
    source = _write_pdf(
        tmp_path / "repeated.pdf",
        [
            [((20, 20, 148, 100), repeated, {})],
            [((42, 50, 170, 130), repeated, {})],
        ],
    )

    unique_outputs = extract_pdf_images(
        source,
        tmp_path / "deduplicated",
        mode="original",
        deduplicate=True,
        write_manifest=False,
    )
    occurrence_outputs = extract_pdf_images(
        source,
        tmp_path / "occurrences",
        mode="original",
        deduplicate=False,
        write_manifest=False,
    )

    assert len(_image_outputs(unique_outputs)) == 1
    assert len(_image_outputs(occurrence_outputs)) == 2


def test_encrypted_pdf_accepts_the_correct_password_and_rejects_bad_passwords(
    tmp_path: Path,
) -> None:
    png = _encoded_image("PNG", (42, 26), (70, 130, 210))
    plain = _write_pdf(
        tmp_path / "plain.pdf",
        [[((20, 20, 146, 98), png, {})]],
    )
    encrypted = _encrypt_pdf(plain, tmp_path / "encrypted.pdf", password="open-me")

    outputs = extract_pdf_images(
        encrypted,
        tmp_path / "unlocked",
        mode="original",
        password="open-me",
        write_manifest=False,
    )
    assert len(_image_outputs(outputs)) == 1

    for supplied_password in (None, "wrong-password"):
        with pytest.raises(PDFPasswordError, match="密码"):
            extract_pdf_images(
                encrypted,
                tmp_path / f"rejected-{supplied_password or 'missing'}",
                mode="original",
                password=supplied_password,
                write_manifest=False,
            )


def test_original_mode_extracts_a_true_inline_xref_zero_image(tmp_path: Path) -> None:
    source = _write_inline_image_pdf(tmp_path / "inline.pdf")
    with pymupdf.open(str(source)) as document:
        image_info = document[0].get_image_info(hashes=True, xrefs=True)
    assert len(image_info) == 1
    assert image_info[0]["xref"] == 0

    outputs = extract_pdf_images(
        source,
        tmp_path / "inline-output",
        mode="original",
        image_format="auto",
        write_manifest=False,
    )

    images = _image_outputs(outputs)
    assert len(images) == 1
    with Image.open(images[0]) as extracted:
        assert extracted.size == (37, 23)
        red, green, blue = extracted.convert("RGB").getpixel((18, 11))
    assert green > 140
    assert red < 60
    assert blue < 120


def test_inline_filter_matches_same_bbox_images_by_intrinsic_size(
    tmp_path: Path,
) -> None:
    source = tmp_path / "overlaid-inline.pdf"
    small = Image.new("RGB", (10, 10), (240, 20, 20))
    large = Image.new("RGB", (40, 30), (20, 220, 30))
    try:
        document = Canvas(str(source), pagesize=(200, 150))
        document.drawInlineImage(small, 20, 30, width=100, height=80)
        document.drawInlineImage(large, 20, 30, width=100, height=80)
        document.save()
    finally:
        small.close()
        large.close()

    outputs = extract_pdf_images(
        source,
        tmp_path / "overlaid-output",
        mode="original",
        min_width=20,
        deduplicate=False,
        write_manifest=False,
    )

    images = _image_outputs(outputs)
    assert len(images) == 1
    with Image.open(images[0]) as extracted:
        assert extracted.size == (40, 30)
        red, green, blue = extracted.convert("RGB").getpixel((20, 15))
    assert green > 180
    assert red < 80
    assert blue < 80


def test_original_xref_images_do_not_decode_the_page_inline_catalog(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    png = _encoded_image("PNG", (32, 20), (20, 100, 210))
    source = _write_pdf(
        tmp_path / "xref-only.pdf",
        [[((20, 20, 116, 80), png, {})]],
    )

    def fail_if_called(page: object) -> list[dict[str, Any]]:
        raise AssertionError("xref-only pages must not decode inline image blocks")

    monkeypatch.setattr(pdf_images_module, "_inline_blocks", fail_if_called)
    outputs = extract_pdf_images(
        source,
        tmp_path / "xref-only-output",
        mode="original",
        write_manifest=False,
    )
    assert len(_image_outputs(outputs)) == 1


def test_page_filter_extracts_only_the_requested_page(tmp_path: Path) -> None:
    red = _encoded_image("PNG", (36, 24), (240, 20, 20))
    green = _encoded_image("PNG", (36, 24), (20, 220, 30))
    blue = _encoded_image("PNG", (36, 24), (20, 40, 230))
    source = _write_pdf(
        tmp_path / "pages.pdf",
        [
            [((20, 20, 128, 92), red, {})],
            [((20, 20, 128, 92), green, {})],
            [((20, 20, 128, 92), blue, {})],
        ],
    )

    outputs = extract_pdf_images(
        source,
        tmp_path / "page-two",
        mode="original",
        pages="2",
        image_format="png",
        write_manifest=False,
    )

    images = _image_outputs(outputs)
    assert len(images) == 1
    with Image.open(images[0]) as extracted:
        red_value, green_value, blue_value = extracted.convert("RGB").getpixel(
            (extracted.width // 2, extracted.height // 2)
        )
    assert green_value > 180
    assert red_value < 80
    assert blue_value < 80


def test_visible_mode_renders_the_actual_placed_rectangle_as_png(
    tmp_path: Path,
) -> None:
    jpeg = _encoded_image("JPEG", (60, 30), (220, 110, 20))
    source = _write_pdf(
        tmp_path / "rotated.pdf",
        [[((40, 30, 160, 150), jpeg, {"rotate": 90})]],
        page_size=(200, 180),
    )
    with pymupdf.open(str(source)) as document:
        bbox = pymupdf.Rect(document[0].get_image_info()[0]["bbox"])
        expected = document[0].get_pixmap(
            dpi=144,
            clip=bbox,
            alpha=True,
            annots=False,
        )
        expected_size = (expected.width, expected.height)

    outputs = extract_pdf_images(
        source,
        tmp_path / "visible",
        mode="visible",
        image_format="png",
        dpi=144,
        region_padding=0,
        include_annotations=False,
        write_manifest=False,
    )

    images = _image_outputs(outputs)
    assert len(images) == 1
    with Image.open(images[0]) as extracted:
        assert extracted.format == "PNG"
        assert extracted.size == expected_size
        assert extracted.getbbox() is not None


def test_visible_size_filter_uses_the_clipped_output_dimensions(
    tmp_path: Path,
) -> None:
    png = _encoded_image("PNG", (100, 20), (230, 40, 40))
    source = _write_pdf(
        tmp_path / "partly-outside.pdf",
        [[((-90, 10, 10, 30), png, {})]],
        page_size=(100, 100),
    )

    with pytest.raises(PDFProcessorError, match=r"未发现.*尺寸.*位图"):
        extract_pdf_images(
            source,
            tmp_path / "clipped-output",
            mode="visible",
            dpi=72,
            region_padding=0,
            min_width=20,
            write_manifest=False,
        )


def test_disabled_deduplication_and_manifest_skip_payload_hashing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    png = _encoded_image("PNG", (30, 18), (30, 150, 210))
    source = _write_pdf(
        tmp_path / "no-hash.pdf",
        [[((20, 20, 110, 74), png, {})]],
    )

    def fail_if_hashed(payload: bytes) -> str:
        raise AssertionError("hashing is unnecessary when both consumers are disabled")

    monkeypatch.setattr(pdf_images_module, "_bytes_digest", fail_if_hashed)
    outputs = extract_pdf_images(
        source,
        tmp_path / "no-hash-output",
        deduplicate=False,
        write_manifest=False,
    )
    assert len(_image_outputs(outputs)) == 1


def test_smart_mode_merges_adjacent_image_fragments_into_one_region(
    tmp_path: Path,
) -> None:
    red = _encoded_image("PNG", (20, 20), (245, 20, 20))
    blue = _encoded_image("PNG", (20, 20), (20, 40, 245))
    source = _write_pdf(
        tmp_path / "tiles.pdf",
        [
            [
                ((10, 12, 30, 32), red, {}),
                ((32, 12, 52, 32), blue, {}),
            ]
        ],
        page_size=(90, 60),
    )

    outputs = extract_pdf_images(
        source,
        tmp_path / "smart",
        mode="smart",
        image_format="png",
        dpi=72,
        merge_gap=4,
        region_padding=0,
        deduplicate=False,
        write_manifest=False,
    )

    images = _image_outputs(outputs)
    assert len(images) == 1
    with Image.open(images[0]) as extracted:
        rgb = extracted.convert("RGB")
        assert 41 <= rgb.width <= 43
        assert 19 <= rgb.height <= 21
        pixels = list(rgb.get_flattened_data())
    assert any(
        red_value > 200 and green < 80 and blue_value < 80
        for red_value, green, blue_value in pixels
    )
    assert any(
        blue_value > 200 and red_value < 80 and green < 100
        for red_value, green, blue_value in pixels
    )


@pytest.mark.parametrize(("mode", "expected_count"), [("both", 2), ("all", 3)])
def test_combination_modes_emit_the_documented_extraction_variants(
    tmp_path: Path,
    mode: str,
    expected_count: int,
) -> None:
    png = _encoded_image("PNG", (24, 18), (80, 150, 220))
    source = _write_pdf(
        tmp_path / f"{mode}.pdf",
        [[((20, 20, 92, 74), png, {})]],
    )

    outputs = extract_pdf_images(
        source,
        tmp_path / mode,
        mode=mode,
        image_format="png",
        dpi=72,
        region_padding=0,
        deduplicate=False,
        write_manifest=False,
    )

    assert len(_image_outputs(outputs)) == expected_count


def test_original_mode_reports_a_clear_error_when_pdf_has_no_raster_images(
    tmp_path: Path,
) -> None:
    source = tmp_path / "vectors-only.pdf"
    document = pymupdf.open()
    try:
        page = document.new_page(width=240, height=180)
        page.insert_text((24, 50), "Vector and text only")
        page.draw_rect(pymupdf.Rect(24, 72, 180, 132), color=(0, 0, 0))
        document.save(str(source))
    finally:
        document.close()

    with pytest.raises(DocuForgeError, match=r"未发现.*(?:图片|位图)"):
        extract_pdf_images(
            source,
            tmp_path / "none",
            mode="original",
            write_manifest=False,
        )


@pytest.mark.parametrize(
    "size_filter",
    [{"min_width": 21}, {"min_height": 21}],
)
def test_minimum_dimensions_report_a_clear_error_when_every_image_is_filtered(
    tmp_path: Path,
    size_filter: dict[str, int],
) -> None:
    png = _encoded_image("PNG", (20, 20), (120, 70, 210))
    source = _write_pdf(
        tmp_path / "too-small.pdf",
        [[((10, 10, 50, 50), png, {})]],
    )

    with pytest.raises(PDFProcessorError, match=r"未发现.*尺寸.*位图"):
        extract_pdf_images(
            source,
            tmp_path / "filtered",
            mode="original",
            write_manifest=False,
            **size_filter,
        )


@pytest.mark.parametrize(
    "parameters",
    [
        {"mode": "unknown"},
        {"image_format": "gif"},
        {"mode": "visible", "dpi": 0},
        {"jpeg_quality": 0},
        {"min_width": 0},
        {"min_height": 0},
        {"mode": "smart", "merge_gap": -0.01},
        {"mode": "visible", "region_padding": -0.01},
    ],
)
def test_extraction_rejects_invalid_parameters(
    tmp_path: Path,
    parameters: dict[str, Any],
) -> None:
    png = _encoded_image("PNG", (20, 20), (30, 80, 160))
    source = _write_pdf(
        tmp_path / "validation.pdf",
        [[((10, 10, 50, 50), png, {})]],
    )

    with pytest.raises(ValidationError):
        extract_pdf_images(
            source,
            tmp_path / "invalid",
            write_manifest=False,
            **parameters,
        )


def test_manifest_is_valid_json_and_references_source_and_outputs(
    tmp_path: Path,
) -> None:
    png = _encoded_image("PNG", (44, 28), (40, 170, 90))
    source = _write_pdf(
        tmp_path / "manifest-source.pdf",
        [[((18, 22, 150, 106), png, {})]],
    )

    outputs = extract_pdf_images(
        source,
        tmp_path / "manifest",
        mode="original",
        write_manifest=True,
    )

    images = _image_outputs(outputs)
    manifests = _json_outputs(outputs)
    assert len(images) == 1
    assert len(manifests) == 1
    manifest = json.loads(manifests[0].read_text(encoding="utf-8"))
    assert _contains_json_value(manifest, source.name)
    assert _contains_json_value(manifest, images[0].name)
    assert _contains_json_value(manifest, "page")
