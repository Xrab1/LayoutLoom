from __future__ import annotations

import threading
from pathlib import Path
from unittest.mock import patch

import pytest
from PIL import Image

from docuforge.processors import image as image_processor
from docuforge.processors.image import (
    add_border,
    add_image_watermark,
    add_text_watermark,
    adjust_images,
    apply_filter,
    batch_rename,
    compress_images,
    convert_format,
    crop_images,
    flip_images,
    mosaic_images,
    overlay_images,
    remove_exif,
    resize_images,
    rotate_images,
    scale_images,
    stitch_images,
)


def make_image(
    path: Path,
    size: tuple[int, int] = (80, 60),
    color: tuple[int, ...] = (220, 40, 30),
    mode: str = "RGB",
    *,
    image_format: str | None = None,
    exif: Image.Exif | None = None,
) -> Path:
    image = Image.new(mode, size, color)
    try:
        options = {"exif": exif} if exif is not None else {}
        image.save(path, format=image_format, **options)
    finally:
        image.close()
    return path


def read_size(path: Path) -> tuple[int, int]:
    with Image.open(path) as image:
        return image.size


def test_mpo_still_photo_uses_largest_embedded_image_without_animation_warning(
    tmp_path: Path,
    recwarn: pytest.WarningsRecorder,
) -> None:
    source = tmp_path / "phone_photo.png"
    auxiliary = Image.new("RGB", (40, 30), (10, 20, 200))
    primary = Image.new("RGB", (80, 60), (220, 30, 20))
    try:
        auxiliary.save(
            source,
            format="MPO",
            save_all=True,
            append_images=[primary],
        )
    finally:
        auxiliary.close()
        primary.close()

    loaded, source_format = image_processor._load_image(source)
    try:
        assert source_format == "MPO"
        assert loaded.size == (80, 60)
        pixel = loaded.getpixel((0, 0))
        assert all(abs(actual - expected) <= 2 for actual, expected in zip(pixel, (220, 30, 20)))
        assert not any("仅处理首帧" in str(item.message) for item in recwarn.list)
    finally:
        loaded.close()


def test_convert_format_handles_alpha_chinese_paths_and_avoids_overwrite(
    tmp_path: Path,
) -> None:
    source = make_image(tmp_path / "透明图片.png", color=(10, 20, 30, 100), mode="RGBA")

    first = convert_format(source, "jpg", background="#ffffff")
    second = convert_format(source, ".jpeg", background="#ffffff")

    assert first == [tmp_path / "透明图片.jpg"]
    assert second == [tmp_path / "透明图片_1.jpg"]
    assert source.exists()
    with Image.open(first[0]) as converted:
        assert converted.format == "JPEG"
        assert converted.mode == "RGB"
        assert converted.size == (80, 60)


def test_resize_scale_crop_rotate_and_flip(tmp_path: Path) -> None:
    source = make_image(tmp_path / "source.png", size=(80, 40))

    resized = resize_images(source, (30, 30))[0]
    fitted = resize_images(source, (30, 30), keep_aspect=True)[0]
    scaled = scale_images(source, 0.5)[0]
    cropped = crop_images(source, (10, 5, 50, 25))[0]
    rotated = rotate_images(source, 90)[0]
    flipped = flip_images(source, "vertical")[0]

    assert read_size(resized) == (30, 30)
    assert read_size(fitted) == (30, 15)
    assert read_size(scaled) == (40, 20)
    assert read_size(cropped) == (40, 20)
    assert read_size(rotated) == (40, 80)
    assert read_size(flipped) == (80, 40)


def test_proportional_resize_upscales_and_right_angle_rotation_is_lossless(
    tmp_path: Path,
) -> None:
    small = make_image(tmp_path / "small.png", size=(10, 5), color=(20, 40, 60))
    enlarged = resize_images(small, (40, 40), keep_aspect=True)[0]
    assert read_size(enlarged) == (40, 20)

    pixels = Image.new("RGB", (3, 2))
    pixels.putdata(
        [
            (1, 2, 3),
            (4, 5, 6),
            (7, 8, 9),
            (10, 11, 12),
            (13, 14, 15),
            (16, 17, 18),
        ]
    )
    source = tmp_path / "pixels.png"
    expected = pixels.transpose(Image.Transpose.ROTATE_90)
    pixels.save(source)
    pixels.close()
    rotated = rotate_images(source, 90)[0]
    try:
        with Image.open(rotated) as actual:
            assert actual.size == expected.size
            assert actual.tobytes() == expected.tobytes()
    finally:
        expected.close()


def test_parallel_image_pipeline_preserves_input_order(tmp_path: Path) -> None:
    sources = [
        make_image(tmp_path / f"input_{index}.png", color=(index, 20, 30))
        for index in range(6)
    ]
    original_load = image_processor._load_image
    barrier = threading.Barrier(3)
    counter_lock = threading.Lock()
    call_count = 0
    worker_names: set[str] = set()

    def synchronized_load(path: Path):
        nonlocal call_count
        with counter_lock:
            call_count += 1
            sequence = call_count
            worker_names.add(threading.current_thread().name)
        if sequence <= 3:
            barrier.wait(timeout=3)
        return original_load(path)

    with patch.object(
        image_processor, "optimal_worker_count", return_value=3
    ), patch.object(image_processor, "_load_image", side_effect=synchronized_load):
        outputs = image_processor.resize_images(sources, (24, 18), tmp_path / "out")

    assert [path.name for path in outputs] == [
        f"input_{index}_resized.png" for index in range(6)
    ]
    assert len(worker_names) == 3


def test_compress_images_honors_max_bytes(tmp_path: Path) -> None:
    source = tmp_path / "noise.jpg"
    image = Image.effect_noise((320, 240), 100).convert("RGB")
    try:
        image.save(source, quality=100)
    finally:
        image.close()

    result = compress_images(source, quality=85, max_bytes=12_000)[0]

    assert result.stat().st_size <= 12_000
    with Image.open(result) as compressed:
        compressed.verify()


def test_remove_exif_removes_metadata_but_keeps_pixels(tmp_path: Path) -> None:
    exif = Image.Exif()
    exif[0x010E] = "private description"
    source = make_image(
        tmp_path / "metadata.jpg",
        color=(20, 80, 140),
        image_format="JPEG",
        exif=exif,
    )

    result = remove_exif(source)[0]

    with Image.open(source) as original, Image.open(result) as cleaned:
        assert original.getexif().get(0x010E) == "private description"
        assert len(cleaned.getexif()) == 0
        assert cleaned.size == original.size


def test_adjust_and_filters_generate_valid_outputs(tmp_path: Path) -> None:
    source = make_image(tmp_path / "filter.png", color=(100, 120, 140))

    adjusted = adjust_images(source, brightness=1.2, contrast=1.1, saturation=0.5)[0]
    grayscale = apply_filter(source, "grayscale")[0]
    sepia = apply_filter(source, "vintage", intensity=0.6)[0]
    blurred = apply_filter(source, "blur", intensity=2)[0]

    with Image.open(adjusted) as image:
        assert image.getpixel((0, 0)) != (100, 120, 140)
    with Image.open(grayscale) as image:
        assert image.mode == "L"
    with Image.open(sepia) as image:
        assert image.mode == "RGB"
    with Image.open(blurred) as image:
        assert image.size == (80, 60)


def test_palette_and_grayscale_modes_are_handled(tmp_path: Path) -> None:
    palette_path = tmp_path / "palette.png"
    palette = Image.new("P", (24, 16))
    palette.putpalette([channel for value in range(256) for channel in (value, 0, 0)])
    palette.putdata([(x + y) % 256 for y in range(16) for x in range(24)])
    palette.save(palette_path)
    palette.close()

    gray_path = tmp_path / "gray.png"
    gray = Image.new("L", (20, 10), 100)
    gray.save(gray_path)
    gray.close()

    adjusted = adjust_images(palette_path, brightness=1.1)[0]
    filtered = apply_filter(palette_path, "blur")[0]
    bordered = add_border(gray_path, 2, color="white")[0]
    rotated = rotate_images(gray_path, 30, fillcolor="black")[0]

    with Image.open(adjusted) as image:
        assert image.mode == "RGB"
    with Image.open(filtered) as image:
        assert image.size == (24, 16)
    with Image.open(bordered) as image:
        assert image.mode == "L"
        assert image.getpixel((0, 0)) == 255
    with Image.open(rotated) as image:
        assert image.mode == "L"


def test_watermarks_overlay_border_and_mosaic(tmp_path: Path) -> None:
    source = make_image(
        tmp_path / "base.png", size=(100, 80), color=(0, 0, 0, 255), mode="RGBA"
    )
    mark = make_image(
        tmp_path / "mark.png", size=(20, 10), color=(255, 0, 0, 255), mode="RGBA"
    )

    text_result = add_text_watermark(
        source,
        "TEST",
        position="center",
        font_size=18,
        opacity=1,
    )[0]
    watermark_result = add_image_watermark(
        source,
        mark,
        position="top-left",
        scale=None,
        opacity=1,
        margin=0,
    )[0]
    overlay_result = overlay_images(source, mark, position=(30, 20), opacity=1)[0]
    border_result = add_border(source, (2, 3), color="blue")[0]

    gradient = Image.new("RGB", (60, 60))
    gradient.putdata(
        [(x * 4, y * 4, (x + y) * 2) for y in range(60) for x in range(60)]
    )
    gradient_path = tmp_path / "gradient.png"
    gradient.save(gradient_path)
    gradient.close()
    mosaic_result = mosaic_images(gradient_path, (10, 10, 50, 50), block_size=10)[0]

    with Image.open(text_result) as image:
        extrema = image.convert("RGB").getextrema()
        assert any(maximum > 0 for _minimum, maximum in extrema)
    with Image.open(watermark_result) as image:
        assert image.convert("RGB").getpixel((0, 0))[0] > 200
    with Image.open(overlay_result) as image:
        assert image.convert("RGB").getpixel((30, 20))[0] > 200
    assert read_size(border_result) == (104, 86)
    with Image.open(mosaic_result) as image:
        assert image.getpixel((11, 11)) == image.getpixel((18, 18))


def test_stitch_images_supports_both_directions(tmp_path: Path) -> None:
    first = make_image(tmp_path / "first.png", size=(30, 20), color=(255, 0, 0))
    second = make_image(tmp_path / "second.png", size=(10, 40), color=(0, 255, 0))

    vertical = stitch_images(
        [first, second],
        tmp_path / "vertical.png",
        direction="vertical",
        spacing=5,
        alignment="end",
    )[0]
    horizontal = stitch_images(
        [first, second],
        tmp_path / "horizontal.jpg",
        direction="horizontal",
        spacing=3,
    )[0]

    assert read_size(vertical) == (30, 65)
    assert read_size(horizontal) == (43, 40)


def test_stitch_images_preserves_alpha_without_squaring_it(tmp_path: Path) -> None:
    first = make_image(
        tmp_path / "alpha.png", size=(2, 2), color=(255, 0, 0, 128), mode="RGBA"
    )
    second = make_image(
        tmp_path / "opaque.png", size=(2, 2), color=(0, 0, 255, 255), mode="RGBA"
    )
    output = stitch_images(
        [first, second],
        tmp_path / "alpha_stitched.png",
        direction="horizontal",
        spacing=1,
        background=(0, 0, 0, 0),
        alignment="start",
    )[0]

    with Image.open(output) as stitched:
        assert stitched.mode == "RGBA"
        assert stitched.getpixel((0, 0))[3] == 128
        assert stitched.getpixel((2, 0))[3] == 0
        assert stitched.getpixel((3, 0)) == (0, 0, 255, 255)


def test_batch_rename_can_copy_move_and_avoid_collisions(tmp_path: Path) -> None:
    first = make_image(tmp_path / "一.png", color=(1, 2, 3))
    second = make_image(tmp_path / "二.jpg", color=(4, 5, 6))
    copied_dir = tmp_path / "copies"

    copied = batch_rename(
        [first, second],
        "photo_{index:02d}",
        copied_dir,
        start=7,
        move=False,
    )
    collision = batch_rename(first, "photo_07", copied_dir, move=False)

    assert copied == [copied_dir / "photo_07.png", copied_dir / "photo_08.jpg"]
    assert collision == [copied_dir / "photo_07_1.png"]
    assert first.exists() and second.exists()

    moved = batch_rename([first, second], "归档_{index}", tmp_path / "moved")
    assert all(path.exists() for path in moved)
    assert not first.exists() and not second.exists()


def test_batch_rename_rejects_duplicate_overwrite_destinations(tmp_path: Path) -> None:
    first = make_image(tmp_path / "first.png")
    second = make_image(tmp_path / "second.png")

    with pytest.raises(ValueError, match="duplicate destinations"):
        batch_rename(
            [first, second],
            "same.png",
            tmp_path / "output",
            overwrite=True,
        )

    assert first.exists() and second.exists()


@pytest.mark.parametrize(
    ("call", "message"),
    [
        (lambda path: resize_images(path, (0, 20)), "positive integer"),
        (lambda path: flip_images(path, "diagonal"), "horizontal"),
        (lambda path: apply_filter(path, "unknown"), "Unsupported filter"),
        (lambda path: mosaic_images(path, (200, 200, 210, 210)), "does not overlap"),
    ],
)
def test_validation_errors_are_actionable(tmp_path: Path, call, message: str) -> None:
    source = make_image(tmp_path / "input.png")
    with pytest.raises(ValueError, match=message):
        call(source)
