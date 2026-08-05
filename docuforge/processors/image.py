"""High quality, dependency-light image processing helpers.

The functions in this module deliberately accept ordinary Python arguments and
return ``list[pathlib.Path]``.  This keeps them easy to call from a CLI, a GUI,
or a registry without coupling the image processor to application core types.

Pillow is the only image dependency.  Every generated file is first written to
a temporary file in the destination directory and then atomically replaced.
Unless ``overwrite=True`` is explicitly supplied, a destination is reserved
with an exclusive create and a numeric suffix is used to avoid overwriting an
existing file.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence
from concurrent.futures import Future, ThreadPoolExecutor
from io import BytesIO
import math
import os
from pathlib import Path
import shutil
import tempfile
from typing import Any, TypeAlias
import warnings

from PIL import (
    Image,
    ImageColor,
    ImageDraw,
    ImageEnhance,
    ImageFilter,
    ImageFont,
    ImageOps,
)

from ..utils import optimal_worker_count


PathLike: TypeAlias = str | os.PathLike[str]
PathInputs: TypeAlias = PathLike | Iterable[PathLike]
ColorValue: TypeAlias = str | tuple[int, int, int] | tuple[int, int, int, int]
Position: TypeAlias = str | tuple[int, int]


__all__ = [
    "convert_format",
    "resize_images",
    "scale_images",
    "crop_images",
    "rotate_images",
    "flip_images",
    "compress_images",
    "remove_exif",
    "strip_exif",
    "enhance_images",
    "adjust_images",
    "apply_filter",
    "add_text_watermark",
    "add_image_watermark",
    "add_border",
    "mosaic_images",
    "stitch_images",
    "overlay_images",
    "batch_rename",
]


_FORMAT_ALIASES: dict[str, tuple[str, str]] = {
    "JPG": ("JPEG", ".jpg"),
    "JPEG": ("JPEG", ".jpg"),
    "PNG": ("PNG", ".png"),
    "BMP": ("BMP", ".bmp"),
    "WEBP": ("WEBP", ".webp"),
    "TIF": ("TIFF", ".tif"),
    "TIFF": ("TIFF", ".tiff"),
    "GIF": ("GIF", ".gif"),
    "ICO": ("ICO", ".ico"),
    "PPM": ("PPM", ".ppm"),
}

_RESAMPLING = {
    "nearest": Image.Resampling.NEAREST,
    "box": Image.Resampling.BOX,
    "bilinear": Image.Resampling.BILINEAR,
    "hamming": Image.Resampling.HAMMING,
    "bicubic": Image.Resampling.BICUBIC,
    "lanczos": Image.Resampling.LANCZOS,
}


def _paths(inputs: PathInputs) -> list[Path]:
    if isinstance(inputs, (str, os.PathLike)):
        result = [Path(inputs).expanduser()]
    else:
        result = [Path(item).expanduser() for item in inputs]
    if not result:
        raise ValueError("At least one input image is required")
    for path in result:
        if not path.is_file():
            raise FileNotFoundError(f"Image does not exist or is not a file: {path}")
    return result


def _positive_int(value: int, name: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ValueError(f"{name} must be a positive integer")
    return value


def _unit_interval(value: float, name: str) -> float:
    value = float(value)
    if not 0.0 <= value <= 1.0:
        raise ValueError(f"{name} must be between 0 and 1")
    return value


def _resample(value: str | Image.Resampling) -> Image.Resampling:
    if isinstance(value, Image.Resampling):
        return value
    try:
        return _RESAMPLING[value.strip().lower()]
    except (AttributeError, KeyError) as exc:
        choices = ", ".join(sorted(_RESAMPLING))
        raise ValueError(
            f"Unknown resampling method {value!r}; choose one of {choices}"
        ) from exc


def _normalise_format(value: str) -> tuple[str, str]:
    key = value.strip().lstrip(".").upper()
    if key in _FORMAT_ALIASES:
        return _FORMAT_ALIASES[key]

    Image.init()
    extension = f".{key.lower()}"
    pillow_format = Image.registered_extensions().get(extension)
    if pillow_format and pillow_format in Image.SAVE:
        return pillow_format, extension
    raise ValueError(f"Unsupported output image format: {value!r}")


def _format_for_path(path: Path, fallback: str | None = None) -> str:
    Image.init()
    result = Image.registered_extensions().get(path.suffix.lower())
    if result and result in Image.SAVE:
        return result
    if fallback and fallback in Image.SAVE:
        return fallback
    raise ValueError(f"Cannot infer a writable image format from: {path}")


def _load_image(path: Path) -> tuple[Image.Image, str | None]:
    from docuforge.runner import check_cancelled

    check_cancelled("任务已取消；已完成的图片会保留")
    try:
        with Image.open(path) as opened:
            frame_count = int(getattr(opened, "n_frames", 1))
            source_format = opened.format
            # MPO is a still-photo container used by many phones and cameras.
            # It commonly stores one full-resolution photograph plus a smaller
            # auxiliary/preview image.  Pillow exposes those images through the
            # same ``n_frames`` API used by GIF/APNG, but treating an MPO as an
            # animation produces a misleading warning.  Select the largest
            # embedded photograph as the authentic source image instead.
            if frame_count > 1 and str(source_format or "").upper() == "MPO":
                primary_index = 0
                primary_area = -1
                for frame_index in range(frame_count):
                    opened.seek(frame_index)
                    width, height = opened.size
                    area = int(width) * int(height)
                    if area > primary_area:
                        primary_index = frame_index
                        primary_area = area
                opened.seek(primary_index)
            elif frame_count > 1:
                warnings.warn(
                    f"{path.name} 包含 {frame_count} 帧；当前图片编辑操作仅处理首帧。",
                    UserWarning,
                    stacklevel=2,
                )
                opened.seek(0)
            else:
                opened.seek(0)
            prepared = ImageOps.exif_transpose(opened)
            prepared.load()
            result = prepared.copy()
            result.info = dict(prepared.info)
            return result, source_format
    except (OSError, ValueError) as exc:
        raise ValueError(f"Unable to read image {path}: {exc}") from exc


def _metadata(image: Image.Image, image_format: str) -> dict[str, Any]:
    result: dict[str, Any] = {}
    if image_format in {"JPEG", "PNG", "TIFF", "WEBP"}:
        for key in ("icc_profile", "exif"):
            if image.info.get(key):
                result[key] = image.info[key]
    if image_format in {"JPEG", "PNG", "TIFF"} and image.info.get("dpi"):
        result["dpi"] = image.info["dpi"]
    return result


def _encoding_defaults(image_format: str) -> dict[str, Any]:
    if image_format == "JPEG":
        return {
            "quality": 95,
            "subsampling": 0,
            "optimize": True,
            "progressive": True,
        }
    if image_format == "PNG":
        return {"optimize": False, "compress_level": 6}
    if image_format == "WEBP":
        return {"quality": 95, "method": 6}
    if image_format == "TIFF":
        return {"compression": "tiff_deflate"}
    return {}


def _jpeg_subsampling(quality: int) -> int:
    """Prefer color fidelity at high quality and smaller files at lower quality."""

    return 0 if quality >= 85 else 1 if quality >= 70 else 2


def _rgba(color: ColorValue, opacity: float = 1.0) -> tuple[int, int, int, int]:
    if isinstance(color, str):
        parsed = ImageColor.getcolor(color, "RGBA")
    elif len(color) == 3:
        parsed = (int(color[0]), int(color[1]), int(color[2]), 255)
    elif len(color) == 4:
        parsed = tuple(int(channel) for channel in color)  # type: ignore[assignment]
    else:
        raise ValueError("Color tuples must contain three (RGB) or four (RGBA) values")
    if any(channel < 0 or channel > 255 for channel in parsed):
        raise ValueError("Color channels must be between 0 and 255")
    return parsed[0], parsed[1], parsed[2], round(parsed[3] * opacity)


def _color_for_mode(color: ColorValue, mode: str) -> Any:
    """Convert an RGB/RGBA color into a value Pillow accepts for ``mode``."""

    rgba = _rgba(color)
    if mode == "RGBA":
        return rgba
    if mode == "RGB":
        return rgba[:3]
    swatch = Image.new("RGBA", (1, 1), rgba)
    converted: Image.Image | None = None
    try:
        converted = swatch.convert(mode)
        return converted.getpixel((0, 0))
    finally:
        if converted is not None:
            converted.close()
        swatch.close()


def _contains_alpha(image: Image.Image) -> bool:
    return "A" in image.getbands() or (
        image.mode == "P" and image.info.get("transparency") is not None
    )


def _flatten(image: Image.Image, background: ColorValue) -> Image.Image:
    rgba = image.convert("RGBA")
    color = _rgba(background)
    canvas = Image.new("RGBA", rgba.size, color)
    canvas.alpha_composite(rgba)
    return canvas.convert("RGB")


def _prepare_for_format(
    image: Image.Image,
    image_format: str,
    background: ColorValue = "white",
) -> Image.Image:
    if image_format == "JPEG":
        if _contains_alpha(image):
            return _flatten(image, background)
        if image.mode not in {"RGB", "L", "CMYK"}:
            return image.convert("RGB")
    elif image_format in {"BMP", "PPM"}:
        if _contains_alpha(image):
            return _flatten(image, background)
        if image.mode not in {"RGB", "L"}:
            return image.convert("RGB")
    elif image_format == "WEBP" and image.mode not in {"RGB", "RGBA", "L", "LA"}:
        return image.convert("RGBA" if _contains_alpha(image) else "RGB")
    elif image_format == "PNG" and image.mode in {"CMYK", "YCbCr"}:
        return image.convert("RGB")
    elif image_format == "ICO" and image.mode not in {"RGB", "RGBA"}:
        return image.convert("RGBA" if _contains_alpha(image) else "RGB")
    return image


def _reserve_path(path: Path, overwrite: bool) -> tuple[Path, bool]:
    path.parent.mkdir(parents=True, exist_ok=True)
    if overwrite:
        return path, False

    candidate = path
    number = 1
    while True:
        try:
            descriptor = os.open(candidate, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            candidate = path.with_name(f"{path.stem}_{number}{path.suffix}")
            number += 1
        else:
            os.close(descriptor)
            return candidate, True


def _destination(
    source: Path,
    output_dir: PathLike | None,
    marker: str,
    *,
    extension: str | None = None,
    overwrite: bool,
) -> tuple[Path, bool]:
    directory = (
        Path(output_dir).expanduser() if output_dir is not None else source.parent
    )
    suffix = extension if extension is not None else source.suffix
    return _reserve_path(directory / f"{source.stem}{marker}{suffix}", overwrite)


def _cleanup_reservation(destination: Path, reserved: bool) -> None:
    if reserved:
        destination.unlink(missing_ok=True)


def _atomic_write_bytes(data: bytes, destination: Path, reserved: bool) -> None:
    descriptor, temporary_name = tempfile.mkstemp(
        dir=destination.parent,
        prefix=f".{destination.stem}-",
        suffix=".tmp",
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, destination)
    except Exception:
        temporary.unlink(missing_ok=True)
        _cleanup_reservation(destination, reserved)
        raise


def _atomic_save(
    image: Image.Image,
    destination: Path,
    image_format: str,
    reserved: bool,
    *,
    save_options: dict[str, Any] | None = None,
) -> None:
    descriptor, temporary_name = tempfile.mkstemp(
        dir=destination.parent,
        prefix=f".{destination.stem}-",
        suffix=".tmp",
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        image.save(temporary, format=image_format, **(save_options or {}))
        os.replace(temporary, destination)
    except Exception:
        temporary.unlink(missing_ok=True)
        _cleanup_reservation(destination, reserved)
        raise


def _save_processed(
    result: Image.Image,
    source: Image.Image,
    destination: Path,
    image_format: str,
    reserved: bool,
    *,
    background: ColorValue = "white",
    preserve_metadata: bool = True,
    save_options: dict[str, Any] | None = None,
) -> None:
    prepared = _prepare_for_format(result, image_format, background)
    options = _encoding_defaults(image_format)
    if preserve_metadata:
        options.update(_metadata(source, image_format))
    if save_options:
        options.update(save_options)
    try:
        _atomic_save(
            prepared, destination, image_format, reserved, save_options=options
        )
    finally:
        if prepared is not result:
            prepared.close()


def _ordered_parallel_paths(
    sources: list[Path], worker: Callable[[Path], Path]
) -> list[Path]:
    workers = optimal_worker_count(len(sources), cap=4)
    if workers == 1:
        return [worker(source) for source in sources]

    futures: list[Future[Path]] = []
    with ThreadPoolExecutor(
        max_workers=workers, thread_name_prefix="docuforge-image"
    ) as executor:
        try:
            futures = [executor.submit(worker, source) for source in sources]
            return [future.result() for future in futures]
        except Exception:
            for future in futures:
                future.cancel()
            raise


def _map_images(
    inputs: PathInputs,
    output_dir: PathLike | None,
    marker: str,
    transform: Callable[[Image.Image, Path], Image.Image],
    *,
    overwrite: bool,
    background: ColorValue = "white",
    preserve_metadata: bool = True,
) -> list[Path]:
    sources = _paths(inputs)

    def process(source_path: Path) -> Path:
        image, source_format = _load_image(source_path)
        result: Image.Image | None = None
        destination: Path | None = None
        reserved = False
        try:
            result = transform(image, source_path)
            if not isinstance(result, Image.Image):
                raise TypeError("Image transform must return a PIL.Image.Image")
            image_format = _format_for_path(source_path, source_format)
            destination, reserved = _destination(
                source_path,
                output_dir,
                marker,
                overwrite=overwrite,
            )
            _save_processed(
                result,
                image,
                destination,
                image_format,
                reserved,
                background=background,
                preserve_metadata=preserve_metadata,
            )
            return destination
        except Exception:
            if destination is not None:
                _cleanup_reservation(destination, reserved)
            raise
        finally:
            if result is not None and result is not image:
                result.close()
            image.close()

    return _ordered_parallel_paths(sources, process)


def convert_format(
    inputs: PathInputs,
    target_format: str,
    output_dir: PathLike | None = None,
    *,
    quality: int = 95,
    background: ColorValue = "white",
    overwrite: bool = False,
) -> list[Path]:
    """Convert one or more images to ``target_format``.

    Alpha is composited over ``background`` when converting to a format such as
    JPEG that cannot store transparency.
    """

    image_format, extension = _normalise_format(target_format)
    if not 1 <= int(quality) <= 100:
        raise ValueError("quality must be between 1 and 100")

    sources = _paths(inputs)

    def process(source_path: Path) -> Path:
        image, _ = _load_image(source_path)
        destination: Path | None = None
        reserved = False
        prepared: Image.Image | None = None
        try:
            destination, reserved = _destination(
                source_path,
                output_dir,
                "",
                extension=extension,
                overwrite=overwrite,
            )
            prepared = _prepare_for_format(image, image_format, background)
            options = _metadata(image, image_format)
            if image_format == "JPEG":
                options.update(
                    quality=int(quality),
                    subsampling=_jpeg_subsampling(int(quality)),
                    optimize=True,
                    progressive=True,
                )
            elif image_format == "WEBP":
                options.update(quality=int(quality), method=6)
            elif image_format == "PNG":
                options.update(optimize=False, compress_level=6)
            _atomic_save(
                prepared, destination, image_format, reserved, save_options=options
            )
            return destination
        except Exception:
            if destination is not None:
                _cleanup_reservation(destination, reserved)
            raise
        finally:
            if prepared is not None and prepared is not image:
                prepared.close()
            image.close()

    return _ordered_parallel_paths(sources, process)


def resize_images(
    inputs: PathInputs,
    size: tuple[int, int],
    output_dir: PathLike | None = None,
    *,
    keep_aspect: bool = False,
    resample: str | Image.Resampling = "lanczos",
    overwrite: bool = False,
) -> list[Path]:
    """Resize images to an exact size or fit them inside it proportionally."""

    width = _positive_int(size[0], "width")
    height = _positive_int(size[1], "height")
    method = _resample(resample)

    def transform(image: Image.Image, _: Path) -> Image.Image:
        if keep_aspect:
            return ImageOps.contain(image, (width, height), method)
        return image.resize((width, height), method)

    return _map_images(inputs, output_dir, "_resized", transform, overwrite=overwrite)


def scale_images(
    inputs: PathInputs,
    scale: float | tuple[float, float],
    output_dir: PathLike | None = None,
    *,
    resample: str | Image.Resampling = "lanczos",
    overwrite: bool = False,
) -> list[Path]:
    """Scale images by one proportional factor or separate x/y factors."""

    if isinstance(scale, Sequence) and not isinstance(scale, (str, bytes)):
        if len(scale) != 2:
            raise ValueError("scale tuples must contain exactly two factors")
        scale_x, scale_y = float(scale[0]), float(scale[1])
    else:
        scale_x = scale_y = float(scale)
    if scale_x <= 0 or scale_y <= 0:
        raise ValueError("scale factors must be greater than zero")
    method = _resample(resample)

    def transform(image: Image.Image, _: Path) -> Image.Image:
        width = max(1, round(image.width * scale_x))
        height = max(1, round(image.height * scale_y))
        return image.resize((width, height), method)

    return _map_images(inputs, output_dir, "_scaled", transform, overwrite=overwrite)


def crop_images(
    inputs: PathInputs,
    box: tuple[int, int, int, int],
    output_dir: PathLike | None = None,
    *,
    overwrite: bool = False,
) -> list[Path]:
    """Crop ``(left, top, right, bottom)`` from every image."""

    left, top, right, bottom = (int(value) for value in box)
    if right <= left or bottom <= top:
        raise ValueError("crop box must have right > left and bottom > top")

    def transform(image: Image.Image, _: Path) -> Image.Image:
        return image.crop((left, top, right, bottom))

    return _map_images(inputs, output_dir, "_cropped", transform, overwrite=overwrite)


def rotate_images(
    inputs: PathInputs,
    angle: float,
    output_dir: PathLike | None = None,
    *,
    expand: bool = True,
    fillcolor: ColorValue | None = None,
    resample: str | Image.Resampling = "bicubic",
    overwrite: bool = False,
) -> list[Path]:
    """Rotate counter-clockwise by ``angle`` degrees."""

    method = _resample(resample)

    def transform(image: Image.Image, _: Path) -> Image.Image:
        normalized_angle = float(angle) % 360
        if expand and math.isclose(normalized_angle, 0, abs_tol=1e-9):
            return image.copy()
        if expand and math.isclose(normalized_angle, 90, abs_tol=1e-9):
            return image.transpose(Image.Transpose.ROTATE_90)
        if expand and math.isclose(normalized_angle, 180, abs_tol=1e-9):
            return image.transpose(Image.Transpose.ROTATE_180)
        if expand and math.isclose(normalized_angle, 270, abs_tol=1e-9):
            return image.transpose(Image.Transpose.ROTATE_270)
        working = image
        if fillcolor is not None and image.mode == "P":
            working = image.convert("RGBA" if _contains_alpha(image) else "RGB")
        fill = None
        if fillcolor is not None:
            fill = _color_for_mode(fillcolor, working.mode)
        try:
            return working.rotate(float(angle), method, expand=expand, fillcolor=fill)
        finally:
            if working is not image:
                working.close()

    return _map_images(inputs, output_dir, "_rotated", transform, overwrite=overwrite)


def flip_images(
    inputs: PathInputs,
    direction: str = "horizontal",
    output_dir: PathLike | None = None,
    *,
    overwrite: bool = False,
) -> list[Path]:
    """Flip images horizontally or vertically."""

    direction_key = direction.strip().lower().replace("_", "-")
    horizontal = {"horizontal", "h", "left-right", "mirror"}
    vertical = {"vertical", "v", "top-bottom"}
    if direction_key not in horizontal | vertical:
        raise ValueError("direction must be 'horizontal' or 'vertical'")

    def transform(image: Image.Image, _: Path) -> Image.Image:
        return (
            ImageOps.mirror(image)
            if direction_key in horizontal
            else ImageOps.flip(image)
        )

    return _map_images(inputs, output_dir, "_flipped", transform, overwrite=overwrite)


def _compression_options(
    image: Image.Image,
    image_format: str,
    quality: int,
    optimize: bool,
) -> dict[str, Any]:
    options = _metadata(image, image_format)
    if image_format == "JPEG":
        options.update(
            quality=quality,
            subsampling=_jpeg_subsampling(quality),
            optimize=optimize,
            progressive=True,
        )
    elif image_format == "WEBP":
        options.update(quality=quality, method=6)
    elif image_format == "PNG":
        options.update(optimize=optimize, compress_level=9 if optimize else 6)
    elif image_format == "TIFF":
        options.update(compression="tiff_deflate")
    return options


def _encode(
    image: Image.Image,
    source: Image.Image,
    image_format: str,
    quality: int,
    optimize: bool,
) -> bytes:
    prepared = _prepare_for_format(image, image_format)
    try:
        stream = BytesIO()
        prepared.save(
            stream,
            format=image_format,
            **_compression_options(source, image_format, quality, optimize),
        )
        return stream.getvalue()
    finally:
        if prepared is not image:
            prepared.close()


def _encode_to_limit(
    image: Image.Image,
    source: Image.Image,
    image_format: str,
    quality: int,
    min_quality: int,
    max_bytes: int,
    optimize: bool,
    allow_resize: bool,
) -> bytes:
    quality_formats = {"JPEG", "WEBP"}
    working = image
    owns_working = False
    try:
        for _ in range(12):
            if image_format in quality_formats:
                low, high = min_quality, quality
                best: bytes | None = None
                smallest: bytes | None = None
                while low <= high:
                    candidate_quality = (low + high) // 2
                    encoded = _encode(
                        working, source, image_format, candidate_quality, optimize
                    )
                    if smallest is None or len(encoded) < len(smallest):
                        smallest = encoded
                    if len(encoded) <= max_bytes:
                        best = encoded
                        low = candidate_quality + 1
                    else:
                        high = candidate_quality - 1
                if best is not None:
                    return best
                encoded = smallest or _encode(
                    working, source, image_format, min_quality, optimize
                )
            else:
                encoded = _encode(working, source, image_format, quality, optimize)

            if len(encoded) <= max_bytes:
                return encoded
            if not allow_resize or working.width == 1 and working.height == 1:
                break

            ratio = min(0.9, math.sqrt(max_bytes / len(encoded)) * 0.95)
            next_size = (
                max(1, math.floor(working.width * ratio)),
                max(1, math.floor(working.height * ratio)),
            )
            if next_size == working.size:
                next_size = (max(1, working.width - 1), max(1, working.height - 1))
            resized = working.resize(next_size, Image.Resampling.LANCZOS)
            if owns_working:
                working.close()
            working = resized
            owns_working = True
        raise ValueError(
            f"Unable to compress image below {max_bytes} bytes with the requested settings"
        )
    finally:
        if owns_working:
            working.close()


def compress_images(
    inputs: PathInputs,
    output_dir: PathLike | None = None,
    *,
    quality: int = 80,
    max_bytes: int | None = None,
    min_quality: int = 25,
    optimize: bool = True,
    allow_resize: bool = True,
    overwrite: bool = False,
) -> list[Path]:
    """Compress images, optionally enforcing a maximum encoded byte size.

    JPEG and WebP use the highest quality that fits ``max_bytes``.  If quality
    reduction alone cannot meet the limit and ``allow_resize`` is true, pixel
    dimensions are reduced proportionally until the target is met.
    """

    quality = int(quality)
    min_quality = int(min_quality)
    if not 1 <= min_quality <= quality <= 100:
        raise ValueError(
            "quality values must satisfy 1 <= min_quality <= quality <= 100"
        )
    if max_bytes is not None:
        _positive_int(max_bytes, "max_bytes")

    sources = _paths(inputs)

    def process(source_path: Path) -> Path:
        image, source_format = _load_image(source_path)
        destination: Path | None = None
        reserved = False
        try:
            image_format = _format_for_path(source_path, source_format)
            destination, reserved = _destination(
                source_path,
                output_dir,
                "_compressed",
                overwrite=overwrite,
            )
            if max_bytes is None:
                effective_optimize = optimize if image_format != "PNG" else False
                data = _encode(image, image, image_format, quality, effective_optimize)
            else:
                data = _encode_to_limit(
                    image,
                    image,
                    image_format,
                    quality,
                    min_quality,
                    max_bytes,
                    optimize,
                    allow_resize,
                )
            _atomic_write_bytes(data, destination, reserved)
            return destination
        except Exception:
            if destination is not None:
                _cleanup_reservation(destination, reserved)
            raise
        finally:
            image.close()

    return _ordered_parallel_paths(sources, process)


def remove_exif(
    inputs: PathInputs,
    output_dir: PathLike | None = None,
    *,
    overwrite: bool = False,
) -> list[Path]:
    """Write image copies without EXIF or other embedded metadata."""

    def transform(image: Image.Image, _: Path) -> Image.Image:
        result = image.copy()
        result.info.clear()
        return result

    return _map_images(
        inputs,
        output_dir,
        "_no_exif",
        transform,
        overwrite=overwrite,
        preserve_metadata=False,
    )


def strip_exif(
    inputs: PathInputs,
    output_dir: PathLike | None = None,
    *,
    overwrite: bool = False,
) -> list[Path]:
    """Alias with an explicit function signature for :func:`remove_exif`."""

    return remove_exif(inputs, output_dir, overwrite=overwrite)


def enhance_images(
    inputs: PathInputs,
    output_dir: PathLike | None = None,
    *,
    mode: str = "auto",
    content_type: str = "auto",
    scale: int = 2,
    max_dimension: int = 4096,
    tile_size: int = 256,
    output_format: str = "png",
    overwrite: bool = False,
) -> list[Path]:
    """Enhance images with Real-ESRGAN plus conservative structural checks.

    PNG is the default because writing an AI-enhanced image back through JPEG
    would immediately add new compression artefacts.  Alpha is resampled from
    the authentic source and never synthesized by the enhancement model.
    """

    try:
        import cv2  # type: ignore[import-not-found]
        import numpy as np  # type: ignore[import-not-found]
    except ImportError as exc:  # pragma: no cover - capability probe covers this
        raise RuntimeError("高清图像增强需要 OpenCV 与 NumPy") from exc
    from docuforge.runner import check_cancelled, report_progress
    from .image_enhancement import enhance_bgr

    sources = _paths(inputs)
    image_format, extension = _normalise_format(output_format)
    if image_format not in {"PNG", "JPEG", "WEBP"}:
        raise ValueError("高清增强输出格式仅支持 PNG、JPG 或 WebP")
    results: list[Path] = []
    for index, source_path in enumerate(sources, start=1):
        check_cancelled("已取消高清图像增强；已完成的图片会保留")
        report_progress(
            (index - 1) / max(1, len(sources)),
            f"高保真预处理与 AI 清晰增强：{source_path.name}",
            current_file=index,
            total_files=len(sources),
        )
        source, _source_format = _load_image(source_path)
        destination: Path | None = None
        reserved = False
        result_image: Image.Image | None = None
        try:
            rgba = source.convert("RGBA")
            rgb = np.asarray(rgba.convert("RGB"), dtype=np.uint8)
            bgr = cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)
            enhanced = enhance_bgr(
                bgr,
                mode=mode,
                content_type=content_type,
                scale=int(scale),
                max_dimension=int(max_dimension),
                tile_size=int(tile_size),
            )
            enhanced_rgb = cv2.cvtColor(enhanced.image, cv2.COLOR_BGR2RGB)
            result_image = Image.fromarray(enhanced_rgb, "RGB")
            if _contains_alpha(source) and image_format in {"PNG", "WEBP"}:
                alpha = rgba.getchannel("A").resize(
                    result_image.size, Image.Resampling.LANCZOS
                )
                with_alpha = result_image.convert("RGBA")
                with_alpha.putalpha(alpha)
                alpha.close()
                result_image.close()
                result_image = with_alpha
            rgba.close()

            destination, reserved = _destination(
                source_path,
                output_dir,
                "_高清增强",
                extension=extension,
                overwrite=overwrite,
            )
            _save_processed(
                result_image,
                source,
                destination,
                image_format,
                reserved,
                preserve_metadata=True,
            )
            results.append(destination)
            if mode in {"auto", "gpu_ai", "high_fidelity"} and not enhanced.ai_accepted:
                warnings.warn(
                    f"{source_path.name}：{enhanced.reason}；已自动使用可验证的安全结果。",
                    UserWarning,
                    stacklevel=2,
                )
            elif enhanced.fallback_blocks:
                warnings.warn(
                    f"{source_path.name}：AI 二检将 {enhanced.fallback_blocks}/"
                    f"{enhanced.total_blocks} 个高风险区域自动恢复为真实像素放大结果。",
                    UserWarning,
                    stacklevel=2,
                )
        except Exception:
            if destination is not None:
                _cleanup_reservation(destination, reserved)
            raise
        finally:
            if result_image is not None:
                result_image.close()
            source.close()
        report_progress(
            index / max(1, len(sources)),
            f"已完成高清增强：{source_path.name}",
            current_file=index,
            total_files=len(sources),
        )
    return results


def adjust_images(
    inputs: PathInputs,
    output_dir: PathLike | None = None,
    *,
    brightness: float = 1.0,
    contrast: float = 1.0,
    saturation: float = 1.0,
    overwrite: bool = False,
) -> list[Path]:
    """Adjust brightness, contrast and color saturation with Pillow factors."""

    factors = (float(brightness), float(contrast), float(saturation))
    if any(factor < 0 for factor in factors):
        raise ValueError("brightness, contrast and saturation must be non-negative")

    def transform(image: Image.Image, _: Path) -> Image.Image:
        working = image
        if image.mode not in {"RGB", "RGBA", "L", "LA"}:
            working = image.convert("RGBA" if _contains_alpha(image) else "RGB")
        try:
            result = ImageEnhance.Brightness(working).enhance(factors[0])
            contrasted = ImageEnhance.Contrast(result).enhance(factors[1])
            result.close()
            colored = ImageEnhance.Color(contrasted).enhance(factors[2])
            contrasted.close()
            return colored
        finally:
            if working is not image:
                working.close()

    return _map_images(inputs, output_dir, "_adjusted", transform, overwrite=overwrite)


def _with_original_alpha(source: Image.Image, result: Image.Image) -> Image.Image:
    if not _contains_alpha(source):
        return result
    rgba_source = source.convert("RGBA")
    alpha = rgba_source.getchannel("A")
    converted = result.convert("RGBA")
    converted.putalpha(alpha)
    alpha.close()
    rgba_source.close()
    if converted is not result:
        result.close()
    return converted


def apply_filter(
    inputs: PathInputs,
    filter_name: str,
    output_dir: PathLike | None = None,
    *,
    intensity: float = 1.0,
    overwrite: bool = False,
) -> list[Path]:
    """Apply grayscale, black/white, sepia, blur, sharpen, or classic filters."""

    name = filter_name.strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "grey": "grayscale",
        "gray": "grayscale",
        "blackwhite": "black_white",
        "bw": "black_white",
        "vintage": "sepia",
        "gaussian": "gaussian_blur",
        "edge": "edge_enhance",
    }
    name = aliases.get(name, name)
    intensity = float(intensity)
    if intensity < 0:
        raise ValueError("intensity must be non-negative")

    pillow_filters: dict[str, ImageFilter.Filter] = {
        "contour": ImageFilter.CONTOUR,
        "detail": ImageFilter.DETAIL,
        "emboss": ImageFilter.EMBOSS,
        "edge_enhance": ImageFilter.EDGE_ENHANCE_MORE,
        "find_edges": ImageFilter.FIND_EDGES,
        "smooth": ImageFilter.SMOOTH_MORE,
    }
    supported = {
        "grayscale",
        "black_white",
        "sepia",
        "blur",
        "gaussian_blur",
        "sharpen",
        *pillow_filters,
    }
    if name not in supported:
        raise ValueError(
            f"Unsupported filter {filter_name!r}; choose from {sorted(supported)}"
        )

    def transform(image: Image.Image, _: Path) -> Image.Image:
        if name == "grayscale":
            return _with_original_alpha(image, ImageOps.grayscale(image))
        if name == "black_white":
            grayscale = ImageOps.grayscale(image)
            threshold = max(0, min(255, round(128 * max(intensity, 0.01))))
            result = grayscale.point(
                lambda value: 255 if value >= threshold else 0, mode="1"
            )
            grayscale.close()
            return _with_original_alpha(image, result)
        if name == "sepia":
            grayscale = ImageOps.grayscale(image)
            result = ImageOps.colorize(grayscale, "#2b170b", "#f4d7a1")
            grayscale.close()
            if intensity != 1.0:
                original = image.convert("RGB")
                blended = Image.blend(original, result, min(1.0, intensity))
                original.close()
                result.close()
                result = blended
            return _with_original_alpha(image, result)

        working = image
        if image.mode not in {"RGB", "RGBA", "L", "LA"}:
            working = image.convert("RGBA" if _contains_alpha(image) else "RGB")
        try:
            if name in {"blur", "gaussian_blur"}:
                return working.filter(ImageFilter.GaussianBlur(radius=intensity))
            if name == "sharpen":
                return working.filter(
                    ImageFilter.UnsharpMask(
                        radius=2, percent=round(100 * intensity), threshold=3
                    )
                )
            return working.filter(pillow_filters[name])
        finally:
            if working is not image:
                working.close()

    return _map_images(inputs, output_dir, f"_{name}", transform, overwrite=overwrite)


def _font(font_size: int, font_path: PathLike | None) -> ImageFont.ImageFont:
    _positive_int(font_size, "font_size")
    if font_path is not None:
        path = Path(font_path).expanduser()
        if not path.is_file():
            raise FileNotFoundError(f"Font does not exist: {path}")
        return ImageFont.truetype(str(path), font_size)

    candidates = [
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
        Path("/System/Library/Fonts/PingFang.ttc"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]
    for candidate in candidates:
        if candidate.is_file():
            try:
                return ImageFont.truetype(str(candidate), font_size)
            except OSError:
                pass
    try:
        return ImageFont.load_default(size=font_size)
    except TypeError:  # Pillow before 10.1
        return ImageFont.load_default()


def _position(
    canvas_size: tuple[int, int],
    item_size: tuple[int, int],
    position: Position,
    margin: int,
) -> tuple[int, int]:
    if isinstance(position, tuple):
        if len(position) != 2:
            raise ValueError("position tuples must contain x and y")
        return int(position[0]), int(position[1])
    key = position.strip().lower().replace("_", "-").replace(" ", "-")
    aliases = {
        "top-left": (margin, margin),
        "top": ((canvas_size[0] - item_size[0]) // 2, margin),
        "top-right": (canvas_size[0] - item_size[0] - margin, margin),
        "left": (margin, (canvas_size[1] - item_size[1]) // 2),
        "center": (
            (canvas_size[0] - item_size[0]) // 2,
            (canvas_size[1] - item_size[1]) // 2,
        ),
        "right": (
            canvas_size[0] - item_size[0] - margin,
            (canvas_size[1] - item_size[1]) // 2,
        ),
        "bottom-left": (margin, canvas_size[1] - item_size[1] - margin),
        "bottom": (
            (canvas_size[0] - item_size[0]) // 2,
            canvas_size[1] - item_size[1] - margin,
        ),
        "bottom-right": (
            canvas_size[0] - item_size[0] - margin,
            canvas_size[1] - item_size[1] - margin,
        ),
    }
    try:
        return aliases[key]
    except KeyError as exc:
        raise ValueError(f"Unknown position: {position!r}") from exc


def add_text_watermark(
    inputs: PathInputs,
    text: str,
    output_dir: PathLike | None = None,
    *,
    position: Position = "bottom-right",
    font_size: int = 36,
    font_path: PathLike | None = None,
    color: ColorValue = "white",
    opacity: float = 0.7,
    margin: int = 20,
    stroke_width: int = 0,
    stroke_color: ColorValue = "black",
    overwrite: bool = False,
) -> list[Path]:
    """Add a Unicode text watermark at a named position or exact coordinates."""

    if not text:
        raise ValueError("watermark text cannot be empty")
    opacity = _unit_interval(opacity, "opacity")
    if margin < 0 or stroke_width < 0:
        raise ValueError("margin and stroke_width must be non-negative")
    selected_font = _font(font_size, font_path)
    fill = _rgba(color, opacity)
    stroke_fill = _rgba(stroke_color, opacity)

    def transform(image: Image.Image, _: Path) -> Image.Image:
        canvas = image.convert("RGBA")
        overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        bounds = draw.textbbox(
            (0, 0),
            text,
            font=selected_font,
            stroke_width=stroke_width,
        )
        text_size = (bounds[2] - bounds[0], bounds[3] - bounds[1])
        x, y = _position(canvas.size, text_size, position, margin)
        draw.text(
            (x - bounds[0], y - bounds[1]),
            text,
            font=selected_font,
            fill=fill,
            stroke_width=stroke_width,
            stroke_fill=stroke_fill,
        )
        canvas.alpha_composite(overlay)
        overlay.close()
        return canvas

    return _map_images(
        inputs, output_dir, "_watermarked", transform, overwrite=overwrite
    )


def _overlay_image(
    base: Image.Image,
    layer: Image.Image,
    *,
    position: Position,
    opacity: float,
    scale: float | None,
    margin: int,
) -> Image.Image:
    canvas = base.convert("RGBA")
    mark = layer.convert("RGBA")
    if scale is not None:
        if scale <= 0:
            mark.close()
            canvas.close()
            raise ValueError("scale must be greater than zero")
        target_width = max(1, round(canvas.width * scale))
        target_height = max(1, round(mark.height * target_width / mark.width))
        resized = mark.resize((target_width, target_height), Image.Resampling.LANCZOS)
        mark.close()
        mark = resized
    if opacity < 1.0:
        alpha = mark.getchannel("A").point(lambda value: round(value * opacity))
        mark.putalpha(alpha)
        alpha.close()
    location = _position(canvas.size, mark.size, position, margin)
    canvas.alpha_composite(mark, location)
    mark.close()
    return canvas


def _overlay_many(
    inputs: PathInputs,
    layer_path: PathLike,
    output_dir: PathLike | None,
    marker: str,
    *,
    position: Position,
    opacity: float,
    scale: float | None,
    margin: int,
    overwrite: bool,
) -> list[Path]:
    opacity = _unit_interval(opacity, "opacity")
    if margin < 0:
        raise ValueError("margin must be non-negative")
    path = Path(layer_path).expanduser()
    if not path.is_file():
        raise FileNotFoundError(f"Overlay image does not exist: {path}")
    layer, _ = _load_image(path)
    try:
        return _map_images(
            inputs,
            output_dir,
            marker,
            lambda image, _: _overlay_image(
                image,
                layer,
                position=position,
                opacity=opacity,
                scale=scale,
                margin=margin,
            ),
            overwrite=overwrite,
        )
    finally:
        layer.close()


def add_image_watermark(
    inputs: PathInputs,
    watermark_path: PathLike,
    output_dir: PathLike | None = None,
    *,
    position: Position = "bottom-right",
    opacity: float = 0.7,
    scale: float | None = 0.2,
    margin: int = 20,
    overwrite: bool = False,
) -> list[Path]:
    """Place a proportionally sized image watermark over every input."""

    return _overlay_many(
        inputs,
        watermark_path,
        output_dir,
        "_watermarked",
        position=position,
        opacity=opacity,
        scale=scale,
        margin=margin,
        overwrite=overwrite,
    )


def overlay_images(
    inputs: PathInputs,
    overlay_path: PathLike,
    output_dir: PathLike | None = None,
    *,
    position: Position = "center",
    opacity: float = 1.0,
    scale: float | None = None,
    margin: int = 0,
    overwrite: bool = False,
) -> list[Path]:
    """Composite a sticker/layer image over each base image."""

    return _overlay_many(
        inputs,
        overlay_path,
        output_dir,
        "_overlay",
        position=position,
        opacity=opacity,
        scale=scale,
        margin=margin,
        overwrite=overwrite,
    )


def _border_widths(
    border: int | tuple[int, int] | tuple[int, int, int, int]
) -> tuple[int, int, int, int]:
    values: tuple[int, int, int, int]
    if isinstance(border, int):
        values = (border, border, border, border)
    elif len(border) == 2:
        values = (int(border[0]), int(border[1]), int(border[0]), int(border[1]))
    elif len(border) == 4:
        values = (
            int(border[0]),
            int(border[1]),
            int(border[2]),
            int(border[3]),
        )
    else:
        raise ValueError(
            "border must be an int, (horizontal, vertical), or four values"
        )
    if any(value < 0 for value in values):
        raise ValueError("border widths must be non-negative")
    return values


def add_border(
    inputs: PathInputs,
    border: int | tuple[int, int] | tuple[int, int, int, int],
    output_dir: PathLike | None = None,
    *,
    color: ColorValue = "black",
    overwrite: bool = False,
) -> list[Path]:
    """Add solid borders as int, (horizontal, vertical), or L/T/R/B widths."""

    widths = _border_widths(border)

    def transform(image: Image.Image, _: Path) -> Image.Image:
        working = image
        if image.mode == "P":
            working = image.convert("RGBA" if _contains_alpha(image) else "RGB")
        try:
            fill = _color_for_mode(color, working.mode)
            return ImageOps.expand(working, border=widths, fill=fill)
        finally:
            if working is not image:
                working.close()

    return _map_images(inputs, output_dir, "_bordered", transform, overwrite=overwrite)


def mosaic_images(
    inputs: PathInputs,
    box: tuple[int, int, int, int] | None = None,
    output_dir: PathLike | None = None,
    *,
    block_size: int = 12,
    overwrite: bool = False,
) -> list[Path]:
    """Pixelate a rectangular sensitive region, or the whole image if omitted."""

    block_size = _positive_int(block_size, "block_size")

    def transform(image: Image.Image, _: Path) -> Image.Image:
        region_box = box or (0, 0, image.width, image.height)
        left, top, right, bottom = (int(value) for value in region_box)
        left, top = max(0, left), max(0, top)
        right, bottom = min(image.width, right), min(image.height, bottom)
        if right <= left or bottom <= top:
            raise ValueError("mosaic box does not overlap the image")
        result = image.copy()
        region = result.crop((left, top, right, bottom))
        small_size = (
            max(1, math.ceil(region.width / block_size)),
            max(1, math.ceil(region.height / block_size)),
        )
        small = region.resize(small_size, Image.Resampling.BOX)
        pixelated = small.resize(region.size, Image.Resampling.NEAREST)
        result.paste(pixelated, (left, top))
        region.close()
        small.close()
        pixelated.close()
        return result

    return _map_images(inputs, output_dir, "_mosaic", transform, overwrite=overwrite)


def stitch_images(
    inputs: PathInputs,
    output_path: PathLike | None = None,
    *,
    direction: str = "vertical",
    spacing: int = 0,
    background: ColorValue = "white",
    alignment: str = "center",
    overwrite: bool = False,
) -> list[Path]:
    """Stitch multiple images horizontally or vertically into one image."""

    source_paths = _paths(inputs)
    if len(source_paths) < 2:
        raise ValueError("stitch_images requires at least two input images")
    direction = direction.strip().lower()
    if direction not in {"horizontal", "vertical"}:
        raise ValueError("direction must be 'horizontal' or 'vertical'")
    alignment = alignment.strip().lower()
    if alignment not in {"start", "center", "end"}:
        raise ValueError("alignment must be 'start', 'center', or 'end'")
    if spacing < 0:
        raise ValueError("spacing must be non-negative")

    destination: Path | None = None
    reserved = False
    canvas: Image.Image | None = None
    prepared: Image.Image | None = None
    try:
        image_sizes: list[tuple[int, int]] = []
        color = _rgba(background)
        use_alpha = color[3] < 255
        for source in source_paths:
            image, _ = _load_image(source)
            try:
                image_sizes.append(image.size)
                use_alpha = use_alpha or _contains_alpha(image)
            finally:
                image.close()

        if output_path is None:
            requested = source_paths[0].with_name(
                f"{source_paths[0].stem}_stitched.png"
            )
        else:
            requested = Path(output_path).expanduser()
            if not requested.suffix:
                requested = requested.with_suffix(".png")
        destination, reserved = _reserve_path(requested, overwrite)
        image_format = _format_for_path(destination)

        mode = "RGBA" if use_alpha else "RGB"
        if direction == "horizontal":
            canvas_size = (
                sum(width for width, _height in image_sizes)
                + spacing * (len(image_sizes) - 1),
                max(height for _width, height in image_sizes),
            )
        else:
            canvas_size = (
                max(width for width, _height in image_sizes),
                sum(height for _width, height in image_sizes)
                + spacing * (len(image_sizes) - 1),
            )
        canvas = Image.new(mode, canvas_size, color if mode == "RGBA" else color[:3])
        cursor = 0
        for source in source_paths:
            image, _ = _load_image(source)
            converted = image.convert(mode)
            try:
                if direction == "horizontal":
                    if alignment == "start":
                        cross = 0
                    elif alignment == "center":
                        cross = (canvas.height - converted.height) // 2
                    else:
                        cross = canvas.height - converted.height
                    if mode == "RGBA":
                        canvas.alpha_composite(converted, (cursor, cross))
                    else:
                        canvas.paste(converted, (cursor, cross))
                    cursor += converted.width + spacing
                else:
                    if alignment == "start":
                        cross = 0
                    elif alignment == "center":
                        cross = (canvas.width - converted.width) // 2
                    else:
                        cross = canvas.width - converted.width
                    if mode == "RGBA":
                        canvas.alpha_composite(converted, (cross, cursor))
                    else:
                        canvas.paste(converted, (cross, cursor))
                    cursor += converted.height + spacing
            finally:
                converted.close()
                image.close()

        prepared = _prepare_for_format(canvas, image_format, background)
        _atomic_save(
            prepared,
            destination,
            image_format,
            reserved,
            save_options=_encoding_defaults(image_format),
        )
        return [destination]
    except Exception:
        if destination is not None:
            _cleanup_reservation(destination, reserved)
        raise
    finally:
        if prepared is not None and prepared is not canvas:
            prepared.close()
        if canvas is not None:
            canvas.close()


def _atomic_copy(source: Path, destination: Path, reserved: bool) -> None:
    descriptor, temporary_name = tempfile.mkstemp(
        dir=destination.parent,
        prefix=f".{destination.stem}-",
        suffix=".tmp",
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        shutil.copy2(source, temporary)
        os.replace(temporary, destination)
    except Exception:
        temporary.unlink(missing_ok=True)
        _cleanup_reservation(destination, reserved)
        raise


def batch_rename(
    inputs: PathInputs,
    pattern: str = "photo_{index:03d}",
    output_dir: PathLike | None = None,
    *,
    start: int = 1,
    move: bool = True,
    overwrite: bool = False,
) -> list[Path]:
    """Rename or copy files using ``index``, ``stem`` and ``suffix`` fields.

    If ``pattern`` has no extension, each source extension is preserved.  The
    default performs a move only after all destination copies are complete, so
    a copy failure leaves every original intact.  Use ``move=False`` to retain
    originals.
    """

    if not pattern:
        raise ValueError("rename pattern cannot be empty")
    if isinstance(start, bool) or not isinstance(start, int):
        raise ValueError("start must be an integer")
    sources = _paths(inputs)
    if len({path.resolve() for path in sources}) != len(sources):
        raise ValueError("input paths must be unique")

    reservations: list[tuple[Path, bool]] = []
    outputs: list[Path] = []
    copies_complete = False
    try:
        for offset, source in enumerate(sources):
            try:
                rendered = pattern.format(
                    index=start + offset,
                    stem=source.stem,
                    suffix=source.suffix,
                )
            except (KeyError, ValueError, IndexError) as exc:
                raise ValueError(f"Invalid rename pattern {pattern!r}: {exc}") from exc
            rendered_path = Path(rendered)
            if rendered_path.name != rendered or rendered in {".", ".."}:
                raise ValueError("rename pattern must produce a filename, not a path")
            if not rendered_path.suffix:
                rendered_path = rendered_path.with_suffix(source.suffix)
            directory = (
                Path(output_dir).expanduser()
                if output_dir is not None
                else source.parent
            )
            requested = directory / rendered_path.name
            destination, reserved = _reserve_path(requested, overwrite)
            if overwrite:
                destination_resolved = destination.resolve()
                if any(
                    destination_resolved == planned.resolve()
                    for planned, _ in reservations
                ):
                    raise ValueError("rename pattern produced duplicate destinations")
                for other_source in sources:
                    if (
                        destination_resolved == other_source.resolve()
                        and other_source != source
                    ):
                        raise ValueError(
                            "overwrite destinations cannot replace another input file"
                        )
            reservations.append((destination, reserved))

        completed: list[Path] = []
        try:
            for source, (destination, reserved) in zip(
                sources, reservations, strict=True
            ):
                if source.resolve() == destination.resolve():
                    completed.append(destination)
                    continue
                _atomic_copy(source, destination, reserved)
                completed.append(destination)
        except Exception:
            if not overwrite:
                for destination in completed:
                    if all(
                        destination.resolve() != source.resolve() for source in sources
                    ):
                        destination.unlink(missing_ok=True)
            raise

        copies_complete = True
        if move:
            for source, destination in zip(sources, completed, strict=True):
                if source.resolve() != destination.resolve():
                    source.unlink()
        outputs.extend(completed)
        return outputs
    except Exception:
        # Once all copies exist they are the recovery copies if deleting an
        # original fails.  Never remove those and risk data loss.
        if not copies_complete:
            for destination, reserved in reservations:
                _cleanup_reservation(destination, reserved)
        raise
