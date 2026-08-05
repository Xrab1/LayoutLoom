"""Certificate-backed PDF signatures using the optional :mod:`pyhanko` engine.

The adapter deliberately imports pyHanko only when :func:`sign_pdf` is called,
so the rest of LayoutLoom remains usable without the optional signing stack.
Private-key material and certificate passwords are never logged or included in
user-facing exception messages.
"""

from __future__ import annotations

import importlib
import inspect
import math
import os
import tempfile
from contextlib import contextmanager
from pathlib import Path
from types import ModuleType
from typing import Any, Iterator, NamedTuple, Sequence
from urllib.parse import urlsplit

from ..models import DocuForgeError, MissingEngineError, ValidationError


PathLike = str | os.PathLike[str]


class PDFSignatureError(DocuForgeError):
    """Raised when a certificate could not be loaded or a PDF could not be signed."""


class _PyHankoAPI(NamedTuple):
    signers: ModuleType
    fields: ModuleType
    incremental_writer: ModuleType
    timestamps: ModuleType | None


__all__ = ["PDFSignatureError", "sign_pdf"]


def _existing_file(value: PathLike, label: str, suffixes: set[str]) -> Path:
    try:
        path = Path(value).expanduser().resolve()
    except (TypeError, ValueError, OSError) as exc:
        raise ValidationError(f"{label}路径无效") from exc
    if not path.is_file():
        raise ValidationError(f"{label}不存在或不是文件：{path}")
    if path.suffix.lower() not in suffixes:
        allowed = "、".join(sorted(suffix.lstrip(".").upper() for suffix in suffixes))
        raise ValidationError(f"{label}格式无效；仅支持 {allowed}")
    return path


def _output_file(value: PathLike) -> Path:
    try:
        path = Path(value).expanduser()
        path = (
            (Path.cwd() / path).resolve() if not path.is_absolute() else path.resolve()
        )
    except (TypeError, ValueError, OSError) as exc:
        raise ValidationError("输出路径无效") from exc
    if path.suffix.lower() != ".pdf":
        raise ValidationError("数字签名输出文件必须使用 .pdf 扩展名")
    if path.exists() and path.is_dir():
        raise ValidationError(f"输出位置是文件夹而不是文件：{path}")
    if path.parent.exists() and not path.parent.is_dir():
        raise ValidationError(f"输出文件的父路径不是文件夹：{path.parent}")
    return path


def _same_path(left: Path, right: Path) -> bool:
    try:
        return os.path.samefile(left, right)
    except (FileNotFoundError, OSError):
        return os.path.normcase(str(left.resolve())) == os.path.normcase(
            str(right.resolve())
        )


def _normalise_text(value: str | None, label: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise ValidationError(f"{label}必须是文本或 None")
    if "\x00" in value:
        raise ValidationError(f"{label}不能包含空字符")
    stripped = value.strip()
    return stripped or None


def _normalise_box(box: Sequence[float]) -> tuple[float, float, float, float]:
    if isinstance(box, (str, bytes, bytearray)):
        raise ValidationError("签名框 box 必须包含四个数字")
    try:
        values = tuple(box)
    except TypeError as exc:
        raise ValidationError("签名框 box 必须包含四个数字") from exc
    if len(values) != 4:
        raise ValidationError("签名框 box 必须按 (左, 下, 右, 上) 提供四个数字")

    coordinates: list[float] = []
    for value in values:
        if isinstance(value, bool):
            raise ValidationError("签名框坐标必须是有限数字")
        try:
            coordinate = float(value)
        except (TypeError, ValueError, OverflowError) as exc:
            raise ValidationError("签名框坐标必须是有限数字") from exc
        if not math.isfinite(coordinate):
            raise ValidationError("签名框坐标必须是有限数字")
        coordinates.append(coordinate)

    left, bottom, right, top = coordinates
    if right <= left or top <= bottom:
        raise ValidationError("签名框必须满足右坐标大于左坐标、上坐标大于下坐标")
    return left, bottom, right, top


def _normalise_timestamp_url(value: str | None) -> str | None:
    url = _normalise_text(value, "时间戳服务地址")
    if url is None:
        return None
    parsed = urlsplit(url)
    if parsed.scheme.lower() not in {"http", "https"} or not parsed.netloc:
        raise ValidationError("时间戳服务地址必须是有效的 HTTP 或 HTTPS URL")
    return url


def _certificate_passphrase(value: str | bytes | bytearray | None) -> bytes | None:
    if value is None:
        return None
    if isinstance(value, str):
        return value.encode("utf-8")
    if isinstance(value, (bytes, bytearray)):
        return bytes(value)
    raise ValidationError("证书密码必须是文本、字节或 None")


def _load_pyhanko(timestamp_required: bool) -> _PyHankoAPI:
    try:
        signers = importlib.import_module("pyhanko.sign.signers")
        fields = importlib.import_module("pyhanko.sign.fields")
        incremental_writer = importlib.import_module(
            "pyhanko.pdf_utils.incremental_writer"
        )
    except (ImportError, ModuleNotFoundError):
        raise MissingEngineError(
            "PDF 证书数字签名需要可选组件 pyHanko；请先安装：pip install pyHanko"
        ) from None

    required = (
        (signers, "SimpleSigner"),
        (signers, "PdfSignatureMetadata"),
        (signers, "PdfSigner"),
        (fields, "SigFieldSpec"),
        (incremental_writer, "IncrementalPdfFileWriter"),
    )
    if any(not hasattr(module, attribute) for module, attribute in required):
        raise MissingEngineError("已安装的 pyHanko 版本缺少数字签名 API；请升级：pip install -U pyHanko")

    timestamps: ModuleType | None = None
    if timestamp_required:
        try:
            timestamps = importlib.import_module("pyhanko.sign.timestamps")
            if not hasattr(timestamps, "HTTPTimeStamper"):
                timestamps = importlib.import_module(
                    "pyhanko.sign.timestamps.requests_client"
                )
        except (ImportError, ModuleNotFoundError):
            raise MissingEngineError(
                "当前 pyHanko 安装不支持 HTTP 时间戳；请升级 pyHanko 并安装其网络依赖"
            ) from None
        if not hasattr(timestamps, "HTTPTimeStamper"):
            raise MissingEngineError("当前 pyHanko 版本缺少 HTTPTimeStamper；请升级 pyHanko")

    return _PyHankoAPI(signers, fields, incremental_writer, timestamps)


def _supports_keyword(callable_object: Any, keyword: str) -> bool | None:
    """Report keyword support, or ``None`` when a signature cannot be inspected."""

    try:
        parameters = inspect.signature(callable_object).parameters.values()
    except (TypeError, ValueError):
        return None
    for parameter in parameters:
        if parameter.kind is inspect.Parameter.VAR_KEYWORD:
            return True
        if (
            parameter.name == keyword
            and parameter.kind is not inspect.Parameter.POSITIONAL_ONLY
        ):
            return True
    return False


def _load_certificate_signer(
    simple_signer: Any, certificate_path: Path, passphrase: bytes | None
) -> Any:
    file_loader = getattr(simple_signer, "load_pkcs12", None)
    data_loader = getattr(simple_signer, "load_pkcs12_data", None)
    try:
        if callable(file_loader):
            pfx_keyword = _supports_keyword(file_loader, "pfx_file")
            password_keyword = _supports_keyword(file_loader, "passphrase")
            kwargs: dict[str, Any] = {}
            if password_keyword is not False:
                kwargs["passphrase"] = passphrase
            if pfx_keyword is not False:
                signer = file_loader(pfx_file=str(certificate_path), **kwargs)
            else:
                signer = file_loader(str(certificate_path), **kwargs)
        elif callable(data_loader):
            password_keyword = _supports_keyword(data_loader, "passphrase")
            kwargs = {}
            if password_keyword is not False:
                kwargs["passphrase"] = passphrase
            signer = data_loader(certificate_path.read_bytes(), **kwargs)
        else:
            raise MissingEngineError("当前 pyHanko 版本无法加载 PKCS#12 证书；请升级 pyHanko")
    except MissingEngineError:
        raise
    except Exception:
        # Do not chain certificate-loader errors: a third-party exception could
        # contain sensitive passphrase or private-key diagnostics.
        raise PDFSignatureError("无法加载 PKCS#12 证书；请检查证书文件、密码和证书有效性") from None
    finally:
        passphrase = None

    if signer is None:
        raise PDFSignatureError("无法加载 PKCS#12 证书；请检查证书文件、密码和证书有效性")
    return signer


def _metadata(
    signers: ModuleType, field_name: str, values: dict[str, str | None]
) -> Any:
    metadata_type = signers.PdfSignatureMetadata
    kwargs: dict[str, Any] = {"field_name": field_name}
    kwargs.update({key: value for key, value in values.items() if value is not None})
    try:
        return metadata_type(**kwargs)
    except TypeError:
        # A few older pyHanko releases did not expose all descriptive fields.
        try:
            accepted = {
                key: value
                for key, value in kwargs.items()
                if _supports_keyword(metadata_type, key) is not False
            }
            return metadata_type(**accepted)
        except Exception:
            raise MissingEngineError("已安装的 pyHanko 版本与当前签名适配器不兼容；请升级 pyHanko") from None
    except Exception:
        raise PDFSignatureError("无法创建 PDF 数字签名元数据") from None


def _append_legacy_field(fields: ModuleType, writer: Any, field_spec: Any) -> None:
    append_field = getattr(fields, "append_signature_field", None)
    if not callable(append_field):
        raise MissingEngineError("当前 pyHanko 版本不支持在指定页面创建签名域；请升级 pyHanko")
    try:
        if _supports_keyword(append_field, "sig_field_spec") is not False:
            append_field(writer, sig_field_spec=field_spec)
        else:
            append_field(writer, field_spec)
    except Exception:
        raise PDFSignatureError("无法在指定页面创建 PDF 签名域") from None


def _pdf_signer(
    api: _PyHankoAPI,
    writer: Any,
    signature_metadata: Any,
    certificate_signer: Any,
    field_spec: Any,
    timestamper: Any,
) -> Any:
    pdf_signer_type = api.signers.PdfSigner
    kwargs: dict[str, Any] = {}
    timestamp_support = _supports_keyword(pdf_signer_type, "timestamper")
    field_support = _supports_keyword(pdf_signer_type, "new_field_spec")

    if timestamper is not None:
        if timestamp_support is False:
            raise MissingEngineError("当前 pyHanko 版本不支持为 PDF 签名附加时间戳；请升级 pyHanko")
        kwargs["timestamper"] = timestamper

    if field_support is False:
        _append_legacy_field(api.fields, writer, field_spec)
    else:
        kwargs["new_field_spec"] = field_spec

    try:
        return pdf_signer_type(signature_metadata, certificate_signer, **kwargs)
    except Exception:
        raise PDFSignatureError("无法初始化 PDF 数字签名器") from None


def _sign_to_stream(pdf_signer: Any, writer: Any, output_stream: Any) -> None:
    sign_method = getattr(pdf_signer, "sign_pdf", None)
    if not callable(sign_method):
        raise MissingEngineError("当前 pyHanko 版本缺少 PdfSigner.sign_pdf；请升级 pyHanko")

    try:
        if _supports_keyword(sign_method, "output") is not False:
            sign_method(writer, output=output_stream)
        elif _supports_keyword(sign_method, "output_stream"):
            sign_method(writer, output_stream=output_stream)
        else:
            raise MissingEngineError("当前 pyHanko 输出 API 不受支持；请升级 pyHanko")
    except MissingEngineError:
        raise
    except Exception:
        # Suppress third-party exception context so certificate passwords or
        # private-key diagnostics can never be echoed in a traceback.
        raise PDFSignatureError("PDF 数字签名失败；请检查输入 PDF、证书权限和时间戳服务") from None


@contextmanager
def _atomic_output(target: Path, overwrite: bool) -> Iterator[Path]:
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and not overwrite:
        raise FileExistsError(f"输出文件已存在：{target}；如需替换请设置 overwrite=True")

    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{target.stem}.", suffix=".tmp.pdf", dir=target.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        yield temporary
        if not temporary.is_file() or temporary.stat().st_size == 0:
            raise PDFSignatureError("签名引擎没有生成有效的 PDF 输出")
        if target.exists() and not overwrite:
            raise FileExistsError(f"输出文件已存在：{target}；如需替换请设置 overwrite=True")
        os.replace(temporary, target)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def sign_pdf(
    input_pdf: PathLike,
    output_path: PathLike,
    certificate_path: PathLike,
    certificate_password: str | bytes | bytearray | None,
    *,
    field_name: str = "Signature1",
    page: int = 1,
    box: Sequence[float] = (36, 36, 220, 100),
    reason: str | None = None,
    location: str | None = None,
    contact_info: str | None = None,
    timestamp_url: str | None = None,
    overwrite: bool = False,
) -> list[Path]:
    """Digitally sign a PDF with a PKCS#12/PFX/P12 certificate.

    ``page`` is one-based.  ``box`` uses PDF point coordinates in
    ``(left, bottom, right, top)`` order.  The signature is cryptographic and
    is produced exclusively by pyHanko; this function never substitutes a
    visual-only mark for a digital signature.
    """

    source = _existing_file(input_pdf, "输入 PDF", {".pdf"})
    certificate = _existing_file(
        certificate_path, "PKCS#12 证书", {".p12", ".pfx", ".pkcs12"}
    )
    target = _output_file(output_path)

    if _same_path(source, target):
        raise ValidationError("输出文件不能覆盖输入 PDF；请选择新的输出路径")
    if _same_path(certificate, target):
        raise ValidationError("输出文件不能覆盖证书文件；请选择新的输出路径")
    if not isinstance(overwrite, bool):
        raise ValidationError("overwrite 必须是布尔值")
    if target.exists() and not overwrite:
        raise FileExistsError(f"输出文件已存在：{target}；如需替换请设置 overwrite=True")

    if not isinstance(field_name, str) or not field_name.strip():
        raise ValidationError("签名域名称 field_name 不能为空")
    field_name = field_name.strip()
    if "\x00" in field_name:
        raise ValidationError("签名域名称不能包含空字符")
    if isinstance(page, bool) or not isinstance(page, int) or page < 1:
        raise ValidationError("签名页码 page 必须是从 1 开始的整数")

    normalised_box = _normalise_box(box)
    metadata_values = {
        "reason": _normalise_text(reason, "签名原因"),
        "location": _normalise_text(location, "签名地点"),
        "contact_info": _normalise_text(contact_info, "联系信息"),
    }
    normalised_timestamp_url = _normalise_timestamp_url(timestamp_url)
    passphrase = _certificate_passphrase(certificate_password)

    api = _load_pyhanko(normalised_timestamp_url is not None)
    certificate_signer = _load_certificate_signer(
        api.signers.SimpleSigner, certificate, passphrase
    )
    passphrase = None

    signature_metadata = _metadata(api.signers, field_name, metadata_values)
    try:
        field_spec = api.fields.SigFieldSpec(
            sig_field_name=field_name,
            on_page=page - 1,
            box=normalised_box,
        )
    except Exception:
        raise PDFSignatureError("无法创建 PDF 签名域参数") from None

    timestamper = None
    if normalised_timestamp_url is not None:
        assert api.timestamps is not None
        try:
            timestamper = api.timestamps.HTTPTimeStamper(normalised_timestamp_url)
        except Exception:
            raise PDFSignatureError("无法初始化 HTTP 时间戳服务") from None

    writer_type = api.incremental_writer.IncrementalPdfFileWriter
    with _atomic_output(target, overwrite) as temporary:
        try:
            with source.open("rb") as input_stream, temporary.open(
                "w+b"
            ) as output_stream:
                writer = writer_type(input_stream)
                pdf_signer = _pdf_signer(
                    api,
                    writer,
                    signature_metadata,
                    certificate_signer,
                    field_spec,
                    timestamper,
                )
                _sign_to_stream(pdf_signer, writer, output_stream)
                output_stream.flush()
                os.fsync(output_stream.fileno())
        except (MissingEngineError, PDFSignatureError):
            raise
        except Exception:
            raise PDFSignatureError("PDF 数字签名失败；请确认输入文件是可读取的有效 PDF") from None

    return [target]
