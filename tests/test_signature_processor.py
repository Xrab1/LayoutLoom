from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from docuforge.models import MissingEngineError, ValidationError
from docuforge.processors import signature as signature_processor


class _FakePyHanko:
    def __init__(self, *, fail_signing: bool = False, legacy: bool = False) -> None:
        self.calls: dict[str, object] = {}
        calls = self.calls

        class SimpleSigner:
            @classmethod
            def load_pkcs12(cls, pfx_file: str, passphrase: bytes | None = None):
                calls["certificate_path"] = pfx_file
                calls["passphrase"] = passphrase
                return object()

        class PdfSignatureMetadata:
            def __init__(self, **kwargs: object) -> None:
                calls["metadata"] = kwargs

        class SigFieldSpec:
            def __init__(self, **kwargs: object) -> None:
                calls["field_spec"] = kwargs

        class IncrementalPdfFileWriter:
            def __init__(self, stream: object) -> None:
                calls["input_stream"] = stream

        class HTTPTimeStamper:
            def __init__(self, url: str) -> None:
                calls["timestamp_url"] = url

        if legacy:

            class PdfSigner:
                def __init__(
                    self,
                    signature_meta: object,
                    signer: object,
                    timestamper: object | None = None,
                ) -> None:
                    calls["pdf_signer"] = (signature_meta, signer, timestamper)

                def sign_pdf(self, writer: object, output: object) -> None:
                    calls["writer"] = writer
                    output.write(b"%PDF-1.7\nlegacy-cryptographic-signature")

        else:

            class PdfSigner:
                def __init__(
                    self,
                    signature_meta: object,
                    signer: object,
                    *,
                    timestamper: object | None = None,
                    new_field_spec: object | None = None,
                ) -> None:
                    calls["pdf_signer"] = (
                        signature_meta,
                        signer,
                        timestamper,
                        new_field_spec,
                    )

                def sign_pdf(self, writer: object, output: object) -> None:
                    calls["writer"] = writer
                    if fail_signing:
                        raise RuntimeError("third-party failure containing top-secret")
                    output.write(b"%PDF-1.7\ncryptographic-signature")

        def append_signature_field(writer: object, sig_field_spec: object) -> None:
            calls["legacy_field"] = (writer, sig_field_spec)

        self.modules = {
            "pyhanko.sign.signers": SimpleNamespace(
                SimpleSigner=SimpleSigner,
                PdfSignatureMetadata=PdfSignatureMetadata,
                PdfSigner=PdfSigner,
            ),
            "pyhanko.sign.fields": SimpleNamespace(
                SigFieldSpec=SigFieldSpec,
                append_signature_field=append_signature_field,
            ),
            "pyhanko.pdf_utils.incremental_writer": SimpleNamespace(
                IncrementalPdfFileWriter=IncrementalPdfFileWriter
            ),
            "pyhanko.sign.timestamps": SimpleNamespace(HTTPTimeStamper=HTTPTimeStamper),
        }

    def import_module(self, name: str):
        return self.modules[name]


class SignatureProcessorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name) / "中文签名路径"
        self.root.mkdir()
        self.source = self.root / "待签合同.pdf"
        self.source.write_bytes(b"%PDF-1.7\ninput")
        self.certificate = self.root / "企业证书.p12"
        self.certificate.write_bytes(b"pkcs12-placeholder")

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_missing_pyhanko_has_actionable_error(self) -> None:
        with mock.patch.object(
            signature_processor.importlib,
            "import_module",
            side_effect=ModuleNotFoundError("pyhanko"),
        ):
            with self.assertRaisesRegex(MissingEngineError, "pyHanko"):
                signature_processor.sign_pdf(
                    self.source,
                    self.root / "已签.pdf",
                    self.certificate,
                    "certificate-password",
                )
        self.assertFalse((self.root / "已签.pdf").exists())

    def test_validation_happens_without_loading_optional_engine(self) -> None:
        with mock.patch.object(
            signature_processor.importlib, "import_module"
        ) as importer:
            with self.assertRaises(ValidationError):
                signature_processor.sign_pdf(
                    self.source,
                    self.root / "页码无效.pdf",
                    self.certificate,
                    "secret",
                    page=0,
                )
            with self.assertRaises(ValidationError):
                signature_processor.sign_pdf(
                    self.source,
                    self.root / "签名框无效.pdf",
                    self.certificate,
                    "secret",
                    box=(10, 10, 5, 20),
                )
            with self.assertRaises(ValidationError):
                signature_processor.sign_pdf(
                    self.source,
                    self.root / "时间戳无效.pdf",
                    self.certificate,
                    "secret",
                    timestamp_url="ftp://tsa.example.test",
                )
            with self.assertRaises(ValidationError):
                signature_processor.sign_pdf(
                    self.source,
                    self.root / "覆盖参数无效.pdf",
                    self.certificate,
                    "secret",
                    overwrite="yes",  # type: ignore[arg-type]
                )
        importer.assert_not_called()

    def test_unicode_paths_metadata_timestamp_and_page_conversion(self) -> None:
        fake = _FakePyHanko()
        output = self.root / "正式签署件.pdf"
        with mock.patch.object(
            signature_processor.importlib,
            "import_module",
            side_effect=fake.import_module,
        ):
            result = signature_processor.sign_pdf(
                self.source,
                output,
                self.certificate,
                "证书口令",
                field_name="审批签名",
                page=2,
                box=(20, 30, 240, 120),
                reason="合同审批",
                location="上海",
                contact_info="legal@example.test",
                timestamp_url="https://tsa.example.test/api",
            )

        self.assertEqual(result, [output.resolve()])
        self.assertEqual(output.read_bytes(), b"%PDF-1.7\ncryptographic-signature")
        self.assertEqual(
            fake.calls["certificate_path"], str(self.certificate.resolve())
        )
        self.assertEqual(fake.calls["passphrase"], "证书口令".encode("utf-8"))
        self.assertEqual(
            fake.calls["metadata"],
            {
                "field_name": "审批签名",
                "reason": "合同审批",
                "location": "上海",
                "contact_info": "legal@example.test",
            },
        )
        self.assertEqual(
            fake.calls["field_spec"],
            {
                "sig_field_name": "审批签名",
                "on_page": 1,
                "box": (20.0, 30.0, 240.0, 120.0),
            },
        )
        self.assertEqual(fake.calls["timestamp_url"], "https://tsa.example.test/api")

    def test_existing_output_is_not_overwritten_by_default(self) -> None:
        output = self.root / "已存在.pdf"
        output.write_bytes(b"keep-this")
        with mock.patch.object(
            signature_processor.importlib, "import_module"
        ) as importer:
            with self.assertRaises(FileExistsError):
                signature_processor.sign_pdf(
                    self.source, output, self.certificate, "secret"
                )
        self.assertEqual(output.read_bytes(), b"keep-this")
        importer.assert_not_called()

    def test_atomic_overwrite_preserves_old_file_when_signing_fails(self) -> None:
        fake = _FakePyHanko(fail_signing=True)
        output = self.root / "原有签署件.pdf"
        output.write_bytes(b"original")
        with mock.patch.object(
            signature_processor.importlib,
            "import_module",
            side_effect=fake.import_module,
        ):
            with self.assertRaises(signature_processor.PDFSignatureError) as captured:
                signature_processor.sign_pdf(
                    self.source,
                    output,
                    self.certificate,
                    "top-secret",
                    overwrite=True,
                )

        self.assertEqual(output.read_bytes(), b"original")
        self.assertNotIn("top-secret", str(captured.exception))
        self.assertTrue(captured.exception.__suppress_context__)
        self.assertEqual(list(self.root.glob(".原有签署件.*.tmp.pdf")), [])

    def test_legacy_pyhanko_field_api_is_supported(self) -> None:
        fake = _FakePyHanko(legacy=True)
        output = self.root / "兼容签署件.pdf"
        with mock.patch.object(
            signature_processor.importlib,
            "import_module",
            side_effect=fake.import_module,
        ):
            signature_processor.sign_pdf(
                self.source, output, self.certificate, b"password"
            )
        self.assertTrue(output.is_file())
        self.assertIn("legacy_field", fake.calls)

    def test_output_cannot_replace_input_pdf(self) -> None:
        with mock.patch.object(
            signature_processor.importlib, "import_module"
        ) as importer:
            with self.assertRaises(ValidationError):
                signature_processor.sign_pdf(
                    self.source,
                    self.source,
                    self.certificate,
                    "secret",
                    overwrite=True,
                )
        importer.assert_not_called()


if __name__ == "__main__":
    unittest.main()
