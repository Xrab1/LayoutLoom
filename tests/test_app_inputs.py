from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from types import SimpleNamespace

from docuforge.app import (
    DND_COPY,
    DocuForgeApp,
    canonical_path_key,
    collect_input_files,
    natural_path_key,
    normalize_input_path,
    parse_drop_payload,
)


def test_parse_drop_payload_preserves_unicode_and_spaces(tmp_path: Path) -> None:
    first = tmp_path / "论文 文件.pdf"
    second = tmp_path / "资料" / "表格 02.xlsx"
    interpreter = tk.Tcl()
    payload = f"{{{first}}} {{{second}}}"

    assert parse_drop_payload(payload, interpreter.splitlist) == (
        str(first),
        str(second),
    )


def test_parse_drop_payload_handles_literal_braces_and_unc_paths() -> None:
    interpreter = tk.Tcl()
    interpreter.tk.wantobjects(False)
    paths = (
        r"C:\资料\a {draft}.pdf",
        r"C:\资料\中文 空格.pdf",
        r"\\server\share\report.pdf",
    )
    payload = interpreter.call("list", *paths)

    assert isinstance(payload, str)
    assert parse_drop_payload(payload, interpreter.splitlist) == paths


def test_drop_enter_always_advertises_copy_even_if_source_requests_move() -> None:
    class DropStub:
        @staticmethod
        def _set_drop_hint(*_args, **_kwargs) -> None:
            return None

    result = DocuForgeApp._on_drop_enter(
        DropStub(),
        SimpleNamespace(action="move"),
    )

    assert result == str(DND_COPY)


def test_normalize_input_path_accepts_file_uri_and_has_stable_key(
    tmp_path: Path,
) -> None:
    source = tmp_path / "带 空格.pdf"
    source.write_bytes(b"pdf")

    normalized = normalize_input_path(source.as_uri())

    assert normalized == source.resolve()
    assert canonical_path_key(normalized) == canonical_path_key(source.resolve())


def test_normalize_input_path_is_lexical_and_does_not_call_resolve(
    tmp_path: Path,
    monkeypatch,
) -> None:
    raw = tmp_path / "missing" / ".." / "report.pdf"
    expected = Path(os.path.abspath(os.path.normpath(str(raw))))

    def fail_resolve(*_args, **_kwargs):
        raise AssertionError("normalization must not probe path components")

    monkeypatch.setattr(Path, "resolve", fail_resolve)

    assert normalize_input_path(raw) == expected


def test_collect_input_files_recurses_filters_sorts_and_deduplicates(
    tmp_path: Path,
) -> None:
    folder = tmp_path / "批量目录"
    nested = folder / "子目录"
    nested.mkdir(parents=True)
    first = folder / "报告1.pdf"
    second = folder / "报告2.PDF"
    tenth = nested / "报告10.pdf"
    unsupported = nested / "说明.txt"
    for path in (first, second, tenth, unsupported):
        path.write_bytes(path.name.encode("utf-8"))
    missing = tmp_path / "不存在.pdf"

    result = collect_input_files(
        (folder, first, missing),
        allowed_extensions=(".pdf",),
    )

    assert result.files == tuple(
        sorted(
            (first.resolve(), second.resolve(), tenth.resolve()), key=natural_path_key
        )
    )
    assert result.scanned_directories == 1
    assert result.duplicate_files == 1
    assert result.unsupported_files == 1
    assert result.missing_paths == (str(missing.resolve()),)
    assert result.errors == ()


def test_collect_input_files_prefilters_unsupported_directory_entries(
    tmp_path: Path,
    monkeypatch,
) -> None:
    folder = tmp_path / "mixed"
    folder.mkdir()
    supported = folder / "report.pdf"
    unsupported = folder / "large-video.mp4"
    supported.write_bytes(b"pdf")
    unsupported.write_bytes(b"video")
    original_is_file = Path.is_file
    checked_paths: list[Path] = []

    def tracking_is_file(path: Path) -> bool:
        checked_paths.append(path)
        return original_is_file(path)

    monkeypatch.setattr(Path, "is_file", tracking_is_file)

    result = collect_input_files((folder,), allowed_extensions=(".pdf",))

    assert result.files == (supported.resolve(),)
    assert result.unsupported_files == 1
    assert supported in checked_paths
    assert unsupported not in checked_paths


def test_collect_input_files_respects_existing_list_and_multiple_folder_drops(
    tmp_path: Path,
) -> None:
    first_folder = tmp_path / "first"
    second_folder = tmp_path / "second"
    first_folder.mkdir()
    second_folder.mkdir()
    existing = first_folder / "same.pdf"
    duplicate = second_folder / "same-copy.pdf"
    new_file = second_folder / "new.pdf"
    existing.write_bytes(b"one")
    duplicate.write_bytes(b"two")
    new_file.write_bytes(b"three")

    result = collect_input_files(
        (first_folder, second_folder, second_folder),
        allowed_extensions=("pdf",),
        existing_paths=(existing,),
    )

    assert result.files == tuple(
        sorted((duplicate.resolve(), new_file.resolve()), key=natural_path_key)
    )
    assert result.scanned_directories == 3
    assert result.duplicate_files == 3


def test_collect_input_files_handles_a_large_flat_batch_in_natural_order(
    tmp_path: Path,
) -> None:
    folder = tmp_path / "large-batch"
    folder.mkdir()
    expected = []
    for index in range(512, 0, -1):
        path = folder / f"document-{index}.pdf"
        path.touch()
        expected.append(path)

    result = collect_input_files((folder,), allowed_extensions=(".pdf",))

    assert len(result.files) == 512
    assert result.files == tuple(sorted(expected, key=natural_path_key))
