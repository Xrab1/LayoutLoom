from __future__ import annotations

import hashlib
from pathlib import Path

import cv2
import numpy as np
import pytest
from docx import Document

from docuforge.models import ValidationError
from docuforge.processors import video_slide_repair, video_slides


def _make_pptx(tmp_path: Path, images: list[np.ndarray], name: str = "source.pptx") -> Path:
    paths: list[Path] = []
    for index, image in enumerate(images, start=1):
        path = tmp_path / f"page-{index}.png"
        video_slides._write_png(path, image)
        paths.append(path)
    target = tmp_path / name
    video_slides._write_pptx(paths, target)
    return target


def test_local_temporal_repair_changes_only_the_user_box() -> None:
    height, width = 120, 220
    background = np.full((height, width, 3), (36, 48, 70), dtype=np.uint8)
    frames: list[np.ndarray] = []
    for index in range(9):
        frame = background.copy()
        x = 55 + index * 10
        cv2.rectangle(frame, (x, 25), (x + 13, 42), (245, 245, 245), -1)
        frames.append(frame)
    base = frames[2].copy()
    region = (20.0, 12.0, 58.0, 35.0)

    repaired = video_slide_repair._local_temporal_repair(base, frames, region)

    video_slide_repair._validate_outside_region_unchanged(base, repaired, region)
    assert np.array_equal(repaired[:, :40], base[:, :40])
    before_bright = int(np.count_nonzero(np.max(base[20:50, 40:175], axis=2) > 220))
    after_bright = int(np.count_nonzero(np.max(repaired[20:50, 40:175], axis=2) > 220))
    assert after_bright < before_bright


def test_repair_ppt_preserves_source_and_unselected_pages(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    first = np.full((120, 220, 3), (30, 60, 120), dtype=np.uint8)
    second = np.full((120, 220, 3), (120, 60, 30), dtype=np.uint8)
    source = _make_pptx(tmp_path, [first, second])
    source_digest = hashlib.sha256(source.read_bytes()).hexdigest()
    video = tmp_path / "lecture.mp4"
    video.write_bytes(b"test-video-placeholder")
    region = (10.0, 15.0, 25.0, 20.0)

    def fake_repair(
        base: np.ndarray,
        _video_path: Path,
        _timestamp: float,
        selected_region: tuple[float, float, float, float],
        *,
        method: str,
        colour: str,
    ) -> np.ndarray:
        assert method == "temporal"
        assert colour == "#FFFFFF"
        result = base.copy()
        left, top, right, bottom = video_slide_repair._pixel_region(
            selected_region, base.shape[1], base.shape[0]
        )
        result[top:bottom, left:right] = (0, 255, 0)
        return result

    monkeypatch.setattr(video_slide_repair, "repair_region_image", fake_repair)
    plan = video_slide_repair.make_plan(
        [
            {
                "kind": "repair_region",
                "page": 1,
                "timestamp": 12.5,
                "region": list(region),
                "method": "temporal",
                "colour": "#FFFFFF",
            }
        ]
    )

    output = video_slide_repair.repair_video_ppt(source, video, tmp_path, plan)

    assert output != source
    assert hashlib.sha256(source.read_bytes()).hexdigest() == source_digest
    assert video_slide_repair.ppt_slide_count(output) == 2
    repaired = video_slide_repair.read_ppt_slide_image(output, 1)
    untouched = video_slide_repair.read_ppt_slide_image(output, 2)
    video_slide_repair._validate_outside_region_unchanged(first, repaired, region)
    assert np.array_equal(untouched, second)


def test_missing_page_is_inserted_at_the_confirmed_position(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    first = np.full((90, 160, 3), (20, 30, 40), dtype=np.uint8)
    second = np.full((90, 160, 3), (80, 90, 100), dtype=np.uint8)
    inserted = np.full((90, 160, 3), (150, 160, 170), dtype=np.uint8)
    source = _make_pptx(tmp_path, [first, second])
    video = tmp_path / "lecture.mp4"
    video.write_bytes(b"test-video-placeholder")

    monkeypatch.setattr(
        video_slide_repair,
        "find_stable_aligned_frame",
        lambda *_args, **_kwargs: video_slide_repair.StableFrame(8.4, inserted.copy(), 7),
    )
    plan = video_slide_repair.make_plan(
        [
            {
                "kind": "insert_page",
                "position": 2,
                "timestamp": 8.0,
                "method": "temporal",
                "colour": "#FFFFFF",
            }
        ]
    )

    output = video_slide_repair.repair_video_ppt(source, video, tmp_path, plan)

    assert video_slide_repair.ppt_slide_count(output) == 3
    assert np.array_equal(video_slide_repair.read_ppt_slide_image(output, 1), first)
    assert np.array_equal(video_slide_repair.read_ppt_slide_image(output, 2), inserted)
    assert np.array_equal(video_slide_repair.read_ppt_slide_image(output, 3), second)


def test_manual_best_frame_replaces_page_and_cleans_multiple_watermarks(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    original = np.full((100, 180, 3), (30, 40, 50), dtype=np.uint8)
    untouched = np.full((100, 180, 3), (70, 80, 90), dtype=np.uint8)
    selected = np.full((100, 180, 3), (110, 120, 130), dtype=np.uint8)
    cv2.rectangle(selected, (18, 10), (35, 24), (245, 245, 245), -1)
    cv2.rectangle(selected, (125, 70), (150, 88), (245, 245, 245), -1)
    source = _make_pptx(tmp_path, [original, untouched])
    video = tmp_path / "lecture.mp4"
    video.write_bytes(b"test-video-placeholder")
    regions = [(8.0, 7.0, 15.0, 20.0), (67.0, 66.0, 20.0, 26.0)]
    monkeypatch.setattr(
        video_slide_repair,
        "read_aligned_video_frame",
        lambda *_args, **_kwargs: selected.copy(),
    )
    plan = video_slide_repair.make_plan(
        [
            {
                "kind": "replace_page_frame",
                "page": 1,
                "timestamp": 22.4,
                "regions": [list(region) for region in regions],
                "method": "color",
                "colour": "#102030",
            }
        ]
    )

    output = video_slide_repair.repair_video_ppt(source, video, tmp_path, plan)

    rebuilt = video_slide_repair.read_ppt_slide_image(output, 1)
    expected = video_slide_repair.repair_regions_on_selected_frame(
        selected,
        video,
        22.4,
        regions,
        method="color",
        colour="#102030",
    )
    assert np.array_equal(rebuilt, expected)
    video_slide_repair._validate_outside_regions_unchanged(selected, rebuilt, regions)
    assert np.array_equal(
        video_slide_repair.read_ppt_slide_image(output, 2), untouched
    )


def test_workbench_page_preview_applies_confirmed_actions_in_order(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    original = np.full((80, 140, 3), (20, 30, 40), dtype=np.uint8)
    selected = np.full((80, 140, 3), (100, 110, 120), dtype=np.uint8)
    source = _make_pptx(tmp_path, [original])
    video = tmp_path / "lecture.mp4"
    video.write_bytes(b"test-video-placeholder")
    region = (10.0, 10.0, 20.0, 20.0)

    monkeypatch.setattr(
        video_slide_repair,
        "read_aligned_video_frame",
        lambda *_args, **_kwargs: selected.copy(),
    )

    def fake_repair(
        base: np.ndarray,
        _video_path: Path,
        _timestamp: float,
        selected_region: tuple[float, float, float, float],
        *,
        method: str,
        colour: str,
    ) -> np.ndarray:
        assert method == "color"
        assert colour == "#00FF00"
        result = base.copy()
        left, top, right, bottom = video_slide_repair._pixel_region(
            selected_region, base.shape[1], base.shape[0]
        )
        result[top:bottom, left:right] = (0, 255, 0)
        return result

    monkeypatch.setattr(video_slide_repair, "repair_region_image", fake_repair)
    actions = [
        {
            "kind": "replace_page_frame",
            "page": 1,
            "timestamp": 4.2,
            "regions": [],
            "method": "temporal",
            "colour": "#FFFFFF",
        },
        {
            "kind": "repair_region",
            "page": 1,
            "timestamp": 4.4,
            "region": list(region),
            "method": "color",
            "colour": "#00FF00",
        },
    ]

    rendered = video_slide_repair.render_page_after_actions(
        source, video, 1, actions
    )

    expected = fake_repair(
        selected,
        video,
        4.4,
        region,
        method="color",
        colour="#00FF00",
    )
    assert np.array_equal(rendered, expected)


def test_repair_plan_rejects_invalid_page_and_region() -> None:
    with pytest.raises(ValidationError, match="页码超出范围"):
        video_slide_repair.normalize_plan(
            {
                "schema": video_slide_repair.PLAN_SCHEMA,
                "actions": [
                    {
                        "kind": "repair_region",
                        "page": 3,
                        "timestamp": 1.0,
                        "region": [0, 0, 10, 10],
                        "method": "temporal",
                    }
                ],
            },
            2,
        )
    with pytest.raises(ValidationError, match="0–100%"):
        video_slide_repair.normalize_plan(
            {
                "schema": video_slide_repair.PLAN_SCHEMA,
                "actions": [
                    {
                        "kind": "insert_page",
                        "position": 2,
                        "timestamp": 1.0,
                        "region": [95, 0, 10, 10],
                        "method": "background",
                    }
                ],
            },
            2,
        )


def test_companion_report_automatically_locates_the_selected_page(
    tmp_path: Path,
) -> None:
    source = _make_pptx(
        tmp_path,
        [np.full((60, 100, 3), 80, dtype=np.uint8)],
        name="课程_高清幻灯片.pptx",
    )
    report = Document()
    table = report.add_table(rows=2, cols=3)
    table.rows[0].cells[0].text = "PPT 页码"
    table.rows[0].cells[1].text = "实际采用帧"
    table.rows[0].cells[2].text = "状态"
    table.rows[1].cells[0].text = "1"
    table.rows[1].cells[1].text = "00:03:12.345"
    table.rows[1].cells[2].text = "通过"
    report.save(tmp_path / "课程_提取报告.docx")

    timestamp = video_slide_repair.companion_report_timestamp(source, 1)

    assert timestamp == pytest.approx(192.345)
