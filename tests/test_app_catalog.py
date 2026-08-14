from __future__ import annotations

from collections import Counter
from dataclasses import fields, is_dataclass
import math
from pathlib import Path
import re
from types import SimpleNamespace

import docuforge.app as app_module
from docuforge.app import (
    CATALOG_STRUCTURE,
    OFFICE_COMPATIBILITY_NOTICE,
    OFFICE_ENGINE_VALUES,
    UI_THEME_LABELS,
    UI_THEME_PALETTES,
    UI_THEME_PALETTE_KEYS,
    WPS_COMPATIBILITY_NOTICE,
    DocuForgeApp,
    RoundedCard,
    append_color_value,
    build_live_image_preview,
    background_watermark_plan,
    catalog_sidebar_width,
    canvas_point_to_source,
    canvas_selection_to_percent_region,
    catalog_order_key,
    click_particle_frame_plan,
    click_particle_specs,
    color_dialog_initial,
    composite_hex_colour,
    ease_out_back,
    ease_out_cubic,
    fitted_dialog_size,
    fitted_media_rect,
    format_percent_region,
    initial_window_size,
    interpolate_hex_colour,
    mousewheel_scroll_units,
    motion_button_text,
    motion_effect_timing,
    motion_frame_delay,
    normalize_motion_mode,
    normalize_particle_effects_enabled,
    normalize_ui_theme,
    operation_description_text,
    operation_display_name,
    operation_catalog_path,
    office_engine_button_text,
    office_engine_parameter_spec,
    parameter_help_text,
    particle_effect_button_text,
    progress_percent,
    progress_status_text,
    repair_preview_layout,
    q_bounce_transition_plan,
    responsive_layout_mode,
    responsive_wraplength,
    result_dialog_geometry,
    robust_frame_rgb,
    short_window_layout,
    smooth_progress_step,
    load_ui_preferences,
    save_ui_preferences,
    result_dialog_entrance_plan,
    supports_live_image_preview,
    task_result_presentation,
    unavailable_operation_presentation,
    windows_redraw_flags,
)
from docuforge.models import Capability, Operation, ParameterSpec, TaskFailure, TaskResult
from docuforge.registry import get_operations


def test_office_compatibility_notice_is_balanced_and_actionable() -> None:
    assert WPS_COMPATIBILITY_NOTICE == OFFICE_COMPATIBILITY_NOTICE
    assert "WPS Office 与 Microsoft Office" in OFFICE_COMPATIBILITY_NOTICE
    assert "独立识别和定向适配" in OFFICE_COMPATIBILITY_NOTICE
    assert "WPS → Microsoft Office → LibreOffice" in OFFICE_COMPATIBILITY_NOTICE
    assert "锁定单一引擎" in OFFICE_COMPATIBILITY_NOTICE
    assert "LibreOffice 作为兼容回退" in OFFICE_COMPATIBILITY_NOTICE


def test_motion_preferences_are_normalized_and_performance_aware() -> None:
    assert normalize_motion_mode("RICH") == "rich"
    assert normalize_motion_mode("unknown") == "rich"
    assert normalize_particle_effects_enabled(True) is True
    assert normalize_particle_effects_enabled("on") is True
    assert normalize_particle_effects_enabled(False) is False
    assert normalize_particle_effects_enabled("false") is False
    assert particle_effect_button_text(True) == "粒子动效：开"
    assert particle_effect_button_text(False, compact=True) == "粒子：关"
    assert motion_button_text("off") == "粒子动效：关"
    assert motion_frame_delay("off") == 0
    assert motion_frame_delay("rich", busy=True) > motion_frame_delay("rich")
    assert motion_frame_delay("light") > motion_frame_delay("rich")
    assert motion_frame_delay("rich", minimized=True) == 650
    assert ease_out_back(0.0) == 0.0
    assert ease_out_back(1.0) == 1.0
    assert ease_out_back(0.8) > 0.8


def test_ui_themes_are_colour_only_complete_and_restore_the_real_original() -> None:
    assert UI_THEME_LABELS == {
        "tech": "科技风",
        "original": "原版画风",
        "cream": "奶油画风",
    }
    expected_keys = set(UI_THEME_PALETTE_KEYS)
    assert expected_keys
    for palette in UI_THEME_PALETTES.values():
        assert set(palette) == expected_keys
        assert all(re.fullmatch(r"#[0-9A-F]{6}", colour) for colour in palette.values())

    original = UI_THEME_PALETTES["original"]
    assert original["BG"] == "#F3F6FB"
    assert original["PANEL"] == "#FFFFFF"
    assert original["TEXT"] == "#172033"
    assert original["ACCENT"] == "#4F6BED"
    assert original["BORDER"] == "#DDE4EF"


def test_ui_theme_preferences_are_normalized_atomic_and_damage_tolerant(
    tmp_path: Path,
) -> None:
    preferences = tmp_path / "nested" / "preferences.json"
    assert normalize_ui_theme("CREAM") == "cream"
    assert normalize_ui_theme("unknown") == "tech"
    assert load_ui_preferences(preferences) == {
        "ui_theme": "tech",
        "particle_effects": True,
    }

    saved = save_ui_preferences(
        {"ui_theme": "original", "particle_effects": False}, preferences
    )
    assert saved == preferences
    assert load_ui_preferences(preferences) == {
        "ui_theme": "original",
        "particle_effects": False,
    }
    assert not list(preferences.parent.glob("*.tmp"))

    preferences.write_text("{not valid json", encoding="utf-8")
    assert load_ui_preferences(preferences) == {
        "ui_theme": "tech",
        "particle_effects": True,
    }


def test_q_bounce_transition_is_smooth_monotonic_and_settles_exactly() -> None:
    rich = q_bounce_transition_plan("rich")
    light = q_bounce_transition_plan("light")
    off = q_bounce_transition_plan("off")
    assert len(rich) > len(light) > len(off)
    assert rich[0].progress == light[0].progress == 0.0
    assert all(0.0 <= frame.progress <= 1.0 for frame in rich + light)
    assert all(
        first.progress <= second.progress
        for plan in (rich, light)
        for first, second in zip(plan, plan[1:])
    )
    assert max(frame.scale for frame in rich) <= 1.04
    assert all(frame.scale == 1.0 for frame in light)
    assert rich[-1].progress == 1.0
    assert rich[-1].scale == 1.0
    assert off[-1].progress == off[-1].scale == 1.0


def test_background_watermarks_are_deterministic_small_and_edge_bound() -> None:
    narrow = background_watermark_plan(820, 620, "narrow")
    wide = background_watermark_plan(1480, 920, "wide")
    assert narrow == background_watermark_plan(820, 620, "narrow")
    assert len(wide) > len(narrow) >= 3
    assert {item.kind for item in wide} >= {"star", "file", "node", "satellite"}
    for item in wide:
        x1, y1, x2, y2 = item.box
        assert 0 <= x1 < x2 <= 1480
        assert 0 <= y1 < y2 <= 920
        assert 14 <= max(x2 - x1, y2 - y1) <= 28
        assert x1 < 60 or y1 < 60 or x2 > 1420 or y2 > 860
    assert background_watermark_plan(-1, 10, "wide") == ()


def test_global_office_engine_selector_only_targets_real_office_parameters() -> None:
    detected = {
        operation.id: spec.key
        for operation in get_operations()
        if (spec := office_engine_parameter_spec(operation)) is not None
    }

    assert detected == {
        "word.to_pdf": "engine",
        "word.full_compatibility": "verification_engine",
        "excel.to_pdf": "engine",
        "ppt.to_pdf": "engine",
        "legacy.doc_to_docx": "engine",
        "legacy.xls_to_xlsx": "engine",
        "ppt.to_images": "renderer",
    }
    assert set(OFFICE_ENGINE_VALUES) == {
        "auto",
        "wps",
        "microsoft_office",
        "libreoffice",
    }


def test_global_office_engine_selector_rejects_same_named_non_office_choices() -> None:
    unrelated = Operation(
        "image.fake_engine",
        "测试",
        "图像引擎",
        "测试同名参数不会冲突",
        lambda _paths, _output, _parameters: [],
        parameters=(
            ParameterSpec(
                "engine",
                "图像增强引擎",
                "choice",
                "opencv",
                choices=(("opencv", "OpenCV"), ("realesrgan", "Real-ESRGAN")),
            ),
        ),
    )

    assert office_engine_parameter_spec(unrelated) is None
    assert office_engine_button_text("wps", active=True) == "当前引擎：WPS  ▾"
    assert office_engine_button_text("microsoft_office", compact=True) == "引擎：Microsoft  ▾"


def test_public_operations_are_grouped_into_three_compact_catalog_roots() -> None:
    operations = get_operations()
    paths = [operation_catalog_path(operation) for operation in operations]

    assert len(operations) == 73
    assert Counter(root for root, _section in paths) == {
        "文档工具": 47,
        "图片工具": 18,
        "视频工具": 8,
    }
    assert set(paths) == {
        ("文档工具", "PDF 转换与生成"),
        ("文档工具", "PDF 页面整理"),
        ("文档工具", "PDF 压缩与安全"),
        ("文档工具", "PDF 水印与页码"),
        ("文档工具", "Office 格式转换"),
        ("文档工具", "Word 文档处理"),
        ("文档工具", "Excel 数据处理"),
        ("文档工具", "PowerPoint 演示处理"),
        ("文档工具", "兼容修复 / 高级工具"),
        ("图片工具", "格式转换与批量"),
        ("图片工具", "尺寸裁剪与几何"),
        ("图片工具", "压缩优化与隐私"),
        ("图片工具", "编辑增强与合成"),
        ("视频工具", "视频生成"),
        ("视频工具", "转码压缩与提取"),
    }
    extract_images = next(
        item for item in operations if item.id == "pdf.extract_images"
    )
    assert operation_catalog_path(extract_images) == (
        "文档工具",
        "PDF 转换与生成",
    )
    legacy_word_upgrade = next(
        item for item in operations if item.id == "word.full_compatibility"
    )
    assert operation_catalog_path(legacy_word_upgrade) == (
        "文档工具",
        "兼容修复 / 高级工具",
    )
    lecture_slides = next(
        item for item in operations if item.id == "video.extract_slides_ppt"
    )
    assert operation_catalog_path(lecture_slides) == (
        "视频工具",
        "视频生成",
    )
    repair_slides = next(
        item for item in operations if item.id == "video.repair_slides_ppt"
    )
    assert operation_catalog_path(repair_slides) == (
        "视频工具",
        "视频生成",
    )
    image_enhance = next(item for item in operations if item.id == "image.enhance")
    image_mode = next(item for item in image_enhance.parameters if item.key == "mode")
    assert image_mode.default == "auto"
    assert {value for value, _label in image_mode.choices} == {
        "auto",
        "compatible",
        "gpu_ai",
    }
    video_mode = next(
        item for item in lecture_slides.parameters if item.key == "enhancement_mode"
    )
    assert video_mode.default == "auto"
    assert {value for value, _label in video_mode.choices} == {
        "auto",
        "compatible",
        "gpu_ai",
        "off",
    }


def test_task_result_presentation_uses_unambiguous_three_colour_states() -> None:
    source = Path("input.pdf")
    success = TaskResult(outputs=[Path("output.docx")], completed_inputs=[source])
    partial = TaskResult(
        outputs=[Path("output.docx")],
        completed_inputs=[source],
        failed_inputs=[TaskFailure(Path("broken.pdf"), "Error", "损坏")],
    )
    failure = TaskResult(
        failed_inputs=[TaskFailure(source, "Error", "损坏")],
    )

    success_view = task_result_presentation(success)
    partial_view = task_result_presentation(partial)
    failure_view = task_result_presentation(failure)

    assert success_view.title == "处理成功"
    assert success_view.gradient_start == "#0F9D70"
    assert partial_view.title == "部分完成"
    assert partial_view.gradient_start == "#D79518"
    assert failure_view.title == "处理失败"
    assert failure_view.gradient_start == "#CF3F4C"


def test_unavailable_office_operation_dialog_explains_local_engine_options() -> None:
    operation = Operation(
        "word.to_pdf",
        "Office 格式转换",
        "Word 转 PDF（高保真）",
        "测试",
        lambda _paths, _out, _params: [],
    )
    presentation = unavailable_operation_presentation(
        operation,
        Capability(
            "unavailable",
            "未检测到可用的 WPS 或 Microsoft Office COM",
            "Office 渲染器",
        ),
    )

    assert presentation.title == "缺少 Office 转换引擎"
    assert "WPS Office 与 Microsoft Office" in presentation.message
    assert "Microsoft Office" in presentation.message
    assert "LibreOffice 作为兼容回退" in presentation.message
    assert "COM 自动化接口" in presentation.message


def test_catalog_structure_has_a_stable_readable_order() -> None:
    flattened = [
        (root_name, section_name)
        for root_name, sections in CATALOG_STRUCTURE
        for section_name in sections
    ]

    assert sorted(flattened, key=lambda item: catalog_order_key(*item)) == flattened
    assert catalog_order_key("其他工具") > catalog_order_key("视频工具")


def test_parameter_help_combines_explanation_bounds_and_default() -> None:
    numeric = ParameterSpec(
        "dpi",
        "清晰度",
        "integer",
        220,
        help_text="影响文字边缘清晰度",
        minimum=72,
        maximum=600,
    )
    choice = ParameterSpec(
        "mode",
        "模式",
        "choice",
        "balanced",
        choices=(("small", "体积优先"), ("balanced", "均衡（推荐）")),
    )

    assert parameter_help_text(numeric) == (
        "影响文字边缘清晰度；允许范围：72–600；默认：220。"
    )
    assert parameter_help_text(choice) == "默认：均衡（推荐）；请从下拉菜单选择。"


def test_multi_color_picker_accepts_hex_and_rgb_without_duplicates() -> None:
    assert color_dialog_initial("#00AEEF;#ff0000") == "#00AEEF"
    assert color_dialog_initial("0, 174, 239;255,0,0") == "#00aeef"
    assert color_dialog_initial("300,0,0") == "#000000"
    assert append_color_value("#00AEEF", "#ff0000") == "#00AEEF;#ff0000"
    assert append_color_value("#00AEEF;#ff0000", "#FF0000") == (
        "#00AEEF;#ff0000"
    )


def test_video_picker_maps_preview_pixels_and_regions_to_source_space() -> None:
    display = fitted_media_rect(1920, 1080, 800, 600)
    assert display == (0, 75, 800, 525)
    assert canvas_point_to_source(400, 300, display, 1920, 1080) == (960, 540)
    assert canvas_point_to_source(400, 50, display, 1920, 1080) is None
    region = canvas_selection_to_percent_region((80, 120), (400, 300), display)
    assert region == (10.0, 10.0, 40.0, 40.0)
    assert format_percent_region(region) == "10,10,40,40"


def test_video_eyedropper_uses_median_patch_instead_of_one_noisy_pixel() -> None:
    import numpy as np

    frame = np.zeros((9, 9, 3), dtype=np.uint8)
    frame[:, :] = (30, 80, 210)  # BGR -> RGB 210,80,30
    frame[4, 4] = (255, 255, 255)

    assert robust_frame_rgb(frame, 4, 4, radius=2) == (210, 80, 30)


def test_responsive_wraplength_uses_available_width_with_a_safe_floor() -> None:
    assert responsive_wraplength(1600, reserved_width=180, padding=24) == 1396
    assert (
        responsive_wraplength(
            420,
            reserved_width=180,
            padding=24,
            minimum=320,
        )
        == 320
    )
    assert responsive_wraplength(-10, minimum=1) == 1


def test_mousewheel_scroll_units_supports_wheels_and_small_touchpad_deltas() -> None:
    assert mousewheel_scroll_units(120) == -3
    assert mousewheel_scroll_units(-240) == 6
    assert mousewheel_scroll_units(30) == -1
    assert mousewheel_scroll_units(0) == 0
    assert mousewheel_scroll_units(float("nan")) == 0


def test_responsive_layout_mode_uses_narrow_compact_and_wide_tiers() -> None:
    assert responsive_layout_mode(820) == "narrow"
    assert responsive_layout_mode(979) == "narrow"
    assert responsive_layout_mode(980) == "compact"
    assert responsive_layout_mode(1319) == "compact"
    assert responsive_layout_mode(1320) == "wide"
    assert responsive_layout_mode(-1) == "narrow"


def test_rounded_card_resize_coalesces_paint_and_uses_latest_dimensions() -> None:
    class FakeRoundedCard:
        _schedule_redraw = RoundedCard._schedule_redraw
        _run_scheduled_redraw = RoundedCard._run_scheduled_redraw

        def __init__(self) -> None:
            self._redraw_job = None
            self._pending_draw_size = (0, 0)
            self.root_resizing = False
            self.inner_sizes: list[tuple[int, int]] = []
            self.surface_sizes: list[tuple[int, int]] = []
            self.scheduled: dict[str, tuple[int, object]] = {}
            self.cancelled: list[str] = []

        def _root_is_resizing(self) -> bool:
            return self.root_resizing

        def _resize_inner_window(self, width: int, height: int) -> None:
            self.inner_sizes.append((width, height))

        def _draw_card_surface(self, width: int, height: int) -> None:
            self.surface_sizes.append((width, height))

        def after(self, delay: int, callback: object) -> str:
            job = f"resize-job-{len(self.scheduled) + len(self.cancelled) + 1}"
            self.scheduled[job] = (delay, callback)
            return job

        def after_cancel(self, job: str) -> None:
            self.cancelled.append(job)
            self.scheduled.pop(job, None)

    card = FakeRoundedCard()
    card._schedule_redraw(SimpleNamespace(width=800, height=500))
    card._schedule_redraw(SimpleNamespace(width=930, height=640))

    assert card.inner_sizes == []
    assert card.surface_sizes == []
    assert card.cancelled == ["resize-job-1"]
    assert len(card.scheduled) == 1
    delay, callback = next(iter(card.scheduled.values()))
    assert delay == 28
    assert card._pending_draw_size == (930, 640)

    assert callable(callback)
    callback()
    assert card._redraw_job is None
    assert card.inner_sizes == [(930, 640)]
    assert card.surface_sizes == [(930, 640)]


def test_rounded_card_resize_freezes_child_geometry_during_root_resize() -> None:
    class FakeRoundedCard:
        _schedule_redraw = RoundedCard._schedule_redraw

        def __init__(self) -> None:
            self._redraw_job = None
            self._pending_draw_size = (0, 0)
            self.scheduled: list[tuple[int, object]] = []

        def _root_is_resizing(self) -> bool:
            return True

        def after(self, delay: int, callback: object) -> str:
            self.scheduled.append((delay, callback))
            return "unexpected-resize-job"

        def after_cancel(self, _job: str) -> None:
            raise AssertionError("resizing must not leave a pending redraw job")

    card = FakeRoundedCard()
    card._schedule_redraw(SimpleNamespace(width=1040, height=720))

    assert card._pending_draw_size == (1040, 720)
    assert card._redraw_job is None
    assert card.scheduled == []


def test_windows_redraw_flags_invalidate_children_without_erasing() -> None:
    assert windows_redraw_flags() == 0x0081


def test_window_configure_ignores_moves_and_debounces_real_resizes() -> None:
    class FakeWindowApp:
        _on_window_configure = DocuForgeApp._on_window_configure

        def __init__(self) -> None:
            self._closing = False
            self._window_mapped = True
            self._window_restoring = False
            self._window_resizing = False
            self._last_window_configure_size = None
            self._pending_window_width = None
            self._pending_window_height = None
            self._window_restore_job = None
            self._window_restore_finalize_job = None
            self._window_resize_finish_job = None
            self._window_layout_job = None
            self.scheduled: dict[str, tuple[int, object]] = {}
            self.cancelled: list[str] = []

        def _logical_window_width(self, width: int) -> int:
            return width

        def _logical_window_height(self, height: int) -> int:
            return height

        def state(self) -> str:
            return "normal"

        def after(self, delay: int, callback: object) -> str:
            job = f"window-job-{len(self.scheduled) + 1}"
            self.scheduled[job] = (delay, callback)
            return job

        def after_cancel(self, job: str) -> None:
            self.cancelled.append(job)
            self.scheduled.pop(job, None)

        def _finish_window_resize(self) -> None:
            return None

        def _flush_window_layout(self) -> None:
            return None

    app = FakeWindowApp()
    app._on_window_configure(SimpleNamespace(widget=app, width=1000, height=700))

    assert app._last_window_configure_size == (1000, 700)
    assert app._window_resizing is False
    assert app.scheduled == {}

    app._on_window_configure(SimpleNamespace(widget=app, width=1000, height=700))
    assert app._window_resizing is False
    assert app.scheduled == {}

    app._on_window_configure(SimpleNamespace(widget=app, width=1040, height=720))
    assert app._window_resizing is True
    assert app._pending_window_width == 1040
    assert app._pending_window_height == 720
    assert sorted(delay for delay, _callback in app.scheduled.values()) == [158]


def test_window_configure_during_restore_only_records_latest_size() -> None:
    class FakeWindowApp:
        _on_window_configure = DocuForgeApp._on_window_configure

        def __init__(self) -> None:
            self._closing = False
            self._window_mapped = True
            self._window_restoring = True
            self._window_resizing = True
            self._last_window_configure_size = (1000, 700)
            self._pending_window_width = None
            self._pending_window_height = None
            self._window_restore_job = "restore-job"
            self._window_restore_finalize_job = None
            self._window_resize_finish_job = None

        def state(self) -> str:
            return "normal"

        def _logical_window_width(self, width: int) -> int:
            return width

        def _logical_window_height(self, height: int) -> int:
            return height

        def after(self, _delay: int, _callback: object) -> str:
            raise AssertionError("restore Configure must not schedule resize work")

        def after_cancel(self, _job: str) -> None:
            raise AssertionError("restore Configure must not cancel restore work")

    app = FakeWindowApp()
    app._on_window_configure(SimpleNamespace(widget=app, width=1220, height=820))

    assert app._last_window_configure_size == (1220, 820)
    assert app._pending_window_width == 1220
    assert app._pending_window_height == 820
    assert app._window_restore_job == "restore-job"


def test_unmapped_configure_does_not_replace_last_visible_size() -> None:
    class FakeWindowApp:
        _on_window_configure = DocuForgeApp._on_window_configure

        def __init__(self) -> None:
            self._closing = False
            self._window_mapped = False
            self._last_window_configure_size = (1000, 700)
            self._pending_window_width = None
            self._pending_window_height = None

    app = FakeWindowApp()
    app._on_window_configure(SimpleNamespace(widget=app, width=1, height=1))

    assert app._last_window_configure_size == (1000, 700)
    assert app._pending_window_width is None
    assert app._pending_window_height is None


def test_force_flush_setup_canvas_width_uses_live_width_over_stale_pending() -> None:
    class FakeCanvas:
        def __init__(self) -> None:
            self.configured: list[tuple[object, int]] = []

        def winfo_width(self) -> int:
            return 777

        def itemconfigure(self, window: object, *, width: int) -> None:
            self.configured.append((window, width))

    class FakeWindowApp:
        _flush_setup_canvas_width = DocuForgeApp._flush_setup_canvas_width

        def __init__(self) -> None:
            self._setup_canvas_resize_job = "stale-job"
            self._closing = False
            self._window_resizing = True
            self._pending_setup_canvas_width = 333
            self._setup_canvas_width = 320
            self.setup_canvas = FakeCanvas()
            self.setup_window = object()
            self.scroll_refreshes = 0

        def _schedule_setup_scroll_refresh(self) -> None:
            self.scroll_refreshes += 1

    app = FakeWindowApp()
    app._flush_setup_canvas_width(force=True)

    assert app.setup_canvas.configured == [(app.setup_window, 777)]
    assert app._setup_canvas_width == 777
    assert app._pending_setup_canvas_width is None
    assert app._setup_canvas_resize_job is None
    assert app.scroll_refreshes == 1


def test_window_map_defers_restore_without_synchronous_native_redraw(
    monkeypatch,
) -> None:
    native_redraw_calls: list[object] = []
    monkeypatch.setattr(
        app_module,
        "_force_windows_window_redraw",
        lambda window: native_redraw_calls.append(window),
    )

    class FakeWindowApp:
        _on_window_map = DocuForgeApp._on_window_map

        def __init__(self) -> None:
            self._closing = False
            self._window_mapped = False
            self._window_restoring = False
            self._window_restore_attempts = 99
            self._window_resizing = False
            self._window_restore_job = None
            self._window_restore_finalize_job = None
            self.idle_callbacks: list[object] = []

        def after_idle(self, callback: object) -> str:
            self.idle_callbacks.append(callback)
            return "restore-after-idle"

        def after_cancel(self, _job: str) -> None:
            return None

        def _restore_window_after_map(self) -> None:
            return None

    app = FakeWindowApp()
    app._on_window_map(SimpleNamespace(widget=app))

    assert native_redraw_calls == []
    assert len(app.idle_callbacks) == 1
    assert app._window_restore_job == "restore-after-idle"
    assert app._window_mapped is True
    assert app._window_restoring is True
    assert app._window_restore_attempts == 0
    assert app._window_resizing is True


def test_tk_configure_width_tracks_monitor_scale() -> None:
    app = object.__new__(DocuForgeApp)
    app._display_scale = 2.0
    assert app._logical_window_width(1640) == 820
    app._display_scale = 1.0
    assert app._logical_window_width(820) == 820


def test_catalog_sidebar_width_adapts_and_preserves_main_content_space() -> None:
    assert catalog_sidebar_width(1000, preferred_width=460) == 300
    assert catalog_sidebar_width(1200, preferred_width=460) == 360
    assert catalog_sidebar_width(1480, preferred_width=460) == 459
    assert catalog_sidebar_width(1920, preferred_width=600) == 480
    assert catalog_sidebar_width(1480, preferred_width=460, user_width=340) == 340
    assert catalog_sidebar_width(1480, preferred_width=460, user_width=900) == 459
    assert catalog_sidebar_width(820, preferred_width=460) == 802


def test_initial_window_size_stays_inside_the_display() -> None:
    assert initial_window_size(1920, 1080) == (1480, 920)
    assert initial_window_size(1366, 768) == (1270, 672)
    assert initial_window_size(700, 500) == (700, 500)
    assert initial_window_size(800, 600) == (704, 504)


def test_short_window_layout_uses_a_stable_height_threshold() -> None:
    assert short_window_layout(679) is True
    assert short_window_layout(680) is False
    assert short_window_layout(900) is False


def test_dialog_size_respects_preference_margins_and_small_displays() -> None:
    assert fitted_dialog_size(
        1920,
        1080,
        preferred_width=1500,
        preferred_height=900,
        minimum_width=860,
        minimum_height=620,
    ) == (1500, 900)
    assert fitted_dialog_size(
        1366,
        768,
        preferred_width=1500,
        preferred_height=900,
        minimum_width=860,
        minimum_height=620,
    ) == (1294, 680)
    assert fitted_dialog_size(
        700,
        500,
        preferred_width=1080,
        preferred_height=780,
        minimum_width=720,
        minimum_height=560,
    ) == (700, 500)


def test_result_dialog_reserves_room_for_clear_actions_and_small_windows() -> None:
    assert result_dialog_geometry(1480, 920) == (820, 662)
    assert result_dialog_geometry(820, 620) == (650, 520)
    assert result_dialog_geometry(2400, 1600) == (820, 700)


def test_ui_motion_helpers_are_bounded_and_colour_safe() -> None:
    assert ease_out_cubic(-1) == 0.0
    assert ease_out_cubic(0) == 0.0
    assert ease_out_cubic(1) == 1.0
    assert 0.8 < ease_out_cubic(0.5) < 0.9
    assert interpolate_hex_colour("#000000", "#FFFFFF", 0.5) == "#808080"
    assert interpolate_hex_colour("#4F6BED", "#172033", 1) == "#172033"


def test_click_particle_plan_is_slow_smooth_and_fully_fades_out() -> None:
    frames = click_particle_frame_plan()

    assert len(frames) == 19
    assert all(is_dataclass(frame) and not isinstance(frame, type) for frame in frames)
    assert all(0.0 <= frame.spread <= 1.0 for frame in frames)
    assert all(0.0 <= frame.opacity <= 1.0 for frame in frames)
    assert all(frame.scale > 0.0 and math.isfinite(frame.scale) for frame in frames)
    assert all(first.spread <= second.spread for first, second in zip(frames, frames[1:]))
    assert frames[-1].spread == 1.0
    assert frames[-1].opacity == 0.0
    assert frames[-1].curve == 0.0
    # Tk renders 18 intervals for 19 frames: deliberately around 0.49 s.
    assert 450 <= (len(frames) - 1) * 27 <= 540


def test_click_particle_variants_are_rich_deterministic_and_bounded() -> None:
    variants = tuple(click_particle_specs(index) for index in range(3))

    assert click_particle_specs(3) == variants[0]
    assert all(len(items) == 11 for items in variants)
    assert len(set(variants)) == 3
    kinds = {particle.kind for items in variants for particle in items}
    assert kinds >= {"star", "moon", "sparkle", "diamond", "ring", "comet", "hex", "dot"}
    for items in variants:
        assert all(-1.0 <= particle.tangent <= 1.0 for particle in items)
        assert all(3.0 <= particle.size <= 9.0 for particle in items)


class _ParticleCanvasRecorder:
    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[object, ...], dict[str, object]]] = []

    def create_arc(self, *coords: object, **options: object) -> None:
        self.calls.append(("arc", coords, options))

    def create_line(self, *coords: object, **options: object) -> None:
        self.calls.append(("line", coords, options))

    def create_oval(self, *coords: object, **options: object) -> None:
        self.calls.append(("oval", coords, options))

    def create_polygon(self, *coords: object, **options: object) -> None:
        self.calls.append(("polygon", coords, options))


def test_particle_ring_uses_broken_orbit_arcs_not_a_full_oval() -> None:
    canvas = _ParticleCanvasRecorder()
    size = 10.0

    DocuForgeApp._draw_particle_symbol(canvas, "ring", 40.0, 30.0, size, "#123456")

    arcs = [(coords, options) for kind, coords, options in canvas.calls if kind == "arc"]
    ovals = [(coords, options) for kind, coords, options in canvas.calls if kind == "oval"]
    assert len(arcs) >= 2
    assert all(options.get("style") == "arc" for _, options in arcs)
    assert sum(abs(float(options["extent"])) for _, options in arcs) < 360.0
    assert not any(
        max(
            abs(float(coords[2]) - float(coords[0])),
            abs(float(coords[3]) - float(coords[1])),
        )
        >= size * 1.5
        for coords, _ in ovals
    )


def test_particle_pulse_uses_four_round_short_rays_without_closed_circle() -> None:
    canvas = _ParticleCanvasRecorder()
    size = 10.0

    DocuForgeApp._draw_particle_symbol(canvas, "pulse", 40.0, 30.0, size, "#654321")

    lines = [(coords, options) for kind, coords, options in canvas.calls if kind == "line"]
    assert len(lines) == 4
    assert all(options.get("capstyle") == "round" for _, options in lines)
    assert all(
        0.0
        < math.hypot(
            float(coords[2]) - float(coords[0]),
            float(coords[3]) - float(coords[1]),
        )
        < size
        for coords, _ in lines
    )
    assert not any(kind in {"arc", "oval", "polygon"} for kind, _, _ in canvas.calls)


def test_composite_hex_colour_handles_alpha_and_non_finite_values() -> None:
    background = "#102030"
    foreground = "#90A0B0"

    assert composite_hex_colour(background, foreground, 0) == background
    assert composite_hex_colour(background, foreground, 1) == foreground
    assert composite_hex_colour(background, foreground, 0.25) == "#304050"
    for non_finite in (float("nan"), float("inf"), float("-inf")):
        assert composite_hex_colour(background, foreground, non_finite) == background


def test_motion_effect_timing_is_a_bounded_performance_aware_dataclass() -> None:
    effects = ("click", "transition", "title", "progress", "dialog")
    modes = ("rich", "light", "off")

    for mode in modes:
        for effect in effects:
            normal = motion_effect_timing(mode, effect)
            busy = motion_effect_timing(mode, effect, busy=True)
            minimized = motion_effect_timing(mode, effect, minimized=True)
            busy_minimized = motion_effect_timing(
                mode,
                effect,
                busy=True,
                minimized=True,
            )

            for timing in (normal, busy, minimized, busy_minimized):
                assert is_dataclass(timing) and not isinstance(timing, type)
                assert {field.name for field in fields(timing)} >= {
                    "frames",
                    "step_ms",
                    "enabled",
                }
                assert isinstance(timing.frames, int) and timing.frames >= 0
                assert isinstance(timing.step_ms, int) and timing.step_ms >= 0
                assert isinstance(timing.enabled, bool)

            if mode == "off":
                assert all(
                    not timing.enabled or timing.frames * timing.step_ms == 0
                    for timing in (normal, busy, minimized, busy_minimized)
                )
                continue

            assert normal.enabled is True
            assert busy.enabled is True
            assert 0 < normal.frames * normal.step_ms <= 260
            assert 0 < busy.frames * busy.step_ms <= 260
            assert busy.frames <= normal.frames
            assert not minimized.enabled or minimized.frames * minimized.step_ms == 0
            assert (
                not busy_minimized.enabled
                or busy_minimized.frames * busy_minimized.step_ms == 0
            )


def test_result_dialog_entrance_plans_have_safe_exact_endpoints() -> None:
    rich = result_dialog_entrance_plan("rich")
    light = result_dialog_entrance_plan("light")
    off = result_dialog_entrance_plan("off")

    assert rich[0] == 0.88
    assert light[0] == 0.94
    assert rich[-1] == light[-1] == 1.0
    assert off == (1.0,)
    for plan in (rich, light):
        assert len(plan) > 1
        assert all(math.isfinite(alpha) and 0.0 <= alpha <= 1.0 for alpha in plan)
        assert all(first <= second for first, second in zip(plan, plan[1:]))


def test_live_image_preview_is_truthful_and_bounded(tmp_path) -> None:
    from PIL import Image

    source = tmp_path / "source.png"
    Image.new("RGB", (100, 80), "#D05050").save(source)
    assert supports_live_image_preview("image.crop") is True
    assert supports_live_image_preview("image.enhance") is False

    original, result, original_info, result_info, source_size = build_live_image_preview(
        source,
        "image.crop",
        {"left": 10, "top": 15, "right": 60, "bottom": 55},
        max_dimension=200,
    )
    try:
        assert original.size == (100, 80)
        assert result.size == (50, 40)
        assert source_size == (100, 80)
        assert "100 × 80" in original_info
        assert "50 × 40" in result_info
        assert "蓝框" in result_info
    finally:
        original.close()
        result.close()


def test_repair_preview_layout_keeps_all_three_panels_complete() -> None:
    assert repair_preview_layout(1500, 900) == "three_columns"
    assert repair_preview_layout(1100, 820) == "two_rows"
    assert repair_preview_layout(960, 650) == "three_columns"


def test_compact_operation_description_collapses_optional_precision_details() -> None:
    collapsed = operation_description_text(
        "正文尽量可编辑。",
        "复杂表格会局部高清保留。",
        "pdf2docx",
        "本地引擎可用",
        compact=True,
        expanded=False,
    )
    assert "正文尽量可编辑" in collapsed
    assert "复杂表格" not in collapsed
    assert "了解更多" in collapsed

    expanded = operation_description_text(
        "正文尽量可编辑。",
        "复杂表格会局部高清保留。",
        "pdf2docx",
        "本地引擎可用",
        compact=True,
        expanded=True,
    )
    assert "更多说明：复杂表格会局部高清保留" in expanded
    assert "技术信息：pdf2docx" in expanded

    wide_collapsed = operation_description_text(
        "正文尽量可编辑。",
        "复杂表格会局部高清保留。",
        "pdf2docx",
        "本地引擎可用",
        compact=False,
        expanded=False,
    )
    assert "正文尽量可编辑" in wide_collapsed
    assert "复杂表格" not in wide_collapsed
    assert "了解更多" in wide_collapsed


def test_catalog_display_names_hide_trailing_technical_qualifiers() -> None:
    assert operation_display_name("PDF 转 Word（版式混合 / 全文可编辑）") == "PDF 转 Word"
    assert operation_display_name("Word 转 PDF（高保真）") == "Word 转 PDF"
    assert operation_display_name("PDF 合并") == "PDF 合并"


def test_progress_display_helpers_are_bounded_and_explain_long_stages() -> None:
    assert progress_percent(0.375) == 37.5
    assert progress_percent(-1) == 0.0
    assert progress_percent(2) == 100.0
    assert progress_percent(float("nan")) == 0.0
    assert progress_status_text("阶段：转换", 3.9) == "阶段：转换 · 已用时 3 秒"
    assert (
        progress_status_text("阶段：转换", 12, seconds_since_update=8)
        == "阶段：转换 · 已用时 12 秒 · 当前阶段仍在运行"
    )


def test_progress_smoothing_is_monotonic_bounded_and_tracks_new_targets() -> None:
    values = [0.0]
    for _ in range(36):
        values.append(smooth_progress_step(values[-1], 72.0, 16.0))
    assert all(first <= second for first, second in zip(values, values[1:]))
    assert all(0.0 <= value <= 72.0 for value in values)
    assert values[-1] > 70.0
    assert smooth_progress_step(values[-1], 93.0, 16.0) > values[-1]
    assert smooth_progress_step(93.0, 0.0, 16.0) == 0.0
    assert smooth_progress_step(-20.0, 150.0, 16.0) <= 100.0
