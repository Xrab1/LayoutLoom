from __future__ import annotations

from collections import Counter
from pathlib import Path

from docuforge.app import (
    CATALOG_STRUCTURE,
    WPS_COMPATIBILITY_NOTICE,
    append_color_value,
    build_live_image_preview,
    catalog_sidebar_width,
    canvas_point_to_source,
    canvas_selection_to_percent_region,
    catalog_order_key,
    color_dialog_initial,
    ease_out_cubic,
    fitted_dialog_size,
    fitted_media_rect,
    format_percent_region,
    initial_window_size,
    interpolate_hex_colour,
    mousewheel_scroll_units,
    operation_description_text,
    operation_catalog_path,
    parameter_help_text,
    progress_percent,
    progress_status_text,
    repair_preview_layout,
    responsive_layout_mode,
    responsive_wraplength,
    robust_frame_rgb,
    supports_live_image_preview,
    task_result_presentation,
    unavailable_operation_presentation,
)
from docuforge.models import Capability, Operation, ParameterSpec, TaskFailure, TaskResult
from docuforge.registry import get_operations


def test_wps_compatibility_notice_is_explicit_and_actionable() -> None:
    assert "国内拥有庞大的用户基础" in WPS_COMPATIBILITY_NOTICE
    assert "Microsoft Office 与 LibreOffice" in WPS_COMPATIBILITY_NOTICE
    assert "安装桌面版 WPS Office" in WPS_COMPATIBILITY_NOTICE


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
    assert "WPS Office（推荐）" in presentation.message
    assert "Microsoft Office" in presentation.message
    assert "自行安装 LibreOffice" in presentation.message
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


def test_ui_motion_helpers_are_bounded_and_colour_safe() -> None:
    assert ease_out_cubic(-1) == 0.0
    assert ease_out_cubic(0) == 0.0
    assert ease_out_cubic(1) == 1.0
    assert 0.8 < ease_out_cubic(0.5) < 0.9
    assert interpolate_hex_colour("#000000", "#FFFFFF", 0.5) == "#808080"
    assert interpolate_hex_colour("#4F6BED", "#172033", 1) == "#172033"


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
    assert "点击查看" in collapsed

    expanded = operation_description_text(
        "正文尽量可编辑。",
        "复杂表格会局部高清保留。",
        "pdf2docx",
        "本地引擎可用",
        compact=True,
        expanded=True,
    )
    assert "精度说明：复杂表格会局部高清保留" in expanded
    assert "引擎：pdf2docx" in expanded

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
    assert "点击查看" in wide_collapsed


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
