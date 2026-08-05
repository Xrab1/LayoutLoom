from __future__ import annotations

import json
from dataclasses import replace
from pathlib import Path
from typing import Any, Callable, Iterator, Mapping, Sequence

from . import engines
from .models import (
    Capability,
    MissingEngineError,
    Operation,
    ParameterSpec,
    TaskResult,
    ValidationError,
)
from .utils import safe_filename, unique_directory, unique_path

P = ParameterSpec
PDF_EXT = (".pdf",)
IMAGE_EXT = (
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
    ".gif",
    ".ico",
    ".psd",
)
WORD_EXT = (".docx", ".docm", ".dotx", ".dotm")
EXCEL_EXT = (".xlsx", ".xlsm", ".xltx", ".xltm")
PPT_EXT = (".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm")
VIDEO_EXT = (
    ".mp4",
    ".mkv",
    ".mov",
    ".avi",
    ".webm",
    ".wmv",
    ".m4v",
    ".flv",
    ".mpeg",
    ".mpg",
    ".ts",
    ".mts",
    ".m2ts",
    ".3gp",
)


def _target(
    output_dir: Path, source: Path, label: str, suffix: str | None = None
) -> Path:
    return output_dir / f"{source.stem}_{label}{suffix or source.suffix}"


def _filename_with_suffix(value: object, suffix: str, fallback: str) -> str:
    filename = safe_filename(str(value), fallback=fallback)
    if filename.lower().endswith(suffix.lower()):
        return filename
    return f"{filename}{suffix}"


def _batch(
    inputs: Sequence[Path],
    output_dir: Path,
    label: str,
    suffix: str,
    invoke: Callable[[Path, Path], Sequence[Path]],
) -> list[Path]:
    from .runner import check_cancelled, progress_scope, report_progress

    outputs: list[Path] = []
    total = len(inputs)
    for index, source in enumerate(inputs, start=1):
        check_cancelled("任务已取消；已完成的文件会保留")
        report_progress(
            (index - 1) / max(1, total),
            f"处理 {source.name}",
            current_file=index,
            total_files=total,
        )
        with progress_scope(
            (index - 1) / max(1, total),
            1.0 / max(1, total),
            current_file=index,
            total_files=total,
        ):
            outputs.extend(
                invoke(source, unique_path(_target(output_dir, source, label, suffix)))
            )
    check_cancelled("任务已取消；已完成的文件会保留")
    report_progress(
        1.0,
        "核对批量输出",
        current_file=total if total else None,
        total_files=total,
    )
    return outputs


def _batch_to_directory(
    inputs: Sequence[Path],
    output_dir: Path,
    label: str,
    invoke: Callable[[Path, Path], Sequence[Path]],
) -> list[Path]:
    from .runner import check_cancelled, progress_scope, report_progress

    outputs: list[Path] = []
    total = len(inputs)
    for index, source in enumerate(inputs, start=1):
        check_cancelled("任务已取消；已完成的文件会保留")
        report_progress(
            (index - 1) / max(1, total),
            f"处理 {source.name}",
            current_file=index,
            total_files=total,
        )
        target_dir = unique_directory(output_dir / f"{source.stem}_{label}")
        with progress_scope(
            (index - 1) / max(1, total),
            1.0 / max(1, total),
            current_file=index,
            total_files=total,
        ):
            outputs.extend(invoke(source, target_dir))
    check_cancelled("任务已取消；已完成的文件会保留")
    report_progress(
        1.0,
        "核对批量输出",
        current_file=total if total else None,
        total_files=total,
    )
    return outputs


def _progress_paths(inputs: Sequence[Path]) -> Iterator[Path]:
    """Yield per-file inputs while reporting progress for compact handlers."""

    from .runner import check_cancelled, progress_scope, report_progress

    total = len(inputs)
    for index, source in enumerate(inputs, start=1):
        check_cancelled("任务已取消；已完成的文件会保留")
        report_progress(
            (index - 1) / max(1, total),
            f"处理 {source.name}",
            current_file=index,
            total_files=total,
        )
        with progress_scope(
            (index - 1) / max(1, total),
            1.0 / max(1, total),
            current_file=index,
            total_files=total,
        ):
            yield source


def _missing_handler(message: str):
    def handler(_inputs: Sequence[Path], _output: Path, _params: Mapping[str, Any]):
        raise MissingEngineError(message)

    return handler


def _fixed_unavailable(reason: str, engine: str):
    return lambda: Capability("unavailable", reason, engine)


def _json_object(value: str, label: str) -> dict[str, Any]:
    try:
        result = json.loads(value)
    except json.JSONDecodeError as exc:
        from .models import ValidationError

        raise ValidationError(f"{label}不是有效 JSON：{exc}") from exc
    if not isinstance(result, dict):
        from .models import ValidationError

        raise ValidationError(f"{label}必须是 JSON 对象")
    return result


def _json_records(value: str) -> list[dict[str, Any]] | dict[str, Any]:
    try:
        result = json.loads(value)
    except json.JSONDecodeError as exc:
        from .models import ValidationError

        raise ValidationError(f"邮件合并数据不是有效 JSON：{exc}") from exc
    if isinstance(result, dict):
        return result
    if isinstance(result, list) and all(isinstance(item, dict) for item in result):
        return result
    from .models import ValidationError

    raise ValidationError("邮件合并数据必须是 JSON 对象或对象数组")


def _csv_values(value: str) -> list[str] | None:
    items = [
        item.strip()
        for item in str(value).replace("，", ",").split(",")
        if item.strip()
    ]
    return items or None


def _smart_value(value: str) -> Any:
    text = str(value).strip()
    if not text:
        return ""
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


def _rgb(value: str) -> tuple[int, int, int]:
    try:
        from PIL import ImageColor

        return tuple(ImageColor.getrgb(value))  # type: ignore[return-value]
    except Exception as exc:
        from .models import ValidationError

        raise ValidationError(f"无效颜色：{value}") from exc


def _optional_int(value: str, label: str) -> int | None:
    text = str(value).strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError as exc:
        from .models import ValidationError

        raise ValidationError(f"{label}必须是整数") from exc


def _convert_office_document(
    source: Path, output_dir: Path, target_format: str, engine: str
) -> list[Path]:
    from .processors.office import convert_with_office

    requested = engine.strip().lower()
    if requested not in {"auto", "microsoft_office", "wps", "libreoffice"}:
        from .models import ValidationError

        raise ValidationError(
            "Office 转换引擎必须是 auto、microsoft_office、wps 或 libreoffice"
        )
    return convert_with_office(source, output_dir, target_format, engine=requested)


def _convert_excel_to_pdf(
    source: Path,
    output_dir: Path,
    engine: str,
    *,
    excel_pdf_layout: str,
    excel_pdf_paper: str,
    excel_pdf_orientation: str,
    excel_pdf_margin: str,
) -> list[Path]:
    from .processors.office import convert_with_office

    requested = engine.strip().lower()
    if requested not in {"auto", "microsoft_office", "wps", "libreoffice"}:
        from .models import ValidationError

        raise ValidationError(
            "Office 转换引擎必须是 auto、microsoft_office、wps 或 libreoffice"
        )
    return convert_with_office(
        source,
        output_dir,
        "pdf",
        engine=requested,
        excel_pdf_layout=excel_pdf_layout,
        excel_pdf_paper=excel_pdf_paper,
        excel_pdf_orientation=excel_pdf_orientation,
        excel_pdf_margin=excel_pdf_margin,
    )


def _pdf_operations() -> list[Operation]:
    from .processors import conversion, pdf, pdf_images, signature

    core = engines.pdf_core_capability
    render = engines.pdf_render_capability
    extract = engines.pdf_extract_capability
    operations: list[Operation] = []

    operations.extend(
        [
            Operation(
                "pdf.to_word",
                "PDF 格式转换",
                "PDF 转 Word（混合保真 / 全文可编辑 / 整篇原样）",
                "将数字 PDF 转为可编辑、混合保真或整篇高清原样的 Word 文档。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    {
                        "hybrid": "混合保真Word",
                        "editable": "可编辑Word",
                        "visual": "高清原样Word",
                    }[p["mode"]],
                    ".docx",
                    lambda src, target: conversion.pdf_to_docx(
                        src,
                        target,
                        password=p["password"] or None,
                        mode=p["mode"],
                        dpi=p["dpi"],
                        low_quality_policy=p["low_quality_policy"],
                        hybrid_force_visual_pages=p["hybrid_force_visual_pages"],
                        column_layout=p.get("column_layout", "auto"),
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "mode",
                        "Word 内容模式",
                        "choice",
                        "hybrid",
                        choices=(
                            ("hybrid", "版式优先混合（推荐）"),
                            ("editable", "全文可编辑重建"),
                            ("visual", "整篇高清原样（不可编辑）"),
                        ),
                        help_text=(
                            "混合模式将公式、图表和复杂表格局部高清保留，正文尽量可编辑；"
                            "双栏论文和设计型简历会自动切换为一源页一 Word 页的固定坐标布局；"
                            "定位节点较多时会在真实 WPS 复检保护下自动压缩为少量可编辑文字区域，失败则保留原高精度布局；"
                            "无法安全分区或质量校验失败的页面仍可能整页高清兜底；"
                            "全文可编辑模式尽量重建全部页面；整篇原样模式视觉最稳定"
                        ),
                    ),
                    P(
                        "column_layout",
                        "PDF 分栏结构",
                        "choice",
                        "auto",
                        choices=(
                            ("auto", "自动识别（推荐）"),
                            ("single", "全文单栏"),
                            ("double", "全文双栏"),
                            ("mixed", "混合分栏（单双栏并存）"),
                        ),
                        help_text=(
                            "用于确定可编辑正文的阅读顺序和 Word 分栏版式。整篇都是一列或两列时请选择全文单栏/全文双栏；"
                            "标题、摘要为单栏而正文为双栏，或不同页面分栏不同，请选择混合分栏；不确定时保留自动识别。"
                            "整篇高清原样模式不受此选项影响。"
                        ),
                    ),
                    P(
                        "hybrid_force_visual_pages",
                        "混合模式强制整页高清页码",
                        "text",
                        "",
                        help_text=(
                            "仅版式优先混合模式生效；可手动指定必须整页高清保留的页码，"
                            "例如：1,3-5；留空则由程序自动判断"
                        ),
                    ),
                    P(
                        "low_quality_policy",
                        "质量校验未通过时",
                        "choice",
                        "discard",
                        choices=(
                            ("discard", "不保留（推荐）"),
                            ("keep", "仍保留并警告"),
                        ),
                        help_text=(
                            "全文可编辑和版式优先混合模式生效；选择不保留时，文本完整度、排版结构或 WPS 实际分页"
                            "任一最终校验失败都会停止保存；混合模式会先尽量使用局部高清或整页高清兜底。"
                            "选择仍保留时，即使最终排版或分页复检未通过也会保存成品并明确警告，供人工复核。"
                        ),
                    ),
                    P(
                        "dpi",
                        "高清保留清晰度 DPI",
                        "integer",
                        300,
                        help_text=(
                            "用于混合模式的局部图像、强制整页高清页和整篇高清原样模式；"
                            "数值越高越清晰，文件也越大"
                        ),
                        minimum=150,
                        maximum=600,
                    ),
                    P("password", "PDF 打开密码", "password", ""),
                ),
                fidelity="editable",
                capability_probe=engines.pdf_to_word_capability,
                notes=(
                    "版式优先混合是默认模式；可由用户明确指定单栏、双栏或混合分栏，减少多栏论文的阅读顺序和排版误判。"
                    "检测到双栏论文、复杂简历或设计型单页时，会自动采用固定坐标可编辑布局，"
                    "保留原页面尺寸、照片、装饰和图表位置，避免 Word 重排造成整页拆分或大面积空白。"
                    "复杂固定布局会先保留原逐行坐标版作为精度基线，再尝试区域级最终化；只有结构、文字、图片、"
                    "页数、重叠、WPS 实际版面和资源预算全部通过时才采用区域结果，否则自动回退原版。"
                    "可靠页面保持可编辑，其中公式、图表、复杂表格和高密度图形"
                    "以局部高清图像保留，其余可靠正文尽量保持可编辑。局部图像内的文字、公式和表格单元格不能单独编辑；"
                    "页面无法可靠分区、可编辑正文质量校验未通过或由用户指定时，仍会整页高清兜底。"
                    "全文可编辑模式会自动校验字符、英文词与词序；校验过低时可选择不保留，或仍保留并警告。"
                    "整篇高清原样模式不可编辑，但版式与画面最稳定。"
                ),
            ),
            Operation(
                "pdf.to_excel",
                "PDF 格式转换",
                "PDF 转 Excel（表格提取）",
                "识别数字版 PDF 中的表格，并按页写入 Excel 工作簿。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "表格",
                    ".xlsx",
                    lambda src, target: pdf.pdf_to_excel(
                        src, target, password=p["password"] or None
                    ),
                ),
                PDF_EXT,
                (P("password", "PDF 打开密码", "password", ""),),
                fidelity="extract",
                capability_probe=extract,
                notes="适合有真实文本和表格线的 PDF；扫描表格需要 OCR，输出后应复核。",
            ),
            Operation(
                "pdf.to_ppt",
                "PDF 格式转换",
                "PDF 转 PPT（视觉原样）",
                "将每个 PDF 页面高清渲染为一张幻灯片，最大限度保持版式。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "转PPT",
                    ".pptx",
                    lambda src, target: pdf.pdf_to_ppt(
                        src,
                        target,
                        dpi=p["dpi"],
                        password=p["password"] or None,
                        poppler_path=engines.poppler_bin_path(),
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "dpi",
                        "渲染清晰度 DPI",
                        "integer",
                        240,
                        minimum=150,
                        maximum=600,
                    ),
                    P("password", "PDF 打开密码", "password", ""),
                ),
                fidelity="visual",
                capability_probe=engines.pdf_to_ppt_capability,
                notes=(
                    "视觉保真高，但每页是整张图片，文字和图形不可单独编辑；"
                    "PPT 一份文件只能使用统一画布，混合横竖或不同尺寸 PDF 会按画布适配。"
                ),
            ),
            Operation(
                "pdf.to_images",
                "PDF 格式转换",
                "PDF 转图片",
                "按指定 DPI 将每页导出为 PNG 或 JPG。",
                lambda paths, out, p: _batch_to_directory(
                    paths,
                    out,
                    "页面",
                    lambda src, target_dir: pdf.pdf_to_images(
                        src,
                        target_dir,
                        image_format=p["format"],
                        dpi=p["dpi"],
                        first_page=_optional_int(p["first_page"], "起始页"),
                        last_page=_optional_int(p["last_page"], "结束页"),
                        password=p["password"] or None,
                        poppler_path=engines.poppler_bin_path(),
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "format",
                        "图片格式",
                        "choice",
                        "png",
                        choices=(("png", "PNG"), ("jpg", "JPG")),
                    ),
                    P("dpi", "清晰度 DPI", "integer", 300, minimum=150, maximum=1200),
                    P("first_page", "起始页（留空为第 1 页）", "text", ""),
                    P("last_page", "结束页（留空为最后页）", "text", ""),
                    P("password", "PDF 打开密码", "password", ""),
                ),
                fidelity="visual",
                capability_probe=render,
            ),
            Operation(
                "pdf.extract_images",
                "PDF 格式转换",
                "提取 PDF 图片（原始 / 可见 / 智能）",
                "直接提取 PDF 内嵌位图，也可按页面可见位置高清还原或智能合并相邻图块。",
                lambda paths, out, p: _batch_to_directory(
                    paths,
                    out,
                    "PDF图片",
                    lambda src, target_dir: pdf_images.extract_pdf_images(
                        src,
                        target_dir,
                        mode=p["mode"],
                        pages=p["pages"],
                        image_format=p["format"],
                        dpi=p["dpi"],
                        jpeg_quality=p["jpeg_quality"],
                        min_width=p["min_width"],
                        min_height=p["min_height"],
                        merge_gap=p["merge_gap"],
                        region_padding=p["region_padding"],
                        deduplicate=p["deduplicate"],
                        include_annotations=p["include_annotations"],
                        write_manifest=p["write_manifest"],
                        password=p["password"] or None,
                        overwrite=False,
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "mode",
                        "提取模式",
                        "choice",
                        "original",
                        choices=(
                            ("original", "原始资源无重采样（最快，推荐）"),
                            ("visible", "逐个可见位置高清还原"),
                            ("smart", "相邻碎片智能合并"),
                            ("both", "原始资源 + 智能合并"),
                            ("all", "原始 + 可见 + 智能（输出最多）"),
                        ),
                        help_text=(
                            "原始模式直接导出 PDF 内嵌位图；可见模式按每个图片在页面上的"
                            "实际位置渲染；智能模式将相邻碎片聚合成更完整的图区；"
                            "组合模式会同时生成多类结果"
                        ),
                    ),
                    P(
                        "pages",
                        "页码范围",
                        "text",
                        "全部",
                        help_text="例如 1-3,5 或 3-；页码从 1 开始，输入“全部”处理所有页面",
                    ),
                    P(
                        "format",
                        "输出格式",
                        "choice",
                        "auto",
                        choices=(
                            ("auto", "自动（推荐）"),
                            ("png", "PNG"),
                            ("jpg", "JPG"),
                        ),
                        help_text=(
                            "自动模式尽量保留原始图片编码，高清渲染区域则使用 PNG；"
                            "PNG 无损且支持透明，JPG 体积通常更小"
                        ),
                    ),
                    P(
                        "dpi",
                        "可见 / 智能模式清晰度 DPI",
                        "integer",
                        300,
                        minimum=72,
                        maximum=1200,
                        help_text=(
                            "仅影响可见、智能及包含它们的组合模式；300 DPI 适合通用提取，"
                            "600 DPI 适合小字或细线，越高越耗时且占用更多内存"
                        ),
                    ),
                    P(
                        "jpeg_quality",
                        "JPG 质量",
                        "integer",
                        95,
                        minimum=30,
                        maximum=100,
                        help_text="仅在输出或转换为 JPG 时生效；越高越清晰，文件也越大",
                    ),
                    P(
                        "min_width",
                        "最小图片宽度（像素）",
                        "integer",
                        1,
                        minimum=1,
                        maximum=100000,
                        help_text=(
                            "小于此原生宽度的图片会被过滤；可见 / 智能模式按所选 DPI "
                            "的渲染宽度判断，设为 32–64 可过滤小图标"
                        ),
                    ),
                    P(
                        "min_height",
                        "最小图片高度（像素）",
                        "integer",
                        1,
                        minimum=1,
                        maximum=100000,
                        help_text=(
                            "小于此原生高度的图片会被过滤；可见 / 智能模式按所选 DPI "
                            "的渲染高度判断，保留 1 可避免漏图"
                        ),
                    ),
                    P(
                        "merge_gap",
                        "相邻图块合并距离（PDF 点）",
                        "number",
                        4,
                        minimum=0,
                        maximum=72,
                        help_text=(
                            "仅影响智能及包含智能的组合模式；1 PDF 点 = 1/72 英寸，"
                            "数值越大越容易拼合碎片，过大可能合并无关图片"
                        ),
                    ),
                    P(
                        "region_padding",
                        "提取区域外扩边距（PDF 点）",
                        "number",
                        2,
                        minimum=0,
                        maximum=72,
                        help_text=(
                            "仅影响可见、智能及包含它们的组合模式；适当外扩可避免裁掉"
                            "边缘，过大则可能带入邻近文字"
                        ),
                    ),
                    P(
                        "deduplicate",
                        "去除重复图片",
                        "boolean",
                        True,
                        help_text="按图片字节内容在整个 PDF 和所选模式间去重，清单仍会记录重复出现位置",
                    ),
                    P(
                        "include_annotations",
                        "可见 / 智能模式包含批注",
                        "boolean",
                        False,
                        help_text="仅在可见、智能及包含它们的组合模式中，将页面批注一并渲染",
                    ),
                    P(
                        "write_manifest",
                        "生成图片提取清单",
                        "boolean",
                        True,
                        help_text="生成 JSON 清单，记录页码、位置、尺寸、去重引用和提取警告",
                    ),
                    P(
                        "password",
                        "PDF 打开密码",
                        "password",
                        "",
                        help_text="仅加密 PDF 需要；密码不会写入图片提取清单",
                    ),
                ),
                fidelity="extract",
                capability_probe=lambda: engines.require_modules(
                    "PyMuPDF + Pillow", "pymupdf", "PIL"
                ),
                notes=(
                    "“原始资源”尽量保持 PDF 内嵌图片的原编码；“可见位置”和“智能合并”会按所选 DPI 渲染。"
                    "文字和纯矢量图不是内嵌位图；如需整页图片，请使用“PDF 转图片”。"
                    "每个输入 PDF 会自动建立“<文件名>_PDF图片”独立目录；重名时自动避让，不覆盖历史结果。"
                ),
            ),
            Operation(
                "pdf.to_text",
                "PDF 格式转换",
                "PDF 转 TXT",
                "提取数字版 PDF 的文字并按页分隔。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "文本",
                    ".txt",
                    lambda src, target: pdf.pdf_to_text(
                        src, target, password=p["password"] or None, layout=p["layout"]
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "layout",
                        "尽量保留空格布局",
                        "boolean",
                        False,
                        help_text="适合代码或简单表格",
                    ),
                    P("password", "PDF 打开密码", "password", ""),
                ),
                fidelity="extract",
                capability_probe=extract,
                notes="扫描页不会凭空产生文字；此时请使用 OCR。",
            ),
            Operation(
                "pdf.to_html",
                "PDF 格式转换",
                "PDF 转 HTML",
                "生成可搜索、适合阅读的语义化网页，并可提取表格。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "网页",
                    ".html",
                    lambda src, target: conversion.pdf_to_html(
                        src,
                        target,
                        password=p["password"] or None,
                        include_tables=p["include_tables"],
                    ),
                ),
                PDF_EXT,
                (
                    P("include_tables", "提取表格", "boolean", True),
                    P("password", "PDF 打开密码", "password", ""),
                ),
                fidelity="extract",
                capability_probe=extract,
                notes="语义网页便于编辑和搜索，但不会逐像素复刻 PDF 固定版式。",
            ),
            Operation(
                "image.to_pdf",
                "PDF 格式转换",
                "多张图片合并为 PDF",
                "按文件列表顺序生成一个 PDF，自动处理透明背景。",
                lambda paths, out, p: pdf.images_to_pdf(
                    paths,
                    unique_path(out / f"{safe_filename(p['filename'])}.pdf"),
                    dpi=p["dpi"],
                    background=_rgb(p["background"]),
                ),
                IMAGE_EXT,
                (
                    P("filename", "输出文件名", "text", "图片合集", required=True),
                    P("dpi", "图片 DPI", "number", 96, minimum=36, maximum=1200),
                    P("background", "透明区域背景色", "color", "white"),
                ),
                min_inputs=1,
                fidelity="visual",
                capability_probe=lambda: engines.require_modules(
                    "Pillow + pypdf", "PIL", "pypdf"
                ),
            ),
        ]
    )

    operations.extend(
        [
            Operation(
                "pdf.merge",
                "PDF 页面与安全",
                "PDF 合并",
                "按文件列表顺序合并多个 PDF。",
                lambda paths, out, p: pdf.merge_pdfs(
                    paths, unique_path(out / f"{safe_filename(p['filename'])}.pdf")
                ),
                PDF_EXT,
                (P("filename", "输出文件名", "text", "合并文档", required=True),),
                min_inputs=2,
                fidelity="transform",
                capability_probe=core,
                notes="页面内容不会重新渲染；书签、表单、附件、标签结构等文档级目录不保证跨文件合并。",
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.split",
                "PDF 页面与安全",
                "PDF 拆分",
                "按页码段拆成多个 PDF，例如 1-3,4-7,8-。",
                lambda paths, out, p: _batch_to_directory(
                    paths,
                    out,
                    "拆分",
                    lambda src, target_dir: pdf.split_pdf(
                        src,
                        p["ranges"],
                        target_dir,
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "ranges",
                        "拆分页码段",
                        "text",
                        "1-",
                        required=True,
                        help_text="多个范围用逗号分隔，如 1-3,4-7",
                    ),
                ),
                fidelity="transform",
                capability_probe=core,
                notes="拆分页面本身不会重新渲染；书签、表单、附件、标签结构等文档级目录不会承诺完整保留。",
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.extract_pages",
                "PDF 页面与安全",
                "提取 PDF 页面",
                "把指定页提取为新的 PDF。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "提取页",
                    ".pdf",
                    lambda src, target: pdf.extract_pages(src, p["pages"], target),
                ),
                PDF_EXT,
                (
                    P(
                        "pages",
                        "页码",
                        "text",
                        "1",
                        required=True,
                        help_text="如 1-3,5,8-",
                    ),
                ),
                fidelity="transform",
                capability_probe=core,
                notes="提取页面本身不会重新渲染；书签、表单、附件、标签结构等文档级目录不会承诺完整保留。",
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.delete_pages",
                "PDF 页面与安全",
                "删除 PDF 页面",
                "移除指定页并另存为新文件。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "已删页",
                    ".pdf",
                    lambda src, target: pdf.delete_pages(src, p["pages"], target),
                ),
                PDF_EXT,
                (P("pages", "要删除的页码", "text", "1", required=True),),
                fidelity="transform",
                capability_probe=core,
                notes="删除页面本身不会重新渲染；书签、表单、附件、标签结构等文档级目录不会承诺完整保留。",
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.insert_pages",
                "PDF 页面与安全",
                "插入 PDF 页面",
                "选择两个 PDF：第一个为主文档，第二个插入到指定位置。",
                lambda paths, out, p: pdf.insert_pages(
                    paths[0],
                    paths[1],
                    p["position"],
                    unique_path(_target(out, paths[0], "已插页", ".pdf")),
                    pages=p["pages"] or None,
                ),
                PDF_EXT,
                (
                    P("position", "插入到第几页之前", "integer", 1, minimum=1),
                    P("pages", "插入文档页码（留空为全部）", "text", ""),
                ),
                min_inputs=2,
                max_inputs=2,
                fidelity="transform",
                capability_probe=core,
                notes="插入页面本身不会重新渲染；书签、表单、附件、标签结构等文档级目录不会承诺完整保留。",
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.rotate",
                "PDF 页面与安全",
                "旋转 PDF 页面",
                "按 90° 倍数旋转指定页面，不重新渲染。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "已旋转",
                    ".pdf",
                    lambda src, target: pdf.rotate_pages(
                        src,
                        p["pages"],
                        int(p["angle"]),
                        target,
                    ),
                ),
                PDF_EXT,
                (
                    P("pages", "页码", "text", "全部"),
                    P(
                        "angle",
                        "旋转角度",
                        "choice",
                        "90",
                        choices=(
                            ("90", "顺时针 90°"),
                            ("180", "180°"),
                            ("270", "顺时针 270°"),
                        ),
                    ),
                ),
                fidelity="lossless",
                capability_probe=core,
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.compress",
                "PDF 页面与安全",
                "PDF 无损压缩",
                "压缩可安全解码的内容流并清理重复对象，不降低图片分辨率。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "已压缩",
                    ".pdf",
                    lambda src, target: pdf.compress_pdf(
                        src, target, level=int(p["level"])
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "level",
                        "结构压缩强度",
                        "choice",
                        "9",
                        choices=(
                            ("0", "仅整理结构（最快）"),
                            ("6", "标准压缩"),
                            ("9", "最大压缩（推荐）"),
                        ),
                        help_text="只影响 PDF 内容流的无损 Deflate 强度；不会降低图片清晰度。",
                    ),
                ),
                fidelity="lossless",
                capability_probe=core,
                notes="超大内容流会保留原编码并继续处理其余页面；无损优化幅度仍取决于源文件结构。",
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.compress_lossy",
                "PDF 页面与安全",
                "PDF 高精度有损压缩",
                "默认仅优化高成本图片并保留文字、字体、矢量、链接与书签；扫描件可选整页兼容压缩。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "高精度有损压缩",
                    ".pdf",
                    lambda src, target: pdf.compress_pdf_lossy(
                        src,
                        target,
                        strategy=p["strategy"],
                        dpi=int(p["dpi"]),
                        jpeg_quality=int(p["jpeg_quality"]),
                        color_mode=p["color_mode"],
                        password=p["password"] or None,
                        poppler_path=engines.poppler_bin_path(),
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "strategy",
                        "压缩策略",
                        "choice",
                        "smart",
                        choices=(
                            ("smart", "结构保留智能压缩（推荐）"),
                            ("raster", "整页栅格兼容压缩（扫描件）"),
                        ),
                        help_text=(
                            "结构保留模式只优化高 DPI 位图，文字仍可搜索复制，并保留字体、矢量、链接、"
                            "书签、表单与批注；整页栅格模式兼容疑难或纯扫描 PDF，但所有页面会变成图片。"
                        ),
                    ),
                    P(
                        "dpi",
                        "目标图片清晰度",
                        "choice",
                        "220",
                        choices=(
                            ("150", "150 DPI｜屏幕阅读 / 更小体积"),
                            ("180", "180 DPI｜均衡"),
                            ("220", "220 DPI｜高清（推荐）"),
                            ("300", "300 DPI｜打印 / 最大清晰度"),
                        ),
                        help_text=(
                            "结构保留模式仅下采样明显高于此值的图片，不改变文字和矢量；"
                            "整页栅格模式则按此 DPI 渲染整页。"
                        ),
                    ),
                    P(
                        "jpeg_quality",
                        "照片 JPEG 质量",
                        "choice",
                        "88",
                        choices=(
                            ("72", "72｜体积优先"),
                            ("82", "82｜均衡"),
                            ("88", "88｜高清（推荐）"),
                            ("93", "93｜极高清"),
                        ),
                        help_text=(
                            "只影响照片类有损位图；PNG/Flate 图示保持无损编码，黑白线稿不转 JPEG。"
                            "通常 82–88 已适合文档。"
                        ),
                    ),
                    P(
                        "color_mode",
                        "色彩模式",
                        "choice",
                        "color",
                        choices=(
                            ("color", "彩色（保留原色）"),
                            ("grayscale", "灰度（黑白文档更省空间）"),
                        ),
                        help_text="灰度适合纯文字和黑白扫描件；彩色图表、印章和照片应选彩色。",
                    ),
                    P(
                        "password",
                        "PDF 打开密码（可选）",
                        "password",
                        "",
                        help_text="仅在源 PDF 已加密时填写。",
                    ),
                ),
                fidelity="visual",
                capability_probe=core,
                notes=(
                    "推荐模式不整页栅格化，也不使用 ReportLab/ASCII85 重建；会启用 PDF 对象清理、"
                    "Deflate 与对象流压缩。纯扫描件或结构异常文件可手动选择整页栅格兼容模式。"
                ),
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.encrypt",
                "PDF 页面与安全",
                "PDF 加密",
                "设置打开密码和独立所有者密码。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "已加密",
                    ".pdf",
                    lambda src, target: pdf.encrypt_pdf(
                        src,
                        target,
                        p["user_password"],
                        owner_password=p["owner_password"] or None,
                        algorithm=p["algorithm"],
                        allow_print=p["allow_print"],
                        allow_modify=p["allow_modify"],
                        allow_copy=p["allow_copy"],
                        allow_annotate=p["allow_annotate"],
                        allow_fill_forms=p["allow_fill_forms"],
                        allow_assemble=p["allow_assemble"],
                        password=p["source_password"] or None,
                    ),
                ),
                PDF_EXT,
                (
                    P("user_password", "新打开密码", "password", "", required=True),
                    P(
                        "owner_password",
                        "所有者 / 权限密码",
                        "password",
                        "",
                        help_text="限制打印、复制或编辑权限时，必须与打开密码不同。",
                    ),
                    P(
                        "algorithm",
                        "加密算法",
                        "choice",
                        "AES-256-R5",
                        choices=(
                            ("AES-256-R5", "AES-256"),
                            ("AES-128", "AES-128"),
                        ),
                    ),
                    P("allow_print", "允许打印", "boolean", True),
                    P("allow_modify", "允许修改内容", "boolean", True),
                    P("allow_copy", "允许复制 / 提取内容", "boolean", True),
                    P("allow_annotate", "允许批注", "boolean", True),
                    P("allow_fill_forms", "允许填写表单", "boolean", True),
                    P("allow_assemble", "允许插入 / 删除 / 旋转页面", "boolean", True),
                    P("source_password", "原 PDF 密码", "password", ""),
                ),
                fidelity="lossless",
                capability_probe=core,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.decrypt",
                "PDF 页面与安全",
                "PDF 解密（已知密码）",
                "使用已知密码移除保护，不提供密码破解。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "已解密",
                    ".pdf",
                    lambda src, target: pdf.decrypt_pdf(src, target, p["password"]),
                ),
                PDF_EXT,
                (P("password", "现有密码", "password", "", required=True),),
                fidelity="lossless",
                capability_probe=core,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.watermark",
                "PDF 页面与安全",
                "PDF 添加水印",
                "添加半透明文字或图片水印；每页可设置 1–100 个并自动均匀铺排。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "水印",
                    ".pdf",
                    lambda src, target: pdf.add_watermark(
                        src,
                        target,
                        text=p["text"] or None,
                        image_path=p["image_path"],
                        pages=p["pages"],
                        opacity=p["opacity"],
                        angle=p["angle"],
                        scale=p["scale"],
                        font_size=p["font_size"],
                        count=p["count"],
                    ),
                ),
                PDF_EXT,
                (
                    P("text", "水印文字", "text", "仅供内部使用"),
                    P("image_path", "水印图片（可选）", "path", None),
                    P("pages", "页码", "text", "全部"),
                    P(
                        "count",
                        "每页水印数量",
                        "integer",
                        1,
                        minimum=1,
                        maximum=100,
                        help_text="大字号建议 1 个，中小字号可选 4、6、9 个；多个时自动均匀铺排并缩放防重叠",
                    ),
                    P("opacity", "透明度 0-1", "number", 0.2, minimum=0, maximum=1),
                    P("angle", "旋转角度", "number", 45, minimum=-360, maximum=360),
                    P("scale", "图片宽度比例", "number", 0.35, minimum=0.02, maximum=1),
                    P("font_size", "文字字号", "number", 48, minimum=6, maximum=300),
                ),
                fidelity="transform",
                capability_probe=core,
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.header_footer",
                "PDF 页面与安全",
                "PDF 页眉页脚与页码",
                "批量添加页眉、页脚和页码。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "页眉页脚",
                    ".pdf",
                    lambda src, target: pdf.add_header_footer(
                        src,
                        target,
                        header=p["header"] or None,
                        footer=p["footer"] or None,
                        add_page_numbers=p["page_numbers"],
                        page_number_format=p["page_format"],
                        pages=p["pages"],
                        font_size=p["font_size"],
                        margin=p["margin"],
                    ),
                ),
                PDF_EXT,
                (
                    P("header", "页眉文字", "text", ""),
                    P("footer", "页脚文字", "text", ""),
                    P("page_numbers", "添加页码", "boolean", True),
                    P(
                        "page_format",
                        "页码格式",
                        "text",
                        "{page}/{total}",
                        help_text="可用 {page} 与 {total}",
                    ),
                    P("pages", "页码范围", "text", "全部"),
                    P("font_size", "字号", "number", 10, minimum=6, maximum=72),
                    P("margin", "页边距（点）", "number", 24, minimum=0, maximum=200),
                ),
                fidelity="transform",
                capability_probe=core,
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.visual_signature",
                "PDF 页面与安全",
                "PDF 手写签名图片",
                "把透明签名图片叠加到指定页面坐标。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "已签名",
                    ".pdf",
                    lambda src, target: conversion.add_visual_signature(
                        src,
                        p["image_path"],
                        target,
                        pages=p["pages"],
                        x=p["x"],
                        y=p["y"],
                        width=p["width"],
                        opacity=p["opacity"],
                    ),
                ),
                PDF_EXT,
                (
                    P("image_path", "签名图片", "path", None, required=True),
                    P("pages", "页码", "text", "1"),
                    P("x", "左下角 X（点）", "number", 36),
                    P("y", "左下角 Y（点）", "number", 36),
                    P("width", "签名宽度（点）", "number", 120, minimum=10),
                    P("opacity", "透明度", "number", 1, minimum=0, maximum=1),
                ),
                fidelity="transform",
                capability_probe=core,
                notes="这是视觉签名，不等同于使用证书的密码学数字签名。",
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.add_note",
                "PDF 页面与安全",
                "PDF 添加文本批注",
                "在指定页添加支持 Unicode 内容的标准文本便笺批注。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "批注",
                    ".pdf",
                    lambda src, target: conversion.add_pdf_note(
                        src,
                        target,
                        page=p["page"],
                        text=p["text"],
                        x=p["x"],
                        y=p["y"],
                        width=p["width"],
                        height=p["height"],
                    ),
                ),
                PDF_EXT,
                (
                    P("page", "页码", "integer", 1, minimum=1),
                    P("text", "批注文字", "text", "请复核此处", required=True),
                    P("x", "X 坐标", "number", 36),
                    P("y", "Y 坐标", "number", 36),
                    P("width", "宽度", "number", 220, minimum=20),
                    P("height", "高度", "number", 90, minimum=20),
                ),
                fidelity="transform",
                capability_probe=core,
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.add_markup",
                "PDF 页面与安全",
                "PDF 高亮 / 下划线 / 删除线",
                "在指定页面矩形区域添加标准 PDF 文本标记批注。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "文本标记",
                    ".pdf",
                    lambda src, target: conversion.add_pdf_markup(
                        src,
                        target,
                        page=p["page"],
                        kind=p["kind"],
                        x=p["x"],
                        y=p["y"],
                        width=p["width"],
                        height=p["height"],
                        color=p["color"],
                        opacity=p["opacity"],
                        comment=p["comment"],
                    ),
                ),
                PDF_EXT,
                (
                    P("page", "页码", "integer", 1, minimum=1),
                    P(
                        "kind",
                        "标记类型",
                        "choice",
                        "highlight",
                        choices=(
                            ("highlight", "高亮"),
                            ("underline", "下划线"),
                            ("strikeout", "删除线"),
                        ),
                    ),
                    P("x", "左下角 X", "number", 36, minimum=0),
                    P("y", "左下角 Y", "number", 36, minimum=0),
                    P("width", "区域宽度", "number", 220, minimum=1),
                    P("height", "区域高度", "number", 24, minimum=1),
                    P("color", "标记颜色", "color", "#ffff00"),
                    P("opacity", "透明度", "number", 0.45, minimum=0, maximum=1),
                    P("comment", "批注说明（可选）", "text", ""),
                ),
                fidelity="transform",
                capability_probe=core,
                notes="坐标单位为 PDF 点（1/72 英寸），原点在页面左下角。",
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.fill_form",
                "PDF 页面与安全",
                "填写 PDF 表单",
                "通过 JSON 字段名和值填写 AcroForm 表单，可选扁平化。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "已填写",
                    ".pdf",
                    lambda src, target: conversion.fill_pdf_form(
                        src,
                        target,
                        p["fields"],
                        flatten=p["flatten"],
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "fields",
                        "字段 JSON",
                        "text",
                        "{}",
                        required=True,
                        help_text='例如 {"姓名":"张三"}',
                    ),
                    P("flatten", "填写后扁平化", "boolean", False),
                ),
                fidelity="transform",
                capability_probe=core,
                notes="支持标准 AcroForm；Adobe XFA 动态表单不在本地开源引擎保证范围内。",
                reject_encrypted_pdf_inputs=True,
                reject_signed_pdf_inputs=True,
            ),
            Operation(
                "pdf.ocr",
                "PDF 页面与安全",
                "PDF OCR 识别",
                "把扫描 PDF 识别为可编辑 Word 或纯文本。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "OCR",
                    f".{p['format']}",
                    lambda src, target: conversion.ocr_pdf(
                        src,
                        target,
                        output_format=p["format"],
                        language=p["language"],
                        dpi=p["dpi"],
                        password=p["password"] or None,
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "format",
                        "输出格式",
                        "choice",
                        "docx",
                        choices=(("docx", "Word DOCX"), ("txt", "TXT")),
                    ),
                    P(
                        "language",
                        "识别语言",
                        "text",
                        "chi_sim+eng",
                        help_text="Tesseract 示例：chi_sim+eng",
                    ),
                    P(
                        "dpi",
                        "识别清晰度 DPI",
                        "integer",
                        300,
                        minimum=150,
                        maximum=600,
                    ),
                    P("password", "PDF 打开密码（可选）", "password", ""),
                ),
                fidelity="extract",
                capability_probe=engines.ocr_capability,
                notes=(
                    "中文扫描表格建议使用带版面分析的 PaddleOCR；"
                    "当前 DOCX/TXT 输出不携带逐行置信度，关键数据必须人工复核。"
                ),
            ),
            Operation(
                "pdf.digital_signature",
                "PDF 页面与安全",
                "PDF 证书数字签名",
                "使用证书、私钥和时间戳服务生成可信 PAdES 签名。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "数字签名",
                    ".pdf",
                    lambda src, target: signature.sign_pdf(
                        src,
                        target,
                        p["certificate_path"],
                        p["certificate_password"],
                        field_name=p["field_name"],
                        page=p["page"],
                        box=(p["left"], p["bottom"], p["right"], p["top"]),
                        reason=p["reason"] or None,
                        location=p["location"] or None,
                        contact_info=p["contact_info"] or None,
                        timestamp_url=p["timestamp_url"] or None,
                    ),
                ),
                PDF_EXT,
                (
                    P(
                        "certificate_path",
                        "P12 / PFX 证书",
                        "path",
                        None,
                        required=True,
                    ),
                    P(
                        "certificate_password",
                        "证书密码",
                        "password",
                        "",
                        required=True,
                    ),
                    P("field_name", "签名域名称", "text", "Signature1", required=True),
                    P("page", "签名页码", "integer", 1, minimum=1),
                    P("left", "签名框左坐标", "number", 36),
                    P("bottom", "签名框下坐标", "number", 36),
                    P("right", "签名框右坐标", "number", 220),
                    P("top", "签名框上坐标", "number", 100),
                    P("reason", "签名原因（可选）", "text", ""),
                    P("location", "签名地点（可选）", "text", ""),
                    P("contact_info", "联系信息（可选）", "text", ""),
                    P("timestamp_url", "时间戳服务 URL（可选）", "text", ""),
                ),
                fidelity="transform",
                capability_probe=lambda: engines.format_capability(
                    "pyhanko", "pyHanko PAdES"
                ),
                notes="证书数字签名与手写图片签名不同；软件不会代替用户生成或保管私钥。",
                reject_encrypted_pdf_inputs=True,
            ),
        ]
    )
    return operations


def _image_operations() -> list[Operation]:
    from .processors import conversion, image

    ready = engines.image_capability
    position_choices = (
        ("top-left", "左上"),
        ("top", "上中"),
        ("top-right", "右上"),
        ("center", "居中"),
        ("bottom-left", "左下"),
        ("bottom", "下中"),
        ("bottom-right", "右下"),
    )
    return [
        Operation(
            "image.convert",
            "图片格式与批处理",
            "图片格式转换",
            "批量转换 JPG、PNG、BMP、WebP、TIFF，并可扁平化读取 PSD。",
            lambda paths, out, p: image.convert_format(
                paths,
                p["format"],
                out,
                quality=p["quality"],
                background=p["background"],
            ),
            IMAGE_EXT,
            (
                P(
                    "format",
                    "目标格式",
                    "choice",
                    "png",
                    choices=(
                        ("jpg", "JPG"),
                        ("png", "PNG"),
                        ("webp", "WebP"),
                        ("bmp", "BMP"),
                        ("tiff", "TIFF"),
                    ),
                ),
                P("quality", "有损格式质量", "integer", 95, minimum=1, maximum=100),
                P("background", "透明转 JPG 背景色", "color", "white"),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.resize",
            "图片尺寸与裁剪",
            "调整图片尺寸",
            "批量设置像素宽高，可选择等比例缩放。",
            lambda paths, out, p: image.resize_images(
                paths, (p["width"], p["height"]), out, keep_aspect=p["keep_aspect"]
            ),
            IMAGE_EXT,
            (
                P(
                    "width",
                    "宽度（像素）",
                    "integer",
                    1920,
                    minimum=1,
                    maximum=100000,
                    help_text="结果允许占用的最大宽度；开启等比例时不会拉伸图片",
                ),
                P(
                    "height",
                    "高度（像素）",
                    "integer",
                    1080,
                    minimum=1,
                    maximum=100000,
                    help_text="结果允许占用的最大高度；右侧预览会同步显示实际比例",
                ),
                P(
                    "keep_aspect",
                    "保持原始宽高比",
                    "boolean",
                    True,
                    help_text="推荐开启，避免人物、文字和图形被横向或纵向拉伸",
                ),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.scale",
            "图片尺寸与裁剪",
            "按百分比缩放",
            "按原始尺寸百分比批量缩放。",
            lambda paths, out, p: image.scale_images(paths, p["percent"] / 100.0, out),
            IMAGE_EXT,
            (
                P(
                    "percent",
                    "缩放百分比",
                    "number",
                    50,
                    minimum=1,
                    maximum=1000,
                    help_text="100 表示原尺寸，50 表示宽高各缩小一半，200 表示宽高各放大一倍",
                ),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.crop",
            "图片尺寸与裁剪",
            "裁剪图片",
            "按左、上、右、下像素坐标裁剪。",
            lambda paths, out, p: image.crop_images(
                paths, (p["left"], p["top"], p["right"], p["bottom"]), out
            ),
            IMAGE_EXT,
            (
                P(
                    "left",
                    "左边界（像素）",
                    "integer",
                    0,
                    minimum=0,
                    help_text="从原图左侧向右数多少像素开始保留；0 表示从最左侧开始",
                ),
                P(
                    "top",
                    "上边界（像素）",
                    "integer",
                    0,
                    minimum=0,
                    help_text="从原图顶部向下数多少像素开始保留；0 表示从最顶部开始",
                ),
                P(
                    "right",
                    "右边界（像素）",
                    "integer",
                    800,
                    minimum=1,
                    help_text="保留区域结束位置，不是要裁掉的宽度；必须大于左边界",
                ),
                P(
                    "bottom",
                    "下边界（像素）",
                    "integer",
                    600,
                    minimum=1,
                    help_text="保留区域结束位置，不是要裁掉的高度；必须大于上边界",
                ),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.rotate",
            "图片尺寸与裁剪",
            "旋转图片",
            "支持任意角度旋转，可扩展画布防止裁切。",
            lambda paths, out, p: image.rotate_images(
                paths,
                p["angle"],
                out,
                expand=p["expand"],
                fillcolor=p["background"] or None,
            ),
            IMAGE_EXT,
            (
                P(
                    "angle",
                    "逆时针角度",
                    "number",
                    90,
                    minimum=-3600,
                    maximum=3600,
                    help_text="正数逆时针旋转，负数顺时针旋转；90、180、270 度最清晰",
                ),
                P(
                    "expand",
                    "扩展画布",
                    "boolean",
                    True,
                    help_text="推荐开启，旋转后自动扩大画布，避免图片四角被切掉",
                ),
                P(
                    "background",
                    "空白区域颜色",
                    "color",
                    "white",
                    help_text="仅在旋转后出现空白区域时使用",
                ),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.flip",
            "图片尺寸与裁剪",
            "翻转图片",
            "批量水平或垂直翻转。",
            lambda paths, out, p: image.flip_images(paths, p["direction"], out),
            IMAGE_EXT,
            (
                P(
                    "direction",
                    "方向",
                    "choice",
                    "horizontal",
                    choices=(("horizontal", "水平翻转"), ("vertical", "垂直翻转")),
                ),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.compress",
            "图片压缩与优化",
            "图片批量压缩",
            "按质量压缩，或迭代逼近目标文件大小。",
            lambda paths, out, p: image.compress_images(
                paths,
                out,
                quality=p["quality"],
                max_bytes=p["target_kb"] * 1024 if p["target_kb"] else None,
                min_quality=p["min_quality"],
                allow_resize=p["allow_resize"],
            ),
            IMAGE_EXT,
            (
                P("quality", "初始质量", "integer", 88, minimum=1, maximum=100),
                P("target_kb", "目标大小 KB（0 为不限定）", "integer", 0, minimum=0),
                P("min_quality", "最低质量", "integer", 45, minimum=1, maximum=100),
                P("allow_resize", "必要时缩小像素", "boolean", False),
            ),
            capability_probe=ready,
            notes="目标大小是尽力逼近；固定体积与完全无损无法同时保证。",
        ),
        Operation(
            "image.remove_exif",
            "图片压缩与优化",
            "去除图片元数据",
            "删除 EXIF 拍摄时间、地点和相机参数，并固化显示方向。",
            lambda paths, out, p: image.remove_exif(paths, out),
            IMAGE_EXT,
            capability_probe=ready,
        ),
        Operation(
            "image.enhance",
            "图片效果与编辑",
            "高保真 AI 清晰增强",
            "Real-ESRGAN NCNN Vulkan 结合保守去噪、局部对比度、限幅锐化和增强后二重检测；异常区域自动恢复为真实像素放大结果。",
            lambda paths, out, p: image.enhance_images(
                paths,
                out,
                mode=p["mode"],
                content_type=p["content_type"],
                scale=int(p["scale"]),
                max_dimension=p["max_dimension"],
                output_format=p["output_format"],
            ),
            IMAGE_EXT,
            (
                P(
                    "mode",
                    "增强方式",
                    "choice",
                    "auto",
                    choices=(
                        ("auto", "自动：有可用显卡则 AI，否则兼容增强（推荐）"),
                        ("compatible", "无独显兼容增强（CPU / OpenCV）"),
                        ("gpu_ai", "GPU AI 高清增强（Real-ESRGAN）"),
                    ),
                    help_text="没有独立显卡请选择兼容增强；自动模式仅在检测到 Vulkan GPU 时调用 AI，失败也会安全回退。",
                ),
                P(
                    "content_type",
                    "图片内容",
                    "choice",
                    "auto",
                    choices=(
                        ("auto", "自动判断（推荐）"),
                        ("document", "文字 / 课件 / 表格 / 扫描件"),
                        ("photo", "照片 / 普通图像"),
                    ),
                    help_text="严谨文字、公式和表格建议显式选择“文字 / 课件”，加强真实笔画保护。",
                ),
                P(
                    "scale",
                    "AI 放大倍率",
                    "choice",
                    "2",
                    choices=(("2", "2 倍（推荐）"), ("4", "4 倍（更慢、文件更大）")),
                    help_text="2 倍更适合文字与课件；4 倍主要用于分辨率很低的普通图片。",
                ),
                P(
                    "max_dimension",
                    "输出最长边上限",
                    "integer",
                    4096,
                    minimum=1024,
                    maximum=8192,
                    help_text=(
                        "最终尺寸同时受放大倍率和此上限约束；例如 4032 像素原图要完整输出 2 倍，"
                        "需将上限设为 8192。保持 4096 时只进行接近原尺寸的细节增强，可避免超大文件。"
                    ),
                ),
                P(
                    "output_format",
                    "输出格式",
                    "choice",
                    "png",
                    choices=(("png", "PNG 无损（推荐）"), ("jpg", "JPG"), ("webp", "WebP")),
                    help_text="建议 PNG，避免增强完成后再次产生 JPEG 压缩噪声。",
                ),
            ),
            fidelity="visual",
            capability_probe=engines.image_enhancement_capability,
            notes="AI 不能恢复原图中从未存在的信息；软件会保护文字结构并对高风险区域自动回退，关键数字、公式和合同仍应抽查。",
        ),
        Operation(
            "image.adjust",
            "图片效果与编辑",
            "亮度 / 对比度 / 饱和度",
            "使用倍率批量调整画面。1.0 表示不变。",
            lambda paths, out, p: image.adjust_images(
                paths,
                out,
                brightness=p["brightness"],
                contrast=p["contrast"],
                saturation=p["saturation"],
            ),
            IMAGE_EXT,
            (
                P(
                    "brightness",
                    "亮度倍率",
                    "number",
                    1,
                    minimum=0,
                    maximum=5,
                    help_text="1.0 不变，小于 1 变暗，大于 1 变亮",
                ),
                P(
                    "contrast",
                    "对比度倍率",
                    "number",
                    1,
                    minimum=0,
                    maximum=5,
                    help_text="1.0 不变，提高后明暗区分更明显，过高会丢失细节",
                ),
                P(
                    "saturation",
                    "饱和度倍率",
                    "number",
                    1,
                    minimum=0,
                    maximum=5,
                    help_text="0 为灰度，1.0 不变，大于 1 会让颜色更鲜艳",
                ),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.filter",
            "图片效果与编辑",
            "图片滤镜",
            "批量应用黑白、复古、模糊、锐化、浮雕等滤镜。",
            lambda paths, out, p: image.apply_filter(
                paths, p["filter"], out, intensity=p["intensity"]
            ),
            IMAGE_EXT,
            (
                P(
                    "filter",
                    "滤镜",
                    "choice",
                    "grayscale",
                    choices=(
                        ("grayscale", "灰度"),
                        ("black_white", "黑白"),
                        ("sepia", "复古"),
                        ("gaussian_blur", "高斯模糊"),
                        ("sharpen", "锐化"),
                        ("emboss", "浮雕"),
                        ("find_edges", "边缘检测"),
                        ("smooth", "平滑"),
                    ),
                ),
                P(
                    "intensity",
                    "强度",
                    "number",
                    1,
                    minimum=0,
                    maximum=20,
                    help_text="不同滤镜含义略有差异，请以右侧实时效果为准；建议从 1.0 小幅调整",
                ),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.text_watermark",
            "图片效果与编辑",
            "图片文字水印",
            "批量添加支持中文的半透明文字水印。",
            lambda paths, out, p: image.add_text_watermark(
                paths,
                p["text"],
                out,
                position=p["position"],
                font_size=p["font_size"],
                font_path=p["font_path"],
                color=p["color"],
                opacity=p["opacity"],
                margin=p["margin"],
            ),
            IMAGE_EXT,
            (
                P("text", "水印文字", "text", "仅供内部使用", required=True),
                P(
                    "position",
                    "位置",
                    "choice",
                    "bottom-right",
                    choices=position_choices,
                ),
                P("font_size", "字号", "integer", 36, minimum=6, maximum=500),
                P("font_path", "字体文件（可选）", "path", None),
                P("color", "文字颜色", "color", "white"),
                P("opacity", "透明度", "number", 0.7, minimum=0, maximum=1),
                P("margin", "边距", "integer", 20, minimum=0),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.image_watermark",
            "图片效果与编辑",
            "图片图标水印",
            "批量叠加 Logo、印章或贴纸。",
            lambda paths, out, p: image.add_image_watermark(
                paths,
                p["watermark_path"],
                out,
                position=p["position"],
                opacity=p["opacity"],
                scale=p["scale"],
                margin=p["margin"],
            ),
            IMAGE_EXT,
            (
                P("watermark_path", "水印图片", "path", None, required=True),
                P(
                    "position",
                    "位置",
                    "choice",
                    "bottom-right",
                    choices=position_choices,
                ),
                P("opacity", "透明度", "number", 0.7, minimum=0, maximum=1),
                P("scale", "相对底图宽度比例", "number", 0.2, minimum=0.01, maximum=2),
                P("margin", "边距", "integer", 20, minimum=0),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.border",
            "图片效果与编辑",
            "添加图片边框",
            "为图片四周添加指定颜色和宽度的边框。",
            lambda paths, out, p: image.add_border(
                paths, p["width"], out, color=p["color"]
            ),
            IMAGE_EXT,
            (
                P(
                    "width",
                    "边框宽度",
                    "integer",
                    12,
                    minimum=0,
                    help_text="单位为原图像素，边框会添加在图片四周并增大最终尺寸",
                ),
                P("color", "边框颜色", "color", "black"),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.mosaic",
            "图片效果与编辑",
            "马赛克 / 打码",
            "对指定矩形区域进行像素化处理；留空坐标可处理整图。",
            lambda paths, out, p: image.mosaic_images(
                paths,
                (
                    (p["left"], p["top"], p["right"], p["bottom"])
                    if p["right"] > p["left"] and p["bottom"] > p["top"]
                    else None
                ),
                out,
                block_size=p["block_size"],
            ),
            IMAGE_EXT,
            (
                P("left", "左", "integer", 0, minimum=0),
                P("top", "上", "integer", 0, minimum=0),
                P("right", "右（0 表示整图）", "integer", 0, minimum=0),
                P("bottom", "下（0 表示整图）", "integer", 0, minimum=0),
                P("block_size", "马赛克块大小", "integer", 12, minimum=2, maximum=500),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.stitch",
            "图片效果与编辑",
            "图片拼接 / 长图",
            "把多张图片按列表顺序横向或纵向拼成一张长图。",
            lambda paths, out, p: image.stitch_images(
                paths,
                out / f"{safe_filename(p['filename'])}.{p['format']}",
                direction=p["direction"],
                spacing=p["spacing"],
                background=p["background"],
                alignment=p["alignment"],
            ),
            IMAGE_EXT,
            (
                P("filename", "输出文件名", "text", "拼接长图", required=True),
                P(
                    "format",
                    "格式",
                    "choice",
                    "png",
                    choices=(("png", "PNG"), ("jpg", "JPG"), ("webp", "WebP")),
                ),
                P(
                    "direction",
                    "方向",
                    "choice",
                    "vertical",
                    choices=(("vertical", "纵向"), ("horizontal", "横向")),
                ),
                P("spacing", "间距", "integer", 0, minimum=0),
                P("background", "背景色", "color", "white"),
                P(
                    "alignment",
                    "对齐",
                    "choice",
                    "center",
                    choices=(
                        ("start", "起始边"),
                        ("center", "居中"),
                        ("end", "结束边"),
                    ),
                ),
            ),
            min_inputs=2,
            capability_probe=ready,
        ),
        Operation(
            "image.overlay",
            "图片效果与编辑",
            "图片叠加",
            "把同一张贴纸或图层叠加到多张底图。",
            lambda paths, out, p: image.overlay_images(
                paths,
                p["overlay_path"],
                out,
                position=p["position"],
                opacity=p["opacity"],
                scale=p["scale"] or None,
                margin=p["margin"],
            ),
            IMAGE_EXT,
            (
                P("overlay_path", "叠加图片", "path", None, required=True),
                P("position", "位置", "choice", "center", choices=position_choices),
                P("opacity", "透明度", "number", 1, minimum=0, maximum=1),
                P(
                    "scale",
                    "宽度比例（0 保持原尺寸）",
                    "number",
                    0,
                    minimum=0,
                    maximum=5,
                ),
                P("margin", "边距", "integer", 0, minimum=0),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.rename",
            "图片格式与批处理",
            "批量重命名",
            "按模板统一命名，支持 {index}、{stem}、{suffix}。默认复制以保护原文件。",
            lambda paths, out, p: image.batch_rename(
                paths, p["pattern"], out, start=p["start"], move=p["move"]
            ),
            IMAGE_EXT,
            (
                P("pattern", "命名模板", "text", "photo_{index:03d}", required=True),
                P("start", "起始序号", "integer", 1, minimum=0),
                P("move", "移动原文件而非复制", "boolean", False),
            ),
            capability_probe=ready,
        ),
        Operation(
            "image.remove_background",
            "图片效果与编辑",
            "AI 抠图 / 去背景",
            "使用本地 AI 模型输出透明 PNG。",
            lambda paths, out, p: conversion.remove_background(paths, out),
            IMAGE_EXT,
            capability_probe=engines.background_removal_capability,
            notes="毛发、透明物体和前景背景颜色相近时仍需人工复核。",
        ),
        Operation(
            "image.heic",
            "图片格式与批处理",
            "HEIC 转 JPG / PNG",
            "转换苹果照片格式，并处理 EXIF 方向。",
            lambda paths, out, p: conversion.heic_to_images(
                paths, out, target_format=p["format"]
            ),
            (".heic", ".heif"),
            (
                P(
                    "format",
                    "目标格式",
                    "choice",
                    "jpg",
                    choices=(("jpg", "JPG"), ("png", "PNG")),
                ),
            ),
            capability_probe=lambda: engines.format_capability(
                "pillow_heif", "libheif"
            ),
        ),
        Operation(
            "image.raw",
            "图片格式与批处理",
            "RAW 转 JPG / PNG",
            "使用相机白平衡冲洗常见 RAW 文件。",
            lambda paths, out, p: conversion.raw_to_images(
                paths, out, target_format=p["format"]
            ),
            (".dng", ".cr2", ".cr3", ".nef", ".arw", ".rw2", ".orf", ".raf"),
            (
                P(
                    "format",
                    "目标格式",
                    "choice",
                    "jpg",
                    choices=(("jpg", "JPG"), ("png", "PNG")),
                ),
            ),
            capability_probe=lambda: engines.format_capability("rawpy", "LibRaw"),
            notes="RAW 冲洗没有唯一正确结果；首版采用相机白平衡和标准色调。",
        ),
        Operation(
            "image.svg",
            "图片格式与批处理",
            "SVG 转 PNG / JPG",
            "将矢量图按指定倍率渲染为位图。",
            lambda paths, out, p: conversion.svg_to_images(
                paths, out, target_format=p["format"], scale=p["scale"]
            ),
            (".svg",),
            (
                P(
                    "format",
                    "目标格式",
                    "choice",
                    "png",
                    choices=(("png", "PNG"), ("jpg", "JPG")),
                ),
                P("scale", "渲染倍率", "number", 2, minimum=0.1, maximum=20),
            ),
            capability_probe=lambda: engines.format_capability("cairosvg", "CairoSVG"),
        ),
    ]


def _pdf_video_capability() -> Capability:
    renderer = engines.pdf_render_capability()
    video = engines.slideshow_video_capability()
    if not renderer.runnable:
        return Capability("unavailable", renderer.reason, "PDF 视频流水线")
    if not video.runnable:
        return Capability("unavailable", video.reason, "PDF 视频流水线")
    return Capability(
        "external",
        "使用 Poppler 渲染页面，再由 FFmpeg 编码视频",
        "Poppler + FFmpeg",
    )


def _video_operations() -> list[Operation]:
    from .processors import video

    def extract_lecture_slides(
        paths: Sequence[Path], output_dir: Path, parameters: Mapping[str, Any]
    ) -> list[Path]:
        from .processors import video_slides

        return _batch_to_directory(
            paths,
            output_dir,
            "高清幻灯片",
            lambda source, target_dir: video_slides.extract_slides_to_pptx(
                source,
                target_dir,
                scan_mode=parameters["scan_mode"],
                change_sensitivity=parameters["change_sensitivity"],
                crop_mode=parameters["crop_mode"],
                crop_rect=parameters["crop_rect"],
                watermark_search=parameters["watermark_search"],
                watermark_rect=parameters["watermark_rect"],
                watermark_text_hint=parameters["watermark_text_hint"],
                annotation_color_mode=parameters["annotation_color_mode"],
                annotation_colors=parameters["annotation_colors"],
                annotation_color_tolerance=parameters["annotation_color_tolerance"],
                fixed_watermark_regions=parameters["fixed_watermark_regions"],
                fixed_watermark_fill=parameters["fixed_watermark_fill"],
                fixed_watermark_fill_color=parameters["fixed_watermark_fill_color"],
                presenter_policy=parameters["presenter_policy"],
                presenter_rect=parameters["presenter_rect"],
                enhancement_mode=parameters["enhancement_mode"],
                keep_images=parameters["keep_images"],
                keep_report=parameters["keep_report"],
            ),
        )

    def repair_lecture_slides(
        paths: Sequence[Path], output_dir: Path, parameters: Mapping[str, Any]
    ) -> list[Path]:
        from .processors import video_slide_repair

        pptx_files = [path for path in paths if path.suffix.casefold() == ".pptx"]
        video_files = [path for path in paths if path.suffix.casefold() in VIDEO_EXT]
        if len(pptx_files) != 1 or len(video_files) != 1:
            raise ValidationError("请恰好添加一份已提取 PPTX 和一份对应原视频")
        return [
            video_slide_repair.repair_video_ppt(
                pptx_files[0],
                video_files[0],
                output_dir,
                parameters["repair_plan"],
            )
        ]

    resolution_choices = (
        ("720p", "720p"),
        ("1080p", "1080p"),
        ("1440p", "1440p"),
        ("2160p", "4K"),
    )
    shared_parameters = (
        P(
            "slide_duration",
            "每张 / 每页时长（秒）",
            "number",
            3,
            minimum=0.1,
            maximum=3600,
        ),
        P("fps", "帧率", "integer", 30, minimum=1, maximum=120),
        P(
            "resolution",
            "分辨率",
            "choice",
            "1080p",
            choices=resolution_choices,
        ),
        P(
            "transition",
            "转场",
            "choice",
            "fade",
            choices=(("none", "无"), ("fade", "淡入淡出")),
        ),
        P(
            "transition_duration",
            "转场时长（秒）",
            "number",
            0.5,
            minimum=0,
            maximum=30,
        ),
        P("background", "背景色", "color", "black"),
        P("audio_path", "背景音频（可选）", "path", None),
        P(
            "quality",
            "画质参数（越小越清晰）",
            "integer",
            18,
            minimum=0,
            maximum=40,
        ),
    )
    return [
        Operation(
            "video.extract_slides_ppt",
            "视频生成",
            "讲解视频提取高清幻灯片（不可编辑 PPT）",
            "识别换页后只保留同页印刷内容最完整的最终状态；按有效首次出现时间组合为一份完整 PPT，自动延后正式讲解前的极短误翻，通过同页多帧回溯与用户指定的 RGB 颜色辅助去除后加手写批注、移动/固定水印和鼠标，并输出高清 PNG、不可编辑 PPT 与 Word 二检报告。",
            extract_lecture_slides,
            VIDEO_EXT,
            (
                P(
                    "scan_mode",
                    "分析精度",
                    "choice",
                    "accurate",
                    choices=(
                        ("accurate", "高精度（推荐，约每 0.5 秒分析）"),
                        ("balanced", "均衡（约每 1 秒分析）"),
                        ("fast", "快速（约每 2 秒分析）"),
                    ),
                    help_text="一般保持“高精度”即可；只在长视频需要更快预览时降低。",
                    section="① 提取与换页",
                ),
                P(
                    "change_sensitivity",
                    "页面变化程度",
                    "choice",
                    "balanced",
                    choices=(
                        ("conservative", "保守：减少把动画误判为换页"),
                        ("balanced", "均衡（推荐）"),
                        ("sensitive", "灵敏：适合页面差异较小的课件"),
                    ),
                    help_text="普通课件选“均衡”；只有相邻页面非常相似且出现漏页时才选“灵敏”。",
                    section="① 提取与换页",
                ),
                P(
                    "crop_mode",
                    "课件画面范围",
                    "choice",
                    "auto",
                    choices=(
                        ("auto", "自动去除外侧黑边（推荐）"),
                        ("full", "保留完整视频画面"),
                        ("custom", "在视频画面中框选课件区域"),
                    ),
                    help_text="推荐自动。需要只保留视频中的某一块课件区域时选择“框选”。",
                    section="① 提取与换页",
                ),
                P(
                    "crop_rect",
                    "框选课件区域",
                    "region",
                    "",
                    help_text="点击右侧按钮，在视频画面中拖框；软件自动填写坐标。",
                    section="① 提取与换页",
                    visible_when=("crop_mode", ("custom",)),
                ),
                P(
                    "watermark_search",
                    "移动水印搜索范围",
                    "choice",
                    "auto",
                    choices=(
                        ("auto", "自动：检测顶部 42%（推荐，覆盖跳位广告）"),
                        ("top", "仅顶部 20%"),
                        ("bottom", "仅底部 20%"),
                        ("full", "全画面（更慢、误判风险较高）"),
                        ("custom", "在视频画面中框选水印范围"),
                        ("off", "关闭移动水印清理"),
                    ),
                    help_text="通常选自动；水印会移动到中部或底部时选“全画面”。",
                    section="③ 水印与讲师画面",
                ),
                P(
                    "watermark_rect",
                    "框选移动水印范围",
                    "region",
                    "",
                    help_text="点击右侧按钮，在水印可能经过的完整范围拖框。",
                    section="③ 水印与讲师画面",
                    visible_when=("watermark_search", ("custom",)),
                ),
                P(
                    "watermark_text_hint",
                    "水印文字提示（可选）",
                    "text",
                    "",
                    help_text="可填写反复出现的微信号或短语，仅用于报告和辅助判断。",
                    section="高级：固定水印精细修复",
                    advanced=True,
                ),
                P(
                    "annotation_color_mode",
                    "手写批注处理",
                    "choice",
                    "auto",
                    choices=(
                        ("auto", "自动识别（推荐）"),
                        ("manual", "使用视频取色笔（颜色明确时更准）"),
                        ("off", "关闭颜色辅助，仅使用多帧时序"),
                    ),
                    help_text="先用自动识别；若仍残留明显彩色笔迹，再选择视频取色笔。",
                    section="② 手写批注清理",
                ),
                P(
                    "annotation_colors",
                    "手写笔迹基准色",
                    "colors",
                    "#00AEEF",
                    help_text="拖动视频时间轴找到清晰笔迹，悬停放大观察，单击即可自动读取 RGB；可连续添加多种笔色。",
                    section="② 手写批注清理",
                    visible_when=("annotation_color_mode", ("manual",)),
                ),
                P(
                    "annotation_color_tolerance",
                    "颜色容差",
                    "integer",
                    24,
                    minimum=0,
                    maximum=100,
                    help_text="推荐 24。视频压缩严重、笔迹边缘仍残留时可逐步调至 30–36。",
                    section="② 手写批注清理",
                    visible_when=("annotation_color_mode", ("manual",)),
                ),
                P(
                    "fixed_watermark_regions",
                    "固定水印区域",
                    "region",
                    "",
                    help_text="点击“添加框选”可依次添加多个区域。只紧密框住水印笔画附近，避免覆盖标题正文。",
                    section="高级：固定水印精细修复",
                    advanced=True,
                ),
                P(
                    "fixed_watermark_fill",
                    "固定水印填充方式",
                    "choice",
                    "temporal",
                    choices=(
                        ("temporal", "多帧真实像素 → 背景建模（推荐）"),
                        ("background", "直接用周边背景建模覆盖"),
                        ("color", "使用下方指定纯色填充"),
                    ),
                    help_text="优先使用多帧真实像素；只有背景平滑且真实像素无法恢复时再改用背景建模。",
                    section="高级：固定水印精细修复",
                    advanced=True,
                ),
                P(
                    "fixed_watermark_fill_color",
                    "固定水印指定填充色",
                    "color",
                    "#FFFFFF",
                    help_text="仅选择“指定纯色填充”时显示。",
                    section="高级：固定水印精细修复",
                    advanced=True,
                    visible_when=("fixed_watermark_fill", ("color",)),
                ),
                P(
                    "presenter_policy",
                    "讲师画面处理",
                    "choice",
                    "auto_crop",
                    choices=(
                        ("auto_crop", "自动：保留课件全画面并处理常见右下讲师区"),
                        ("keep", "保留讲师画面"),
                        ("right_bottom", "处理右下角 22%×33%"),
                        ("custom", "在视频画面中框选讲师区域"),
                    ),
                    help_text="推荐自动；若讲师窗口位置特殊，可选择“框选讲师区域”。",
                    section="③ 水印与讲师画面",
                ),
                P(
                    "presenter_rect",
                    "框选讲师区域",
                    "region",
                    "",
                    help_text="点击右侧按钮，在视频画面中直接拖框。",
                    section="③ 水印与讲师画面",
                    visible_when=("presenter_policy", ("custom",)),
                ),
                P(
                    "enhancement_mode",
                    "幻灯片清晰增强",
                    "choice",
                    "auto",
                    choices=(
                        (
                            "auto",
                            "自动：有可用显卡则 AI，否则兼容增强（推荐）",
                        ),
                        (
                            "compatible",
                            "无独显兼容：多帧融合＋CPU高保真预处理",
                        ),
                        (
                            "gpu_ai",
                            "GPU AI：多帧融合＋Real-ESRGAN 2倍＋二检",
                        ),
                        ("off", "关闭清晰增强，使用旧版单帧结果"),
                    ),
                    help_text="无独立显卡可直接选择兼容模式；自动模式不会强制调用 GPU，AI 失败也会立即回退，不影响 PPT 输出。",
                    section="④ 清晰增强与输出",
                ),
                P(
                    "keep_images",
                    "同时保留高清 PNG 序列",
                    "boolean",
                    True,
                    help_text="PPT 始终生成；需要单独使用图片时保留此项。",
                    section="④ 清晰增强与输出",
                ),
                P(
                    "keep_report",
                    "保存 Word 提取报告",
                    "boolean",
                    True,
                    help_text="建议保留，便于核对每页采用时间、水印清理和低可信区域。",
                    section="④ 清晰增强与输出",
                ),
            ),
            fidelity="visual",
            capability_probe=engines.video_slide_extraction_capability,
            notes=(
                "输出一份按有效首次出现时间稳定排序的一页一图不可编辑 PPT；正式讲解前的极短误翻不会抢占页序，同页动画/逐条出现内容只保留印刷内容最完整的最终状态，后续重复页的全部出现时间写入 Word 报告。"
                "手写批注以出现时间、逐笔增长和可选 RGB 基准色联合识别，并回溯同页真实早期帧；颜色辅助不是全页硬删色，从首帧起始终存在且覆盖正文时会优先保护正文。"
                "永久固定遮挡、从首帧即存在的手写或讲师覆盖正文，没有可回溯真值时无法精确恢复，报告会明确标注。"
                "请仅处理本人或已获授权的视频。"
            ),
        ),
        Operation(
            "video.repair_slides_ppt",
            "视频生成",
            "视频 PPT 快速补修（水印 / 漏页）",
            "同时添加已有提取 PPT 与对应原视频，在一个窗口中直接选择页面、定位画面、拖框和预览。手写标注过多时可由用户逐帧选择最佳画面，并连续框选多个水印一次处理。确认后立即生成新 PPT；原文件和未框选区域保持不变。",
            repair_lecture_slides,
            (".pptx", *VIDEO_EXT),
            (
                P(
                    "repair_plan",
                    "打开快速补修",
                    "text",
                    "",
                    required=True,
                    help_text="先添加一份 PPTX 和对应原视频，再打开快速补修窗口。选择页面后会优先根据相邻提取报告自动定位原视频时间，无需填写坐标或方案。",
                    section="页面补修",
                ),
            ),
            min_inputs=2,
            max_inputs=2,
            fidelity="visual",
            capability_probe=engines.video_slide_extraction_capability,
            notes=(
                "自动缺页检测与时间排序仍由常规提取功能负责；本工具只用于自动结果中少量仍不满意的页面。"
                "多帧增强清理严格裁剪在用户框选区域内，并在写入前逐像素验证区域外完全不变。"
                "漏页插入会在用户时间点附近再次搜索稳定帧，避开尚未完成的转场；手选最佳帧模式则严格采用用户当前选中的确切画面，不再自动改选。"
            ),
        ),
        Operation(
            "image.to_video",
            "视频生成",
            "图片序列转视频",
            "按文件列表顺序生成高清 MP4，支持淡入淡出和背景音频。",
            lambda paths, out, p: video.images_to_video(
                paths,
                out / _filename_with_suffix(p["filename"], ".mp4", "图片视频"),
                slide_duration=p["slide_duration"],
                fps=p["fps"],
                resolution=p["resolution"],
                transition=p["transition"],
                transition_duration=p["transition_duration"],
                background=p["background"],
                audio_path=p["audio_path"],
                encoder="auto",
                quality=p["quality"],
            ),
            IMAGE_EXT,
            (P("filename", "输出文件名", "text", "图片视频", required=True),)
            + shared_parameters,
            min_inputs=1,
            fidelity="visual",
            capability_probe=engines.slideshow_video_capability,
        ),
        Operation(
            "pdf.to_video",
            "视频生成",
            "PDF 转视频",
            "将 PDF 页面高精度渲染后编码为 MP4 幻灯片。",
            lambda paths, out, p: _batch(
                paths,
                out,
                "视频",
                ".mp4",
                lambda src, target: video.pdf_to_video(
                    src,
                    target,
                    dpi=p["dpi"],
                    password=p["password"] or None,
                    slide_duration=p["slide_duration"],
                    fps=p["fps"],
                    resolution=p["resolution"],
                    transition=p["transition"],
                    transition_duration=p["transition_duration"],
                    background=p["background"],
                    audio_path=p["audio_path"],
                    encoder="auto",
                    quality=p["quality"],
                ),
            ),
            PDF_EXT,
            (
                P("dpi", "PDF 渲染 DPI", "integer", 240, minimum=150, maximum=600),
                P("password", "PDF 打开密码", "password", ""),
            )
            + shared_parameters,
            fidelity="visual",
            capability_probe=_pdf_video_capability,
        ),
        Operation(
            "video.transcode",
            "视频处理",
            "视频格式转换 / 转码",
            "在 MP4、MKV、MOV 之间转换，并可调整分辨率、帧率和主音轨。",
            lambda paths, out, p: _batch(
                paths,
                out,
                "转码",
                f".{p['format']}",
                lambda src, target: video.video_transcode(
                    src,
                    target,
                    video_codec=p["video_codec"],
                    audio_codec=p["audio_codec"],
                    quality=p["quality"],
                    resolution=(
                        None if p["resolution"] == "original" else p["resolution"]
                    ),
                    fps=_optional_int(p["target_fps"], "目标帧率"),
                    audio_bitrate=p["audio_bitrate"],
                ),
            ),
            VIDEO_EXT,
            (
                P(
                    "format",
                    "输出容器",
                    "choice",
                    "mp4",
                    choices=(
                        ("mp4", "MP4"),
                        ("mkv", "MKV"),
                        ("mov", "MOV"),
                    ),
                ),
                P(
                    "video_codec",
                    "视频编码",
                    "choice",
                    "auto",
                    choices=(
                        ("auto", "自动"),
                        ("h264", "H.264"),
                        ("copy", "复制原视频流"),
                    ),
                ),
                P(
                    "audio_codec",
                    "音频编码",
                    "choice",
                    "auto",
                    choices=(
                        ("auto", "自动"),
                        ("aac", "AAC"),
                        ("copy", "复制原音频流"),
                        ("none", "移除音频"),
                    ),
                ),
                P(
                    "quality",
                    "画质参数（越小越清晰）",
                    "integer",
                    20,
                    minimum=0,
                    maximum=40,
                ),
                P(
                    "resolution",
                    "目标分辨率",
                    "choice",
                    "original",
                    choices=(
                        ("original", "保持原分辨率"),
                        ("720p", "720p"),
                        ("1080p", "1080p"),
                        ("1440p", "1440p"),
                        ("2160p", "4K"),
                    ),
                ),
                P("target_fps", "目标帧率（留空保持）", "text", ""),
                P(
                    "audio_bitrate",
                    "音频码率 kbps",
                    "integer",
                    192,
                    minimum=16,
                    maximum=1024,
                ),
            ),
            fidelity="transform",
            capability_probe=engines.video_transform_capability,
            notes=(
                "容器与编码器必须兼容；选择流复制时不能改变分辨率或帧率。"
                "默认仅保留第一条视频流和第一条音轨，字幕、附件/数据流及额外流不会自动保留。"
            ),
        ),
        Operation(
            "video.compress",
            "视频处理",
            "视频压缩（H.264）",
            "使用兼容性最好的 H.264 重新编码，在清晰度与体积之间取舍。",
            lambda paths, out, p: _batch(
                paths,
                out,
                "压缩",
                f".{p['format']}",
                lambda src, target: video.video_compress(
                    src,
                    target,
                    codec="h264",
                    encoder="auto",
                    quality=p["quality"],
                    resolution=(
                        None if p["resolution"] == "original" else p["resolution"]
                    ),
                    fps=_optional_int(p["target_fps"], "目标帧率"),
                    audio_bitrate=p["audio_bitrate"],
                ),
            ),
            VIDEO_EXT,
            (
                P(
                    "format",
                    "输出容器",
                    "choice",
                    "mp4",
                    choices=(("mp4", "MP4"), ("mkv", "MKV")),
                ),
                P(
                    "quality",
                    "压缩强度（越小越清晰）",
                    "integer",
                    26,
                    minimum=18,
                    maximum=35,
                ),
                P(
                    "resolution",
                    "目标分辨率",
                    "choice",
                    "original",
                    choices=(
                        ("original", "保持原分辨率"),
                        ("720p", "720p"),
                        ("1080p", "1080p"),
                        ("1440p", "1440p"),
                        ("2160p", "4K"),
                    ),
                ),
                P("target_fps", "目标帧率（留空保持）", "text", ""),
                P(
                    "audio_bitrate",
                    "音频码率 kbps",
                    "integer",
                    128,
                    minimum=16,
                    maximum=1024,
                ),
            ),
            fidelity="transform",
            capability_probe=engines.video_transform_capability,
            notes="默认仅保留第一条视频流和第一条音轨；字幕、附件、数据流及额外音轨不会自动保留。",
        ),
        Operation(
            "video.trim",
            "视频处理",
            "视频裁剪 / 截取片段",
            "按时间精准截取视频并重新编码，避免关键帧流复制造成起止时间偏差。",
            lambda paths, out, p: _batch(
                paths,
                out,
                "裁剪",
                f".{p['format']}",
                lambda src, target: video.video_trim(
                    src,
                    target,
                    start=p["start"],
                    end=p["end"] or None,
                    duration=p["duration"] or None,
                    mode="precise",
                    video_codec="auto",
                    audio_codec="auto",
                    quality=p["quality"],
                ),
            ),
            VIDEO_EXT,
            (
                P(
                    "format",
                    "输出容器",
                    "choice",
                    "mp4",
                    choices=(
                        ("mp4", "MP4"),
                        ("mkv", "MKV"),
                        ("mov", "MOV"),
                    ),
                ),
                P("start", "开始时间", "text", "00:00:00", required=True),
                P("end", "结束时间（与持续时长二选一）", "text", ""),
                P("duration", "持续时长（可选）", "text", ""),
                P(
                    "quality",
                    "画质参数（越小越清晰）",
                    "integer",
                    20,
                    minimum=0,
                    maximum=40,
                ),
            ),
            fidelity="transform",
            capability_probe=engines.video_transform_capability,
            notes="默认仅保留第一条视频流和第一条音轨，字幕、附件/数据流及额外流不会自动保留。",
        ),
        Operation(
            "video.extract_audio",
            "视频处理",
            "提取无损 WAV 音频",
            "提取第一条音轨并输出无损 WAV，避免有损转码和编码器兼容问题。",
            lambda paths, out, p: _batch(
                paths,
                out,
                "音频",
                ".wav",
                lambda src, target: video.video_extract_audio(
                    src,
                    target,
                    audio_codec="auto",
                    bitrate=192,
                    sample_rate=_optional_int(p["sample_rate"], "采样率"),
                    channels=_optional_int(p["channels"], "声道数"),
                ),
            ),
            VIDEO_EXT,
            (
                P("sample_rate", "采样率 Hz（留空保持）", "text", ""),
                P("channels", "声道数（留空保持）", "text", ""),
            ),
            fidelity="extract",
            capability_probe=engines.audio_extraction_capability,
        ),
    ]


CORE_OPERATION_IDS: tuple[str, ...] = (
    # PDF：只保留视觉保真转换、文本提取和可预测的页面处理。
    "pdf.to_word",
    "pdf.to_ppt",
    "pdf.to_images",
    "pdf.extract_images",
    "pdf.to_text",
    "image.to_pdf",
    "pdf.merge",
    "pdf.split",
    "pdf.extract_pages",
    "pdf.delete_pages",
    "pdf.insert_pages",
    "pdf.rotate",
    "pdf.compress",
    "pdf.compress_lossy",
    "pdf.encrypt",
    "pdf.decrypt",
    "pdf.watermark",
    "pdf.header_footer",
    # Office：保留真实 Office/WPS/LibreOffice 转换及稳定的数据导出。
    "word.to_pdf",
    "word.full_compatibility",
    "excel.to_pdf",
    "ppt.to_pdf",
    "word.to_text",
    "excel.to_csv",
    "excel.to_json",
    "excel.to_txt",
    "legacy.doc_to_docx",
    "legacy.xls_to_xlsx",
    "ppt.to_images",
    # Word：保留边界清晰、容易复核的批处理。
    "word.replace",
    "word.remove_blank_lines",
    "word.remove_images",
    "word.typography",
    "word.headers_footers",
    "word.extract_images",
    "word.remove_hyperlinks",
    # Excel：移除可能破坏公式、表结构或合并区域的高风险入口。
    "excel.sort",
    "excel.filter",
    "excel.deduplicate",
    "excel.replace",
    "excel.split_column",
    "excel.conditional_format",
    "excel.extract_images",
    # PPT：不承诺母版重建、长图或动画视频复刻。
    "ppt.replace_fonts",
    "ppt.watermark",
    "ppt.extract_media",
    "ppt.compress_images",
    # 图片：保留确定性的像素级操作。
    "image.convert",
    "image.resize",
    "image.scale",
    "image.crop",
    "image.rotate",
    "image.flip",
    "image.compress",
    "image.remove_exif",
    "image.enhance",
    "image.adjust",
    "image.filter",
    "image.text_watermark",
    "image.image_watermark",
    "image.border",
    "image.mosaic",
    "image.stitch",
    "image.overlay",
    "image.rename",
    # 视频：统一走兼容性更高、结果更容易复核的 FFmpeg 路径。
    "image.to_video",
    "pdf.to_video",
    "video.extract_slides_ppt",
    "video.repair_slides_ppt",
    "video.transcode",
    "video.compress",
    "video.trim",
    "video.extract_audio",
)


def get_operations() -> list[Operation]:
    all_operations = _pdf_operations()
    all_operations.extend(_office_operations())
    all_operations.extend(_image_operations())
    all_operations.extend(_video_operations())

    allowed = set(CORE_OPERATION_IDS)
    operations = [item for item in all_operations if item.id in allowed]
    if len(operations) != len(CORE_OPERATION_IDS):
        found = {item.id for item in operations}
        missing = ", ".join(item for item in CORE_OPERATION_IDS if item not in found)
        raise RuntimeError(f"核心任务目录不完整：{missing}")
    combined_input_operations = {
        # These operations intentionally consume the whole ordered input set
        # to create one aggregate result and must never be split per file.
        "image.to_pdf",
        "pdf.merge",
        "pdf.insert_pages",
        "image.stitch",
        "image.rename",
        "image.to_video",
        "video.repair_slides_ppt",
    }
    return [
        replace(
            operation,
            independent_inputs=operation.id not in combined_input_operations,
        )
        for operation in operations
    ]


def _office_operations() -> list[Operation]:
    from .processors import excel_pivot, office, office_com, video, word_compat

    structure = engines.office_structure_capability
    word_renderer = lambda: engines.office_render_capability("word")
    excel_renderer = lambda: engines.office_render_capability("excel")
    powerpoint_renderer = lambda: engines.office_render_capability("powerpoint")
    engine_choices = (
        ("auto", "自动：WPS → Microsoft Office → LibreOffice"),
        ("wps", "WPS Office COM"),
        ("microsoft_office", "Microsoft Office COM（仅显式选择）"),
        ("libreoffice", "LibreOffice"),
    )
    compatibility_engine_choices = engine_choices + (
        ("none", "仅结构校验（跳过桌面渲染复检）"),
    )
    excel_pdf_layout_choices = (
        ("smart", "智能纸张与分页（推荐）"),
        ("preserve", "完全保留原设置"),
        ("fit_width", "固定一页宽，高度自动分页"),
        ("single_page", "整张工作表一页（可能缩小）"),
    )
    excel_pdf_paper_choices = (
        ("auto", "智能选择"),
        ("preserve", "保留工作簿设置"),
        ("a4", "A4"),
        ("a3", "A3"),
        ("letter", "Letter"),
    )
    excel_pdf_orientation_choices = (
        ("auto", "智能选择"),
        ("preserve", "保留工作簿设置"),
        ("portrait", "纵向"),
        ("landscape", "横向"),
    )
    excel_pdf_margin_choices = (
        ("auto", "智能优化（推荐）"),
        ("preserve", "保留原边距"),
        ("narrow", "窄边距"),
    )
    operations: list[Operation] = [
        Operation(
            "word.to_pdf",
            "Office 格式转换",
            "Word 转 PDF（高保真）",
            "调用真实 Office 渲染器导出 PDF，避免自行模拟排版。",
            lambda paths, out, p: sum(
                (
                    _convert_office_document(src, out, "pdf", p["engine"])
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            WORD_EXT + (".doc", ".rtf"),
            (P("engine", "转换引擎", "choice", "auto", choices=engine_choices),),
            fidelity="visual",
            capability_probe=word_renderer,
            notes=(
                "自动模式优先使用 WPS，再尝试 Microsoft Office 与 LibreOffice；"
                "显式选择 Microsoft Office 时失败不会偷换其他引擎。"
            ),
        ),
        Operation(
            "word.full_compatibility",
            "兼容修复 / 高级工具",
            "旧版固定坐标 Word 兼容升级",
            "升级本软件旧版本或其他遗留固定坐标 DOCX，减少手机端重叠并改善区域内复制与编辑。",
            lambda paths, out, p: _batch(
                paths,
                out,
                "旧版固定坐标Word_兼容升级",
                ".docx",
                lambda src, target: word_compat.optimize_word_full_compatibility(
                    src,
                    target,
                    verification_engine=p["verification_engine"],
                ),
            ),
            (".docx",),
            (
                P(
                    "verification_engine",
                    "二重检查渲染引擎",
                    "choice",
                    "auto",
                    choices=compatibility_engine_choices,
                    help_text=(
                        "自动模式优先使用 WPS，再尝试 Microsoft Office 与 LibreOffice；"
                        "第一重检查校验区域、文字、图片、分页和旧定位节点，第二重检查会真实渲染并检测缺字、"
                        "重叠、异常分页和版面差异；首次未通过时会自动用更保守的字体适配重建一次。"
                    ),
                ),
            ),
            fidelity="editable",
            capability_probe=structure,
            notes=(
                "仅建议用于本软件旧版本生成的逐行固定坐标 Word、首次转换时未能完成 WPS 区域优化的文件，"
                "以及其他遗留固定坐标 DOCX；新版 PDF 转 Word 已自动尝试区域优化，新文件通常无需再次处理。"
                "本工具不会修改原 Word，也不会把页面转成图片。"
                "程序按原坐标把逐行定位框合并为段落、分栏和功能区级文本框，区域内部使用普通 Word 段落；"
                "公式、图表与复杂表格从原页面视觉层裁剪后按原坐标回填。"
                "输出保留原件并另存为兼容升级 Word；任务 ID 继续保持 word.full_compatibility，以兼容旧配置。"
            ),
        ),
        Operation(
            "excel.to_pdf",
            "Office 格式转换",
            "Excel 转 PDF（高保真）",
            "使用原生打印与分页引擎导出 PDF，并可智能选择纸张、方向与分页方式。",
            lambda paths, out, p: sum(
                (
                    _convert_excel_to_pdf(
                        src,
                        out,
                        p["engine"],
                        excel_pdf_layout=p["excel_pdf_layout"],
                        excel_pdf_paper=p["excel_pdf_paper"],
                        excel_pdf_orientation=p["excel_pdf_orientation"],
                        excel_pdf_margin=p["excel_pdf_margin"],
                    )
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            EXCEL_EXT + (".xls",),
            (
                P(
                    "engine",
                    "转换引擎",
                    "choice",
                    "auto",
                    choices=engine_choices,
                    help_text=(
                        "自动模式依次尝试 WPS、Microsoft Office 与 LibreOffice，并使用"
                        "首个生成非空有效文件的引擎；显式选择的引擎不会自动偷换。"
                    ),
                ),
                P(
                    "excel_pdf_layout",
                    "页面布局",
                    "choice",
                    "smart",
                    choices=excel_pdf_layout_choices,
                    help_text="智能模式分析每个工作表的有效内容，尽量避免把相邻列拆到不同页面；长表仍会按高度自然分页。",
                ),
                P(
                    "excel_pdf_paper",
                    "纸张",
                    "choice",
                    "auto",
                    choices=excel_pdf_paper_choices,
                    help_text="自动会选择兼顾文字可读性与页数的纸张；保留则沿用工作簿中的纸张设置。",
                ),
                P(
                    "excel_pdf_orientation",
                    "方向",
                    "choice",
                    "auto",
                    choices=excel_pdf_orientation_choices,
                    help_text="自动会比较纵向与横向的有效缩放比例；保留则沿用工作簿中的方向设置。",
                ),
                P(
                    "excel_pdf_margin",
                    "页边距",
                    "choice",
                    "auto",
                    choices=excel_pdf_margin_choices,
                    help_text="自动优先保留安全边距，仅在能明显改善可读性或避免横向拆页时使用安全窄边距。",
                ),
            ),
            fidelity="visual",
            capability_probe=excel_renderer,
            notes=(
                "普通表格建议使用智能布局；人工设计过打印区域或分页的报表，可将页面布局、纸张、方向和页边距均设为保留。"
                "“整张工作表一页”可能把大型工作表缩得很小。"
            ),
        ),
        Operation(
            "ppt.to_pdf",
            "Office 格式转换",
            "PPT 转 PDF（高保真）",
            "自动按 WPS、Microsoft PowerPoint、LibreOffice 的顺序导出 PDF。",
            lambda paths, out, p: sum(
                (
                    _convert_office_document(src, out, "pdf", p["engine"])
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            PPT_EXT + (".ppt",),
            (P("engine", "转换引擎", "choice", "auto", choices=engine_choices),),
            fidelity="visual",
            capability_probe=powerpoint_renderer,
        ),
        Operation(
            "word.to_text",
            "Office 格式转换",
            "Word 转 TXT",
            "提取段落与表格为 UTF-8 纯文本。",
            lambda paths, out, p: sum(
                (
                    office.word_to_txt(src, out, include_tables=p["include_tables"])
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            WORD_EXT,
            (P("include_tables", "包含表格", "boolean", True),),
            fidelity="extract",
            capability_probe=structure,
        ),
        Operation(
            "word.to_markdown",
            "Office 格式转换",
            "Word 转 Markdown",
            "把标题、段落、列表、表格和基础强调转换为 Markdown。",
            lambda paths, out, p: sum(
                (office.word_to_markdown(src, out) for src in _progress_paths(paths)),
                [],
            ),
            WORD_EXT,
            fidelity="extract",
            capability_probe=structure,
        ),
        Operation(
            "word.to_html",
            "Office 格式转换",
            "Word 转 HTML",
            "生成语义化、可搜索的 HTML 文档。",
            lambda paths, out, p: sum(
                (
                    office.word_to_html(src, out, title=p["title"] or None)
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            WORD_EXT,
            (P("title", "网页标题（可选）", "text", ""),),
            fidelity="extract",
            capability_probe=structure,
        ),
        Operation(
            "word.to_epub",
            "Office 格式转换",
            "Word 转 EPUB",
            "生成标准 EPUB 3 电子书，适合流式阅读。",
            lambda paths, out, p: sum(
                (
                    office.word_to_epub(
                        src,
                        out,
                        title=p["title"] or None,
                        language=p["language"],
                        author=p["author"],
                    )
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            WORD_EXT,
            (
                P("title", "书名（可选）", "text", ""),
                P("author", "作者", "text", ""),
                P("language", "语言代码", "text", "zh-CN"),
            ),
            fidelity="extract",
            capability_probe=structure,
            notes="EPUB 是流式布局，不会保留 Word 的固定分页。",
        ),
        Operation(
            "excel.to_csv",
            "Office 格式转换",
            "Excel 转 CSV",
            "每个工作表输出一个 CSV，可指定工作表。",
            lambda paths, out, p: sum(
                (
                    office.excel_to_csv(
                        src,
                        out,
                        sheet_names=_csv_values(p["sheets"]),
                        delimiter=p["delimiter"],
                    )
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            EXCEL_EXT,
            (
                P("sheets", "工作表（逗号分隔，留空为全部）", "text", ""),
                P("delimiter", "分隔符", "text", ",", required=True),
            ),
            fidelity="extract",
            capability_probe=structure,
        ),
        Operation(
            "excel.to_json",
            "Office 格式转换",
            "Excel 转 JSON",
            "按工作表输出 JSON 数组，可将首行作为字段名。",
            lambda paths, out, p: sum(
                (
                    office.excel_to_json(
                        src,
                        out,
                        sheet_names=_csv_values(p["sheets"]),
                        header=p["header"],
                    )
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            EXCEL_EXT,
            (
                P("sheets", "工作表（逗号分隔）", "text", ""),
                P("header", "首行作为字段名", "boolean", True),
            ),
            fidelity="extract",
            capability_probe=structure,
        ),
        Operation(
            "excel.to_xml",
            "Office 格式转换",
            "Excel 转 XML",
            "按工作表输出结构化 XML。",
            lambda paths, out, p: sum(
                (
                    office.excel_to_xml(src, out, sheet_names=_csv_values(p["sheets"]))
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            EXCEL_EXT,
            (P("sheets", "工作表（逗号分隔）", "text", ""),),
            fidelity="extract",
            capability_probe=structure,
        ),
        Operation(
            "excel.to_txt",
            "Office 格式转换",
            "Excel 转 TXT",
            "按工作表输出制表符或自定义分隔文本。",
            lambda paths, out, p: sum(
                (
                    office.excel_to_txt(
                        src,
                        out,
                        sheet_names=_csv_values(p["sheets"]),
                        delimiter=p["delimiter"],
                    )
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            EXCEL_EXT,
            (
                P("sheets", "工作表（逗号分隔）", "text", ""),
                P("delimiter", "分隔符", "text", "\t", required=True),
            ),
            fidelity="extract",
            capability_probe=structure,
        ),
        Operation(
            "legacy.doc_to_docx",
            "Office 格式转换",
            "DOC 转 DOCX",
            "通过真实 Office 引擎把旧版 Word 格式升级为 DOCX。",
            lambda paths, out, p: sum(
                (
                    _convert_office_document(src, out, "docx", p["engine"])
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            (".doc",),
            (P("engine", "转换引擎", "choice", "auto", choices=engine_choices),),
            fidelity="editable",
            capability_probe=word_renderer,
        ),
        Operation(
            "legacy.xls_to_xlsx",
            "Office 格式转换",
            "XLS 转 XLSX",
            "通过真实 Office 引擎把旧版 Excel 格式升级为 XLSX。",
            lambda paths, out, p: sum(
                (
                    _convert_office_document(src, out, "xlsx", p["engine"])
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            (".xls",),
            (P("engine", "转换引擎", "choice", "auto", choices=engine_choices),),
            fidelity="editable",
            capability_probe=excel_renderer,
        ),
        Operation(
            "ppt.to_images",
            "Office 格式转换",
            "PPT 转图片序列",
            "自动模式优先由 WPS 转 PDF 后高清渲染；可显式选择 PowerPoint 原生导出。",
            lambda paths, out, p: sum(
                (
                    video.ppt_to_images(
                        src,
                        out / f"{src.stem}_幻灯片",
                        renderer=p["renderer"],
                        image_format=p["format"],
                        width=p["width"],
                        dpi=p["dpi"],
                    )
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            PPT_EXT + (".ppt",),
            (
                P(
                    "renderer",
                    "渲染引擎",
                    "choice",
                    "auto",
                    choices=engine_choices,
                ),
                P(
                    "format",
                    "图片格式",
                    "choice",
                    "png",
                    choices=(("png", "PNG"), ("jpg", "JPG")),
                ),
                P(
                    "width",
                    "PowerPoint 原生导出宽度",
                    "integer",
                    2560,
                    minimum=1280,
                    maximum=10000,
                ),
                P("dpi", "PDF 页面渲染 DPI", "integer", 220, minimum=150, maximum=600),
            ),
            fidelity="visual",
            capability_probe=engines.ppt_render_capability,
            notes=(
                "自动静态渲染顺序为 WPS → Microsoft Office → LibreOffice；"
                "显式选择 Microsoft Office 时使用 PowerPoint 原生图片导出且不回退。"
            ),
        ),
        Operation(
            "ppt.to_video",
            "Office 格式转换",
            "PPT 转视频",
            "优先使用 PowerPoint 原生视频；没有 PowerPoint 时可由 WPS/LibreOffice 渲染并交给 FFmpeg。",
            lambda paths, out, p: sum(
                (
                    video.ppt_to_video(
                        src,
                        out / f"{src.stem}_视频.mp4",
                        mode=p["mode"],
                        renderer=p["renderer"],
                        use_timings=p["use_timings"],
                        slide_duration=p["slide_duration"],
                        resolution=p["resolution"],
                        fps=p["fps"],
                        quality=p["quality"],
                        transition=p["transition"],
                        transition_duration=p["transition_duration"],
                        audio_path=p["audio_path"],
                        encoder=p["encoder"],
                    )
                    for src in _progress_paths(paths)
                ),
                [],
            ),
            PPT_EXT + (".ppt",),
            (
                P(
                    "mode",
                    "视频模式",
                    "choice",
                    "auto",
                    choices=(
                        ("auto", "自动：原生优先，静态回退"),
                        ("native", "仅 PowerPoint 原生动画视频"),
                        ("static", "静态幻灯片 + FFmpeg"),
                    ),
                ),
                P(
                    "renderer",
                    "静态模式渲染器",
                    "choice",
                    "auto",
                    choices=engine_choices,
                ),
                P("use_timings", "原生模式使用已有计时与旁白", "boolean", True),
                P(
                    "slide_duration",
                    "每页时长（秒）",
                    "number",
                    5,
                    minimum=0.1,
                    maximum=3600,
                ),
                P(
                    "resolution",
                    "分辨率",
                    "choice",
                    "1080p",
                    choices=(
                        ("720p", "720p"),
                        ("1080p", "1080p"),
                        ("1440p", "1440p"),
                        ("2160p", "4K"),
                    ),
                ),
                P("fps", "帧率", "integer", 30, minimum=1, maximum=120),
                P(
                    "transition",
                    "静态转场",
                    "choice",
                    "fade",
                    choices=(("none", "无"), ("fade", "淡入淡出")),
                ),
                P(
                    "transition_duration",
                    "转场时长（秒）",
                    "number",
                    0.5,
                    minimum=0,
                    maximum=30,
                ),
                P("audio_path", "背景音频（可选）", "path", None),
                P(
                    "encoder",
                    "FFmpeg 编码器",
                    "choice",
                    "auto",
                    choices=(
                        ("auto", "自动 / CPU 兼容优先"),
                        ("libx264", "H.264 CPU"),
                        ("h264_nvenc", "NVIDIA NVENC"),
                        ("h264_qsv", "Intel Quick Sync"),
                        ("h264_amf", "AMD AMF"),
                        ("mpeg4", "MPEG-4 兼容模式"),
                    ),
                ),
                P(
                    "quality",
                    "画质参数（越小越清晰）",
                    "integer",
                    20,
                    minimum=0,
                    maximum=40,
                ),
            ),
            fidelity="visual",
            capability_probe=engines.ppt_video_capability,
            notes=(
                "PowerPoint 原生模式可保留动画、转场和旁白；"
                "静态模式按 WPS → Microsoft Office → LibreOffice 渲染，随后经 "
                "PDF → Poppler 页面 → FFmpeg 输出静态页面视频。"
            ),
        ),
    ]

    operations.extend(
        [
            Operation(
                "word.replace",
                "Word 专项处理",
                "Word 批量替换文字",
                "用 JSON 映射一次替换多个词语，并覆盖正文、表格、页眉页脚。",
                lambda paths, out, p: sum(
                    (
                        office.word_replace_text(
                            src,
                            out,
                            _json_object(p["replacements"], "替换映射"),
                            case_sensitive=p["case_sensitive"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                (
                    P(
                        "replacements",
                        "替换 JSON",
                        "text",
                        '{"旧文字":"新文字"}',
                        required=True,
                    ),
                    P("case_sensitive", "区分大小写", "boolean", True),
                ),
                capability_probe=structure,
            ),
            Operation(
                "word.remove_blank_lines",
                "Word 专项处理",
                "删除 Word 空白行",
                "删除没有文字和对象的空段落。",
                lambda paths, out, p: sum(
                    (
                        office.word_remove_blank_paragraphs(src, out)
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                capability_probe=structure,
                notes="真正的空白页通常由分页符、分节符和排版计算产生，需 Word 原生排版引擎才能安全判断。",
            ),
            Operation(
                "word.remove_blank_pages",
                "Word 专项处理",
                "删除 Word 空白页（排版级）",
                "用 Word 原生分页引擎删除可安全确认的显式空白页。",
                lambda paths, out, p: sum(
                    (
                        office_com.word_remove_blank_pages(src, out)
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                capability_probe=engines.microsoft_word_capability,
                notes="遇到分节、多栏、修订跟踪、保护或浮动对象等高风险结构会拒绝处理，避免破坏版式。",
            ),
            Operation(
                "word.remove_images",
                "Word 专项处理",
                "批量删除 Word 图片",
                "删除内嵌和浮动图片关系并另存。",
                lambda paths, out, p: sum(
                    (
                        office.word_remove_images(src, out)
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                capability_probe=structure,
            ),
            Operation(
                "word.revisions",
                "Word 专项处理",
                "清理批注与修订",
                "接受或拒绝修订，并移除批注关系和批注内容。",
                lambda paths, out, p: sum(
                    (
                        office.word_clean_comments_and_revisions(
                            src, out, revision_mode=p["mode"]
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                (
                    P(
                        "mode",
                        "修订处理",
                        "choice",
                        "accept",
                        choices=(
                            ("accept", "接受所有修订"),
                            ("reject", "拒绝所有修订"),
                        ),
                    ),
                ),
                capability_probe=structure,
                notes="基于 OOXML 修订标记处理；复杂域、文本框和第三方扩展文档应先备份抽样验证。",
            ),
            Operation(
                "word.typography",
                "Word 专项处理",
                "统一字体、字号与行距",
                "统一正文和表格文本格式，并设置段落行距。",
                lambda paths, out, p: sum(
                    (
                        office.word_set_typography(
                            src,
                            out,
                            font_name=p["font_name"] or None,
                            font_size_pt=p["font_size"] or None,
                            line_spacing=p["line_spacing"] or None,
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                (
                    P("font_name", "字体名称", "text", "微软雅黑"),
                    P(
                        "font_size",
                        "字号（磅，0 不修改）",
                        "number",
                        12,
                        minimum=0,
                        maximum=200,
                    ),
                    P(
                        "line_spacing",
                        "行距倍数（0 不修改）",
                        "number",
                        1.5,
                        minimum=0,
                        maximum=10,
                    ),
                ),
                capability_probe=structure,
            ),
            Operation(
                "word.headers_footers",
                "Word 专项处理",
                "批量设置页眉页脚",
                "对所有节设置或追加页眉、页脚文字。",
                lambda paths, out, p: sum(
                    (
                        office.word_set_headers_footers(
                            src,
                            out,
                            header_text=p["header"] or None,
                            footer_text=p["footer"] or None,
                            replace=p["replace"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                (
                    P("header", "页眉文字", "text", ""),
                    P("footer", "页脚文字", "text", ""),
                    P("replace", "替换现有内容", "boolean", True),
                ),
                capability_probe=structure,
            ),
            Operation(
                "word.extract_images",
                "Word 专项处理",
                "提取 Word 所有图片",
                "从文档媒体包中原样提取图片资源。",
                lambda paths, out, p: sum(
                    (
                        office.word_extract_images(src, out / f"{src.stem}_图片")
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                capability_probe=structure,
            ),
            Operation(
                "word.mail_merge",
                "Word 专项处理",
                "Word 邮件合并",
                "使用 {{字段名}} 占位符和 JSON 数据批量生成个性化文档。",
                lambda paths, out, p: sum(
                    (
                        office.word_mail_merge(
                            src,
                            out,
                            _json_records(p["records"]),
                            filename_template=p["filename_template"] or None,
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                (
                    P(
                        "records",
                        "数据 JSON",
                        "text",
                        '[{"姓名":"张三"}]',
                        required=True,
                    ),
                    P("filename_template", "文件名模板", "text", "{index:03d}_{姓名}"),
                ),
                max_inputs=1,
                capability_probe=structure,
            ),
            Operation(
                "word.remove_hyperlinks",
                "Word 专项处理",
                "删除 Word 超链接",
                "取消网址链接但保留显示文字。",
                lambda paths, out, p: sum(
                    (
                        office.word_remove_hyperlinks(src, out)
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                WORD_EXT,
                capability_probe=structure,
            ),
        ]
    )

    operations.extend(
        [
            Operation(
                "excel.sort",
                "Excel 专项处理",
                "Excel 数据排序",
                "按列号或列字母排序整行数据，并保留样式。",
                lambda paths, out, p: sum(
                    (
                        office.excel_sort_rows(
                            src,
                            out,
                            column=p["column"],
                            sheet_name=p["sheet"] or None,
                            header=p["header"],
                            reverse=p["reverse"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("column", "排序列（如 A 或 1）", "text", "A", required=True),
                    P("sheet", "工作表（留空为当前表）", "text", ""),
                    P("header", "首行为标题", "boolean", True),
                    P("reverse", "降序", "boolean", False),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.filter",
                "Excel 专项处理",
                "Excel 数据筛选",
                "按条件删除不匹配的行并输出筛选结果。",
                lambda paths, out, p: sum(
                    (
                        office.excel_filter_rows(
                            src,
                            out,
                            column=p["column"],
                            value=_smart_value(p["value"]),
                            operator=p["operator"],
                            sheet_name=p["sheet"] or None,
                            header=p["header"],
                            case_sensitive=p["case_sensitive"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("column", "筛选列", "text", "A", required=True),
                    P(
                        "operator",
                        "条件",
                        "choice",
                        "equals",
                        choices=(
                            ("equals", "等于"),
                            ("not_equals", "不等于"),
                            ("contains", "包含"),
                            ("not_contains", "不包含"),
                            ("starts_with", "开头是"),
                            ("ends_with", "结尾是"),
                            ("greater_than", "大于"),
                            ("greater_equal", "大于等于"),
                            ("less_than", "小于"),
                            ("less_equal", "小于等于"),
                            ("is_blank", "为空"),
                            ("not_blank", "不为空"),
                        ),
                    ),
                    P("value", "比较值", "text", ""),
                    P("sheet", "工作表", "text", ""),
                    P("header", "首行为标题", "boolean", True),
                    P("case_sensitive", "区分大小写", "boolean", False),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.deduplicate",
                "Excel 专项处理",
                "Excel 删除重复值",
                "按指定列组合去重；留空表示比较整行。",
                lambda paths, out, p: sum(
                    (
                        office.excel_remove_duplicates(
                            src,
                            out,
                            columns=_csv_values(p["columns"]),
                            sheet_name=p["sheet"] or None,
                            header=p["header"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("columns", "列（如 A,C；留空为整行）", "text", ""),
                    P("sheet", "工作表", "text", ""),
                    P("header", "首行为标题", "boolean", True),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.remove_blanks",
                "Excel 专项处理",
                "删除 Excel 空白行 / 列",
                "删除完全空白的行和/或列。",
                lambda paths, out, p: sum(
                    (
                        office.excel_remove_blank_rows_columns(
                            src,
                            out,
                            sheet_name=p["sheet"] or None,
                            remove_rows=p["rows"],
                            remove_columns=p["columns"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("sheet", "工作表", "text", ""),
                    P("rows", "删除空白行", "boolean", True),
                    P("columns", "删除空白列", "boolean", True),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.replace",
                "Excel 专项处理",
                "批量替换单元格内容",
                "用 JSON 映射替换一个或多个工作表的文字。",
                lambda paths, out, p: sum(
                    (
                        office.excel_replace_text(
                            src,
                            out,
                            _json_object(p["replacements"], "替换映射"),
                            sheet_names=_csv_values(p["sheets"]),
                            case_sensitive=p["case_sensitive"],
                            exact=p["exact"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P(
                        "replacements",
                        "替换 JSON",
                        "text",
                        '{"旧值":"新值"}',
                        required=True,
                    ),
                    P("sheets", "工作表（逗号分隔）", "text", ""),
                    P("case_sensitive", "区分大小写", "boolean", True),
                    P("exact", "仅替换完全相等单元格", "boolean", False),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.formulas_to_values",
                "Excel 专项处理",
                "删除公式，保留数值",
                "使用工作簿中已缓存的计算结果替换公式。",
                lambda paths, out, p: sum(
                    (
                        office.excel_formulas_to_values(
                            src,
                            out,
                            sheet_names=_csv_values(p["sheets"]),
                            missing_cache=p["missing_cache"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("sheets", "工作表（逗号分隔）", "text", ""),
                    P(
                        "missing_cache",
                        "缺少缓存结果时",
                        "choice",
                        "error",
                        choices=(
                            ("error", "停止，避免丢数据"),
                            ("warn", "保留该公式并警告"),
                        ),
                    ),
                ),
                capability_probe=structure,
                notes="openpyxl 不会计算公式；如需强制重算，必须使用 Excel 或 LibreOffice。",
            ),
            Operation(
                "excel.split_column",
                "Excel 专项处理",
                "Excel 分列",
                "按分隔符把一列拆成多列。",
                lambda paths, out, p: sum(
                    (
                        office.excel_split_column(
                            src,
                            out,
                            column=p["column"],
                            delimiter=p["delimiter"],
                            sheet_name=p["sheet"] or None,
                            header=p["header"],
                            maxsplit=p["maxsplit"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("column", "要拆分的列", "text", "A", required=True),
                    P("delimiter", "分隔符", "text", ",", required=True),
                    P("sheet", "工作表", "text", ""),
                    P("header", "首行为标题", "boolean", True),
                    P("maxsplit", "最多拆分次数（-1 不限）", "integer", -1, minimum=-1),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.merge_cells",
                "Excel 专项处理",
                "合并 Excel 单元格",
                "按 A1:B2 形式合并一个或多个范围。",
                lambda paths, out, p: sum(
                    (
                        office.excel_merge_cells(
                            src,
                            out,
                            _csv_values(p["ranges"]) or [],
                            sheet_name=p["sheet"] or None,
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("ranges", "范围（逗号分隔）", "text", "A1:B2", required=True),
                    P("sheet", "工作表", "text", ""),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.unmerge_cells",
                "Excel 专项处理",
                "拆分 Excel 合并单元格",
                "拆分指定或全部合并范围，可把左上角值填入每个单元格。",
                lambda paths, out, p: sum(
                    (
                        office.excel_unmerge_cells(
                            src,
                            out,
                            _csv_values(p["ranges"]),
                            sheet_name=p["sheet"] or None,
                            fill=p["fill"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("ranges", "范围（留空为全部）", "text", ""),
                    P("sheet", "工作表", "text", ""),
                    P("fill", "将值填入所有拆分单元格", "boolean", False),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.conditional_format",
                "Excel 专项处理",
                "Excel 条件格式",
                "按阈值、公式或色阶高亮区域。",
                lambda paths, out, p: sum(
                    (
                        office.excel_apply_conditional_format(
                            src,
                            out,
                            cell_range=p["range"],
                            rule=p["rule"],
                            operator=p["operator"],
                            threshold=_smart_value(p["threshold"]),
                            formula=p["formula"] or None,
                            fill_color=p["color"].lstrip("#"),
                            sheet_name=p["sheet"] or None,
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("range", "应用范围", "text", "A2:A100", required=True),
                    P(
                        "rule",
                        "规则",
                        "choice",
                        "cell",
                        choices=(
                            ("cell", "单元格值"),
                            ("formula", "自定义公式"),
                            ("color_scale", "三色色阶"),
                        ),
                    ),
                    P(
                        "operator",
                        "比较",
                        "choice",
                        "greaterThan",
                        choices=(
                            ("greaterThan", "大于"),
                            ("greaterThanOrEqual", "大于等于"),
                            ("lessThan", "小于"),
                            ("lessThanOrEqual", "小于等于"),
                            ("equal", "等于"),
                            ("notEqual", "不等于"),
                        ),
                    ),
                    P("threshold", "阈值", "text", "0"),
                    P("formula", "公式", "text", ""),
                    P("color", "高亮颜色", "color", "FFF2CC"),
                    P("sheet", "工作表", "text", ""),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.extract_images",
                "Excel 专项处理",
                "提取 Excel 图片",
                "从工作簿媒体包中提取所有图片。",
                lambda paths, out, p: sum(
                    (
                        office.excel_extract_images(src, out / f"{src.stem}_图片")
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                capability_probe=structure,
            ),
            Operation(
                "excel.manage_sheets",
                "Excel 专项处理",
                "工作表重命名 / 删除 / 复制 / 排序",
                "用 JSON 和逗号列表统一管理工作表。",
                lambda paths, out, p: sum(
                    (
                        office.excel_manage_sheets(
                            src,
                            out,
                            rename=(
                                _json_object(p["rename"], "重命名映射")
                                if p["rename"].strip()
                                else None
                            ),
                            delete=_csv_values(p["delete"]),
                            copy_sheets=(
                                _json_object(p["copy"], "复制映射")
                                if p["copy"].strip()
                                else None
                            ),
                            order=_csv_values(p["order"]),
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("rename", "重命名 JSON", "text", ""),
                    P("delete", "删除工作表（逗号分隔）", "text", ""),
                    P("copy", "复制 JSON（原名→新名）", "text", ""),
                    P("order", "目标顺序（逗号分隔）", "text", ""),
                ),
                capability_probe=structure,
            ),
            Operation(
                "excel.pivot",
                "Excel 专项处理",
                "创建数据透视表",
                "使用 Excel 原生透视表引擎创建并刷新汇总。",
                lambda paths, out, p: sum(
                    (
                        excel_pivot.excel_create_pivot_compatible(
                            src,
                            out,
                            source_sheet=p["source_sheet"],
                            source_range=p["source_range"],
                            target_sheet=p["target_sheet"],
                            target_cell=p["target_cell"],
                            row_fields=_csv_values(p["row_fields"]) or (),
                            column_fields=_csv_values(p["column_fields"]) or (),
                            data_field=p["data_field"] or None,
                            function=p["function"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                EXCEL_EXT,
                (
                    P("source_sheet", "数据源工作表", "text", "Sheet1", required=True),
                    P("source_range", "数据源范围", "text", "A1:D100", required=True),
                    P(
                        "target_sheet",
                        "透视表工作表",
                        "text",
                        "数据透视表",
                        required=True,
                    ),
                    P("target_cell", "放置位置", "text", "A1", required=True),
                    P("row_fields", "行字段（逗号分隔）", "text", ""),
                    P("column_fields", "列字段（逗号分隔）", "text", ""),
                    P("data_field", "数值字段（可选）", "text", ""),
                    P(
                        "function",
                        "汇总方式",
                        "choice",
                        "sum",
                        choices=(
                            ("sum", "求和"),
                            ("count", "计数"),
                            ("average", "平均值"),
                            ("max", "最大值"),
                            ("min", "最小值"),
                        ),
                    ),
                ),
                capability_probe=structure,
                notes=(
                    "检测到 Microsoft Excel 时创建可交互的原生数据透视表；"
                    "其他电脑自动生成兼容 WPS、LibreOffice 和移动端的静态透视汇总表。"
                ),
            ),
        ]
    )

    operations.extend(
        [
            Operation(
                "ppt.replace_fonts",
                "PPT 专项处理",
                "PPT 批量替换字体",
                "替换文本框、表格和图形文字字体。",
                lambda paths, out, p: sum(
                    (
                        office.ppt_replace_fonts(
                            src,
                            out,
                            (
                                _json_object(p["replacements"], "字体映射")
                                if p["replacements"].strip()
                                else None
                            ),
                            default_font=p["default_font"] or None,
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                PPT_EXT,
                (
                    P("replacements", "字体映射 JSON", "text", '{"Arial":"微软雅黑"}'),
                    P("default_font", "统一替换为（可选）", "text", ""),
                ),
                capability_probe=structure,
                notes="SmartArt、主题母版字体和部分图表文字需 PowerPoint 原生引擎处理。",
            ),
            Operation(
                "ppt.watermark",
                "PPT 专项处理",
                "PPT 批量添加水印",
                "在每张幻灯片添加半透明旋转文字。",
                lambda paths, out, p: sum(
                    (
                        office.ppt_add_watermark(
                            src,
                            out,
                            p["text"],
                            font_size_pt=p["font_size"],
                            color=p["color"].lstrip("#"),
                            rotation=p["rotation"],
                            opacity=p["opacity"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                PPT_EXT,
                (
                    P("text", "水印文字", "text", "仅供内部使用", required=True),
                    P("font_size", "字号", "number", 32, minimum=6, maximum=200),
                    P("color", "颜色", "color", "B7B7B7"),
                    P("rotation", "旋转角度", "number", -30, minimum=-360, maximum=360),
                    P("opacity", "透明度", "number", 0.35, minimum=0, maximum=1),
                ),
                capability_probe=structure,
            ),
            Operation(
                "ppt.extract_media",
                "PPT 专项处理",
                "提取 PPT 图片 / 音视频",
                "从演示文稿媒体包中原样提取资源。",
                lambda paths, out, p: sum(
                    (
                        office.ppt_extract_media(
                            src,
                            out / f"{src.stem}_媒体",
                            include_images=p["images"],
                            include_audio=p["audio"],
                            include_video=p["video"],
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                PPT_EXT,
                (
                    P("images", "提取图片", "boolean", True),
                    P("audio", "提取音频", "boolean", True),
                    P("video", "提取视频", "boolean", True),
                ),
                capability_probe=structure,
            ),
            Operation(
                "ppt.compress_images",
                "PPT 专项处理",
                "PPT 压缩图片",
                "重编码媒体包中的位图，并限制最长边。",
                lambda paths, out, p: sum(
                    (
                        office.ppt_compress_images(
                            src,
                            out,
                            quality=p["quality"],
                            max_dimension=p["max_dimension"] or None,
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                PPT_EXT,
                (
                    P("quality", "JPEG 质量", "integer", 88, minimum=1, maximum=100),
                    P(
                        "max_dimension",
                        "最长边像素（0 不缩放）",
                        "integer",
                        2560,
                        minimum=0,
                    ),
                ),
                capability_probe=structure,
                notes="矢量图和不支持的媒体不会被错误重编码；压缩后建议抽样检查透明度和裁剪效果。",
            ),
            Operation(
                "ppt.master",
                "PPT 专项处理",
                "批量修改 PPT 母版",
                "使用 PowerPoint 原生对象模型修改母版和主题。",
                lambda paths, out, p: sum(
                    (
                        office_com.ppt_modify_master(
                            src,
                            out,
                            background_color=p["background_color"] or None,
                            font_name=p["font_name"] or None,
                            footer_text=p["footer_text"] or None,
                        )
                        for src in _progress_paths(paths)
                    ),
                    [],
                ),
                PPT_EXT + (".ppt",),
                (
                    P("background_color", "母版背景色（可选）", "color", ""),
                    P("font_name", "母版字体（可选）", "text", ""),
                    P("footer_text", "母版页脚（可选）", "text", ""),
                ),
                capability_probe=engines.microsoft_powerpoint_capability,
                notes="直接修改 PowerPoint 原生母版；SmartArt、主题变体和第三方插件对象仍需抽样检查。",
            ),
            Operation(
                "ppt.long_image",
                "PPT 专项处理",
                "PPT 转长图",
                "先高精度渲染各页，再按顺序拼接。",
                lambda paths, out, p: _batch(
                    paths,
                    out,
                    "长图",
                    f".{p['format']}",
                    lambda src, target: video.ppt_to_long_image(
                        src,
                        target,
                        renderer=p["renderer"],
                        direction=p["direction"],
                        spacing=p["spacing"],
                        background=p["background"],
                        width=p["width"],
                        dpi=p["dpi"],
                    ),
                ),
                PPT_EXT + (".ppt",),
                (
                    P(
                        "renderer",
                        "渲染引擎",
                        "choice",
                        "auto",
                        choices=engine_choices,
                    ),
                    P(
                        "format",
                        "长图格式",
                        "choice",
                        "png",
                        choices=(("png", "PNG"), ("jpg", "JPG"), ("webp", "WebP")),
                    ),
                    P(
                        "direction",
                        "拼接方向",
                        "choice",
                        "vertical",
                        choices=(("vertical", "纵向"), ("horizontal", "横向")),
                    ),
                    P("spacing", "页面间距", "integer", 0, minimum=0, maximum=500),
                    P("background", "背景色", "color", "white"),
                    P(
                        "width",
                        "PowerPoint 原生导出宽度",
                        "integer",
                        1920,
                        minimum=160,
                        maximum=10000,
                    ),
                    P(
                        "dpi",
                        "PDF 页面渲染 DPI",
                        "integer",
                        160,
                        minimum=72,
                        maximum=600,
                    ),
                ),
                capability_probe=engines.ppt_render_capability,
            ),
        ]
    )
    return operations
