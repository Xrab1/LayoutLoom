from __future__ import annotations

import importlib.util
import os
import shutil
import sys
from functools import lru_cache
from pathlib import Path

from .models import Capability


@lru_cache(maxsize=None)
def has_module(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


@lru_cache(maxsize=None)
def find_executable(*names: str) -> str | None:
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    return None


@lru_cache(maxsize=1)
def poppler_bin_path() -> str | None:
    """Return a directory containing real Poppler executables, not a ``.cmd`` shim."""

    def has_pair(directory: Path) -> bool:
        return any(
            (directory / name).is_file() for name in ("pdftoppm.exe", "pdftoppm")
        ) and any((directory / name).is_file() for name in ("pdfinfo.exe", "pdfinfo"))

    def first_existing(candidates: list[Path]) -> str | None:
        for candidate in candidates:
            if has_pair(candidate):
                return str(candidate)
        return None

    explicit = (
        os.environ.get("LAYOUTLOOM_POPPLER_PATH", "").strip()
        or os.environ.get("DOCUFORGE_POPPLER_PATH", "").strip()
    )
    if explicit:
        resolved = first_existing([Path(explicit).expanduser()])
        if resolved:
            return resolved

    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        bundled_root = Path(bundle_root) / "poppler"
        resolved = first_existing(
            [
                bundled_root,
                bundled_root / "bin",
                bundled_root / "Library/bin",
            ]
        )
        if resolved:
            return resolved

    executable_root = Path(sys.executable).resolve().parent
    resolved = first_existing(
        [
            executable_root / "poppler",
            executable_root / "poppler/bin",
            executable_root / "poppler/Library/bin",
            executable_root / "_internal/poppler",
            executable_root / "_internal/poppler/bin",
            executable_root / "_internal/poppler/Library/bin",
        ]
    )
    if resolved:
        return resolved

    project_root = Path(__file__).resolve().parents[1]
    resolved = first_existing(
        [
            project_root / "third_party/poppler",
            project_root / "third_party/poppler/bin",
            project_root / "third_party/poppler/Library/bin",
        ]
    )
    if resolved:
        return resolved

    found = find_executable("pdftoppm")
    if found:
        candidate = Path(found).resolve()
        if candidate.suffix.lower() == ".exe" and has_pair(candidate.parent):
            return str(candidate.parent)
        for parent in (candidate.parent, *candidate.parents):
            resolved = first_existing(
                [
                    (parent / "Library/bin").resolve(),
                    (parent / "native/poppler/Library/bin").resolve(),
                ]
            )
            if resolved:
                return resolved

    # GUI launches and packaged/venv processes often do not inherit the shell
    # PATH. Probe the local runtime caches used by the installer as a fallback.
    cache_roots: list[Path] = []
    user_home = Path.home()
    cache_roots.append(user_home / ".cache" / "codex-runtimes")
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        cache_roots.append(Path(local_app_data) / ".cache" / "codex-runtimes")
    for cache_root in cache_roots:
        if not cache_root.is_dir():
            continue
        candidates = [
            match / "dependencies/native/poppler/Library/bin"
            for match in cache_root.glob("*/")
        ]
        resolved = first_existing(candidates)
        if resolved:
            return resolved
    return None


def builtin(engine: str) -> Capability:
    return Capability("ready", "本机可直接处理", engine)


def require_modules(engine: str, *modules: str) -> Capability:
    missing = [name for name in modules if not has_module(name)]
    if missing:
        return Capability(
            "unavailable",
            f"缺少组件：{', '.join(missing)}。请先按 requirements.txt 安装依赖。",
            engine,
        )
    return builtin(engine)


@lru_cache(maxsize=1)
def pdf_core_capability() -> Capability:
    return require_modules("pypdf + ReportLab", "pypdf", "reportlab")


@lru_cache(maxsize=1)
def pdf_extract_capability() -> Capability:
    return require_modules("pdfplumber", "pdfplumber")


@lru_cache(maxsize=1)
def pdf_render_capability() -> Capability:
    if not has_module("pdf2image"):
        return Capability("unavailable", "缺少 pdf2image Python 组件", "Poppler")
    poppler_dir = poppler_bin_path()
    if not poppler_dir:
        return Capability(
            "unavailable",
            "未检测到同目录的 Poppler pdftoppm 与 pdfinfo，无法渲染 PDF 页面",
            "Poppler",
        )
    return Capability("external", f"已检测到 {poppler_dir}", "Poppler + pdf2image")


@lru_cache(maxsize=1)
def pdf_to_ppt_capability() -> Capability:
    renderer = pdf_render_capability()
    if not renderer.runnable:
        return renderer
    missing = [name for name in ("PIL", "pptx") if not has_module(name)]
    if missing:
        return Capability(
            "unavailable",
            f"PDF 转 PPT 缺少组件：{', '.join(missing)}",
            "Poppler + Pillow + python-pptx",
        )
    return Capability(
        "external",
        f"{renderer.reason}；Pillow 与 python-pptx 已就绪",
        "Poppler + Pillow + python-pptx",
    )


@lru_cache(maxsize=1)
def pdf_to_word_capability() -> Capability:
    missing = [
        module for module in ("docx", "pdf2docx", "pymupdf") if not has_module(module)
    ]
    if missing:
        return Capability(
            "unavailable",
            f"PDF 转 Word 缺少组件：{', '.join(missing)}。请重新运行安装程序。",
            "pdf2docx + PyMuPDF + python-docx",
        )
    return Capability(
        "ready",
        "混合保真、全文可编辑与整篇高清原样三种模式均已就绪",
        "pdf2docx + PyMuPDF + python-docx",
    )


@lru_cache(maxsize=1)
def image_capability() -> Capability:
    return require_modules("Pillow", "PIL")


@lru_cache(maxsize=1)
def image_enhancement_capability() -> Capability:
    missing = [name for name in ("cv2", "numpy", "PIL") if not has_module(name)]
    if missing:
        return Capability(
            "unavailable",
            f"高清图像增强缺少组件：{', '.join(missing)}",
            "OpenCV + Real-ESRGAN",
        )
    try:
        from .processors.image_enhancement import (
            realesrgan_binding_available,
            realesrgan_executable,
            realesrgan_gpu_available,
        )

        binding = realesrgan_binding_available()
        gpu = realesrgan_gpu_available()
        executable = realesrgan_executable()
    except Exception:
        binding = False
        gpu = False
        executable = None
    if binding and gpu:
        return Capability(
            "external",
            "已检测到 Real-ESRGAN NCNN Vulkan 本地 GPU 绑定",
            "Real-ESRGAN NCNN Vulkan + OpenCV 二检",
        )
    if binding and not gpu:
        return Capability(
            "ready",
            "Real-ESRGAN 已安装但未检测到 Vulkan GPU；无独显兼容增强可直接使用",
            "OpenCV 无独显兼容增强",
        )
    if executable is not None:
        return Capability(
            "external",
            f"已检测到本地 GPU 引擎：{executable.parent}",
            "Real-ESRGAN NCNN Vulkan + OpenCV 二检",
        )
    return Capability(
        "ready",
        "高保真预处理与结构二检可用；未检测到 Real-ESRGAN 时自动安全降级",
        "OpenCV 高保真增强（AI 可选）",
    )


@lru_cache(maxsize=1)
def office_structure_capability() -> Capability:
    return require_modules("Open XML Python 引擎", "docx", "openpyxl", "pptx")


def _known_soffice_paths() -> list[Path]:
    candidates: list[Path] = []
    explicit = (
        os.environ.get("LAYOUTLOOM_LIBREOFFICE_PATH", "").strip()
        or os.environ.get("DOCUFORGE_LIBREOFFICE_PATH", "").strip()
    )
    if explicit:
        configured = Path(explicit).expanduser()
        candidates.append(
            configured / "program/soffice.exe" if configured.is_dir() else configured
        )
    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        candidates.append(Path(bundle_root) / "libreoffice/program/soffice.exe")
    executable_root = Path(sys.executable).resolve().parent
    candidates.extend(
        [
            executable_root / "libreoffice/program/soffice.exe",
            executable_root / "_internal/libreoffice/program/soffice.exe",
        ]
    )
    project_root = Path(__file__).resolve().parents[1]
    candidates.append(
        project_root / "third_party/libreoffice/program/soffice.exe"
    )
    candidates.extend(
        [
        Path(os.environ.get("PROGRAMFILES", "C:/Program Files"))
        / "LibreOffice/program/soffice.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)"))
        / "LibreOffice/program/soffice.exe",
        ]
    )
    return candidates


def _known_microsoft_office_paths(component: str | None) -> list[Path]:
    executable_names = {
        "word": "WINWORD.EXE",
        "writer": "WINWORD.EXE",
        "excel": "EXCEL.EXE",
        "spreadsheets": "EXCEL.EXE",
        "powerpoint": "POWERPNT.EXE",
        "presentation": "POWERPNT.EXE",
        "ppt": "POWERPNT.EXE",
    }
    names = (
        ("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE")
        if component is None
        else (executable_names[component],)
    )
    roots = (
        Path(os.environ.get("PROGRAMFILES", "C:/Program Files")),
        Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)")),
    )
    directories = (
        "Microsoft Office/root/Office16",
        "Microsoft Office/Office16",
        "Microsoft Office/Office15",
        "Microsoft Office/Office14",
    )
    return [
        root / directory / name
        for root in roots
        for directory in directories
        for name in names
    ]


@lru_cache(maxsize=None)
def office_render_capability(component: str | None = None) -> Capability:
    component_name = None if component is None else component.strip().lower()
    microsoft_keys = {
        "word": "microsoft_word",
        "writer": "microsoft_word",
        "excel": "microsoft_excel",
        "spreadsheets": "microsoft_excel",
        "powerpoint": "microsoft_powerpoint",
        "presentation": "microsoft_powerpoint",
        "ppt": "microsoft_powerpoint",
    }
    wps_kinds = {
        "word": "writer",
        "writer": "writer",
        "excel": "spreadsheets",
        "spreadsheets": "spreadsheets",
        "powerpoint": "presentation",
        "presentation": "presentation",
        "ppt": "presentation",
    }
    if component_name is not None and component_name not in microsoft_keys:
        return Capability(
            "unavailable", f"未知 Office 组件：{component}", "Office 渲染器"
        )
    wps_failure_reason = ""
    microsoft_failure_reason = ""
    statuses = {}
    try:
        from .processors.office import detect_office_engines

        statuses = detect_office_engines()
    except Exception:
        statuses = {}

    try:
        from .processors.wps import detect_wps_engines

        wps_statuses = {str(key): value for key, value in detect_wps_engines().items()}
        if component_name is None:
            available_wps = [
                status for status in wps_statuses.values() if status.available
            ]
            unavailable_wps = [
                status for status in wps_statuses.values() if not status.available
            ]
            if unavailable_wps:
                wps_failure_reason = unavailable_wps[0].reason
        else:
            status = wps_statuses.get(wps_kinds[component_name])
            available_wps = [status] if status and status.available else []
            if status and not status.available:
                wps_failure_reason = status.reason
        if available_wps:
            names = "、".join(status.kind for status in available_wps)
            return Capability(
                "external", f"已检测到 WPS 自动化接口：{names}", "WPS Office COM"
            )
    except Exception:
        pass

    microsoft_key = (
        "microsoft_office" if component_name is None else microsoft_keys[component_name]
    )
    microsoft = statuses.get(microsoft_key)
    if microsoft and microsoft.available:
        engine_name = {
            "microsoft_word": "Microsoft Word COM",
            "microsoft_excel": "Microsoft Excel COM",
            "microsoft_powerpoint": "Microsoft PowerPoint COM",
        }.get(microsoft_key, "Microsoft Office COM")
        return Capability(
            "external",
            microsoft.reason or "已检测到 Microsoft Office",
            engine_name,
        )
    if microsoft:
        microsoft_failure_reason = microsoft.reason

    office_paths = _known_microsoft_office_paths(component_name)
    if any(path.is_file() for path in office_paths):
        if has_module("win32com"):
            return Capability(
                "external", "已检测到 Microsoft Office", "Microsoft Office COM"
            )
        microsoft_failure_reason = (
            "检测到 Office，但缺少 pywin32；安装后可启用高保真导出"
        )

    libreoffice = statuses.get("libreoffice")
    if libreoffice and libreoffice.available:
        detail = (
            str(libreoffice.executable)
            if libreoffice.executable
            else libreoffice.reason
        )
        return Capability("external", f"已检测到 LibreOffice：{detail}", "LibreOffice")

    soffice = find_executable("soffice", "libreoffice")
    if not soffice:
        soffice = next(
            (str(path) for path in _known_soffice_paths() if path.is_file()), None
        )
    if soffice:
        return Capability("external", f"已检测到 LibreOffice：{soffice}", "LibreOffice")

    if wps_failure_reason:
        detail = wps_failure_reason
        if microsoft_failure_reason:
            detail += f"；Microsoft Office：{microsoft_failure_reason}"
        return Capability("unavailable", detail, "Office 渲染器")
    return Capability(
        "unavailable",
        "高保真转换需要安装 WPS Office、Microsoft Office 或 LibreOffice",
        "Office 渲染器",
    )


@lru_cache(maxsize=1)
def microsoft_office_capability() -> Capability:
    try:
        from .processors.office import detect_office_engines

        status = detect_office_engines()["microsoft_office"]
    except Exception as exc:
        return Capability(
            "unavailable", f"Office COM 检测失败：{exc}", "Microsoft Office COM"
        )
    if status.available:
        return Capability(
            "external",
            status.reason or "已检测到 Microsoft Office COM",
            "Microsoft Office COM",
        )
    return Capability("unavailable", status.reason, "Microsoft Office COM")


def _microsoft_component_capability(key: str, label: str) -> Capability:
    try:
        from .processors.office import detect_office_engines

        status = detect_office_engines()[key]
    except Exception as exc:
        return Capability(
            "unavailable", f"{label} COM 检测失败：{exc}", f"Microsoft {label} COM"
        )
    if status.available:
        return Capability(
            "external",
            status.reason or f"已检测到 Microsoft {label} COM",
            f"Microsoft {label} COM",
        )
    return Capability("unavailable", status.reason, f"Microsoft {label} COM")


@lru_cache(maxsize=1)
def microsoft_word_capability() -> Capability:
    return _microsoft_component_capability("microsoft_word", "Word")


@lru_cache(maxsize=1)
def microsoft_excel_capability() -> Capability:
    return _microsoft_component_capability("microsoft_excel", "Excel")


@lru_cache(maxsize=1)
def microsoft_powerpoint_capability() -> Capability:
    return _microsoft_component_capability("microsoft_powerpoint", "PowerPoint")


@lru_cache(maxsize=1)
def ppt_render_capability() -> Capability:
    renderer = office_render_capability("powerpoint")
    if not renderer.runnable:
        return renderer
    if renderer.engine in {"Microsoft Office COM", "Microsoft PowerPoint COM"}:
        return Capability(
            "external", "使用 PowerPoint 原生幻灯片导出", "PowerPoint COM"
        )
    pdf_renderer = pdf_render_capability()
    if not pdf_renderer.runnable:
        return Capability(
            "unavailable",
            f"已检测到 {renderer.engine}，但缺少 Poppler 页面渲染器",
            f"{renderer.engine} + Poppler",
        )
    return Capability(
        "external",
        f"使用 {renderer.engine} 转 PDF，再由 Poppler 渲染页面",
        f"{renderer.engine} + Poppler",
    )


def _video_status():
    try:
        from .processors.video import detect_video_engine

        return detect_video_engine(), None
    except Exception as exc:
        return None, Capability("unavailable", f"视频引擎检测失败：{exc}", "FFmpeg")


def _missing_ffmpeg(status) -> Capability | None:
    if status.executable is None:
        return Capability("unavailable", status.reason, "FFmpeg")
    return None


@lru_cache(maxsize=1)
def slideshow_video_capability() -> Capability:
    status, error = _video_status()
    if error is not None:
        return error
    missing = _missing_ffmpeg(status)
    if missing is not None:
        return missing
    supported = tuple(
        encoder
        for encoder in status.encoders
        if encoder in {"libx264", "h264_nvenc", "h264_qsv", "h264_amf", "mpeg4"}
    )
    if not supported:
        return Capability(
            "unavailable",
            "已检测到 FFmpeg，但缺少幻灯片 MP4 所需的 H.264/MPEG-4 编码器",
            "FFmpeg 幻灯片编码",
        )
    return Capability(
        "external",
        f"已检测到 FFmpeg 幻灯片编码器：{'、'.join(supported)}",
        "FFmpeg 幻灯片编码",
    )


@lru_cache(maxsize=1)
def video_transform_capability() -> Capability:
    status, error = _video_status()
    if error is not None:
        return error
    missing = _missing_ffmpeg(status)
    if missing is not None:
        return missing
    supported = tuple(
        encoder
        for encoder in status.encoders
        if encoder in {"libx264", "h264_nvenc", "h264_qsv", "h264_amf"}
    )
    if not supported:
        return Capability(
            "unavailable",
            "已检测到 FFmpeg，但缺少默认 MP4 转码所需的 H.264 编码器",
            "FFmpeg 视频变换",
        )
    audio_supported = any(
        encoder in status.audio_encoders
        for encoder in ("aac", "libfdk_aac", "libvo_aacenc")
    )
    if not audio_supported:
        return Capability(
            "unavailable",
            "已检测到 H.264 编码器，但默认 MP4 转码还需要 AAC 音频编码器",
            "FFmpeg 视频变换",
        )
    return Capability(
        "external",
        f"已检测到默认 MP4/H.264 + AAC 转码能力；视频编码器：{'、'.join(supported)}",
        "FFmpeg 视频变换",
    )


@lru_cache(maxsize=1)
def audio_extraction_capability() -> Capability:
    status, error = _video_status()
    if error is not None:
        return error
    missing = _missing_ffmpeg(status)
    if missing is not None:
        return missing
    supported = tuple(status.audio_encoders)
    if "pcm_s16le" not in supported:
        return Capability(
            "unavailable",
            "已检测到 FFmpeg，但默认 WAV 音频提取需要 pcm_s16le 编码器",
            "FFmpeg 音频提取",
        )
    formats: list[str] = ["WAV"]
    if any(name in supported for name in ("aac", "libfdk_aac", "libvo_aacenc")):
        formats.append("AAC")
    if any(name in supported for name in ("libmp3lame", "libshine")):
        formats.append("MP3")
    return Capability(
        "external",
        f"已检测到独立音频提取能力；可用格式：{'、'.join(formats)}",
        "FFmpeg 音频提取",
    )


@lru_cache(maxsize=1)
def video_capability() -> Capability:
    """Backward-compatible probe for default video transformation tasks."""

    return video_transform_capability()


@lru_cache(maxsize=1)
def video_slide_extraction_capability() -> Capability:
    """Probe the local, decode-only lecture-video slide reconstruction stack."""

    missing: list[str] = []
    for module, label in (
        ("cv2", "OpenCV"),
        ("numpy", "NumPy"),
        ("pptx", "python-pptx"),
        ("docx", "python-docx"),
    ):
        if not has_module(module):
            missing.append(label)
    if missing:
        return Capability(
            "unavailable",
            f"缺少讲解视频提取所需组件：{'、'.join(missing)}",
            "OpenCV 时序重建",
        )
    return Capability(
        "ready",
        "使用 OpenCV 两遍解码、NumPy 多帧分析、python-pptx 生成完整 PPT，并由 python-docx 生成 Word 报告；不依赖云端 OCR",
        "OpenCV 时序重建 + python-pptx + python-docx",
    )


@lru_cache(maxsize=1)
def ppt_video_capability() -> Capability:
    native = microsoft_powerpoint_capability()
    if native.runnable:
        return Capability(
            "external",
            "PowerPoint 原生视频可直接运行并保留动画；静态模式自动按 "
            "WPS → Microsoft Office → LibreOffice 选择渲染器",
            "PowerPoint CreateVideo",
        )
    renderer = office_render_capability("powerpoint")
    if not renderer.runnable:
        return Capability(
            "unavailable",
            "PPT 转视频需要 Microsoft PowerPoint，或 WPS/LibreOffice 渲染器",
            "PPT 视频流水线",
        )
    if renderer.engine in {"Microsoft Office COM", "Microsoft PowerPoint COM"}:
        return Capability(
            "external",
            "PowerPoint 原生视频可直接运行并保留动画；静态模式另需 FFmpeg",
            "PowerPoint CreateVideo",
        )
    pdf_renderer = pdf_render_capability()
    if not pdf_renderer.runnable:
        return Capability(
            "unavailable",
            f"已检测到 {renderer.engine}，但静态视频还需要 PDF 页面渲染：{pdf_renderer.reason}",
            f"{renderer.engine} + Poppler + FFmpeg",
        )
    video = slideshow_video_capability()
    if not video.runnable:
        return Capability(
            "unavailable",
            f"已检测到 {renderer.engine} 与 Poppler，但缺少 FFmpeg 幻灯片编码能力：{video.reason}",
            f"{renderer.engine} + Poppler + FFmpeg",
        )
    return Capability(
        "external",
        f"使用 {renderer.engine} 转 PDF，经 Poppler 渲染页面，再由 FFmpeg 编码静态幻灯片视频",
        f"{renderer.engine} + Poppler + FFmpeg",
    )


@lru_cache(maxsize=1)
def ocr_capability() -> Capability:
    if not has_module("pdf2image") or not poppler_bin_path():
        return Capability(
            "unavailable",
            "OCR 需要 pdf2image 以及同目录的 Poppler pdftoppm/pdfinfo",
            "OCR 渲染流水线",
        )
    if not has_module("docx"):
        return Capability(
            "unavailable", "OCR 输出 Word 需要 python-docx", "OCR 文档输出"
        )
    executable = find_executable("tesseract")
    if executable and has_module("pytesseract"):
        return Capability("external", f"已检测到 {executable}", "Tesseract OCR")
    if has_module("paddleocr"):
        return Capability("external", "已安装 PaddleOCR", "PaddleOCR")
    return Capability(
        "unavailable",
        "扫描件识别需要 PaddleOCR，或 Tesseract + pytesseract；安装后可启用",
        "OCR 引擎",
    )


@lru_cache(maxsize=1)
def background_removal_capability() -> Capability:
    if has_module("rembg"):
        return Capability(
            "external",
            "已安装 rembg；首次运行所选模型时可能需要联网下载，之后可本地复用",
            "rembg",
        )
    return Capability(
        "unavailable", "AI 抠图需要安装 rembg；模型通常在首次运行时下载", "背景移除模型"
    )


@lru_cache(maxsize=None)
def format_capability(module: str, label: str) -> Capability:
    if has_module(module):
        return Capability("external", f"已安装 {module}", label)
    return Capability("unavailable", f"该格式需要可选组件 {module}", label)
