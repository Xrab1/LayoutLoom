"""Shared Excel/WPS page-layout planning for high-fidelity PDF export.

The Office renderers must use the application's own print engine, but many
workbooks carry stale or contradictory page settings (for example ``Zoom=100``
alongside ``FitToPagesWide=1``).  This module makes conservative, in-memory
PageSetup adjustments before export.  The source workbook is always closed
without saving by the caller.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, Literal

from ..models import ValidationError

ExcelPdfLayout = Literal["smart", "preserve", "fit_width", "single_page"]
ExcelPdfPaper = Literal["auto", "preserve", "a4", "a3", "letter"]
ExcelPdfOrientation = Literal["auto", "preserve", "portrait", "landscape"]
ExcelPdfMargin = Literal["auto", "preserve", "narrow"]


_LAYOUTS = {"smart", "preserve", "fit_width", "single_page"}
_PAPERS = {"auto", "preserve", "a4", "a3", "letter"}
_ORIENTATIONS = {"auto", "preserve", "portrait", "landscape"}
_MARGINS = {"auto", "preserve", "narrow"}

# Excel/WPS constants.  Numeric values are used deliberately so the helper does
# not depend on generated COM constant modules being installed.
_XL_PORTRAIT = 1
_XL_LANDSCAPE = 2
_XL_PAPER_LETTER = 1
_XL_PAPER_A3 = 8
_XL_PAPER_A4 = 9

_PAPER_CODES = {
    "letter": _XL_PAPER_LETTER,
    "a3": _XL_PAPER_A3,
    "a4": _XL_PAPER_A4,
}
_PAPER_NAMES = {value: key for key, value in _PAPER_CODES.items()}

# Physical paper dimensions in points (1/72 inch), portrait orientation.
_PAPER_POINTS = {
    "letter": (612.0, 792.0),
    "a4": (595.28, 841.89),
    "a3": (841.89, 1190.55),
}

_TARGET_SCALE = 0.90
_MIN_READABLE_SCALE = 0.68
_MAX_SMART_PAGES_WIDE = 4
_NARROW_LEFT_RIGHT = 18.0  # 0.25 inch
_NARROW_TOP_BOTTOM = 36.0  # 0.50 inch


@dataclass(frozen=True)
class ExcelPdfSheetPlan:
    """The selected layout for one visible worksheet."""

    sheet_name: str
    content_width: float
    content_height: float
    paper: str
    orientation: str
    pages_wide: int
    fit_to_pages_tall: int | bool
    estimated_scale: float


@dataclass(frozen=True)
class _Candidate:
    paper: str
    orientation: str
    margins: tuple[float, float, float, float]
    margin_changed: bool
    printable_width: float
    printable_height: float
    scale: float
    pages_tall: int


def _normalize_choice(value: Any, allowed: set[str], name: str) -> str:
    result = str(value).strip().lower()
    if result not in allowed:
        choices = "、".join(sorted(allowed))
        raise ValidationError(f"{name} 选项无效；可选：{choices}")
    return result


def normalize_excel_pdf_options(
    layout: ExcelPdfLayout | str = "smart",
    paper: ExcelPdfPaper | str = "auto",
    orientation: ExcelPdfOrientation | str = "auto",
    margin: ExcelPdfMargin | str = "auto",
) -> tuple[str, str, str, str]:
    """Validate and normalize public Excel-to-PDF layout choices."""

    return (
        _normalize_choice(layout, _LAYOUTS, "Excel PDF 页面布局"),
        _normalize_choice(paper, _PAPERS, "Excel PDF 纸张"),
        _normalize_choice(orientation, _ORIENTATIONS, "Excel PDF 页面方向"),
        _normalize_choice(margin, _MARGINS, "Excel PDF 页边距"),
    )


def _collection_items(collection: Any) -> list[Any]:
    """Return items from a one-based COM collection or a normal sequence."""

    try:
        count = int(collection.Count)
    except Exception:
        try:
            return list(collection)
        except Exception:
            return []
    result: list[Any] = []
    for index in range(1, count + 1):
        try:
            item = collection.Item(index)
        except Exception:
            try:
                item = collection(index)
            except Exception:
                try:
                    item = collection[index - 1]
                except Exception:
                    continue
        result.append(item)
    return result


def _visible_worksheets(workbook: Any) -> list[Any]:
    try:
        worksheets = workbook.Worksheets
    except Exception:
        return []
    result: list[Any] = []
    for sheet in _collection_items(worksheets):
        try:
            visibility = sheet.Visible
        except Exception:
            visibility = -1
        # Excel: visible=-1, hidden=0, very-hidden=2.  Some WPS builds expose
        # a Boolean instead, so True is accepted as visible as well.
        if visibility is False or visibility == 0 or visibility == 2:
            continue
        result.append(sheet)
    return result


def _number(value: Any, fallback: float) -> float:
    try:
        result = float(value)
    except (TypeError, ValueError, OverflowError):
        return fallback
    if not math.isfinite(result) or result < 0:
        return fallback
    return result


def _page_setup_number(page_setup: Any, name: str, fallback: float) -> float:
    try:
        return _number(getattr(page_setup, name), fallback)
    except Exception:
        return fallback


def _range_for_layout(sheet: Any, page_setup: Any) -> Any:
    try:
        print_area = str(page_setup.PrintArea or "").strip()
    except Exception:
        print_area = ""
    if print_area:
        try:
            return sheet.Range(print_area)
        except Exception:
            pass
    return sheet.UsedRange


def _content_dimensions(sheet: Any, page_setup: Any) -> tuple[float, float]:
    used = _range_for_layout(sheet, page_setup)
    width = max(1.0, _number(getattr(used, "Width", 1.0), 1.0))
    height = max(1.0, _number(getattr(used, "Height", 1.0), 1.0))

    # Shapes do not always expand UsedRange.  When no explicit PrintArea exists,
    # include their physical bounds in the planning dimensions without changing
    # the workbook's actual print area.
    try:
        has_print_area = bool(str(page_setup.PrintArea or "").strip())
    except Exception:
        has_print_area = False
    if has_print_area:
        return width, height

    origin_left = _number(getattr(used, "Left", 0.0), 0.0)
    origin_top = _number(getattr(used, "Top", 0.0), 0.0)
    right = origin_left + width
    bottom = origin_top + height
    try:
        shapes = _collection_items(sheet.Shapes)
    except Exception:
        shapes = []
    for shape in shapes:
        left = _number(getattr(shape, "Left", origin_left), origin_left)
        top = _number(getattr(shape, "Top", origin_top), origin_top)
        shape_width = _number(getattr(shape, "Width", 0.0), 0.0)
        shape_height = _number(getattr(shape, "Height", 0.0), 0.0)
        right = max(right, left + shape_width)
        bottom = max(bottom, top + shape_height)
        origin_left = min(origin_left, left)
        origin_top = min(origin_top, top)
    return max(1.0, right - origin_left), max(1.0, bottom - origin_top)


def _has_header_or_footer(page_setup: Any) -> bool:
    for name in (
        "LeftHeader",
        "CenterHeader",
        "RightHeader",
        "LeftFooter",
        "CenterFooter",
        "RightFooter",
    ):
        try:
            if str(getattr(page_setup, name) or "").strip():
                return True
        except Exception:
            continue
    return False


def _current_margins(page_setup: Any) -> tuple[float, float, float, float]:
    return (
        _page_setup_number(page_setup, "LeftMargin", 36.0),
        _page_setup_number(page_setup, "RightMargin", 36.0),
        _page_setup_number(page_setup, "TopMargin", 54.0),
        _page_setup_number(page_setup, "BottomMargin", 54.0),
    )


def _narrow_margins(
    current: tuple[float, float, float, float], *, keep_vertical: bool
) -> tuple[float, float, float, float]:
    left, right, top, bottom = current
    return (
        min(left, _NARROW_LEFT_RIGHT),
        min(right, _NARROW_LEFT_RIGHT),
        top if keep_vertical else min(top, _NARROW_TOP_BOTTOM),
        bottom if keep_vertical else min(bottom, _NARROW_TOP_BOTTOM),
    )


def _paper_candidates(page_setup: Any, paper: str) -> list[str]:
    if paper == "auto":
        return ["a4", "a3"]
    if paper == "preserve":
        try:
            current = _PAPER_NAMES.get(int(page_setup.PaperSize))
        except Exception:
            current = None
        return [current or "a4"]
    return [paper]


def _orientation_candidates(page_setup: Any, orientation: str) -> list[str]:
    if orientation == "auto":
        return ["portrait", "landscape"]
    if orientation == "preserve":
        try:
            current = int(page_setup.Orientation)
        except Exception:
            current = _XL_PORTRAIT
        return ["landscape" if current == _XL_LANDSCAPE else "portrait"]
    return [orientation]


def _margin_candidates(
    page_setup: Any, margin: str
) -> list[tuple[tuple[float, float, float, float], bool]]:
    current = _current_margins(page_setup)
    narrow = _narrow_margins(current, keep_vertical=_has_header_or_footer(page_setup))
    if margin == "preserve":
        return [(current, False)]
    if margin == "narrow":
        return [(narrow, narrow != current)]
    result = [(current, False)]
    if narrow != current:
        result.append((narrow, True))
    return result


def _make_candidates(
    page_setup: Any,
    content_width: float,
    content_height: float,
    *,
    paper: str,
    orientation: str,
    margin: str,
) -> list[_Candidate]:
    result: list[_Candidate] = []
    for paper_name in _paper_candidates(page_setup, paper):
        base_width, base_height = _PAPER_POINTS[paper_name]
        for orientation_name in _orientation_candidates(page_setup, orientation):
            if orientation_name == "landscape":
                page_width, page_height = base_height, base_width
            else:
                page_width, page_height = base_width, base_height
            for margins, changed in _margin_candidates(page_setup, margin):
                left, right, top, bottom = margins
                printable_width = max(72.0, page_width - left - right)
                printable_height = max(72.0, page_height - top - bottom)
                scale = min(1.0, printable_width / max(1.0, content_width))
                pages_tall = max(
                    1,
                    math.ceil(
                        max(1.0, content_height) * scale / printable_height - 1e-9
                    ),
                )
                result.append(
                    _Candidate(
                        paper_name,
                        orientation_name,
                        margins,
                        changed,
                        printable_width,
                        printable_height,
                        scale,
                        pages_tall,
                    )
                )
    return result


def _select_candidate(candidates: list[_Candidate]) -> _Candidate:
    if not candidates:
        raise ValidationError("无法计算 Excel PDF 页面布局")
    readable = [
        candidate for candidate in candidates if candidate.scale >= _TARGET_SCALE
    ]
    if readable:
        # Prefer the smallest practical paper.  Within the same paper, preserve
        # near-original text size, then minimize vertical pages and avoid margin
        # changes when they do not improve the result.
        return min(
            readable,
            key=lambda candidate: (
                _PAPER_POINTS[candidate.paper][0] * _PAPER_POINTS[candidate.paper][1],
                -candidate.scale,
                candidate.pages_tall,
                candidate.margin_changed,
                candidate.orientation == "landscape",
            ),
        )
    # If no one-page-wide candidate reaches the readability target, use the
    # clearest paper/orientation and let smart mode add horizontal pages below.
    return min(
        candidates,
        key=lambda candidate: (
            -candidate.scale,
            _PAPER_POINTS[candidate.paper][0] * _PAPER_POINTS[candidate.paper][1],
            candidate.pages_tall,
            candidate.margin_changed,
        ),
    )


def _apply_plan(
    page_setup: Any,
    candidate: _Candidate,
    *,
    layout: str,
    paper: str,
    orientation: str,
    margin: str,
    pages_wide: int,
) -> None:
    if paper != "preserve":
        page_setup.PaperSize = _PAPER_CODES[candidate.paper]
    if orientation != "preserve":
        page_setup.Orientation = (
            _XL_LANDSCAPE if candidate.orientation == "landscape" else _XL_PORTRAIT
        )
    if margin != "preserve" and candidate.margin_changed:
        left, right, top, bottom = candidate.margins
        page_setup.LeftMargin = left
        page_setup.RightMargin = right
        page_setup.TopMargin = top
        page_setup.BottomMargin = bottom

    # Zoom must be disabled or Excel/WPS silently ignores FitToPagesWide/Tall.
    page_setup.Zoom = False
    page_setup.FitToPagesWide = pages_wide
    page_setup.FitToPagesTall = 1 if layout == "single_page" else False


def prepare_excel_workbook_for_pdf(
    workbook: Any,
    application: Any | None = None,
    *,
    layout: ExcelPdfLayout | str = "smart",
    paper: ExcelPdfPaper | str = "auto",
    orientation: ExcelPdfOrientation | str = "auto",
    margin: ExcelPdfMargin | str = "auto",
) -> list[ExcelPdfSheetPlan]:
    """Optimize visible worksheets for PDF export without saving the workbook.

    ``smart`` chooses A4/A3 and portrait/landscape per worksheet, fits a normal
    table to one page wide, and allows long sheets to continue vertically.
    Extremely wide sheets may use up to four horizontal pages to avoid illegible
    text.  ``preserve`` performs no PageSetup writes at all.
    """

    layout_name, paper_name, orientation_name, margin_name = (
        normalize_excel_pdf_options(layout, paper, orientation, margin)
    )
    if layout_name == "preserve":
        return []

    plans: list[ExcelPdfSheetPlan] = []
    communication_changed = False
    original_print_communication: Any = True
    if application is not None:
        try:
            original_print_communication = application.PrintCommunication
            application.PrintCommunication = False
            communication_changed = True
        except Exception:
            communication_changed = False
    try:
        for sheet in _visible_worksheets(workbook):
            try:
                page_setup = sheet.PageSetup
                content_width, content_height = _content_dimensions(sheet, page_setup)
            except Exception:
                # Worksheets without a normal grid/PageSetup (or a damaged sheet)
                # are left to the native renderer instead of blocking other tabs.
                continue
            candidates = _make_candidates(
                page_setup,
                content_width,
                content_height,
                paper=paper_name,
                orientation=orientation_name,
                margin=margin_name,
            )
            candidate = _select_candidate(candidates)
            pages_wide = 1
            if layout_name == "smart" and candidate.scale < _MIN_READABLE_SCALE:
                pages_wide = min(
                    _MAX_SMART_PAGES_WIDE,
                    max(1, math.ceil(_MIN_READABLE_SCALE / candidate.scale)),
                )
            _apply_plan(
                page_setup,
                candidate,
                layout=layout_name,
                paper=paper_name,
                orientation=orientation_name,
                margin=margin_name,
                pages_wide=pages_wide,
            )
            plans.append(
                ExcelPdfSheetPlan(
                    str(getattr(sheet, "Name", f"Sheet{len(plans) + 1}")),
                    content_width,
                    content_height,
                    candidate.paper,
                    candidate.orientation,
                    pages_wide,
                    1 if layout_name == "single_page" else False,
                    min(1.0, candidate.scale * pages_wide),
                )
            )
    finally:
        if communication_changed:
            try:
                application.PrintCommunication = original_print_communication
            except Exception:
                pass
    return plans


__all__ = [
    "ExcelPdfSheetPlan",
    "normalize_excel_pdf_options",
    "prepare_excel_workbook_for_pdf",
]
