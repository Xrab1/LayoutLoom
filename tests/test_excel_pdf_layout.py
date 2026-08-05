from __future__ import annotations

import unittest
from typing import Any, Iterator

from docuforge.processors.excel_pdf_layout import prepare_excel_workbook_for_pdf

XL_PAGE_BREAK_AUTOMATIC = -4105
XL_SHEET_VISIBLE = -1
XL_SHEET_HIDDEN = 0


class FakeRange:
    def __init__(self, width: float, height: float) -> None:
        self.Width = width
        self.Height = height
        self.Address = "$A$1:$B$4"


class FakePageBreak:
    def __init__(self, address: str, break_type: int = XL_PAGE_BREAK_AUTOMATIC) -> None:
        self.Type = break_type
        self.Location = type("Location", (), {"Address": address})()


class FakeCollection:
    def __init__(self, items: list[Any]) -> None:
        self._items = list(items)

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> Any:
        return self._items[index - 1]

    def __iter__(self) -> Iterator[Any]:
        return iter(self._items)


class FakePageSetup:
    def __init__(
        self,
        *,
        zoom: int | bool = 100,
        fit_wide: int | bool = 1,
        fit_tall: int | bool = 1,
        paper_size: int = 1,
        orientation: int = 1,
    ) -> None:
        object.__setattr__(self, "writes", [])
        object.__setattr__(self, "_tracking", False)
        self.Zoom = zoom
        self.FitToPagesWide = fit_wide
        self.FitToPagesTall = fit_tall
        self.PaperSize = paper_size
        self.Orientation = orientation
        self.LeftMargin = 54.0
        self.RightMargin = 54.0
        self.TopMargin = 72.0
        self.BottomMargin = 72.0
        self.HeaderMargin = 36.0
        self.FooterMargin = 36.0
        self.PrintArea = ""
        self.PrintTitleRows = ""
        self.PrintTitleColumns = ""
        object.__setattr__(self, "_tracking", True)

    def __setattr__(self, name: str, value: Any) -> None:
        if getattr(self, "_tracking", False) and not name.startswith("_"):
            self.writes.append((name, value))
        object.__setattr__(self, name, value)


class FakeWorksheet:
    def __init__(
        self,
        name: str,
        *,
        width: float,
        height: float,
        visible: int = XL_SHEET_VISIBLE,
        vertical_breaks: int = 0,
    ) -> None:
        self.Name = name
        self.Visible = visible
        self.UsedRange = FakeRange(width, height)
        self.PageSetup = FakePageSetup()
        self.HPageBreaks = FakeCollection([])
        self.VPageBreaks = FakeCollection(
            [FakePageBreak("$B$1") for _ in range(vertical_breaks)]
        )
        self.Shapes = FakeCollection([])


class FakeWorkbook:
    def __init__(self, worksheets: list[FakeWorksheet]) -> None:
        self.Worksheets = FakeCollection(worksheets)
        self.Sheets = self.Worksheets


class ExcelPdfLayoutTests(unittest.TestCase):
    def test_preserve_makes_no_page_setup_writes(self) -> None:
        sheet = FakeWorksheet("登记表", width=800.5, height=346.0, vertical_breaks=1)

        prepare_excel_workbook_for_pdf(
            FakeWorkbook([sheet]),
            layout="preserve",
            paper="auto",
            orientation="auto",
        )

        self.assertEqual(sheet.PageSetup.writes, [])

    def test_fit_width_uses_one_page_wide_and_unlimited_height(self) -> None:
        sheet = FakeWorksheet("宽表", width=800.5, height=346.0)

        prepare_excel_workbook_for_pdf(
            FakeWorkbook([sheet]),
            layout="fit_width",
            paper="preserve",
            orientation="preserve",
            margin="preserve",
        )

        self.assertIs(sheet.PageSetup.Zoom, False)
        self.assertEqual(sheet.PageSetup.FitToPagesWide, 1)
        self.assertIs(sheet.PageSetup.FitToPagesTall, False)

    def test_single_page_fits_both_width_and_height(self) -> None:
        sheet = FakeWorksheet("单页表", width=800.5, height=900.0)

        prepare_excel_workbook_for_pdf(
            FakeWorkbook([sheet]),
            layout="single_page",
            paper="preserve",
            orientation="preserve",
        )

        self.assertIs(sheet.PageSetup.Zoom, False)
        self.assertEqual(sheet.PageSetup.FitToPagesWide, 1)
        self.assertEqual(sheet.PageSetup.FitToPagesTall, 1)

    def test_smart_wide_sheet_uses_a4_landscape_and_one_page_width(self) -> None:
        sheet = FakeWorksheet(
            "高校毕业生登记表",
            width=800.5,
            height=346.0,
            vertical_breaks=1,
        )

        prepare_excel_workbook_for_pdf(
            FakeWorkbook([sheet]), layout="smart", paper="auto", orientation="auto"
        )

        self.assertEqual(sheet.PageSetup.PaperSize, 9)
        self.assertEqual(sheet.PageSetup.Orientation, 2)
        self.assertIs(sheet.PageSetup.Zoom, False)
        self.assertEqual(sheet.PageSetup.FitToPagesWide, 1)
        self.assertIs(sheet.PageSetup.FitToPagesTall, False)
        self.assertEqual(sheet.PageSetup.LeftMargin, 18.0)
        self.assertEqual(sheet.PageSetup.RightMargin, 18.0)
        self.assertEqual(sheet.PageSetup.TopMargin, 36.0)
        self.assertEqual(sheet.PageSetup.BottomMargin, 36.0)

    def test_smart_narrow_tall_sheet_uses_a4_portrait(self) -> None:
        sheet = FakeWorksheet("纵向表", width=300.0, height=900.0)
        sheet.PageSetup.PaperSize = 1
        sheet.PageSetup.Orientation = 2
        sheet.PageSetup.writes.clear()

        prepare_excel_workbook_for_pdf(
            FakeWorkbook([sheet]), layout="smart", paper="auto", orientation="auto"
        )

        self.assertEqual(sheet.PageSetup.PaperSize, 9)
        self.assertEqual(sheet.PageSetup.Orientation, 1)

    def test_hidden_sheet_is_not_modified(self) -> None:
        visible = FakeWorksheet("可见", width=800.5, height=346.0)
        hidden = FakeWorksheet(
            "隐藏", width=800.5, height=346.0, visible=XL_SHEET_HIDDEN
        )

        prepare_excel_workbook_for_pdf(
            FakeWorkbook([visible, hidden]),
            layout="fit_width",
            paper="a4",
            orientation="landscape",
        )

        self.assertTrue(visible.PageSetup.writes)
        self.assertEqual(hidden.PageSetup.writes, [])

    def test_explicit_paper_and_orientation_override_smart_choices(self) -> None:
        sheet = FakeWorksheet("宽表", width=800.5, height=346.0)

        prepare_excel_workbook_for_pdf(
            FakeWorkbook([sheet]),
            layout="smart",
            paper="letter",
            orientation="portrait",
        )

        self.assertEqual(sheet.PageSetup.PaperSize, 1)
        self.assertEqual(sheet.PageSetup.Orientation, 1)

    def test_preserve_margin_keeps_existing_margins(self) -> None:
        sheet = FakeWorksheet("保留边距", width=800.5, height=346.0)

        prepare_excel_workbook_for_pdf(
            FakeWorkbook([sheet]),
            layout="smart",
            paper="auto",
            orientation="auto",
            margin="preserve",
        )

        self.assertEqual(sheet.PageSetup.LeftMargin, 54.0)
        self.assertEqual(sheet.PageSetup.RightMargin, 54.0)
        self.assertEqual(sheet.PageSetup.TopMargin, 72.0)
        self.assertEqual(sheet.PageSetup.BottomMargin, 72.0)
        margin_names = {
            "LeftMargin",
            "RightMargin",
            "TopMargin",
            "BottomMargin",
        }
        self.assertFalse(
            any(name in margin_names for name, _value in sheet.PageSetup.writes)
        )


if __name__ == "__main__":
    unittest.main()
