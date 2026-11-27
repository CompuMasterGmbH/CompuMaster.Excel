Imports Color = System.Drawing.Color
Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Validation
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    <Test>
    Public Sub FormatsRangesInANewWorkbookAndPreservesValuesAndStylesOnReopen()
        Dim output = Path.Combine(_directory, "formatted.xlsx")
        Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, options)
            Dim formatter As IExcelWorksheetFormatting = workbook
            workbook.WriteCellValue("Sheet1", 0, 0, 0.25)
            workbook.WriteCellFormula("Sheet1", 0, 1, "A1*2", False)
            Dim range As New ExcelRange("Sheet1", "A1:B2")
            formatter.ApplyCellFormatting(range, New ExcelCellFormatting With {
                .NumberFormat = "0.000%", .FontName = "Calibri", .FontSize = 12, .Bold = True,
                .FontColor = Color.Navy, .FillColor = Color.LightYellow, .HorizontalAlignment = ExcelHorizontalAlignment.Right,
                .VerticalAlignment = ExcelVerticalAlignment.Center, .WrapText = True,
                .OutsideBorder = New ExcelCellBorder(ExcelBorderLineStyle.Thick, Color.Red),
                .InsideBorder = New ExcelCellBorder(ExcelBorderLineStyle.Thin)
            })
            ' A second patch also applies to the blank cell without resetting its existing style.
            formatter.ApplyCellFormatting(New ExcelRange("Sheet1", "A1:A2"), New ExcelCellFormatting With {.Italic = True})
            formatter.SetColumnWidth("Sheet1", 0, 18.5)
            formatter.SetRowHeight("Sheet1", 1, 27.5)
            Assert.That(workbook.LookupCellFormat("Sheet1", 0, 0), [Is].EqualTo("0.000%"))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using workbook As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            Assert.That(workbook.LookupCellValue(Of Double)("Sheet1", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(0.25))
            Assert.That(workbook.LookupCellFormula(New ExcelCell("Sheet1", "B1", ExcelCell.ValueTypes.All)), [Is].EqualTo("A1*2"))
            Assert.That(workbook.LookupCellFormat("Sheet1", 1, 0), [Is].EqualTo("0.000%"))
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim errors = New OpenXmlValidator().Validate(document).Select(Function(failure) failure.Description).ToArray()
            Assert.That(errors, [Is].Empty, String.Join(Environment.NewLine, errors))
            Dim sheet = document.WorkbookPart.WorksheetParts.Single().Worksheet
            Dim styles = document.WorkbookPart.WorkbookStylesPart.Stylesheet
            Dim cells = sheet.Descendants(Of Cell)().ToDictionary(Function(cell) cell.CellReference.Value)
            Dim firstFormat = styles.CellFormats.Elements(Of CellFormat)().ElementAt(CInt(cells("A1").StyleIndex.Value))
            Dim nextFormat = styles.CellFormats.Elements(Of CellFormat)().ElementAt(CInt(cells("B1").StyleIndex.Value))
            Dim font = styles.Fonts.Elements(Of Font)().ElementAt(CInt(firstFormat.FontId.Value))
            Dim fill = styles.Fills.Elements(Of Fill)().ElementAt(CInt(firstFormat.FillId.Value))
            Dim border = styles.Borders.Elements(Of Border)().ElementAt(CInt(firstFormat.BorderId.Value))
            Assert.That(font.FontName.Val.Value, [Is].EqualTo("Calibri"))
            Assert.That(font.FontSize.Val.Value, [Is].EqualTo(12))
            Assert.That(font.Bold.Val.Value AndAlso font.Italic.Val.Value, [Is].True)
            Assert.That(font.Color.Rgb.Value, [Is].EqualTo("FF000080"))
            Assert.That(fill.PatternFill.ForegroundColor.Rgb.Value, [Is].EqualTo("FFFFFFE0"))
            Assert.That(firstFormat.Alignment.Horizontal.Value, [Is].EqualTo(HorizontalAlignmentValues.Right))
            Assert.That(firstFormat.Alignment.Vertical.Value, [Is].EqualTo(VerticalAlignmentValues.Center))
            Assert.That(firstFormat.Alignment.WrapText.Value, [Is].True)
            Assert.That(firstFormat.NumberFormatId.Value, [Is].EqualTo(nextFormat.NumberFormatId.Value))
            Assert.That(border.LeftBorder.Style.Value, [Is].EqualTo(BorderStyleValues.Thick))
            Assert.That(border.RightBorder.Style.Value, [Is].EqualTo(BorderStyleValues.Thin))
            Assert.That(sheet.Descendants(Of Column)().Single().Width.Value, [Is].EqualTo(18.5))
            Assert.That(sheet.Descendants(Of Row)().Single(Function(row) row.RowIndex.Value = 2UI).Height.Value, [Is].EqualTo(27.5))
        End Using
    End Sub

End Class
