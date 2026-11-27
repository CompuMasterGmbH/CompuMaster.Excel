Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    <Test>
    Public Sub MergesUnmergesAndClearsCellRectanglesAndWorksheetContents()
        Dim input = Path.Combine(_directory, "cell-layout.xlsx")
        Dim output = Path.Combine(_directory, "cell-layout-out.xlsx")
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(input))
            Dim sheet = package.Workbook.Worksheets.Add("Layout")
            sheet.Cells("A1").Value = 8
            sheet.Cells("B2").Value = 4
            sheet.Cells("A1:B2").Style.Numberformat.Format = "0.00"
            sheet.Cells("C1").Formula = "SUM(A1:B2)"
            sheet.Column(3).Width = 28
            sheet.Column(4).Style.Numberformat.Format = "0.000"
            sheet.Row(3).Height = 25
            package.Save()
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .DisableInitialCalculation = True
        }
        Const cached As FormulaCalculationMode = FormulaCalculationMode.UseCachedValueOnly
        Using workbook As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            workbook.MergeCells("Layout", 0, 0, 1, 1)
            workbook.MergeCells("Layout", 0, 0, 1, 1)
            Assert.That(workbook.IsMergedCell("Layout", 1, 1), [Is].True)
            Assert.That(workbook.LookupCellValue(Of Integer)("Layout", 0, 0, cached), [Is].EqualTo(8))
            Assert.That(workbook.LookupCellValueAsObject("Layout", 1, 1, cached), [Is].Null)
            Assert.Throws(Of InvalidOperationException)(Sub() workbook.MergeCells("Layout", 1, 1, 2, 2))
            Assert.Throws(Of InvalidOperationException)(Sub() workbook.ClearCells("Layout", New ExcelCell("Layout", "A1", ExcelCell.ValueTypes.All)))

            workbook.UnMergeCells("Layout", 1, 1)
            Assert.That(workbook.IsMergedCell("Layout", 1, 1), [Is].False)
            Assert.That(workbook.LookupCellValueAsObject("Layout", 1, 1, cached), [Is].Null)
            workbook.RecalculateAll()
            Assert.That(workbook.LookupCellValue(Of Integer)("Layout", 0, 2, cached), [Is].EqualTo(8))

            workbook.MergeCells("Layout", 0, 0, 1, 1)
            workbook.ClearCells("Layout", New ExcelCell("Layout", "A1", ExcelCell.ValueTypes.All), New ExcelCell("Layout", "B2", ExcelCell.ValueTypes.All))
            Assert.That(workbook.IsMergedCell("Layout", 0, 0), [Is].False)
            Assert.That(workbook.LookupCellFormat("Layout", 0, 0), [Is].EqualTo("General"))
            workbook.RecalculateAll()
            Assert.That(workbook.LookupCellValue(Of Integer)("Layout", 0, 2, cached), [Is].Zero)
            workbook.ClearSheet("Layout")
            Assert.That(workbook.LookupCellFormula(New ExcelCell("Layout", "C1", ExcelCell.ValueTypes.All)), [Is].Null)
            Assert.That(workbook.LookupCellFormat("Layout", 9, 3), [Is].EqualTo("General"))
            workbook.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using document = SpreadsheetDocument.Open(output, False)
            Dim sheet = document.WorkbookPart.WorksheetParts.Single().Worksheet
            Assert.That(sheet.Descendants(Of Cell)(), [Is].Empty)
            Assert.That(sheet.Descendants(Of Column)().Single(Function(col) col.Min.Value = 3UI).Width.Value, [Is].EqualTo(28.0))
            Assert.That(sheet.Descendants(Of Row)().Single(Function(row) row.RowIndex.Value = 3UI).Height.Value, [Is].EqualTo(25.0))
        End Using
    End Sub

    <Test>
    Public Sub ClearingRangeResetsInheritedCellFormatsAndKeepsTheSurroundingLayout()
        Dim input = Path.Combine(_directory, "inherited-layout.xlsx")
        Dim output = Path.Combine(_directory, "inherited-layout-cleared.xlsx")
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(input))
            Dim sheet = package.Workbook.Worksheets.Add("Layout")
            sheet.Column(2).Style.Numberformat.Format = "0.000"
            sheet.Row(3).Style.Numberformat.Format = "0%"
            sheet.Cells("B2").Value = 123
            package.Save()
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using workbook As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            Assert.That(workbook.LookupCellFormat("Layout", 1, 1), [Is].EqualTo("0.000"))
            Assert.That(workbook.LookupCellFormat("Layout", 2, 2), [Is].EqualTo("0%"))
            workbook.ClearCells(New ExcelRange("Layout", "B2:C3"))
            workbook.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using workbook As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            Assert.That(workbook.LookupCellValueAsObject("Layout", 1, 1, FormulaCalculationMode.UseCachedValueOnly), [Is].Null)
            Assert.That(workbook.LookupCellFormat("Layout", 1, 1), [Is].EqualTo("General"))
            Assert.That(workbook.LookupCellFormat("Layout", 2, 2), [Is].EqualTo("General"))
            Assert.That(workbook.LookupCellFormat("Layout", 0, 1), [Is].EqualTo("0.000"))
            Assert.That(workbook.LookupCellFormat("Layout", 2, 0), [Is].EqualTo("0%"))
        End Using
    End Sub

    <Test>
    Public Sub FitsColumnsToFormattedTextAndPreservesNeighbouringLayout()
        Dim input = Path.Combine(_directory, "fit.xlsx")
        Dim output = Path.Combine(_directory, "fit-out.xlsx")
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(input))
            Dim sheet = package.Workbook.Worksheets.Add("Layout")
            sheet.Cells("A1").Value = "Wide visible column heading"
            sheet.Cells("B1").Value = "Value"
            sheet.Cells("C1").Value = "Hidden content"
            sheet.Cells("D1").Value = New String("W"c, 150)
            sheet.Cells("D1").Style.WrapText = True
            sheet.Column(2).Width = 30
            sheet.Column(2).Style.Numberformat.Format = "0.000"
            sheet.Column(2).OutlineLevel = 1
            sheet.Column(3).Width = 17
            sheet.Column(3).Hidden = True
            package.Save()
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = True}
        Using workbook As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            workbook.AutoFitColumns("Layout", 0, 12.0)
            Assert.That(workbook.LookupCellFormat("Layout", 4, 1), [Is].EqualTo("0.000"))
            workbook.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim columns = document.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants(Of Column)().ToArray()
            Assert.That(columns.Single(Function(col) col.Min.Value = 1UI).Width.Value, [Is].GreaterThan(12.0))
            Dim neighbour = columns.Single(Function(col) col.Min.Value = 2UI)
            Assert.That(neighbour.Width.Value, [Is].EqualTo(30.0))
            Assert.That(neighbour.OutlineLevel.Value, [Is].EqualTo(1))
        End Using
        Using workbook As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            workbook.AutoFitColumns("Layout", 12.0)
            workbook.SaveCopyAs(input, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(input, False)
            Dim columns = document.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants(Of Column)().ToArray()
            Assert.That(columns.Single(Function(col) col.Min.Value = 2UI).Width.Value, [Is].EqualTo(12.0))
            Dim hidden = columns.Single(Function(col) col.Min.Value = 3UI)
            Assert.That(hidden.Hidden.Value, [Is].True)
            Assert.That(hidden.Width.Value, [Is].EqualTo(17.0))
            Assert.That(columns.Single(Function(col) col.Min.Value = 4UI).Width.Value, [Is].EqualTo(12.0))
        End Using
    End Sub

    <Test>
    Public Sub InsertsAndRemovesColumnsWithStylesWidthsAndOutlineLevels()
        Dim input = Path.Combine(_directory, "grouped-columns.xlsx")
        Dim output = Path.Combine(_directory, "grouped-columns-out.xlsx")
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(input))
            Dim sheet = package.Workbook.Worksheets.Add("Layout")
            sheet.Cells("A1").Value = "Header"
            For i = 2 To 4
                sheet.Column(i).OutlineLevel = 1
            Next
            sheet.Column(2).Style.Numberformat.Format = "0.00"
            sheet.Column(2).Width = 22
            sheet.Column(3).Style.Numberformat.Format = "0.000"
            package.Save()
        End Using
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = True}
        Using workbook As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            workbook.AddColumn("Layout", 2, 1, updateFormulasAndReferences:=True)
            Assert.That(workbook.LookupCellFormat("Layout", 5, 2), [Is].EqualTo("0.00"))
            Assert.That(workbook.LookupCellFormat("Layout", 5, 3), [Is].EqualTo("0.000"))
            workbook.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim layout = document.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants(Of Column)()
            Dim inserted = layout.Single(Function(col) col.Min.Value <= 3UI AndAlso col.Max.Value >= 3UI)
            Assert.That(inserted.OutlineLevel.Value, [Is].EqualTo(1))
            Assert.That(inserted.Width.Value, [Is].EqualTo(22.0))
        End Using
        Using workbook As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            workbook.RemoveColumns("Layout", 2, 1, updateFormulasAndReferences:=True)
            Assert.That(workbook.LookupCellFormat("Layout", 5, 2), [Is].EqualTo("0.000"))
            workbook.SaveCopyAs(input, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(input, False)
            Dim layouts = document.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants(Of Column)()
            Dim remaining = layouts.Single(Function(col) col.Min.Value <= 4UI AndAlso col.Max.Value >= 4UI)
            Assert.That(remaining.OutlineLevel.Value, [Is].EqualTo(1))
        End Using
    End Sub


End Class
