Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    <TestCase(False)>
    <TestCase(True)>
    Public Sub ExposesExcelDateAndBlankValuesWithoutChangingFormulaInputs(use1904Dates As Boolean)
        Dim filePath = Path.Combine(_directory, "cell-types.xlsx")
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(filePath))
            package.Workbook.Date1904 = use1904Dates
            Dim sheet = package.Workbook.Worksheets.Add("Types")
            sheet.Cells("A1").Formula = "EOMONTH(DATE(2024,2,15),0)"
            sheet.Cells("A1").Style.Numberformat.Format = "dd.mm.yyyy"
            sheet.Cells("B1").Formula = "A1+1"
            sheet.Cells("A2").Value = 60.0
            sheet.Cells("A2").Style.Numberformat.Format = "yyyy-mm-dd"
            sheet.Cells("A3").Value = 0.25
            sheet.Cells("A3").Style.Numberformat.Format = "yyyy-mm-dd hh:mm"
            sheet.Cells("A4").Value = 0.25
            sheet.Cells("A4").Style.Numberformat.Format = "hh:mm"
            sheet.Cells("A5").Value = 1.25
            sheet.Cells("A5").Style.Numberformat.Format = "[h]:mm"
            sheet.Cells("A6").Formula = "CHAR(32)"
            sheet.Cells("B6").Formula = "LEN(A6)"
            sheet.Cells("A7").Value = " content "
            sheet.Cells("A8").Value = 45000.0
            sheet.Cells("A8").Style.Numberformat.Format = "0 ""days"""
            package.Save()
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False}
        Const cached As FormulaCalculationMode = FormulaCalculationMode.UseCachedValueOnly
        Using workbook As New OpenXmlExcelDataOperations(filePath, OpenMode.OpenExistingFile, options)
            workbook.RecalculateAll()
            Dim monthEnd As New DateTime(2024, 2, 29)
            Dim monthEndSerial = monthEnd.ToOADate() - If(use1904Dates, 1462.0, 0.0)
            Assert.That(workbook.LookupCellValueAsObject("Types", 0, 0, cached), [Is].EqualTo(monthEnd))
            Assert.That(workbook.LookupCellValue(Of DateTime)("Types", 0, 0, cached), [Is].EqualTo(monthEnd))
            Assert.That(workbook.TryLookupCellValue(Of DateTime)("Types", 0, 0).Value, [Is].EqualTo(monthEnd))
            Assert.That(workbook.LookupCellValue(Of Double)("Types", 0, 0, cached), [Is].EqualTo(monthEndSerial))
            Assert.That(workbook.LookupCellValue(Of Double)("Types", 0, 1, cached), [Is].EqualTo(monthEndSerial + 1.0))
            Assert.That(workbook.LookupCellFormattedText("Types", 0, 0), [Is].EqualTo("29.02.2024"))
            Assert.That(workbook.LookupCellValueAsObject("Types", 1, 0, cached), [Is].EqualTo(
                If(use1904Dates, New DateTime(1904, 3, 1), New DateTime(1900, 2, 28))))
            Assert.That(workbook.LookupCellValueAsObject("Types", 2, 0, cached), [Is].EqualTo(New DateTime(1899, 12, 30, 6, 0, 0)))
            Assert.That(workbook.LookupCellValueAsObject("Types", 3, 0, cached), [Is].EqualTo(0.25))
            Assert.That(workbook.LookupCellValueAsObject("Types", 4, 0, cached), [Is].EqualTo(1.25))
            Assert.That(workbook.LookupCellValueAsObject("Types", 5, 0, cached), [Is].Null)
            Assert.That(workbook.LookupCellValue(Of Decimal)("Types", 5, 0, cached), [Is].EqualTo(0D))
            Assert.That(workbook.TryLookupCellValue(Of Decimal)("Types", 5, 0).HasValue, [Is].False)
            Assert.That(workbook.HasMissingCachedFormulaResult("Types", 5, 0), [Is].False)
            Assert.That(workbook.LookupCellValue(Of Integer)("Types", 5, 1, cached), [Is].EqualTo(1))
            Assert.That(workbook.LookupCellValueAsObject("Types", 6, 0, cached), [Is].EqualTo(" content "))
            Assert.That(workbook.LookupCellValueAsObject("Types", 7, 0, cached), [Is].EqualTo(45000.0))
        End Using
    End Sub

End Class
