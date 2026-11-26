Imports System.IO
Imports System.Xml
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    <Test>
    Public Sub RecalculatesAfterAnEpplus4EditOfItsSavedOutput()
        Dim filePath = Path.Combine(_directory, "roundtrip.xlsx")
        Using epplus As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(filePath))
            Dim sheet = epplus.Workbook.Worksheets.Add("Inputs")
            sheet.Cells("A8").Value = 10
            sheet.Cells("B8").Formula = "A8+1"
            epplus.Save()
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadWrite) With {.DisableCalculationEngine = False, .DisableInitialCalculation = True}
        Using workbook As New OpenXmlExcelDataOperations(filePath, OpenMode.OpenExistingFile, options)
            workbook.RecalculateAll()
            workbook.Save(recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using epplus As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(filePath))
            Dim sheet = epplus.Workbook.Worksheets("Inputs")
            Assert.That(sheet.Cells("B8").Formula, [Is].EqualTo("A8+1"))
            Assert.That(sheet.Cells("B8").GetValue(Of Double)(), [Is].EqualTo(11.0))
            sheet.Cells("A8").Value = 20
            epplus.Save()
        End Using

        Using workbook As New OpenXmlExcelDataOperations(filePath, OpenMode.OpenExistingFile, options)
            Assert.That(workbook.LookupCellFormula(New ExcelCell("Inputs", "B8", ExcelCell.ValueTypes.All)), [Is].EqualTo("A8+1"))
            workbook.RecalculateAll()
            workbook.Save(recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using document = SpreadsheetDocument.Open(filePath, False)
            Dim sheet = document.WorkbookPart.WorksheetParts.Single()
            Using reader = XmlReader.Create(sheet.GetStream())
                reader.MoveToContent()
                Assert.That(reader.LookupNamespace(String.Empty), [Is].EqualTo(New Worksheet().NamespaceUri))
            End Using
            Dim result = sheet.Worksheet.Descendants(Of Cell)().Single(Function(cell) cell.CellReference.Value = "B8")
            Assert.That(result.CellFormula.Text, [Is].EqualTo("A8+1"))
            Assert.That(result.CellValue.Text, [Is].EqualTo("21"))
        End Using
    End Sub

End Class
