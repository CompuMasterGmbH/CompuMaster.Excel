Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlArrayFormulaTest

    <TestCase("SUM(SEQUENCE(3))", 6.0)>
    <TestCase("SUM(SEQUENCE(,3))", 6.0)>
    <TestCase("SUM(SEQUENCE(0))", "#CALC!")>
    <TestCase("SUM(SEQUENCE(1,0))", "#CALC!")>
    <TestCase("SUM(SEQUENCE(-1))", "#VALUE!")>
    <TestCase("SUM(SEQUENCE(2,1,,0))", 2.0)>
    <TestCase("SUM(_xlfn.SEQUENCE(2,3,10,-2))", 30.0)>
    <TestCase("INDEX(SEQUENCE(2,3,10,-2),2,1)", 4.0)>
    <TestCase("SUM(INDEX(SEQUENCE(2,3),0,2))", 7.0)>
    <TestCase("ROWS(TRANSPOSE(SEQUENCE(2,3)))", 3.0)>
    <TestCase("COLUMNS(TRANSPOSE(SEQUENCE(2,3)))", 2.0)>
    <TestCase("SUM(-SEQUENCE(3))", -6.0)>
    <TestCase("SUM(SEQUENCE(2)+SEQUENCE(1,3))", 21.0)>
    <TestCase("SUM(SEQUENCE(2)+SEQUENCE(3))", "#N/A")>
    <TestCase("SUM(FILTER(A1:B3,A1:A3>1))", 55.0)>
    <TestCase("SUM(_xlfn._xlws.FILTER(A1:B3,A1:A3>1))", 55.0)>
    <TestCase("SUM(FILTER(A1:B3,A1:B1>1))", 60.0)>
    <TestCase("COLUMNS(FILTER(A1:B3,A1:A3>1))", 2.0)>
    <TestCase("SUM(FILTER(A1:B3,A1:A3>99,8))", 8.0)>
    <TestCase("SUM(FILTER(A1:B3,A1:A3>99))", "#CALC!")>
    <TestCase("SUM(FILTER(A1:B3,A1:B2>1))", "#VALUE!")>
    <TestCase("SUM(FILTER(B1,""1""))", "#VALUE!")>
    <TestCase("SUM(FILTER(B1,A4))", "#CALC!")>
    <TestCase("SUM(FILTER(A1:B3,1/(A1:A3-2)))", "#DIV/0!")>
    <TestCase("SUM(TRANSPOSE(A1:A4)*SEQUENCE(1,4))", 14.0)>
    Public Sub EvaluatesArraysInsideScalarFormulas(formula As String, expected As Object)
        Using workbook = CreateWorkbook()
            workbook.WriteCellFormula("Sheet1", 0, 4, formula, False)
            Assert.That(workbook.LookupCellValueAsObject("Sheet1", 0, 4, FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(expected), formula)
        End Using
    End Sub

    <Test>
    Public Sub LeavesPreviousValuesUntouchedWhenAnArrayCannotBeWritten()
        Using workbook = CreateWorkbook()
            workbook.WriteCellFormula("Sheet1", 0, 4, "SUM(SEQUENCE(3))", False)
            workbook.WriteCellFormula("Sheet1", 1, 4, "SEQUENCE(3)", False)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.RecalculateSheet("Sheet1"))
            Assert.That(workbook.HasMissingCachedFormulaResult("Sheet1", 0, 4), [Is].True)
            Assert.That(workbook.HasMissingCachedFormulaResult("Sheet1", 1, 4), [Is].True)
        End Using
    End Sub

    <Test>
    Public Sub SavesAndReopensCalcErrorsInANewWorkbook()
        Dim output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Try
            Using workbook = CreateWorkbook()
                workbook.WriteCellFormula("Sheet1", 0, 4, "SUM(_xlfn.SEQUENCE(0))", False)
                workbook.RecalculateCell("Sheet1", 0, 4, throwExceptionOnCalculationError:=False)
                workbook.SaveCopyAs(output, recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using
            Using package = SpreadsheetDocument.Open(output, False)
                Dim part = package.WorkbookPart
                Dim cell = part.WorksheetParts.Single().Worksheet.Descendants(Of Cell)().Single(Function(item) item.CellReference.Value = "E1")
                Assert.That(cell.DataType.Value, [Is].EqualTo(CellValues.Error))
                Assert.That(cell.CellValue.Text, [Is].EqualTo("#VALUE!"))
                Assert.That(cell.ValueMetaIndex?.Value, [Is].EqualTo(1UI))
                Assert.That(part.CellMetadataPart, [Is].Not.Null)
                Assert.That(part.RdRichValueParts.Count(), [Is].EqualTo(1))
            End Using
            Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
            Using reopened As New OpenXmlExcelDataOperations(File.ReadAllBytes(output), options)
                Assert.That(reopened.LookupCellValueAsObject("Sheet1", 0, 4, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("#CALC!"))
                reopened.ResetCellValueFromFormulaCell("Sheet1", 0, 4)
                Assert.That(reopened.HasMissingCachedFormulaResult("Sheet1", 0, 4), [Is].True)
            End Using
        Finally
            If File.Exists(output) Then File.Delete(output)
        End Try
    End Sub

    Private Shared Function CreateWorkbook() As OpenXmlExcelDataOperations
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False}
        Dim workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, options)
        For row = 0 To 2
            workbook.WriteCellValue("Sheet1", row, 0, row + 1)
            workbook.WriteCellValue("Sheet1", row, 1, (row + 1) * 10)
        Next
        Return workbook
    End Function

End Class
