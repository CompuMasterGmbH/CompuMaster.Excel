Imports System.Globalization
Imports System.IO
Imports System.Xml
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

<TestFixture>
Public Class FormulaConformanceFixtureTest

    Private Const ReferenceWorkbook As String = "FormulaReference.xlsx"
    Private Const CompanionWorkbook As String = "FormulaReferenceSource.xlsx"

    ' The checked-in caches were recorded by Microsoft Excel build 16.0.20625.20108.
    <TestCase(ReferenceWorkbook, "Cases", 26)>
    <TestCase(CompanionWorkbook, "Source Data", 1)>
    Public Sub RecalculatesTheExcelReferenceFixture(fileName As String, sheetName As String, caseCount As Integer)
        Dim fixture = FixturePath(fileName)
        Dim expected = ReadFormulaCaches(fixture, sheetName)
        Assert.That(expected.Count, [Is].EqualTo(caseCount))

        Dim output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Try
            Using workbook = OpenForCalculation(fixture)
                workbook.ResetCellValueFromFormulaCell(sheetName)
                For Each address In expected.Keys
                    Assert.That(workbook.HasMissingCachedFormulaResult(New ExcelCell(sheetName, address, ExcelCell.ValueTypes.All)), [Is].True, address)
                Next
                workbook.RecalculateSheet(sheetName)
                workbook.SaveCopyAs(output, recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using

            Assert.That(ReadFormulaCaches(output, sheetName), [Is].EqualTo(expected))
            If fileName = ReferenceWorkbook Then
                Assert.That(ReadFormulaCaches(output, "External cases"), [Is].EqualTo(ReadFormulaCaches(fixture, "External cases")))
            End If
        Finally
            If File.Exists(output) Then File.Delete(output)
        End Try
    End Sub

    ' External results retain their Excel caches, and the companion link is relative to the fixture.
    <Test>
    Public Sub KeepsTheExternalReferenceCasesPortable()
        Dim fixture = FixturePath(ReferenceWorkbook)
        Using package = SpreadsheetDocument.Open(fixture, False)
            Dim links = package.WorkbookPart.ExternalWorkbookParts.Single().ExternalRelationships.ToArray()
            Assert.That(links, [Is].Not.Empty)
            For Each link In links
                Assert.That(link.Uri.OriginalString, [Is].EqualTo(CompanionWorkbook))
                Assert.That(link.Uri.IsAbsoluteUri, [Is].False)
            Next
            Assert.That(File.Exists(Path.Combine(Path.GetDirectoryName(fixture), CompanionWorkbook)), [Is].True)
        End Using

        Using workbook = OpenForCalculation(fixture)
            Const sheetName As String = "External cases"
            Dim expected = ReadFormulaCaches(fixture, sheetName)
            Assert.That(expected.Count, [Is].EqualTo(3))
            For Each item In expected
                Dim address As New ExcelCell(sheetName, item.Key, ExcelCell.ValueTypes.All)
                Assert.That(workbook.LookupCellValueAsObject(address, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(item.Value.Value), item.Key)
            Next

            Dim firstLink As New ExcelCell(sheetName, "C2", ExcelCell.ValueTypes.All)
            workbook.ResetCellValueFromFormulaCell(firstLink)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(
                Sub() workbook.LookupCellValueAsObject(firstLink, FormulaCalculationMode.CalculateIfMissing))
        End Using
    End Sub

    <TestCase(True)>
    <TestCase(False)>
    Public Sub RecalculatesExcelExternalReferencesFromOpenSource(usePaths As Boolean)
        Dim session As New ExcelFormulaSession()
        Dim fixture = FixturePath(ReferenceWorkbook)
        Dim sourcePath = FixturePath(CompanionWorkbook)
        Using workbook = OpenInSession(fixture, session, usePaths)
            Using source = OpenInSession(sourcePath, session, usePaths)
                workbook.ResetCellValueFromFormulaCell("External cases")
                workbook.RecalculateSheet("External cases")
                Assert.That(ExternalValues(workbook), [Is].EqualTo({42.0, 100.0, 46.0}))
                source.WriteCellValue("Source Data", 1, 0, 45.0)
                workbook.RecalculateSheet("External cases")
                Assert.That(ExternalValues(workbook), [Is].EqualTo({45.0, 106.0, 49.0}))
            End Using
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.RecalculateSheet("External cases"))
            Assert.That(ExternalValues(workbook), [Is].EqualTo({45.0, 106.0, 49.0}))
        End Using
    End Sub

    Private Shared Function ExternalValues(workbook As OpenXmlExcelDataOperations) As Double()
        Return Enumerable.Range(1, 3).Select(
            Function(row) workbook.LookupCellValue(Of Double)("External cases", row, 2, FormulaCalculationMode.UseCachedValueOnly)).ToArray()
    End Function

    Private Shared Function OpenInSession(path As String, session As ExcelFormulaSession, usePath As Boolean) As OpenXmlExcelDataOperations
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .FormulaSession = session
        }
        If usePath Then Return New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, options)
        options.FormulaWorkbookName = IO.Path.GetFileName(path)
        Return New OpenXmlExcelDataOperations(File.ReadAllBytes(path), options)
    End Function

    Private Shared Function OpenForCalculation(path As String) As OpenXmlExcelDataOperations
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False}
        Return New OpenXmlExcelDataOperations(File.ReadAllBytes(path), options)
    End Function

    Private Shared Function FixturePath(fileName As String) As String
        Return TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", fileName)
    End Function

    Private Shared Function ReadFormulaCaches(fileName As String, sheetName As String) As Dictionary(Of String, (Kind As CellValues, Value As Object))
        Dim caches As New Dictionary(Of String, (Kind As CellValues, Value As Object))(StringComparer.Ordinal)
        Using package = SpreadsheetDocument.Open(fileName, False)
            Dim workbook = package.WorkbookPart
            Dim sheet = workbook.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = sheetName)
            Dim worksheet = DirectCast(workbook.GetPartById(sheet.Id.Value), WorksheetPart).Worksheet
            For Each cell In worksheet.Descendants(Of Cell)().Where(Function(item) item.CellFormula IsNot Nothing)
                Dim address = cell.CellReference.Value
                Assert.That(cell.CellValue, [Is].Not.Null, sheetName & "!" & address)
                Dim text = If(cell.CellValue.Text, String.Empty)
                Dim kind = If(cell.DataType?.Value, CellValues.Number)
                Dim value As Object = text
                If kind = CellValues.Number Then
                    value = Double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture)
                ElseIf kind = CellValues.Boolean Then
                    value = XmlConvert.ToBoolean(text)
                End If
                caches.Add(address, (kind, value))
            Next
        End Using
        Return caches
    End Function

End Class
