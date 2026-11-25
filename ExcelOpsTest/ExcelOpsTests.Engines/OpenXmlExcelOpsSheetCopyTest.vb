Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Validation
Imports ChartXml = DocumentFormat.OpenXml.Drawing.Charts
Imports CompuMaster.Excel.Test
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    <TestCase(CopyWorksheetFormulaMode.ExcelCompatible)>
    <TestCase(CopyWorksheetFormulaMode.PreserveFormulas)>
    Public Sub WorksheetTransfer_OpenXmlCopiesWithTheRequestedFormulaMode(mode As CopyWorksheetFormulaMode)
        Dim sourcePath = Path.Combine(_directory, "source.xlsx")
        Dim destinationPath = Path.Combine(_directory, "target.xlsx")
        Dim output = Path.Combine(_directory, "copied.xlsx")
        CreateStructuralWorkbook(sourcePath, {ValueCell("B2", 20)},
            {FormulaCell("A1", DataQualifier & "$B$2"), FormulaCell("A2", "Check!A1"), FormulaCell("A3", "A1+N(""Check!A1"")")})
        CreateStructuralWorkbook(destinationPath, {ValueCell("B2", 70)}, {ValueCell("A1", 100)})
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Dim expectedCrossSheet = If(mode = CopyWorksheetFormulaMode.ExcelCompatible, "'[1]O''Brien Data'!$B$2", DataQualifier & "$B$2")
        Dim expectedSelf = If(mode = CopyWorksheetFormulaMode.ExcelCompatible, "'Check (2)'!A1", "Check!A1")
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options)
            Using target As New OpenXmlExcelDataOperations(destinationPath, OpenMode.OpenExistingFile, options)
                target.WriteCellValue(Of Integer)(DataSheet, 0, 0, 17)
                Dim assigned = source.CopySheet("Check", target, mode)
                Assert.That(assigned.TargetSheetName, [Is].EqualTo("Check (2)"))
                Assert.That(assigned.TargetSheetIndex, [Is].EqualTo(3))
                Assert.That(target.LookupCellFormula(assigned.TargetSheetName, 0, 0), [Is].EqualTo(expectedCrossSheet))
                Assert.That(target.LookupCellFormula(assigned.TargetSheetName, 1, 0), [Is].EqualTo(expectedSelf))
                Assert.That(target.LookupCellFormula(assigned.TargetSheetName, 2, 0), [Is].EqualTo("A1+N(""Check!A1"")"))
                Assert.That(target.LookupCellValue(Of Integer)(DataSheet, 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(17))
                Assert.That(source.LookupCellFormula("Check", 0, 0), [Is].EqualTo(DataQualifier & "$B$2"))
                Dim second = source.CopySheet("Check", target, mode)
                Assert.That(second.TargetSheetName, [Is].EqualTo("Check (3)"))
                Assert.That(target.LookupCellFormula(second.TargetSheetName, 0, 0), [Is].EqualTo(expectedCrossSheet))
                target.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using
        End Using
        Using saved = SpreadsheetDocument.Open(output, False)
            Dim links = saved.WorkbookPart.ExternalWorkbookParts.ToArray()
            If mode = CopyWorksheetFormulaMode.ExcelCompatible Then
                Assert.That(links.Single().ExternalRelationships.Single().Uri.LocalPath, [Is].EqualTo(sourcePath))
            Else
                Assert.That(links, [Is].Empty)
            End If
            Assert.That(New OpenXmlValidator().Validate(saved).Select(Function(issue) issue.Description), [Is].Empty)
        End Using
        Using reopened As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            Assert.That(reopened.LookupCellFormula("Check (2)", 0, 0), [Is].EqualTo(expectedCrossSheet))
            Assert.That(reopened.LookupCellFormula("Check (2)", 1, 0), [Is].EqualTo(expectedSelf))
        End Using
    End Sub

    <Test>
    Public Sub WorksheetTransfer_OpenXmlCopiesGroupsWithLocalReferences()
        Dim sourcePath = Path.Combine(_directory, "source.xlsx")
        Dim destinationPath = Path.Combine(_directory, "target.xlsx")
        CreateStructuralWorkbook(sourcePath, {ValueCell("B2", 20)}, {FormulaCell("A1", DataQualifier & "$B$2")})
        CreateStructuralWorkbook(destinationPath, {ValueCell("B2", 70)}, {ValueCell("A1", 100)})
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options)
            Using target As New OpenXmlExcelDataOperations(destinationPath, OpenMode.OpenExistingFile, options)
                Assert.That(source.CopySheets({"Check", DataSheet}, target).Select(Function(item) item.TargetSheetName), [Is].EqualTo({DataSheet & " (2)", "Check (2)"}))
                Assert.That(target.LookupCellFormula("Check (2)", 0, 0), [Is].EqualTo("'O''Brien Data (2)'!$B$2"))
                Assert.That(target.LookupCellValue(Of Integer)(DataSheet & " (2)", 1, 1, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(20))
            End Using
        End Using
    End Sub

    <Test>
    Public Sub WorksheetTransfer_OpenXmlRequiresASavedSourceOnlyWhenCreatingAnExternalLink()
        Dim input = Path.Combine(_directory, "source.xlsx")
        CreateStructuralWorkbook(input, {ValueCell("B2", 20)}, {FormulaCell("A1", DataQualifier & "$B$2")})
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(File.ReadAllBytes(input), options),
              target As New OpenXmlExcelDataOperations(Nothing, OpenMode.CreateFile, options)
            Assert.Throws(Of NotSupportedException)(Sub() source.CopySheet("Check", target))
            Assert.That(target.SheetNames(), [Is].EqualTo({"Sheet1"}))
            Dim copied = source.CopySheet("Check", target, CopyWorksheetFormulaMode.PreserveFormulas)
            Assert.That(target.LookupCellFormula(copied.TargetSheetName, 0, 0), [Is].EqualTo(DataQualifier & "$B$2"))
        End Using
    End Sub

    <TestCase(CopyWorksheetFormulaMode.ExcelCompatible, 0)>
    <TestCase(CopyWorksheetFormulaMode.ExcelCompatible, 3)>
    <TestCase(CopyWorksheetFormulaMode.PreserveFormulas, 1)>
    Public Sub WorksheetTransfer_OpenXmlInsertsMixedGroupsAtTheRequestedIndex(mode As CopyWorksheetFormulaMode, index As Integer)
        Dim input = TestFiles.TestFileChartSheet01.FullName
        Dim targetPath = Path.Combine(_directory, "target-with-charts.xlsx")
        Dim output = Path.Combine(_directory, "mixed-copy.xlsx")
        File.Copy(input, targetPath)
        Using document = SpreadsheetDocument.Open(targetPath, True)
            Dim book = document.WorkbookPart
            Dim tail = book.AddNewPart(Of WorksheetPart)()
            tail.Worksheet = New Worksheet(New SheetData())
            book.Workbook.Sheets.Append(New Sheet With {.SheetId = 4UI, .Name = "Tail", .Id = book.GetIdOfPart(tail)})
            book.Workbook.DefinedNames = New DefinedNames(New DefinedName("Tail!$A$1") With {.Name = "LocalValue", .LocalSheetId = 2UI})
            Dim view = book.Workbook.BookViews.GetFirstChild(Of WorkbookView)()
            view.ActiveTab = 2UI
            view.FirstSheet = 1UI
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Dim expectedOrder As New List(Of String) From {"data", "chart", "Tail"}
        expectedOrder.InsertRange(index, {"data (2)", "chart (2)"})
        Using source As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options),
              target As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, options)
            Assert.Throws(Of ArgumentOutOfRangeException)(Sub() source.CopySheets({"data", "chart"}, target, -1, mode))
            Assert.Throws(Of ArgumentOutOfRangeException)(Sub() source.CopySheets({"data", "chart"}, target, 4, mode))
            Assert.That(target.SheetNames(), [Is].EqualTo({"data", "chart", "Tail"}))

            Dim copied = source.CopySheets({"chart", "data"}, target, index, mode)
            Assert.That(copied.Select(Function(item) item.TargetSheetName), [Is].EqualTo({"data (2)", "chart (2)"}))
            Assert.That(copied.Select(Function(item) item.TargetSheetIndex), [Is].EqualTo({index, index + 1}))
            Assert.That(target.SheetNames(), [Is].EqualTo(expectedOrder))
            Assert.That(target.ChartSheetNames(), [Is].EquivalentTo({"chart", "chart (2)"}))
            Assert.That(target.SheetNames().Count, [Is].EqualTo(5))
            Assert.That(target.WorkSheetNames().Count, [Is].EqualTo(3))
            Assert.That(target.SheetIndex("chart (2)"), [Is].EqualTo(copied(1).TargetSheetIndex))
            Assert.That(target.ChartSheetIndex("chart (2)"), [Is].EqualTo(If(index <= 1, 0, 1)))
            Assert.That(source.SheetNames(), [Is].EqualTo({"data", "chart"}))
            target.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using saved = SpreadsheetDocument.Open(output, False)
            Dim book = saved.WorkbookPart
            Dim expectedScope = If(index <= 2, 4UI, 2UI)
            Assert.That(book.Workbook.DefinedNames.Elements(Of DefinedName)().Single().LocalSheetId.Value, [Is].EqualTo(expectedScope))
            Dim view = book.Workbook.BookViews.GetFirstChild(Of WorkbookView)()
            Assert.That(view.ActiveTab.Value, [Is].EqualTo(expectedScope))
            Assert.That(view.FirstSheet.Value, [Is].EqualTo(If(index <= 1, 3UI, 1UI)))
            Assert.That(book.Workbook.Sheets.Elements(Of Sheet)().Select(Function(sheet) sheet.Name.Value), [Is].EqualTo(expectedOrder))

            Dim prefix = If(mode = CopyWorksheetFormulaMode.ExcelCompatible, "'data (2)'!", "data!")
            Assert.That(CopiedChartFormulas(book, "chart (2)"), [Is].EqualTo({prefix & "$A$6", prefix & "$B$5:$I$5", prefix & "$B$6:$I$6"}))
            Assert.That(CopiedChartFormulas(book, "chart"), [Is].EqualTo({"data!$A$6", "data!$B$5:$I$5", "data!$B$6:$I$6"}))
            Assert.That(book.ExternalWorkbookParts, [Is].Empty)
        End Using
    End Sub

    <Test>
    Public Sub WorksheetTransfer_OpenXmlChartCopiedAloneLinksBackToItsUncopiedData()
        Dim input = TestFiles.TestFileChartSheet01.FullName
        Dim output = Path.Combine(_directory, "chart-only-copy.xlsx")
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options),
              target As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            Dim copied = source.CopySheet("chart", target, 1)
            Assert.That(copied.TargetSheetName, [Is].EqualTo("chart (2)"))
            Assert.That(copied.TargetSheetIndex, [Is].EqualTo(1))
            target.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using saved = SpreadsheetDocument.Open(output, False)
            Dim book = saved.WorkbookPart
            Assert.That(CopiedChartFormulas(book, "chart (2)"), [Is].EqualTo({"[1]data!$A$6", "[1]data!$B$5:$I$5", "[1]data!$B$6:$I$6"}))
            Assert.That(book.ExternalWorkbookParts.Single().ExternalRelationships.Single().Uri.LocalPath, [Is].EqualTo(input))
        End Using
    End Sub

    Private Shared Function CopiedChartFormulas(book As WorkbookPart, name As String) As String()
        Dim sheet = book.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = name)
        Dim part = DirectCast(book.GetPartById(sheet.Id.Value), ChartsheetPart)
        Return part.DrawingsPart.ChartParts.Single().ChartSpace.Descendants(Of ChartXml.Formula)().
            Select(Function(formula) formula.Text).ToArray()
    End Function

End Class
