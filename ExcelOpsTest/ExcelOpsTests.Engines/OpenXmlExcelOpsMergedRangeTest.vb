Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports CompuMaster.Excel.Test.ExcelOpsTests.Engines
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Validation
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    ' Excel moves, expands or preserves the merge and identically shaped absolute defined names.
    <TestCase("Row", 1, "B4:D6")>
    <TestCase("Row", 2, "B4:D6")>
    <TestCase("Row", 3, "B3:D6")>
    <TestCase("Row", 5, "B3:D5")>
    <TestCase("Column", 0, "C3:E5")>
    <TestCase("Column", 1, "C3:E5")>
    <TestCase("Column", 2, "B3:E5")>
    <TestCase("Column", 4, "B3:D5")>
    <TestCase("Down", 1, "B4:D6")>
    <TestCase("Down", 3, "B3:D6")>
    <TestCase("Down", 5, "B3:D5")>
    <TestCase("Right", 0, "C3:E5")>
    <TestCase("Right", 2, "B3:E5")>
    <TestCase("Right", 4, "B3:D5")>
    Public Sub InsertsBeforeInsideAndAfterMergedAndDefinedAreas(operation As String, index As Integer, expectedRange As String)
        Dim source = Path.Combine(_directory, "merge-source.xlsx")
        Dim output = Path.Combine(_directory, "merge-output.xlsx")
        CreateMergedRangeWorkbook(source)
        Dim expectedAnchor = New ExcelCell(DataSheet, expectedRange.Split(":"c)(0), ExcelCell.ValueTypes.All)

        Using workbook = OpenStructuralWorkbook(source)
            Select Case operation
                Case "Row"
                    workbook.AddRow(DataSheet, index, 1, True)
                Case "Column"
                    workbook.AddColumn(DataSheet, index, 1, True)
                Case "Down"
                    workbook.AddCells(DataSheet, index, 1, 1, 3, CellInsertShiftDirection.ShiftCellsDown, True)
                Case "Right"
                    workbook.AddCells(DataSheet, 2, index, 3, 1, CellInsertShiftDirection.ShiftCellsRight, True)
            End Select
            Assert.That(workbook.LookupCellValue(Of Integer)(expectedAnchor, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(17))
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM(" & DataQualifier & expectedRange & ")"))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using reopened = OpenStructuralWorkbook(output)
            Assert.That(reopened.LookupCellValue(Of Integer)(expectedAnchor, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(17))
            Assert.That(reopened.LookupCellFormula("Check", 1, 0), [Is].EqualTo("SUM(WorkbookArea)"))
        End Using
        AssertMergedRangeAndNames(output, expectedRange, expectedRange)
    End Sub

    ' A split shift removes the merge. The same rectangular formula and defined-name references remain fixed.
    <TestCase(1, "B4")>
    <TestCase(3, "B3")>
    Public Sub PartialShiftThroughMergedAreaUnmergesIt(row As Integer, expectedAnchor As String)
        Dim source = Path.Combine(_directory, "split-merge.xlsx")
        Dim output = Path.Combine(_directory, "split-merge-output.xlsx")
        CreateMergedRangeWorkbook(source)
        Using workbook = OpenStructuralWorkbook(source)
            workbook.AddCells(DataSheet, row, 1, 1, 2, CellInsertShiftDirection.ShiftCellsDown, True)
            Assert.That(workbook.LookupCellValue(Of Integer)(New ExcelCell(DataSheet, expectedAnchor, ExcelCell.ValueTypes.All),
                                                             FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(17))
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM(" & DataQualifier & "B3:D5)"))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        AssertMergedRangeAndNames(output, Nothing, "B3:D5")
    End Sub

    <TestCase(2, 1, "B3:D4", False)>
    <TestCase(3, 1, "B3:D4", True)>
    <TestCase(2, 3, Nothing, False)>
    Public Sub DeletesRowsWithinMergedAndDefinedAreas(row As Integer, count As Integer, expectedRange As String, anchorSurvives As Boolean)
        Dim source = Path.Combine(_directory, "delete-merge.xlsx")
        Dim output = Path.Combine(_directory, "delete-merge-output.xlsx")
        CreateMergedRangeWorkbook(source)
        Using workbook = OpenStructuralWorkbook(source)
            workbook.RemoveRows(DataSheet, row, count, True)
            Assert.That(workbook.IsEmptyCell(DataSheet, 2, 1), [Is].EqualTo(Not anchorSurvives))
            Dim expectedReference = If(expectedRange, "#REF!")
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM(" & DataQualifier & expectedReference & ")"))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        AssertMergedRangeAndNames(output, expectedRange, If(expectedRange, "#REF!"))
    End Sub

    <TestCaseSource(GetType(StructuralReferenceCoverage), NameOf(StructuralReferenceCoverage.Cases))>
    Public Sub SharedStructuralReferenceCoverageMatchesExcel(scenario As String)
        Dim source = TestEnvironment.FullPathOfExistingTestFile("test_data", StructuralReferenceCoverage.FixtureName)
        Dim output = Path.Combine(_directory, "structural-reference-coverage.xlsx")
        Using workbook = OpenStructuralWorkbook(source)
            Dim differences = StructuralReferenceCoverage.CompareWithExcel(workbook, scenario, output)
            Assert.That(differences, [Is].Empty, String.Join(Environment.NewLine, differences))
            If scenario = "TableSplit" Then
                workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(HashFile(output), [Is].EqualTo(HashFile(source)))
            End If
        End Using
    End Sub

    <Test>
    Public Sub StructuralChanges_OptInDoesNotBypassOpenXmlTableValidation()
        Dim source = TestEnvironment.FullPathOfExistingTestFile("test_data", StructuralReferenceCoverage.FixtureName)
        Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly) With {
            .AllowKnownStructuralReferenceDifferences = True
        }
        Using workbook As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, options)
            Assert.Throws(Of NotSupportedException)(
                Sub() workbook.AddCells("Data", 4, 2, 1, 2, CellInsertShiftDirection.ShiftCellsDown, True))
        End Using
    End Sub

    ' A second edit in the same session uses the table bounds and names already changed by the first.
    <Test>
    Public Sub RepeatedTableInsertionsUseTheStagedTableBounds()
        Dim source = TestEnvironment.FullPathOfExistingTestFile("test_data", StructuralReferenceCoverage.FixtureName)
        Dim output = Path.Combine(_directory, "repeated-table-insertions.xlsx")
        Using workbook = OpenStructuralWorkbook(source)
            workbook.AddRow("Data", 4, 1, True)
            workbook.AddCells("Data", 7, 1, 2, 3, CellInsertShiftDirection.ShiftCellsDown, True)
            Assert.That(workbook.LookupCellFormula("Checks", 1, 0), [Is].EqualTo("SUM(Data!C4:C9)"))
            Assert.That(workbook.LookupCellFormula("Data", 8, 3), [Is].EqualTo("Sales[[#This Row],[Qty]]*2"))
            Assert.That(workbook.LookupCellFormula("Data", 9, 3), [Is].EqualTo("SUM(Sales[Total])"))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim part = document.WorkbookPart
            Dim data = part.Workbook.Sheets.Elements(Of Sheet)().
                Single(Function(sheet) sheet.Name.Value = "Data")
            Dim table = DirectCast(part.GetPartById(data.Id.Value), WorksheetPart).TableDefinitionParts.Single().Table
            Assert.That(table.Reference.Value, [Is].EqualTo("B3:D10"))
            Assert.That(table.GetFirstChild(Of AutoFilter)().Reference.Value, [Is].EqualTo("B3:D9"))
            Assert.That(part.Workbook.DefinedNames.Elements(Of DefinedName)().Single(Function(item) item.Name.Value = "SalesCells").Text,
                        [Is].EqualTo("Data!$B$4:$D$9"))
        End Using
    End Sub

    ' Shared formulas are translated to their original cell before the structural edit is applied.
    <Test>
    Public Sub ExpandsSharedFormulasWithRelativeRowsAndColumns()
        Dim source = Path.Combine(_directory, "shared-source.xlsx")
        Dim output = Path.Combine(_directory, "shared-output.xlsx")
        Dim master = FormulaCell("D4", "A1+$B1+C$1+$D$1")
        master.CellFormula.FormulaType = CellFormulaValues.Shared
        master.CellFormula.Reference = "D4:E5"
        master.CellFormula.SharedIndex = 0UI
        Dim followers = {"E4", "D5", "E5"}.Select(
            Function(address) New Cell With {
                .CellReference = address,
                .CellFormula = New CellFormula With {.FormulaType = CellFormulaValues.Shared, .SharedIndex = 0UI},
                .CellValue = New CellValue(42)
            })
        CreateStructuralWorkbook(source, {master}.Concat(followers).ToArray(), {})
        Using workbook = OpenStructuralWorkbook(source)
            workbook.AddRow(DataSheet, 0, 1, True)
            Assert.That(workbook.LookupCellFormula(DataSheet, 4, 3), [Is].EqualTo("A2+$B2+C$2+$D$2"))
            Assert.That(workbook.LookupCellFormula(DataSheet, 5, 4), [Is].EqualTo("B3+$B3+D$2+$D$2"))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim sheet = DirectCast(document.WorkbookPart.GetPartById("rIdData"), WorksheetPart)
            Assert.That(sheet.Worksheet.Descendants(Of CellFormula)().Select(Function(item) item.SharedIndex), [Is].All.Null)
        End Using
    End Sub

    Private Shared Function OpenStructuralWorkbook(path As String) As OpenXmlExcelDataOperations
        Return New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
    End Function

    Private Shared Sub CreateMergedRangeWorkbook(path As String)
        CreateStructuralWorkbook(path, {ValueCell("B3", 17)},
            {FormulaCell("A1", "SUM(" & DataQualifier & "B3:D5)"), FormulaCell("A2", "SUM(WorkbookArea)")},
            New DefinedName(DataQualifier & "$B$3:$D$5") With {.Name = "WorkbookArea"},
            New DefinedName("$B$3:$D$5") With {.Name = "LocalArea", .LocalSheetId = 0UI},
            New DefinedName(DataQualifier & "$B$3:$D$5") With {.Name = "OtherLocalArea", .LocalSheetId = 1UI},
            New DefinedName("Check!$B$3:$D$5") With {.Name = "UnrelatedArea"})
        Using document = SpreadsheetDocument.Open(path, True)
            Dim worksheet = DirectCast(document.WorkbookPart.GetPartById("rIdData"), WorksheetPart).Worksheet
            worksheet.SheetDimension.Reference = "B3:D5"
            worksheet.AddChild(New MergeCells(New MergeCell With {.Reference = "B3:D5"}) With {.Count = 1UI})
        End Using
    End Sub

    Private Shared Sub AssertMergedRangeAndNames(path As String, expectedMerge As String, expectedNameRange As String)
        Dim absoluteRange = expectedNameRange
        If absoluteRange <> "#REF!" Then
            absoluteRange = String.Join(":", absoluteRange.Split(":"c).Select(
                Function(address) New ExcelCell(DataSheet, address, ExcelCell.ValueTypes.All).Address(False, True, True)))
        End If
        Using document = SpreadsheetDocument.Open(path, False)
            Dim part = document.WorkbookPart
            Dim worksheet = DirectCast(part.GetPartById("rIdData"), WorksheetPart).Worksheet
            Dim merges = worksheet.GetFirstChild(Of MergeCells)()
            If expectedMerge Is Nothing Then
                Assert.That(merges, [Is].Null)
            Else
                Assert.That(merges.Count.Value, [Is].EqualTo(1UI))
                Assert.That(merges.Elements(Of MergeCell)().Single().Reference.Value, [Is].EqualTo(expectedMerge))
                Assert.That(worksheet.SheetDimension.Reference.Value, [Is].EqualTo(expectedMerge))
            End If
            Dim names = part.Workbook.DefinedNames.Elements(Of DefinedName)().ToDictionary(Function(item) item.Name.Value)
            Assert.That(names("WorkbookArea").Text, [Is].EqualTo(DataQualifier & absoluteRange))
            Assert.That(names("WorkbookArea").LocalSheetId, [Is].Null)
            Assert.That(names("LocalArea").Text, [Is].EqualTo(absoluteRange))
            Assert.That(names("LocalArea").LocalSheetId.Value, [Is].EqualTo(0UI))
            Assert.That(names("OtherLocalArea").Text, [Is].EqualTo(DataQualifier & absoluteRange))
            Assert.That(names("OtherLocalArea").LocalSheetId.Value, [Is].EqualTo(1UI))
            Assert.That(names("UnrelatedArea").Text, [Is].EqualTo("Check!$B$3:$D$5"))
            Assert.That(New OpenXmlValidator().Validate(document).Select(Function(item) item.Description), [Is].Empty)
        End Using
    End Sub

End Class
