Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlSpillFormulaTest
    Private Const SheetName As String = "Spills"

    <TestCase(False)>
    <TestCase(True)>
    Public Sub ExpandsAndShrinksAnExcelSpillWithDependentFormulas(automatic As Boolean)
        Using workbook = OpenReference(automatic)
            Assert.That(Number(workbook, SheetName, "G4"), [Is].EqualTo(90.0))
            Assert.That(Number(workbook, SheetName, "G5"), [Is].EqualTo(21.0))
            Assert.That(Value(workbook, SheetName, "H4"), [Is].EqualTo("#SPILL!"))
            Assert.That(Value(workbook, SheetName, "G8"), [Is].EqualTo("#SPILL!"))

            workbook.WriteCellValue(SheetName, 1, 0, 4.0)
            workbook.WriteCellValue(Of String)(SheetName, 4, 7, Nothing)
            If Not automatic Then workbook.RecalculateAll()
            Assert.That(Number(workbook, SheetName, "C7"), [Is].EqualTo(24.0))
            Assert.That(Number(workbook, SheetName, "G4"), [Is].EqualTo(136.0))
            Assert.That(Number(workbook, SheetName, "G6"), [Is].EqualTo(4.0))
            Assert.That(Number(workbook, SheetName, "G7"), [Is].EqualTo(25.0))
            Assert.That(Number(workbook, SheetName, "G8"), [Is].EqualTo(3.0))
            Assert.That(Number(workbook, "Summary", "B2"), [Is].EqualTo(136.0))
            Assert.That(Number(workbook, "Summary", "B3"), [Is].EqualTo(4.0))
            Assert.Throws(Of NotSupportedException)(Sub() workbook.WriteCellValue(SheetName, 6, 2, 77.0))

            workbook.WriteCellValue(SheetName, 1, 0, 1.0)
            If Not automatic Then workbook.RecalculateAll()
            Assert.That(Value(workbook, SheetName, "C6"), [Is].Null)
            Assert.That(Number(workbook, SheetName, "G4"), [Is].EqualTo(22.0))
            Assert.That(Number(workbook, SheetName, "G5"), [Is].EqualTo(1.0))
            Assert.That(Number(workbook, "Summary", "B3"), [Is].EqualTo(1.0))
            workbook.WriteCellValue(SheetName, 6, 2, 77.0)
            Assert.That(Number(workbook, SheetName, "C7"), [Is].EqualTo(77.0))
        End Using
    End Sub

    <Test>
    Public Sub ClearsAndRecalculatesAnEntireSpillFromAChildCell()
        Using workbook = OpenReference(False)
            Dim anchor = Address(SheetName, "B4")
            Dim child = Address(SheetName, "C6")
            workbook.ResetCellValueFromFormulaCell(anchor)
            Assert.That(workbook.HasMissingCachedFormulaResult(anchor), [Is].True)
            Assert.That(workbook.HasMissingCachedFormulaResult(child), [Is].True)
            Assert.Throws(Of MissingCalculatedCellValueException)(Sub() workbook.LookupCellValueAsObject(child, FormulaCalculationMode.UseCachedValueOnly))
            Assert.That(workbook.LookupCellValueAsObject(child, FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(20.0))
            Assert.That(Number(workbook, SheetName, "B4"), [Is].EqualTo(10.0))
            Assert.That(workbook.HasMissingCachedFormulaResult(child), [Is].False)
        End Using
    End Sub

    <Test>
    Public Sub SavesABlockedSpillAsAnExcelErrorWithoutOverwritingTheBlocker()
        Dim output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Try
            Using workbook = OpenReference(False)
                workbook.WriteCellValue(SheetName, 6, 2, 123.0)
                workbook.WriteCellValue(SheetName, 1, 0, 4.0)
                workbook.RecalculateSheet(SheetName)
                Assert.That(Value(workbook, SheetName, "B4"), [Is].EqualTo("#SPILL!"))
                Assert.That(Value(workbook, SheetName, "G4"), [Is].EqualTo("#SPILL!"))
                Assert.That(Value(workbook, SheetName, "C6"), [Is].Null)
                Assert.That(Number(workbook, SheetName, "C7"), [Is].EqualTo(123.0))
                workbook.SaveCopyAs(output, recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using

            Using document = SpreadsheetDocument.Open(output, False)
                Dim part = document.WorkbookPart
                Dim sheet = part.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = SheetName)
                Dim cells = DirectCast(part.GetPartById(sheet.Id.Value), WorksheetPart).Worksheet.Descendants(Of Cell)().ToArray()
                Dim anchor = cells.Single(Function(item) item.CellReference.Value = "B4")
                Assert.That(anchor.CellFormula.Reference.Value, [Is].EqualTo("B4"))
                Assert.That(anchor.CellValue.Text, [Is].EqualTo("#VALUE!"))
                Assert.That(anchor.ValueMetaIndex, [Is].Not.Null)
            End Using

            Using workbook = OpenWorkbook(output, False)
                Assert.That(Value(workbook, SheetName, "B4"), [Is].EqualTo("#SPILL!"))
                Assert.That(Value(workbook, SheetName, "G4"), [Is].EqualTo("#SPILL!"))
                Assert.That(Number(workbook, SheetName, "C7"), [Is].EqualTo(123.0))
            End Using
        Finally
            If File.Exists(output) Then File.Delete(output)
        End Try
    End Sub

    <Test>
    Public Sub AuthorsDynamicArraysWithoutATemplateAndReusesTheirMetadata()
        Dim output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Try
            Using workbook = CreateBlankWorkbook()
                Dim sheet = workbook.SheetNames()(0)
                workbook.WriteDynamicArrayFormula(sheet, 1, 1, "=SEQUENCE(3,2,10,2)", True)
                Assert.That(Number(workbook, sheet, "C4"), [Is].EqualTo(20.0))

                workbook.WriteDynamicArrayFormula(sheet, 1, 4, "HSTACK(""FILTER(x)"", TAKE({7;8},1))", False)
                Assert.That(workbook.HasMissingCachedFormulaResult(Address(sheet, "E2")), [Is].True)
                workbook.WriteCellValue(sheet, 2, 11, 123.0)
                workbook.WriteDynamicArrayFormula(sheet, 1, 11, "SEQUENCE(2)", True)
                workbook.SaveCopyAs(output, recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using

            Using document = SpreadsheetDocument.Open(output, False)
                Dim part = document.WorkbookPart
                Dim cells = part.WorksheetParts.Single().Worksheet.Descendants(Of Cell)().ToArray()
                Dim first = cells.Single(Function(item) item.CellReference.Value = "B2")
                Dim second = cells.Single(Function(item) item.CellReference.Value = "E2")
                Assert.That(first.CellFormula.Text, [Is].EqualTo("_xlfn.SEQUENCE(3,2,10,2)"))
                Assert.That(first.CellFormula.Reference.Value, [Is].EqualTo("B2:C4"))
                Assert.That(second.CellFormula.Text, [Is].EqualTo("_xlfn.HSTACK(""FILTER(x)"", _xlfn.TAKE({7;8},1))"))
                Assert.That(first.CellMetaIndex.Value, [Is].EqualTo(second.CellMetaIndex.Value))
                Assert.That(part.CellMetadataPart.Metadata.GetFirstChild(Of CellMetadata)().Count.Value, [Is].EqualTo(1UI))
                Assert.That(part.CellMetadataPart.Metadata.Elements(Of FutureMetadata)().Select(Function(item) item.Name.Value),
                    [Is].EquivalentTo({"XLDAPR", "XLRICHVALUE"}))
            End Using

            Using workbook = OpenWorkbook(output, False)
                Dim sheet = workbook.SheetNames()(0)
                Assert.That(Number(workbook, sheet, "C4"), [Is].EqualTo(20.0))
                Assert.That(Value(workbook, sheet, "L2"), [Is].EqualTo("#SPILL!"))
                Assert.That(workbook.LookupCellValueAsObject(Address(sheet, "E2"), FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo("FILTER(x)"))
                Assert.That(Number(workbook, sheet, "F2"), [Is].EqualTo(7.0))
                workbook.WriteCellFormula(sheet, 5, 0, "SUM(_xlfn.ANCHORARRAY(B2))", True)
                Assert.That(Number(workbook, sheet, "A6"), [Is].EqualTo(90.0))
                workbook.WriteDynamicArrayFormula(sheet, 1, 7, "TRANSPOSE('" & sheet & "'!$E$2#)", True)
                Assert.That(Value(workbook, sheet, "H2"), [Is].EqualTo("FILTER(x)"))
                Assert.That(Number(workbook, sheet, "H3"), [Is].EqualTo(7.0))

                workbook.WriteDynamicArrayFormula(sheet, 1, 1, "SEQUENCE(1)", True)
                Assert.That(Number(workbook, sheet, "B2"), [Is].EqualTo(1.0))
                Assert.That(Value(workbook, sheet, "C4"), [Is].Null)

                workbook.WriteDynamicArrayFormula(sheet, 1, 4, "SEQUENCE(2)", False)
                Assert.That(workbook.HasMissingCachedFormulaResult(Address(sheet, "E2")), [Is].True)
                Assert.That(Value(workbook, sheet, "F2"), [Is].Null)
                workbook.RecalculateCell(sheet, 1, 4)
                Assert.That(Number(workbook, sheet, "E3"), [Is].EqualTo(2.0))
            End Using
        Finally
            If File.Exists(output) Then File.Delete(output)
        End Try
    End Sub

    <Test>
    Public Sub AuthorsABlockedAnchorAndKeepsExistingCellsOnUnsupportedCalculation()
        Using workbook = CreateBlankWorkbook()
            Dim sheet = workbook.SheetNames()(0)
            workbook.WriteCellValue(sheet, 0, 0, 17.0)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(
                Sub() workbook.WriteDynamicArrayFormula(sheet, 0, 0, "SERIESSUM(A1,0,1,B1:B2)", True))
            Assert.That(Number(workbook, sheet, "A1"), [Is].EqualTo(17.0))
            workbook.WriteCellValue(sheet, 1, 2, 123.0)
            workbook.WriteDynamicArrayFormula(sheet, 0, 2, "SEQUENCE(2)", True)
            Assert.That(Value(workbook, sheet, "C1"), [Is].EqualTo("#SPILL!"))
            Assert.That(Number(workbook, sheet, "C2"), [Is].EqualTo(123.0))
            workbook.WriteCellValue(Of String)(sheet, 1, 2, Nothing)
            workbook.RecalculateCell(sheet, 0, 2)
            Assert.That(Number(workbook, sheet, "C2"), [Is].EqualTo(2.0))
        End Using
    End Sub

    <Test>
    Public Sub DetectsAReferenceInEitherArrayIfBranchAndKeepsScalarIfLazy()
        Using workbook = CreateBlankWorkbook()
            Dim sheet = workbook.SheetNames()(0)
            workbook.WriteCellFormula(sheet, 0, 0, "IF(TRUE,1,A1)", True)
            Assert.That(Number(workbook, sheet, "A1"), [Is].EqualTo(1.0))
            Assert.Throws(Of CircularReferenceException)(
                Sub() workbook.WriteDynamicArrayFormula(sheet, 0, 0, "IF({TRUE,TRUE},2,A1)", True))
            Assert.That(Number(workbook, sheet, "A1"), [Is].EqualTo(1.0))
        End Using
    End Sub

    <Test>
    Public Sub AuthorsExcelCompatibleIdentityMatrixFormula()
        Using workbook = CreateBlankWorkbook()
            Dim sheet = workbook.SheetNames()(0)
            workbook.WriteDynamicArrayFormula(sheet, 0, 0, "MUNIT(2)", True)
            Assert.That(workbook.LookupCellFormula(sheet, 0, 0), [Is].EqualTo("_xlfn.MUNIT(2)"))
            Assert.That(Number(workbook, sheet, "A2"), [Is].EqualTo(0.0))
            Assert.That(Number(workbook, sheet, "B2"), [Is].EqualTo(1.0))
        End Using
    End Sub

    <Test>
    Public Sub MovesANewDynamicAnchorBeforeTheFirstSave()
        Using workbook = CreateBlankWorkbook()
            Dim sheet = workbook.SheetNames()(0)
            workbook.WriteCellValue(sheet, 0, 0, 3.0)
            workbook.WriteDynamicArrayFormula(sheet, 1, 1, "SEQUENCE(A1)", True)
            workbook.WriteDynamicArrayFormula(sheet, 1, 4, "TRANSPOSE(B2#)", True)
            workbook.AddRow(sheet, 1, 1, updateFormulasAndReferences:=True)
            Assert.That(workbook.LookupCellFormula(sheet, 2, 4), [Is].EqualTo("TRANSPOSE(_xlfn.ANCHORARRAY(B3))"))
            workbook.RecalculateAll()
            Assert.That(Number(workbook, sheet, "B5"), [Is].EqualTo(3.0))
            Assert.That(Number(workbook, sheet, "G3"), [Is].EqualTo(3.0))
        End Using
    End Sub

    Private Shared Function CreateBlankWorkbook() As OpenXmlExcelDataOperations
        Return New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile,
            New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False})
    End Function

    Private Shared Function OpenReference(automatic As Boolean) As OpenXmlExcelDataOperations
        Return OpenWorkbook(TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", "FormulaReferenceSpills.xlsx"), automatic)
    End Function

    Private Shared Function OpenWorkbook(path As String, automatic As Boolean) As OpenXmlExcelDataOperations
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .DisableInitialCalculation = True, .CalculateOnCellEdit = automatic
        }
        Dim workbook As New OpenXmlExcelDataOperations(File.ReadAllBytes(path), options)
        workbook.AutoCalculationEnabledWorkbookSetting = automatic
        Return workbook
    End Function

    Private Shared Function Address(sheet As String, cell As String) As ExcelCell
        Return New ExcelCell(sheet, cell, ExcelCell.ValueTypes.All)
    End Function

    Private Shared Function Number(workbook As OpenXmlExcelDataOperations, sheet As String, cell As String) As Double
        Return workbook.LookupCellValue(Of Double)(Address(sheet, cell), FormulaCalculationMode.UseCachedValueOnly)
    End Function

    Private Shared Function Value(workbook As OpenXmlExcelDataOperations, sheet As String, cell As String) As Object
        Return workbook.LookupCellValueAsObject(Address(sheet, cell), FormulaCalculationMode.UseCachedValueOnly)
    End Function
End Class
