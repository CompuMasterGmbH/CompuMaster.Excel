Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlFormulaCalculationTest

    Private Const SheetName As String = "Grid"

    Private Shared Function Options(Optional calculateOnOpen As Boolean = False) As ExcelDataOperationsOptions
        Return New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .DisableInitialCalculation = Not calculateOnOpen
        }
    End Function

    ' These results include operator precedence and coercion cases checked against Microsoft Excel.
    <TestCase("=-2^2+2^3^2", 68.0)>
    <TestCase("=(2+3)*4-50%", 19.5)>
    <TestCase("=2^-2", 0.25)>
    <TestCase("=4/2+1.5E+1", 17.0)>
    <TestCase("=ABS(-4)", 4.0)>
    <TestCase("=SUM(""2"",TRUE,3)", 6.0)>
    <TestCase("=COUNT(#DIV/0!,2)", 1.0)>
    <TestCase("=SUM(IF(TRUE,'Input''s data'!A3,1))", 0.0)>
    <TestCase("=COUNT(IF(TRUE,'Input''s data'!A3,1))", 0.0)>
    <TestCase("=IF(TRUE,,3)", 0.0)>
    <TestCase("=IF(,1,2)", 2.0)>
    <TestCase("=IF(FALSE,A1,7)", 7.0)>
    <TestCase("=1<>2", True)>
    <TestCase("=IF(FALSE,3)", False)>
    <TestCase("=""a""=""A""", True)>
    <TestCase("=""a""""b""&TRUE", "a""bTRUE")>
    <TestCase("=1/0", "#DIV/0!")>
    <TestCase("=0^0", "#NUM!")>
    <TestCase("=IF(""1"",1,2)", "#VALUE!")>
    Public Sub CalculatesScalarFormulasThroughExplicitLookup(formula As String, expected As Object)
        Using workbook As New OpenXmlExcelDataOperations(CreateWorkbook(), Options())
            workbook.WriteCellFormula(SheetName, 0, 0, formula, immediatelyCalculateCellValue:=False)
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 0), [Is].True)
            Assert.That(ReadResult(workbook, "A1", FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(expected))
            Assert.That(ReadResult(workbook, "A1"), [Is].EqualTo(expected))
        End Using
    End Sub

    <Test>
    Public Sub CalculatesPrecedentsAndAggregatesAcrossWorksheets()
        Using workbook As New OpenXmlExcelDataOperations(CreateWorkbook(), Options())
            workbook.WriteCellFormula(SheetName, 0, 0, "SUM('Input''s data'!$A$1:A3)", immediatelyCalculateCellValue:=False)
            workbook.WriteCellFormula(SheetName, 0, 1, "AVERAGE('Input''s data'!A1:A3)", immediatelyCalculateCellValue:=False)
            workbook.WriteCellFormula(SheetName, 0, 2, "MIN('Input''s data'!A1:A3)+MAX('Input''s data'!A1:A3)+COUNT('Input''s data'!A1:A3)", immediatelyCalculateCellValue:=False)

            workbook.RecalculateSheet("grid")
            Assert.That(ReadResult(workbook, "A1"), [Is].EqualTo(12.0))
            Assert.That(ReadResult(workbook, "B1"), [Is].EqualTo(6.0))
            Assert.That(ReadResult(workbook, "C1"), [Is].EqualTo(14.0))

            workbook.WriteCellValue("Input's data", 0, 0, 8)
            Assert.That(ReadResult(workbook, "A1", FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(12.0))
            workbook.RecalculateCell(SheetName, 0, 0)
            Assert.That(ReadResult(workbook, "A1"), [Is].EqualTo(24.0))
            Assert.That(ReadResult(workbook, "B1"), [Is].EqualTo(6.0))
            workbook.RecalculateAll()
            Assert.That(ReadResult(workbook, "B1"), [Is].EqualTo(12.0))
        End Using
    End Sub

    <Test>
    Public Sub KeepsExistingCachesAndImmediateEditsWhenCalculationIsUnsupported()
        Using workbook As New OpenXmlExcelDataOperations(CreateWorkbook(), Options())
            workbook.WriteCellFormula(SheetName, 0, 0, "1+2", immediatelyCalculateCellValue:=True)
            workbook.WriteCellFormula(SheetName, 0, 1, "A1+4", immediatelyCalculateCellValue:=False)
            workbook.WriteCellFormula(SheetName, 0, 2, "SUMPRODUCT(A1:B1)", immediatelyCalculateCellValue:=False)
            Dim unsupported = Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.RecalculateAll())
            Assert.That(unsupported.Cell.Address, [Is].EqualTo("C1"))
            Assert.That(unsupported.Formula, [Is].EqualTo("SUMPRODUCT(A1:B1)"))
            Assert.That(ReadResult(workbook, "A1"), [Is].EqualTo(3.0))
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 1), [Is].True)
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 2), [Is].True)

            Dim circular = Assert.Throws(Of CircularReferenceException)(
                Sub() workbook.WriteCellFormula(SheetName, 0, 0, "B1", immediatelyCalculateCellValue:=True))
            Assert.That(circular.Cell.Address, [Is].EqualTo("A1"))
            Assert.That(circular.Formula, [Is].EqualTo("B1"))
            Assert.That(ReadResult(workbook, "A1"), [Is].EqualTo(3.0))
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 1), [Is].True)

            workbook.WriteCellFormula(SheetName, 0, 2, "'[external.xlsx]Data'!A1", immediatelyCalculateCellValue:=False)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() ReadResult(workbook, "C1", FormulaCalculationMode.CalculateIfMissing))
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 2), [Is].True)
        End Using
    End Sub

    ' Excel can answer these questions from the reference without calculating its cells.
    <TestCase("ROW(B3)", 3.0)>
    <TestCase("COLUMN($B$3)", 2.0)>
    <TestCase("ROW()", 3.0)>
    <TestCase("COLUMN()", 2.0)>
    <TestCase("ROW(B5:B2)", 2.0)>
    <TestCase("COLUMN(A3:C3)", 1.0)>
    <TestCase("ROWS(B2:B5)", 4.0)>
    <TestCase("COLUMNS(A3:C3)", 3.0)>
    <TestCase("ROW(IF(TRUE,B3,A1))", 3.0)>
    <TestCase("ROW('Input''s data'!A2)", 2.0)>
    <TestCase("COLUMNS(""text"")", 1.0)>
    <TestCase("ROWS(#REF!)", "#REF!")>
    Public Sub AddressFunctionsUseTheReferenceWithoutReadingItsValue(formula As String, expected As Object)
        Using workbook As New OpenXmlExcelDataOperations(CreateWorkbook(), Options())
            workbook.WriteCellFormula(SheetName, 2, 1, formula, immediatelyCalculateCellValue:=False)
            Assert.That(ReadResult(workbook, "B3", FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(expected))
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 0), [Is].True)
            Assert.That(workbook.HasMissingCachedFormulaResult("Input's data", 1, 0), [Is].True)
        End Using
    End Sub

    ' Nested formulas evaluate ROW and COLUMN using their own addresses.
    <Test>
    Public Sub AddressFunctionsUseThePrecedentsLocation()
        Using workbook As New OpenXmlExcelDataOperations(CreateWorkbook(), Options())
            workbook.WriteCellFormula("Input's data", 1, 2, "ROW()+COLUMN()", immediatelyCalculateCellValue:=False)
            workbook.WriteCellFormula(SheetName, 0, 0, "'Input''s data'!C2", immediatelyCalculateCellValue:=True)
            Assert.That(ReadResult(workbook, "A1"), [Is].EqualTo(5.0))
        End Using
    End Sub

    <TestCase("B3+1")>
    <TestCase("SUM(B2:B5)")>
    <TestCase("SUM(#DIV/0!,B3)")>
    <TestCase("IF(TRUE,B3,7)")>
    <TestCase("ROWS(B3+1)")>
    Public Sub ValueCyclesAlwaysRaiseTheCircularReferenceException(formula As String)
        Using workbook As New OpenXmlExcelDataOperations(CreateWorkbook(), Options())
            workbook.WriteCellFormula(SheetName, 2, 1, formula, immediatelyCalculateCellValue:=False)
            Dim cycle = Assert.Throws(Of CircularReferenceException)(
                Sub() workbook.RecalculateCell(SheetName, 2, 1, throwExceptionOnCalculationError:=False))
            Assert.That(cycle.Cell.SheetName, [Is].EqualTo(SheetName))
            Assert.That(cycle.Cell.Address, [Is].EqualTo("B3"))
            Assert.That(cycle.Formula, [Is].EqualTo(formula))
            Assert.Throws(Of CircularReferenceException)(Sub() ReadResult(workbook, "B3", FormulaCalculationMode.CalculateIfMissing))
            Dim htmlOptions As New HtmlSheetExportOptions With {
                .FormulaCalculationMode = FormulaCalculationMode.CalculateIfMissing, .MissingFormulaResultText = ""
            }
            Assert.Throws(Of CircularReferenceException)(Sub() workbook.ExportSheetToHtml(SheetName, htmlOptions))
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 2, 1), [Is].True)
        End Using
    End Sub

    <Test>
    Public Sub CrossWorksheetCycleLeavesStagedCachesUnchanged()
        Using workbook As New OpenXmlExcelDataOperations(CreateWorkbook(), Options())
            workbook.WriteCellFormula(SheetName, 0, 1, "'Input''s data'!A2", immediatelyCalculateCellValue:=False)
            workbook.WriteCellFormula("Input's data", 1, 0, "Grid!B1+1", immediatelyCalculateCellValue:=False)
            Assert.Throws(Of CircularReferenceException)(Sub() workbook.RecalculateAll())
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 0), [Is].True)
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 1), [Is].True)
            Assert.That(workbook.HasMissingCachedFormulaResult("Input's data", 1, 0), [Is].True)
        End Using
    End Sub

    <Test>
    Public Sub ReportsGroupedFormulasWithoutWritingCaches()
        Using workbook As New OpenXmlExcelDataOperations(CreateWorkbook(arrayFormula:=True), Options())
            Dim unsupported = Assert.Throws(Of FormulaComplexityNotSupportedException)(
                Sub() ReadResult(workbook, "A1", FormulaCalculationMode.CalculateIfMissing))
            Assert.That(unsupported.Formula, [Is].EqualTo("1+2"))
            Assert.That(unsupported.Message, Does.Contain("array"))
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.RecalculateCell(SheetName, 1, 0))
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 0), [Is].True)
        End Using
    End Sub

    <Test>
    Public Sub RequiresOptInAndCanCalculateOnOpenWhenRequested()
        Dim bytes = CreateWorkbook()
        Using workbook As New OpenXmlExcelDataOperations(bytes, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.Throws(Of FeatureDisabledException)(Sub() ReadResult(workbook, "A1", FormulaCalculationMode.CalculateIfMissing))
        End Using
        Using workbook As New OpenXmlExcelDataOperations(bytes, Options())
            Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, 0), [Is].True)
        End Using
        Using workbook As New OpenXmlExcelDataOperations(bytes, Options(calculateOnOpen:=True))
            Assert.That(ReadResult(workbook, "A1"), [Is].EqualTo(3.0))
        End Using
    End Sub

    <Test>
    Public Sub SavesTypedCachesAndFormulaTextForCacheOnlyReaders()
        Dim destination = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Try
            Using workbook As New OpenXmlExcelDataOperations(CreateWorkbook(), Options())
                workbook.WriteCellFormula(SheetName, 0, 1, "1=2", immediatelyCalculateCellValue:=True)
                workbook.WriteCellFormula(SheetName, 0, 2, """_x0041_""", immediatelyCalculateCellValue:=True)
                workbook.WriteCellFormula(SheetName, 0, 3, """""", immediatelyCalculateCellValue:=True)
                workbook.WriteCellFormula(SheetName, 0, 4, "1/0", immediatelyCalculateCellValue:=False)
                Assert.Throws(Of InvalidOperationException)(Sub() workbook.RecalculateCell(SheetName, 0, 4))
                Assert.That(ReadResult(workbook, "E1"), [Is].EqualTo("#DIV/0!"))
                workbook.RecalculateAll()
                workbook.SaveCopyAs(destination, recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using
            Using workbook As New OpenXmlExcelDataOperations(File.ReadAllBytes(destination), New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
                Assert.That(ReadResult(workbook, "A1"), [Is].EqualTo(3.0))
                Assert.That(ReadResult(workbook, "B1"), [Is].EqualTo(False))
                Assert.That(ReadResult(workbook, "C1"), [Is].EqualTo("_x0041_"))
                Assert.That(ReadResult(workbook, "D1"), [Is].EqualTo(""))
                Assert.That(ReadResult(workbook, "E1"), [Is].EqualTo("#DIV/0!"))
                For column = 0 To 4
                    Assert.That(workbook.HasMissingCachedFormulaResult(SheetName, 0, column), [Is].False)
                Next
            End Using
            Using file = SpreadsheetDocument.Open(destination, False)
                Dim part = file.WorkbookPart
                Dim sheet = part.Workbook.Sheets.Elements(Of Sheet)().First()
                Dim first = DirectCast(part.GetPartById(sheet.Id.Value), WorksheetPart).Worksheet.Descendants(Of Cell)().First()
                Assert.That(first.CellFormula.Text, [Is].EqualTo("1+2"))
                Assert.That(first.CellValue.Text, [Is].EqualTo("3"))
            End Using
        Finally
            If File.Exists(destination) Then File.Delete(destination)
        End Try
    End Sub

    Private Shared Function ReadResult(workbook As OpenXmlExcelDataOperations, address As String,
                                       Optional mode As FormulaCalculationMode = FormulaCalculationMode.UseCachedValueOnly) As Object
        Return workbook.LookupCellValueAsObject(New ExcelCell(SheetName, address, ExcelCell.ValueTypes.All), mode)
    End Function

    Private Shared Function CreateWorkbook(Optional arrayFormula As Boolean = False) As Byte()
        Using stream As New MemoryStream()
            Using file = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)
                Dim part = file.AddWorkbookPart()
                Dim grid = part.AddNewPart(Of WorksheetPart)()
                Dim input = part.AddNewPart(Of WorksheetPart)()
                part.Workbook = New Workbook(New Sheets(
                    New Sheet With {.Name = SheetName, .SheetId = 1UI, .Id = part.GetIdOfPart(grid)},
                    New Sheet With {.Name = "Input's data", .SheetId = 2UI, .Id = part.GetIdOfPart(input)}))
                Dim formula As New CellFormula("1+2")
                If arrayFormula Then
                    formula.FormulaType = CellFormulaValues.Array
                    formula.Reference = "A1:A2"
                End If
                grid.Worksheet = New Worksheet(New SheetData(
                    New Row(New Cell With {.CellReference = "A1", .CellFormula = formula}) With {.RowIndex = 1UI}))
                input.Worksheet = New Worksheet(New SheetData(
                    New Row(New Cell With {.CellReference = "A1", .CellValue = New CellValue("4")}) With {.RowIndex = 1UI},
                    New Row(New Cell With {.CellReference = "A2", .CellFormula = New CellFormula("A1*2")}) With {.RowIndex = 2UI},
                    New Row(New Cell With {.CellReference = "A3", .DataType = CellValues.Boolean, .CellValue = New CellValue("1")}) With {.RowIndex = 3UI}))
            End Using
            Return stream.ToArray()
        End Using
    End Function

End Class
