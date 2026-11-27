Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlStandardFormulaTest

    <TestCase("ROUND(1.005,2)", 1.01)>
    <TestCase("ROUND(-2.5,0)", -3.0)>
    <TestCase("ROUND(125,-1)", 130.0)>
    <TestCase("ROUNDUP(-1.21,1)", -1.3)>
    <TestCase("ROUNDDOWN(-1.29,1)", -1.2)>
    <TestCase("TRUNC(-1.25)", -1.0)>
    <TestCase("INT(-1.25)", -2.0)>
    <TestCase("MOD(3,-2)", -1.0)>
    <TestCase("MOD(2,0)", "#DIV/0!")>
    <TestCase("SIGN(-12)", -1.0)>
    <TestCase("POWER(2,3)", 8.0)>
    <TestCase("SQRT(-4)", "#NUM!")>
    <TestCase("PRODUCT(2,3,4)", 24.0)>
    <TestCase("TRIM(""  a   b  "")", "a b")>
    <TestCase("UPPER(""Test"")", "TEST")>
    <TestCase("LOWER(""TEST"")", "test")>
    <TestCase("LEN(""two words"")", 9.0)>
    <TestCase("LEFT(""ABC"")", "A")>
    <TestCase("RIGHT(""ABC"",5)", "ABC")>
    <TestCase("MID(""ABC"",2,1)", "B")>
    <TestCase("MID(""ABC"",0,1)", "#VALUE!")>
    <TestCase("CONCATENATE(""a"",TRUE,0)", "aTRUE0")>
    <TestCase("IFERROR(1/0,7)", 7.0)>
    <TestCase("IFNA(1/0,7)", "#DIV/0!")>
    <TestCase("IFNA(NA(),7)", 7.0)>
    <TestCase("NOT(0)", True)>
    <TestCase("AND(1,TRUE)", True)>
    <TestCase("AND(""TRUE"")", "#VALUE!")>
    <TestCase("OR(FALSE,TRUE,1/0)", "#DIV/0!")>
    <TestCase("ISBLANK(B1)", True)>
    <TestCase("ISNUMBER(1/0)", False)>
    <TestCase("ISTEXT("""")", True)>
    <TestCase("ISLOGICAL(FALSE)", True)>
    <TestCase("ISERROR(1/0)", True)>
    <TestCase("ISNA(NA())", True)>
    <TestCase("COUNTA(1,,2)", 3.0)>
    <TestCase("COUNTA({1;2;3})", 3.0)>
    <TestCase("COUNTBLANK(B1:C2)", 4.0)>
    Public Sub CalculatesStandardScalarFunctions(formula As String, expected As Object)
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, Options())
            AssertResult(workbook, formula, expected)
        End Using
    End Sub

    <Test>
    Public Sub ReportsUnstableSubperiodOfDecliningFrenchDepreciation()
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, Options())
            workbook.WriteCellFormula("Sheet1", 0, 0, "AMORDEGRC(1200,DATE(2008,7,1),DATE(2008,12,31),150,0.25,0.15)", False)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(
                Sub() workbook.LookupCellValueAsObject("Sheet1", 0, 0, FormulaCalculationMode.CalculateIfMissing))
        End Using
    End Sub

    <TestCase("Inputs!#REF!", "#REF!")>
    <TestCase("SUM(A1:#REF!)", "#REF!")>
    <TestCase("SUM(#REF!:A1)", "#REF!")>
    <TestCase("SUM(A:#REF!)", "#NAME?")>
    <TestCase("SUM(#REF!:A)", "#REF!")>
    <TestCase("Inputs!A1:#REF!", "#REF!")>
    <TestCase("Inputs!#REF!:A1", "#REF!")>
    <TestCase("IF(FALSE,Inputs!A1:#REF!,12)", 12.0)>
    <TestCase("IFERROR(Inputs!#REF!,7)", 7.0)>
    <TestCase("IFERROR(A1:#REF!,""fallback"")", "fallback")>
    Public Sub CalculatesExcelErrorsFromBrokenReferences(formula As String, expected As Object)
        Using workbook As New OpenXmlExcelDataOperations(DateWorkbook(False, includeInputs:=True), Options())
            AssertResult(workbook, formula, expected)
        End Using
    End Sub

    <TestCase(False, "DATE(2011,7,5)", 40729.0)>
    <TestCase(True, "DATE(2011,7,5)", 39267.0)>
    <TestCase(False, "DATE(1900,2,29)", 60.0)>
    <TestCase(False, "DATE(1900,3,0)", 60.0)>
    <TestCase(False, "DATE(1900,2,30)", 61.0)>
    <TestCase(True, "DATE(1900,2,29)", "#NUM!")>
    <TestCase(True, "YEAR(0)", 1904.0)>
    <TestCase(False, "MONTH(60)", 2.0)>
    <TestCase(False, "DAY(60)", 29.0)>
    <TestCase(False, "DAY(0)", 0.0)>
    <TestCase(True, "DAY(0)", 1.0)>
    <TestCase(False, "HOUR(TIME(46,15,6))", 22.0)>
    <TestCase(False, "MINUTE(TIME(1,-30,0))", 30.0)>
    <TestCase(False, "TIME(32768,0,0)", "#NUM!")>
    <TestCase(False, "SECOND(0.5/86400)", 1.0)>
    <TestCase(False, "SECOND(59.9/86400)", 0.0)>
    <TestCase(False, "MINUTE(59.9/86400)", 1.0)>
    <TestCase(False, "HOUR(-0.25)", "#NUM!")>
    <TestCase(False, "YEAR(B1)", 2011.0)>
    <TestCase(True, "B1", 39267.0)>
    Public Sub CalculatesDatesUsingTheWorkbookDateSystem(date1904 As Boolean, formula As String, expected As Object)
        Using workbook As New OpenXmlExcelDataOperations(DateWorkbook(date1904), Options())
            AssertResult(workbook, formula, expected)
        End Using
    End Sub


    <TestCase(False)>
    <TestCase(True)>
    Public Sub CalculatesDeepAcyclicDependenciesAcrossWorksheets(recalculateAll As Boolean)
        Dim settings = Options()
        Dim clock = New DateTime(2024, 7, 10, 11, 30, 0)
        Dim samples = 0
        settings.FormulaLocalClock =
            Function()
                samples += 1
                Return clock
            End Function
        Using workbook As New OpenXmlExcelDataOperations(DateWorkbook(False, includeInputs:=True), settings)
            For row = 0 To 599
                workbook.WriteCellFormula("Sheet1", row, 0, String.Concat("Inputs!A", row + 1, "+1"), False)
                workbook.WriteCellFormula("Inputs", row, 0, String.Concat("Sheet1!A", row + 2, "+1"), False)
            Next
            workbook.WriteCellFormula("Sheet1", 600, 0, "2*HOUR(NOW())", False)
            If recalculateAll Then
                workbook.RecalculateAll()
            Else
                workbook.RecalculateCell("Sheet1", 0, 0)
            End If
            Dim result = workbook.LookupCellValue(Of Double)("Sheet1", 0, 0, FormulaCalculationMode.UseCachedValueOnly)
            Assert.That(result, [Is].EqualTo(1222.0))
            Assert.That(samples, [Is].EqualTo(1))
        End Using
    End Sub

    <Test>
    Public Sub DetectsCircularDependenciesLongerThanTheRecursiveStack()
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, Options())
            For row = 0 To 599
                workbook.WriteCellFormula("Sheet1", row, 0, String.Concat("A", row + 2), False)
            Next
            workbook.WriteCellFormula("Sheet1", 600, 0, "A1", False)
            Assert.Throws(Of CircularReferenceException)(Sub() workbook.RecalculateCell("Sheet1", 0, 0))
        End Using
    End Sub

    <Test>
    Public Sub UsesOneClockSampleForFiveStagesPerCalculationPass()
        Dim current As New DateTime(2011, 7, 5, 12, 0, 0)
        Dim clockReads = 0
        Dim settings = Options()
        settings.FormulaLocalClock =
            Function()
                clockReads += 1
                Return current
            End Function
        Using workbook As New OpenXmlExcelDataOperations(DateWorkbook(False), settings)
            workbook.WriteCellFormula("Sheet1", 1, 0, "NOW()", False)
            workbook.WriteCellFormula("Sheet1", 1, 1, "A2+1", False)
            workbook.WriteCellFormula("Sheet1", 1, 2, "B2+1", False)
            workbook.WriteCellFormula("Sheet1", 1, 3, "C2+1", False)
            workbook.WriteCellFormula("Sheet1", 1, 4, "D2+1", False)
            workbook.WriteCellFormula("Sheet1", 1, 5, "E2+1+NOW()-TODAY()", False)
            workbook.RecalculateCell("Sheet1", 1, 5)
            Assert.That(workbook.LookupCellValue(Of Double)("Sheet1", 1, 5, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(40735.0))
            Assert.That(clockReads, [Is].EqualTo(1))

            current = current.AddHours(6)
            workbook.RecalculateCell("Sheet1", 1, 5)
            Assert.That(workbook.LookupCellValue(Of Double)("Sheet1", 1, 5, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(40735.5))
            Assert.That(clockReads, [Is].EqualTo(2))
        End Using
    End Sub

    <Test>
    Public Sub UsesExplicitGermanCultureForConversionsButKeepsFormulaSyntaxInvariant()
        Dim settings = Options()
        settings.FormulaCulture = Globalization.CultureInfo.GetCultureInfo("de-DE")
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, settings)
            AssertResult(workbook, "CONCATENATE(""a"",TRUE,1.5)", "aWAHR1,5")
            AssertResult(workbook, """a""&FALSE", "aFALSCH")
            AssertResult(workbook, """1,5""+0.5", 2.0)
        End Using
    End Sub

    Private Shared Sub AssertResult(workbook As OpenXmlExcelDataOperations, formula As String, expected As Object)
        workbook.WriteCellFormula("Sheet1", 0, 0, formula, False)
        Dim result = workbook.LookupCellValueAsObject("Sheet1", 0, 0, FormulaCalculationMode.CalculateIfMissing)
        If TypeOf expected Is Double Then
            Assert.That(result, [Is].EqualTo(expected).Within(1.0E-12))
        Else
            Assert.That(result, [Is].EqualTo(expected))
        End If
    End Sub

    Private Shared Function Options() As ExcelDataOperationsOptions
        Return New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False}
    End Function

    Private Shared Function DateWorkbook(date1904 As Boolean, Optional includeInputs As Boolean = False) As Byte()
        Using memory As New MemoryStream()
            Using doc = SpreadsheetDocument.Create(memory, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
                Dim workbook = doc.AddWorkbookPart()
                Dim sheet = workbook.AddNewPart(Of WorksheetPart)()
                Dim data As New SheetData(New Row(New Cell With {
                    .CellReference = "B1",
                    .DataType = CellValues.Date,
                    .CellValue = New CellValue("2011-07-05T00:00:00")
                }) With {.RowIndex = 1})
                sheet.Worksheet = New Worksheet(data)
                workbook.Workbook = New Workbook(New WorkbookProperties With {.Date1904 = date1904},
                    New Sheets(New Sheet With {.Name = "Sheet1", .SheetId = 1UI, .Id = workbook.GetIdOfPart(sheet)}))
                If includeInputs Then
                    Dim inputs = workbook.AddNewPart(Of WorksheetPart)()
                    inputs.Worksheet = New Worksheet(New SheetData())
                    workbook.Workbook.Sheets.Append(New Sheet With {.Name = "Inputs", .SheetId = 2UI, .Id = workbook.GetIdOfPart(inputs)})
                End If
            End Using
            Return memory.ToArray()
        End Using
    End Function

End Class
