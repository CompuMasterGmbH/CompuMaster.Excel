Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlLookupFormulaTest

    <TestCase("COUNTIF(A1:A9,0)", 2.0)>
    <TestCase("COUNTIF(A1:A9,""=0"")", 2.0)>
    <TestCase("COUNTIF(A1:A9,"""")", 2.0)>
    <TestCase("COUNTIF(A1:A9,""<>"")", 8.0)>
    <TestCase("COUNTIF(A1:A9,TRUE)", 1.0)>
    <TestCase("COUNTIF(A1:A9,""TRUE"")", 0.0)>
    <TestCase("COUNTIF(A1:A9,""#N/A"")", 0.0)>
    <TestCase("COUNTIF(A1:A9,""ap*"")", 2.0)>
    <TestCase("COUNTIF(A1:A9,""ap~?e"")", 1.0)>
    <TestCase("COUNTIF(A1:A9,"">0"")", 1.0)>
    <TestCase("COUNTIF(A1:A9,""<>0"")", 8.0)>
    <TestCase("COUNTIFS(D1:D4,"">=3"",E1:E4,""<70"")", 2.0)>
    <TestCase("SUMIF(D1:D4,3,E1:E1)", 61.0)>
    <TestCase("SUMIF(D1:D4,"">3"")", 7.0)>
    <TestCase("SUMIFS(E1:E4,D1:D4,"">=3"",E1:E4,""<70"")", 61.0)>
    <TestCase("SUMIFS(E1:E4,D1:D3,3)", "#VALUE!")>
    <TestCase("MATCH(2,A9:A9,0)", 1.0)>
    <TestCase("MATCH(0,A4:A4,0)", "#N/A")>
    <TestCase("MATCH(""ap~?e"",A7:A8,0)", 2.0)>
    <TestCase("MATCH(5,D1:D4)", 3.0)>
    <TestCase("INDEX(D1:E4,3,2)", 31.0)>
    <TestCase("SUM(INDEX(D1:E4,2,0))", 33.0)>
    <TestCase("ROW(INDEX(D1:E4,3,2))", 3.0)>
    <TestCase("INDEX(D1:E4,7,2)", "#REF!")>
    <TestCase("VLOOKUP(3,D1:E4,2,FALSE)", 30.0)>
    <TestCase("VLOOKUP(5,D1:E4,2)", 31.0)>
    <TestCase("VLOOKUP(3,D1:E4,3,FALSE)", "#REF!")>
    <TestCase("XLOOKUP(3,D1:D4,E1:E4)", 30.0)>
    <TestCase("_xlfn.XLOOKUP(3,D1:D4,E1:E4,,0,-1)", 31.0)>
    <TestCase("XLOOKUP(5,D1:D4,E1:E4,""missing"")", "missing")>
    <TestCase("XLOOKUP(5,D1:D4,E1:E4,,-1)", 30.0)>
    <TestCase("XLOOKUP(5,D1:D4,E1:E4,,1)", 70.0)>
    Public Sub ComparesCriteriaAndLookupsWithExcelCases(formula As String, expected As Object)
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, Options())
            workbook.WriteCellValue("Sheet1", 0, 0, 0.0)
            workbook.WriteCellFormula("Sheet1", 2, 0, """""", False)
            workbook.WriteCellValue("Sheet1", 3, 0, "0")
            workbook.WriteCellValue("Sheet1", 4, 0, True)
            workbook.WriteCellFormula("Sheet1", 5, 0, "NA()", False)
            workbook.WriteCellValue("Sheet1", 6, 0, "apple")
            workbook.WriteCellValue("Sheet1", 7, 0, "ap?e")
            workbook.WriteCellValue("Sheet1", 8, 0, 2.0)
            Dim keys = {1, 3, 3, 7}
            Dim values = {10, 30, 31, 70}
            For row = 0 To keys.Length - 1
                workbook.WriteCellValue("Sheet1", row, 3, keys(row))
                workbook.WriteCellValue("Sheet1", row, 4, values(row))
            Next
            workbook.WriteCellFormula("Sheet1", 0, 7, formula, False)
            Assert.That(workbook.LookupCellValueAsObject("Sheet1", 0, 7, FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(expected))
        End Using
    End Sub

    <Test>
    Public Sub ResolvesWorkbookNamesAndPrefersWorksheetLocalDefinitions()
        Using workbook As New OpenXmlExcelDataOperations(NamedWorkbook(), Options())
            workbook.WriteCellValue("Inputs", 0, 0, 3)
            workbook.WriteCellValue("Inputs", 1, 0, 8)
            workbook.WriteCellFormula("Inputs", 0, 3, "SUM(Amounts)+Rate", False)
            workbook.WriteCellFormula("Results", 0, 3, "SUM(Amounts)+Rate+Inputs!Rate", False)
            Assert.That(workbook.LookupCellValueAsObject("Inputs", 0, 3, FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(18.0))
            Assert.That(workbook.LookupCellValueAsObject("Results", 0, 3, FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(20.0))

            workbook.WriteCellFormula("Results", 1, 3, "UnknownName", False)
            Assert.That(workbook.LookupCellValueAsObject("Results", 1, 3, FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo("#NAME?"))
            workbook.WriteCellFormula("Results", 2, 3, "RelativeName", False)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.RecalculateCell("Results", 2, 3))
        End Using
    End Sub

    Private Shared Function Options() As ExcelDataOperationsOptions
        Return New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .FormulaEvaluationCulture = System.Globalization.CultureInfo.GetCultureInfo("de-DE")
        }
    End Function

    Private Shared Function NamedWorkbook() As Byte()
        Using memory As New MemoryStream()
            Using doc = SpreadsheetDocument.Create(memory, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
                Dim workbook = doc.AddWorkbookPart()
                Dim inputs = workbook.AddNewPart(Of WorksheetPart)()
                Dim results = workbook.AddNewPart(Of WorksheetPart)()
                inputs.Worksheet = New Worksheet(New SheetData())
                results.Worksheet = New Worksheet(New SheetData())
                workbook.Workbook = New Workbook(New Sheets(
                    New Sheet With {.Name = "Inputs", .SheetId = 1UI, .Id = workbook.GetIdOfPart(inputs)},
                    New Sheet With {.Name = "Results", .SheetId = 2UI, .Id = workbook.GetIdOfPart(results)}),
                    New DefinedNames(
                        New DefinedName With {.Name = "Amounts", .Text = "Inputs!$A$1:$A$2"},
                        New DefinedName With {.Name = "Rate", .Text = "2"},
                        New DefinedName With {.Name = "Rate", .LocalSheetId = 0UI, .Text = "7"},
                        New DefinedName With {.Name = "RelativeName", .Text = "Inputs!A1"}))
            End Using
            Return memory.ToArray()
        End Using
    End Function

End Class
