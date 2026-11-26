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
    <TestCase("FormulaReference1900.xlsx", "Standard", 37)>
    <TestCase("FormulaReference1900.xlsx", "Lookup", 34)>
    <TestCase("FormulaReference1900.xlsx", "Dates", 15)>
    <TestCase("FormulaReference1904.xlsx", "Dates", 5)>
    <TestCase("FormulaReferenceCalendar1900.xlsx", "Calendar", 107)>
    <TestCase("FormulaReferenceCalendar1904.xlsx", "Calendar", 107)>
    <TestCase("FormulaReferenceWorkdays1900.xlsx", "Workdays", 204)>
    <TestCase("FormulaReferenceWorkdays1904.xlsx", "Workdays", 204)>
    <TestCase("FormulaReferenceDateText1900.xlsx", "Date text", 111)>
    <TestCase("FormulaReferenceDateText1904.xlsx", "Date text", 111)>
    <TestCase("FormulaReferenceDateDifference1900.xlsx", "Date differences", 399)>
    <TestCase("FormulaReferenceDateDifference1904.xlsx", "Date differences", 399)>
    <TestCase("FormulaReferenceYearFraction1900.xlsx", "Year fractions", 323)>
    <TestCase("FormulaReferenceYearFraction1904.xlsx", "Year fractions", 323)>
    <TestCase("FormulaReferenceLogical1900.xlsx", "Logical functions", 91)>
    <TestCase("FormulaReferenceLogical1900.xlsx", "Logical arrays", 13)>
    <TestCase("FormulaReferenceInformation1900.xlsx", "Information functions", 241)>
    <TestCase("FormulaReferenceInformation1900.xlsx", "Information arrays", 40)>
    <TestCase("FormulaReferenceStatistics1900.xlsx", "Order statistics", 388)>
    <TestCase("FormulaReferenceStatistics1900.xlsx", "Statistical arrays", 27)>
    <TestCase("FormulaReferenceDispersion1900.xlsx", "Summary statistics", 648)>
    <TestCase("FormulaReferenceDispersion1900.xlsx", "Mode arrays", 10)>
    <TestCase("FormulaReferenceMath1900.xlsx", "Math functions", 546)>
    <TestCase("FormulaReferenceMath1900.xlsx", "Math arrays", 55)>
    <TestCase("FormulaReferenceInteger1900.xlsx", "Integer functions", 801)>
    <TestCase("FormulaReferenceInteger1900.xlsx", "Integer arrays", 49)>
    <TestCase("FormulaReferenceCounting1900.xlsx", "Counting functions", 687)>
    <TestCase("FormulaReferenceCounting1900.xlsx", "Counting arrays", 13)>
    <TestCase("FormulaReferenceDelimiters1900.xlsx", "Delimiters functions", 211)>
    <TestCase("FormulaReferenceDelimiters1900.xlsx", "Delimiters arrays", 66)>
    <TestCase("FormulaReferenceConversion1900.xlsx", "Conversion functions", 448)>
    <TestCase("FormulaReferenceConversion1900.xlsx", "Conversion arrays", 16)>
    <TestCase("FormulaReferenceTextPositionLegacy.xlsx", "Text positions", 79)>
    <TestCase("FormulaReferenceTextPositionLegacy.xlsx", "Position arrays", 4)>
    <TestCase("FormulaReferenceTextPositionUnicode.xlsx", "Text positions", 79)>
    <TestCase("FormulaReferenceTextPositionUnicode.xlsx", "Position arrays", 4)>
    <TestCase("FormulaReferenceHorizontal1900.xlsx", "Horizontal functions", 93)>
    <TestCase("FormulaReferenceHorizontal1900.xlsx", "Horizontal arrays", 16)>
    <TestCase("FormulaReferenceVertical1900.xlsx", "Vertical functions", 25)>
    <TestCase("FormulaReferenceVertical1900.xlsx", "Vertical arrays", 5)>
    <TestCase("FormulaReferenceMatrix1900.xlsx", "Matrix functions", 95)>
    <TestCase("FormulaReferenceMatrix1900.xlsx", "Matrix arrays", 27)>
    <TestCase("FormulaReferenceIf1900.xlsx", "If functions", 33)>
    <TestCase("FormulaReferenceIf1900.xlsx", "If arrays", 30)>
    <TestCase("FormulaReferenceRadix1900.xlsx", "Radix functions", 769)>
    <TestCase("FormulaReferenceRadix1900.xlsx", "Radix arrays", 17)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "Roman functions", 105)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "Roman arrays", 4)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "All Roman forms", 1)>
    <TestCase("FormulaReferenceComplex1900.xlsx", "Complex functions", 501)>
    <TestCase("FormulaReferenceComplex1900.xlsx", "Complex arrays", 7)>
    <TestCase("FormulaReferenceComplexMath1900.xlsx", "ComplexMath functions", 376)>
    <TestCase("FormulaReferenceComplexMath1900.xlsx", "ComplexMath arrays", 3)>
    <TestCase("FormulaReferenceLoan1900.xlsx", "Loan functions", 392)>
    <TestCase("FormulaReferenceLoan1900.xlsx", "Loan arrays", 7)>
    <TestCase("FormulaReferenceLookupVariants.xlsx", "Lookups", 128)>
    <TestCase("FormulaReferenceLookupVariants.xlsx", "Regex", 1)>
    <TestCase("FormulaReferenceRegexVariants.xlsx", "Regex lookups", 33)>
    <TestCase("FormulaReferenceRegexFunctions.xlsx", "Regex functions", 91)>
    <TestCase("FormulaReferenceCriteria.xlsx", "Criteria", 89)>
    <TestCase("FormulaReferenceArrayVariants.xlsx", "Array variants", 65)>
    <TestCase("FormulaReferenceTextVariants.xlsx", "Text variants", 97)>
    <TestCase("FormulaReferenceSortVariants.xlsx", "Sort and unique", 89)>
    <TestCase("FormulaReferenceModern.xlsx", "Arrays", 25)>
    <TestCase("FormulaReferenceModern.xlsx", "Local names", 16)>
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

            Dim actual = ReadFormulaCaches(output, sheetName)
            Assert.That(actual.Keys, [Is].EquivalentTo(expected.Keys), sheetName)
            For Each item In expected
                Dim address = sheetName & "!" & item.Key
                Dim result = actual(item.Key)
                Assert.That(result.Kind, [Is].EqualTo(item.Value.Kind), address)
                If result.Kind = CellValues.Number Then
                    Assert.That(result.Value, [Is].EqualTo(item.Value.Value).Within(1.0E-12), address)
                Else
                    Assert.That(result.Value, [Is].EqualTo(item.Value.Value), address)
                End If
            Next
            If fileName = ReferenceWorkbook Then
                Assert.That(ReadFormulaCaches(output, "External cases"), [Is].EqualTo(ReadFormulaCaches(fixture, "External cases")))
            End If
            If fileName = "FormulaReferenceModern.xlsx" AndAlso sheetName = "Arrays" Then
                AssertSameCalcErrors(fixture, output, sheetName, {"I4", "I5", "I21", "I24"})
            ElseIf fileName = "FormulaReferenceArrayVariants.xlsx" Then
                AssertSameCalcErrors(fixture, output, sheetName, {"S22", "S23", "S31", "S32"})
            ElseIf fileName = "FormulaReferenceTextVariants.xlsx" Then
                AssertSameCalcErrors(fixture, output, sheetName, {"S9", "S37"})
            ElseIf fileName = "FormulaReferenceSortVariants.xlsx" Then
                AssertSameCalcErrors(fixture, output, sheetName, {"S63"})
            End If
        Finally
            If File.Exists(output) Then File.Delete(output)
        End Try
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelRegexResultsAndTheirTypes()
        AssertSpilledFormulaResults("FormulaReferenceRegexFunctions.xlsx", "Array results", "B2", "C2", "D2", "B3", "C3", "D3", "B4", "C4")
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelConditionalResults()
        AssertSpilledFormulaResults("FormulaReferenceLogical1900.xlsx", "Logical arrays",
            "E2", "E6", "D7", "E7", "E10", "E18", "D19", "E19", "E22", "E26", "E30",
            "D31", "E31", "E34", "D43", "E46", "F46", "D47", "E47", "F47")
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelInformationResults()
        AssertSpilledFormulaResults("FormulaReferenceInformation1900.xlsx", "Information arrays",
            "E2", "F2", "G2", "D9", "D10", "D11", "E14", "D15", "E15", "D27", "E32",
            "F32", "G32", "D39", "D40", "D41", "E44", "D45", "E45", "D57", "E62", "F62",
            "G62", "E74", "D75", "E75", "E92", "F92", "G92", "E104", "D105", "E105", "E122",
            "F122", "G122", "E134", "D135", "E135", "E152", "F152", "G152", "E164", "D165", "E165",
            "E212", "F212", "G212", "D219", "D220", "D221", "E224", "D225", "E225", "D237")
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelOrderStatistics()
        AssertSpilledFormulaResults("FormulaReferenceStatistics1900.xlsx", "Statistical arrays",
            "E2", "F2", "D8", "D9", "D13", "D14", "E17", "F17", "D23", "D24",
            "D28", "D29", "E32", "E37", "F37", "D43", "D44", "D48", "D49", "E52",
            "E57", "F57", "D63", "D64", "D68", "D69", "E72", "E77", "F77", "D83",
            "D84", "D88", "D89", "E92", "F92", "G92", "H92", "E97", "F97", "G97",
            "H97", "E102", "F102", "G102", "H102", "E107", "F107", "G107", "H107", "E112",
            "F112", "G112", "H112", "E117", "F117", "G117", "H117")
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelModes()
        AssertSpilledFormulaResults("FormulaReferenceDispersion1900.xlsx", "Mode arrays",
            "D3", "D8", "D13", "D18", "D19", "D23", "D24", "D28", "D33", "D34", "D38", "D39", "D43")
    End Sub

    <TestCase("FormulaReferenceMath1900.xlsx", "Math arrays", 144)>
    <TestCase("FormulaReferenceInteger1900.xlsx", "Integer arrays", 121)>
    <TestCase("FormulaReferenceCounting1900.xlsx", "Counting arrays", 56)>
    <TestCase("FormulaReferenceDelimiters1900.xlsx", "Delimiters arrays", 125)>
    <TestCase("FormulaReferenceConversion1900.xlsx", "Conversion arrays", 31)>
    <TestCase("FormulaReferenceTextPositionLegacy.xlsx", "Position arrays", 10)>
    <TestCase("FormulaReferenceTextPositionUnicode.xlsx", "Position arrays", 10)>
    <TestCase("FormulaReferenceHorizontal1900.xlsx", "Horizontal arrays", 28)>
    <TestCase("FormulaReferenceVertical1900.xlsx", "Vertical arrays", 7)>
    <TestCase("FormulaReferenceMatrix1900.xlsx", "Matrix arrays", 136)>
    <TestCase("FormulaReferenceIf1900.xlsx", "If arrays", 84)>
    <TestCase("FormulaReferenceRadix1900.xlsx", "Radix arrays", 49)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "Roman arrays", 30)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "All Roman forms", 19999)>
    <TestCase("FormulaReferenceComplex1900.xlsx", "Complex arrays", 17)>
    <TestCase("FormulaReferenceComplexMath1900.xlsx", "ComplexMath arrays", 14)>
    <TestCase("FormulaReferenceLoan1900.xlsx", "Loan arrays", 14)>
    Public Sub RecalculatesTheSpilledExcelResults(fileName As String, sheetName As String, expectedChildren As Integer)
        Dim addresses As New List(Of String)
        Using package = SpreadsheetDocument.Open(FixturePath(fileName), False)
            Dim workbook = package.WorkbookPart
            Dim sheet = workbook.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = sheetName)
            Dim worksheet = DirectCast(workbook.GetPartById(sheet.Id.Value), WorksheetPart).Worksheet
            For Each cell In worksheet.Descendants(Of Cell)()
                If cell.CellFormula IsNot Nothing OrElse cell.CellValue Is Nothing Then Continue For
                Dim address As New ExcelCell(sheetName, cell.CellReference.Value, ExcelCell.ValueTypes.All)
                If address.RowIndex > 0 AndAlso address.ColumnIndex >= 3 Then addresses.Add(cell.CellReference.Value)
            Next
        End Using
        Assert.That(addresses, Has.Count.EqualTo(expectedChildren))
        AssertSpilledFormulaResults(fileName, sheetName, addresses.ToArray())
    End Sub

    Private Shared Sub AssertSpilledFormulaResults(fileName As String, sheetName As String, ParamArray addresses As String())
        Dim path = FixturePath(fileName)
        Using expected = OpenForCalculation(path), calculated = OpenForCalculation(path)
            calculated.ResetCellValueFromFormulaCell(sheetName)
            calculated.RecalculateSheet(sheetName)
            For Each address In addresses
                Dim cell As New ExcelCell(sheetName, address, ExcelCell.ValueTypes.All)
                Dim actual = calculated.LookupCellValueAsObject(cell, FormulaCalculationMode.UseCachedValueOnly)
                Dim reference = expected.LookupCellValueAsObject(cell, FormulaCalculationMode.UseCachedValueOnly)
                If TypeOf actual Is Double AndAlso TypeOf reference Is Double Then
                    Assert.That(actual, [Is].EqualTo(reference).Within(1.0E-12), address)
                Else
                    Assert.That(actual, [Is].EqualTo(reference), address)
                End If
            Next
        End Using
    End Sub

    <Test, Explicit("Reports Excel formula coverage for Open XML.")>
    Public Sub ReportsFormulaReferenceCoverageAgainstExcel()
        Dim result = ExcelOpsTests.Engines.FormulaReferenceCoverage.Report("Open XML",
            Function(path, options) New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, options),
            Function(workbook, address) workbook.LookupCellValueAsObject(address, FormulaCalculationMode.CalculateIfMissing))
        Assert.That(result.Matched, [Is].EqualTo(result.Total))
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

    <Test>
    Public Sub RecalculatingASaveCopyUsesSessionSourcesAndKeepsTheirCaches()
        Dim fixture = FixturePath(ReferenceWorkbook)
        Dim sourcePath = FixturePath(CompanionWorkbook)
        Dim outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Dim session As New ExcelFormulaSession()
        session.RegisterFunction("APP.READ",
            Function(context, arguments)
                Assert.That(context.Workbook.Name, [Is].EqualTo(ReferenceWorkbook))
                Dim sourceCell As New ExcelCell("Source Data", "D1", ExcelCell.ValueTypes.All)
                Return context.WorkbookSession.FindWorkbook(CompanionWorkbook).ReadCell(sourceCell)
            End Function)

        Try
            Using target = OpenInSession(fixture, session, usePath:=True), source = OpenInSession(sourcePath, session, usePath:=True)
                source.WriteCellValue("Source Data", 1, 0, 45.0)
                source.WriteCellFormula("Source Data", 0, 3, "A2+3", False)
                target.WriteCellFormula("External cases", 0, 7, "APP.READ()", False)

                target.SaveCopyAs(outputPath, recalculateWorkbook:=True, SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(ExternalValues(target), [Is].EqualTo({42.0, 100.0, 46.0}))
                Assert.That(source.HasMissingCachedFormulaResult("Source Data", 0, 3), [Is].True)
                Assert.That(target.HasMissingCachedFormulaResult("External cases", 0, 7), [Is].True)

                Using output = OpenForCalculation(outputPath)
                    Assert.That(ExternalValues(output), [Is].EqualTo({45.0, 106.0, 49.0}))
                    Assert.That(output.LookupCellValue(Of Double)("External cases", 0, 7, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(48.0))
                End Using
            End Using
        Finally
            If File.Exists(outputPath) Then File.Delete(outputPath)
        End Try
    End Sub

    Private Shared Sub AssertSameCalcErrors(fixture As String, output As String, sheetName As String, addresses As IEnumerable(Of String))
        Using original = OpenForCalculation(fixture), recalculated = OpenForCalculation(output)
            For Each address In addresses
                Dim cell As New ExcelCell(sheetName, address, ExcelCell.ValueTypes.All)
                Assert.That(original.LookupCellValueAsObject(cell, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("#CALC!"), address)
                Assert.That(recalculated.LookupCellValueAsObject(cell, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("#CALC!"), address)
            Next
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
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .FormulaCulture = CultureInfo.GetCultureInfo("de-DE")
        }
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
