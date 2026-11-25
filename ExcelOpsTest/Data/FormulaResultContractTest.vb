Imports System.IO
Imports System.Xml.Linq
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

<TestFixture("epplus4")>
<TestFixture("epplus8")>
<TestFixture("openxml")>
<NonParallelizable>
Public Class FormulaResultContractTest

    Private ReadOnly _engine As String

    Public Sub New(engine As String)
        _engine = engine
    End Sub

    Private Function OpenWorkbook(Optional disableCalculation As Boolean = False, Optional arrayFormula As Boolean = False) As ExcelDataOperationsBase
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableInitialCalculation = True,
            .DisableCalculationEngine = _engine = "openxml" OrElse disableCalculation
        }
        Dim bytes = CreateWorkbook(arrayFormula)
        Select Case _engine
            Case "epplus4"
                Return New EpplusFreeExcelDataOperations(bytes, options)
            Case "epplus8"
                ExcelOpsTests.Engines.EpplusPolyformEditionOpsTest.AssignLicenseContext()
                Return New EpplusPolyformExcelDataOperations(bytes, options)
            Case Else
                Return New OpenXmlExcelDataOperations(bytes, options)
        End Select
    End Function

    ' Asks about the stored cache before and after a read that forbids calculation.
    <Test>
    Public Sub CacheOnlyDoesNotCalculateOrConflateValidResults()
        Dim workbook = OpenWorkbook()
        Try
            Dim target As New ExcelCell("Grid", "A1", ExcelCell.ValueTypes.All)
            Assert.That(workbook.HasMissingCachedFormulaResult(target), [Is].True)
            Dim ex = Assert.Throws(Of MissingCalculatedCellValueException)(
                Sub() workbook.LookupCellValue(Of Integer)(target, FormulaCalculationMode.UseCachedValueOnly))
            Assert.That(ex.Cell.Address, [Is].EqualTo("A1"))
            Assert.That(ex.FoundFormula, [Is].EqualTo("1+2"))
            Assert.That(workbook.HasMissingCachedFormulaResult(target), [Is].True)
            For column = 1 To 5
                Assert.That(workbook.HasMissingCachedFormulaResult("Grid", 0, column), [Is].False)
            Next
            Assert.That(workbook.LookupCellValue(Of Integer)("Grid", 0, 1, FormulaCalculationMode.UseCachedValueOnly), [Is].Zero)
            Assert.That(workbook.LookupCellValue(Of Boolean)("Grid", 0, 2, FormulaCalculationMode.UseCachedValueOnly), [Is].False)
            Assert.That(workbook.LookupCellValue(Of String)("Grid", 0, 3, FormulaCalculationMode.UseCachedValueOnly), [Is].Empty)
        Finally
            workbook.Close()
        End Try
    End Sub

    <Test>
    Public Sub HtmlRequiresExplicitReplacementAndEncodesIt()
        Dim workbook = OpenWorkbook()
        Try
            Dim ex = Assert.Throws(Of MissingCalculatedCellValueException)(
                Sub() workbook.ExportSheetToHtml("Grid", New HtmlSheetExportOptions(), HtmlDocumentExportParts.TableOnly))
            Assert.That(ex.Cell.Address, [Is].EqualTo("A1"))
            Dim options As New HtmlSheetExportOptions With {.MissingFormulaResultText = "<uncomputed>&"}
            Dim html = workbook.ExportSheetToHtml("Grid", options, HtmlDocumentExportParts.TableOnly).ToString()
            Assert.That(html, Does.Contain("&lt;uncomputed&gt;&amp;"))
            Dim cells = XElement.Parse(html).Descendants("td").Select(Function(cell) cell.Value.Trim()).ToArray()
            Dim falseText = If(_engine = "epplus8", "False", "0")
            Assert.That(cells, [Is].EqualTo({"<uncomputed>&", "0", falseText, "", "#DIV/0!", "", "end"}).IgnoreCase)
            options.MissingFormulaResultText = ""
            html = workbook.ExportSheetToHtml("Grid", options, HtmlDocumentExportParts.TableOnly).ToString()
            Assert.That(XElement.Parse(html).Descendants("td").First().Value.Trim(), [Is].Empty)
            Assert.That(workbook.HasMissingCachedFormulaResult("Grid", 0, 0), [Is].True)

            options.FirstColumnIndex = 1
            options.MissingFormulaResultText = Nothing
            Assert.That(workbook.ExportSheetToHtml("Grid", options, HtmlDocumentExportParts.TableOnly).ToString(), Does.Contain("#DIV/0!"))
        Finally
            workbook.Close()
        End Try
    End Sub

    <Test>
    Public Sub CalculationIsOptInAndNativeFailuresStayVisible()
        Dim workbook = OpenWorkbook()
        Try
            Dim options As New HtmlSheetExportOptions With {
                .FormulaCalculationMode = FormulaCalculationMode.CalculateIfMissing, .MissingFormulaResultText = "#Value?"
            }
            If _engine = "openxml" Then
                Assert.Throws(Of NotSupportedException)(Sub() workbook.ExportSheetToHtml("Grid", options))
                Assert.Throws(Of NotSupportedException)(
                    Sub() workbook.LookupCellValueAsObject("Grid", 0, 0, FormulaCalculationMode.CalculateIfMissing))
            Else
                Dim html = workbook.ExportSheetToHtml("Grid", options, HtmlDocumentExportParts.TableOnly).ToString()
                Assert.That(XElement.Parse(html).Descendants("td").First().Value, [Is].EqualTo("3"))
                Assert.That(workbook.LookupCellValue(Of Integer)("Grid", 0, 0, FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(3))
                Assert.That(workbook.HasMissingCachedFormulaResult("Grid", 0, 0), [Is].False)

                workbook.WriteCellFormula("Grid", 0, 0, "1/0", immediatelyCalculateCellValue:=False)
                html = workbook.ExportSheetToHtml("Grid", options, HtmlDocumentExportParts.TableOnly).ToString()
                Assert.That(XElement.Parse(html).Descendants("td").First().Value, [Is].EqualTo("#DIV/0!").IgnoreCase)
                Assert.That(workbook.HasMissingCachedFormulaResult("Grid", 0, 0), [Is].False)
            End If
        Finally
            workbook.Close()
        End Try
    End Sub

    <Test>
    Public Sub DisabledCalculationIsDifferentFromAnUnsupportedFormula()
        If _engine = "openxml" Then Assert.Ignore("This engine has no calculation feature.")
        Dim workbook = OpenWorkbook(disableCalculation:=True)
        Try
            Assert.Throws(Of FeatureDisabledException)(
                Sub() workbook.LookupCellValueAsObject("Grid", 0, 0, FormulaCalculationMode.CalculateIfMissing))
        Finally
            workbook.Close()
        End Try
    End Sub

    <Test>
    Public Sub Epplus4ReportsKnownArrayFormulaLimitation()
        If _engine <> "epplus4" Then Assert.Ignore("The array-formula limitation belongs to EPPlus 4.")
        Dim workbook = OpenWorkbook(arrayFormula:=True)
        Try
            Dim ex = Assert.Throws(Of FormulaComplexityNotSupportedException)(
                Sub() workbook.LookupCellValueAsObject("Grid", 0, 0, FormulaCalculationMode.CalculateIfMissing))
            Assert.That(ex.Cell.Address, [Is].EqualTo("A1"))
            Assert.That(ex.Formula, [Is].EqualTo("1+2"))
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.RecalculateCell("Grid", 0, 0))
            Dim options As New HtmlSheetExportOptions With {
                .FormulaCalculationMode = FormulaCalculationMode.CalculateIfMissing, .MissingFormulaResultText = "#Value?"
            }
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.ExportSheetToHtml("Grid", options))
            Assert.That(workbook.HasMissingCachedFormulaResult("Grid", 0, 0), [Is].True)
        Finally
            workbook.Close()
        End Try
    End Sub

    ' Calls compiled against the earlier signatures still dispatch through the original engine implementation.
    <Test>
    Public Sub LegacyLookupSignaturesRemainCallableAndRequireMigration()
        Dim workbook = OpenWorkbook()
        Try
            Dim target As New ExcelCell("Grid", "B1", ExcelCell.ValueTypes.All)
            Dim range As New ExcelRange("Grid", "B1:B1")
            Dim baseType = GetType(ExcelDataOperationsBase)
            Dim original = baseType.GetMethods().Where(
                Function(method) (method.Name = "LookupCellValue" OrElse method.Name = "LookupCellValueAsObject") AndAlso
                    method.GetParameters().All(Function(parameter) parameter.ParameterType IsNot GetType(FormulaCalculationMode))).ToArray()
            Assert.That(original.Length, [Is].EqualTo(5))
            For Each method In original
                Dim obsolete = DirectCast(Attribute.GetCustomAttribute(method, GetType(ObsoleteAttribute), False), ObsoleteAttribute)
                Assert.That(obsolete?.IsError, [Is].True)
                Assert.That(obsolete.Message, Does.Contain("FormulaCalculationMode"))
                Dim parameters = method.GetParameters()
                Dim firstType = parameters(0).ParameterType
                Dim arguments As Object() = If(firstType Is GetType(String), New Object() {"Grid", 0, 1},
                    If(firstType Is GetType(ExcelRange), New Object() {range}, New Object() {target}))
                Dim legacy = If(method.IsGenericMethodDefinition, method.MakeGenericMethod(GetType(Integer)), method)
                Dim expected = legacy.Invoke(workbook, arguments)
                If firstType Is GetType(ExcelRange) Then
                    Assert.That(workbook.LookupCellValue(Of Integer)(range, FormulaCalculationMode.EngineDefault), [Is].EqualTo(expected))
                ElseIf method.Name = "LookupCellValueAsObject" Then
                    Dim actual = If(firstType Is GetType(ExcelCell), workbook.LookupCellValueAsObject(target, FormulaCalculationMode.EngineDefault),
                        workbook.LookupCellValueAsObject("Grid", 0, 1, FormulaCalculationMode.EngineDefault))
                    Assert.That(actual, [Is].EqualTo(expected))
                Else
                    Dim actual = If(firstType Is GetType(ExcelCell), workbook.LookupCellValue(Of Integer)(target, FormulaCalculationMode.EngineDefault),
                        workbook.LookupCellValue(Of Integer)("Grid", 0, 1, FormulaCalculationMode.EngineDefault))
                    Assert.That(actual, [Is].EqualTo(expected))
                End If
            Next
            Dim engineOverrides = workbook.GetType().GetMethods().Where(
                Function(method) method.DeclaringType Is workbook.GetType() AndAlso
                    (method.Name = "LookupCellValue" OrElse method.Name = "LookupCellValueAsObject")).ToArray()
            Assert.That(engineOverrides.Length, [Is].EqualTo(3))
            For Each method In engineOverrides
                Dim obsolete = DirectCast(Attribute.GetCustomAttribute(method, GetType(ObsoleteAttribute), False), ObsoleteAttribute)
                Assert.That(obsolete?.IsError, [Is].True)
            Next
        Finally
            workbook.Close()
        End Try
    End Sub

    Private Shared Function CreateWorkbook(arrayFormula As Boolean) As Byte()
        Using stream As New MemoryStream()
            Using file = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)
                Dim part = file.AddWorkbookPart()
                Dim sheet = part.AddNewPart(Of WorksheetPart)()
                part.Workbook = New Workbook(New Sheets(New Sheet With {
                    .Name = "Grid", .SheetId = 1UI, .Id = part.GetIdOfPart(sheet)
                }))
                Dim formula As New CellFormula("1+2")
                If arrayFormula Then
                    formula.FormulaType = CellFormulaValues.Array
                    formula.Reference = "A1"
                End If
                sheet.Worksheet = New Worksheet(New SheetData(New Row(
                    New Cell With {.CellReference = "A1", .CellFormula = formula},
                    New Cell With {.CellReference = "B1", .CellFormula = New CellFormula("1-1"), .CellValue = New CellValue("0")},
                    New Cell With {.CellReference = "C1", .CellFormula = New CellFormula("1=2"), .DataType = CellValues.Boolean, .CellValue = New CellValue("0")},
                    New Cell With {.CellReference = "D1", .CellFormula = New CellFormula(""""""), .DataType = CellValues.String, .CellValue = New CellValue("")},
                    New Cell With {.CellReference = "E1", .CellFormula = New CellFormula("1/0"), .DataType = CellValues.Error, .CellValue = New CellValue("#DIV/0!")},
                    New Cell With {.CellReference = "G1", .DataType = CellValues.InlineString, .InlineString = New InlineString(New Text("end"))}
                ) With {.RowIndex = 1UI}))
            End Using
            Return stream.ToArray()
        End Using
    End Function

End Class
