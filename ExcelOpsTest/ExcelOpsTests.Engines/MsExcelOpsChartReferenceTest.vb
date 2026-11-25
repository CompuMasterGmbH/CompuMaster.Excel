Imports System.Globalization
Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework
Imports ChartXml = DocumentFormat.OpenXml.Drawing.Charts
Imports DrawingXml = DocumentFormat.OpenXml.Drawing.Spreadsheet
Imports ExtendedChartXml = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing
Imports NativeExcel = Microsoft.Office.Interop.Excel

Namespace ExcelOpsTests.Engines

    Partial Public Class MsExcelOpsTest

        Private Const ChartDataSheet As String = "Chart Data"
        Private Const ChartDashboardSheet As String = "Dashboard"
        Private Const ChartViewSheet As String = "Chart View"
        Private Const RenamedChartDataSheet As String = "O'Brien Data"

        ' Checks live titles and placement, and saved chart references and point caches for all three placements.
        <TestCase("RowAbove", "$C$3", "$B$4:$B$6", "$C$4:$C$6", "$F$4", "$H$11", "0=10,1=20,2=30")>
        <TestCase("FirstRow", "$C$2", "$B$4:$B$6", "$C$4:$C$6", "$F$4", "$H$11", "0=10,1=20,2=30")>
        <TestCase("InsideRows", "$C$2", "$B$3:$B$6", "$C$3:$C$6", "$F$3", "$H$11", "0=10,2=20,3=30")>
        <TestCase("AfterRows", "$C$2", "$B$3:$B$5", "$C$3:$C$5", "$F$3", "$H$11", "0=10,1=20,2=30")>
        <TestCase("DeleteRow", "$C$2", "$B$3:$B$4", "$C$3:$C$4", "$F$3", "$H$9", "0=10,1=30")>
        <TestCase("DeleteDataRows", "$C$2", Nothing, "#REF!", "#REF!", "$H$7", "0=1")>
        <TestCase("InsertColumn", "$D$2", "$C$3:$C$5", "$D$3:$D$5", "$G$3", "$I$10", "0=10,1=20,2=30")>
        <TestCase("DeleteCategories", "$B$2", Nothing, "$B$3:$B$5", "$E$3", "$G$10", "0=10,1=20,2=30")>
        <TestCase("DeleteValues", Nothing, "$B$3:$B$5", "#REF!", "$E$3", "$G$10", "0=1")>
        <TestCase("ShiftValueCells", "$C$2", "$B$3:$B$5", "$C$3:$C$6", "$F$3", "$H$10", "0=10,2=20,3=30")>
        <TestCase("Rename", "$C$2", "$B$3:$B$5", "$C$3:$C$5", "$F$3", "$H$10", "0=10,1=20,2=30")>
        Public Sub ChartReferences_ExcelStructuralEdits(operation As String, seriesName As String, categories As String,
                                                      values As String, title As String, anchor As String, valueCache As String)
            Dim caseDirectory = CreateChartTestDirectory()
            Dim output = Path.Combine(caseDirectory, "charts.xlsx")
            Dim input = Path.Combine(caseDirectory, "before.xlsx")
            Dim openXmlOutput = Path.Combine(caseDirectory, "openxml.xlsx")
            Dim workbook As MsExcelDataOperations = Nothing
            Try
                workbook = CreateChartWorkbook()
                Dim data = DirectCast(workbook.Workbook.Worksheets.Item(ChartDataSheet), NativeExcel.Worksheet)
                workbook.AddSheet(ChartDashboardSheet, Nothing)
                Dim dashboard = DirectCast(workbook.Workbook.Worksheets.Item(ChartDashboardSheet), NativeExcel.Worksheet)
                Dim dataChart = AddEmbeddedReferenceChart(data)
                Dim dashboardChart = AddEmbeddedReferenceChart(dashboard)
                Dim chartSheet = AddReferenceChartSheet(workbook.Workbook, dashboard)
                ' Create all chart containers before assigning their series so Excel has finished its default data selection.
                ConfigureReferenceChart(dataChart.Chart, NativeExcel.XlChartType.xlColumnClustered)
                ConfigureReferenceChart(dashboardChart.Chart, NativeExcel.XlChartType.xlXYScatter)
                ConfigureReferenceChart(chartSheet, NativeExcel.XlChartType.xlColumnClustered)
                Assert.That(workbook.ChartSheetNames(), [Is].EqualTo({ChartViewSheet}))

                workbook.SaveAs(input, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                ApplyChartDataEdit(workbook, operation)
                Dim qualifier = ChartSheetQualifier(data.Name)
                Assert.That(dataChart.TopLeftCell.Address(), [Is].EqualTo(anchor), "Data-sheet drawing anchor")
                Assert.That(dashboardChart.TopLeftCell.Address(), [Is].EqualTo("$H$10"), "Dashboard drawing anchor")
                For Each chart In {dataChart.Chart, dashboardChart.Chart, chartSheet}
                    Assert.That(chart.ChartTitle.Formula, [Is].EqualTo("=" & qualifier & title))
                Next

                Dim editedSheetName = data.Name
                workbook.SaveAs(output, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                workbook.Close()

                Using document = SpreadsheetDocument.Open(output, False)
                    Dim book = document.WorkbookPart
                    Dim savedData = DirectCast(FindChartSourcePart(book, editedSheetName), WorksheetPart)
                    Dim savedDashboard = DirectCast(FindChartSourcePart(book, ChartDashboardSheet), WorksheetPart)
                    Dim savedChartSheet = DirectCast(FindChartSourcePart(book, ChartViewSheet), ChartsheetPart)
                    AssertSavedChartReferences(savedData.DrawingsPart.ChartParts.Single(), qualifier, seriesName, categories, values, title, valueCache)
                    AssertSavedChartReferences(savedDashboard.DrawingsPart.ChartParts.Single(), qualifier, seriesName, categories, values, title, valueCache)
                    AssertSavedChartReferences(savedChartSheet.DrawingsPart.ChartParts.Single(), qualifier, seriesName, categories, values, title, valueCache)
                End Using

                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
                Using editor As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    ApplyChartDataEdit(editor, operation)
                    editor.SaveCopyAs(openXmlOutput, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using
                Using document = SpreadsheetDocument.Open(openXmlOutput, False)
                    Dim book = document.WorkbookPart
                    Dim savedData = DirectCast(FindChartSourcePart(book, editedSheetName), WorksheetPart)
                    Dim savedDashboard = DirectCast(FindChartSourcePart(book, ChartDashboardSheet), WorksheetPart)
                    Dim savedChartSheet = DirectCast(FindChartSourcePart(book, ChartViewSheet), ChartsheetPart)
                    AssertSavedChartReferences(savedData.DrawingsPart.ChartParts.Single(), qualifier, seriesName, categories, values, title, valueCache)
                    AssertSavedChartReferences(savedDashboard.DrawingsPart.ChartParts.Single(), qualifier, seriesName, categories, values, title, valueCache)
                    AssertSavedChartReferences(savedChartSheet.DrawingsPart.ChartParts.Single(), qualifier, seriesName, categories, values, title, valueCache)
                    Dim marker = savedData.DrawingsPart.WorksheetDrawing.Descendants(Of DrawingXml.FromMarker)().Single()
                    Dim address = New ExcelCell(editedSheetName, Integer.Parse(marker.GetFirstChild(Of DrawingXml.RowId)().Text, CultureInfo.InvariantCulture),
                                                Integer.Parse(marker.GetFirstChild(Of DrawingXml.ColumnId)().Text, CultureInfo.InvariantCulture), ExcelCell.ValueTypes.All)
                    Assert.That(address.Address(False, True, True), [Is].EqualTo(anchor), "Open XML drawing anchor")
                End Using
                workbook = Me.CreateInstance(openXmlOutput, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                Assert.That(workbook.ChartSheetNames(), [Is].EqualTo({ChartViewSheet}))
                Dim reopenedData = DirectCast(workbook.Workbook.Worksheets.Item(editedSheetName), NativeExcel.Worksheet)
                Dim reopenedChart = DirectCast(reopenedData.ChartObjects(1), NativeExcel.ChartObject)
                Assert.That(reopenedChart.TopLeftCell.Address(), [Is].EqualTo(anchor), "Reopened Open XML drawing anchor")
            Finally
                If workbook IsNot Nothing AndAlso Not workbook.IsClosed Then workbook.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        <TestCase(51, "DeleteValues")>
        <TestCase(51, "DeleteDataRows")>
        <TestCase(51, "DeleteCategories")>
        <TestCase(51, "DeleteTitleRow")>
        <TestCase(-4169, "DeleteValues")>
        <TestCase(-4169, "DeleteDataRows")>
        <TestCase(-4169, "DeleteCategories")>
        <TestCase(-4169, "DeleteTitleRow")>
        <TestCase(65, "DeleteValues")>
        <TestCase(65, "DeleteDataRows")>
        <TestCase(65, "DeleteCategories")>
        <TestCase(65, "DeleteTitleRow")>
        Public Sub ChartReferences_OpenXmlMatchesExcelMultiSeriesDeletion(chartType As Integer, operation As String)
            Dim caseDirectory = CreateChartTestDirectory()
            Dim input = Path.Combine(caseDirectory, "before.xlsx")
            Dim expected = Path.Combine(caseDirectory, "excel-edited.xlsx")
            Dim output = Path.Combine(caseDirectory, "openxml-edited.xlsx")
            Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
            Dim workbook As MsExcelDataOperations = Nothing
            Try
                workbook = CreateChartWorkbook()
                Dim data = DirectCast(workbook.Workbook.Worksheets.Item(ChartDataSheet), NativeExcel.Worksheet)
                data.Range("E2").Value2 = "Costs"
                For row = 3 To 5
                    data.Range("E" & row).Value2 = row * 7
                    If chartType <> -4169 Then data.Range("B" & row).Value2 = "Item " & row
                Next
                Dim chart = AddEmbeddedReferenceChart(data).Chart
                chart.ChartType = CType(chartType, NativeExcel.XlChartType)
                Dim series = DirectCast(chart.SeriesCollection(), NativeExcel.SeriesCollection)
                Dim qualifier = ChartSheetQualifier(ChartDataSheet)
                While series.Count > 0
                    series.Item(1).Delete()
                End While
                For Each column As String In {"C", "E"}
                    Dim entry = series.NewSeries()
                    Dim order = If(column = "C", "1", "2")
                    entry.Formula = "=SERIES(" & qualifier & "$" & column & "$2," & qualifier & "$B$3:$B$5," &
                                    qualifier & "$" & column & "$3:$" & column & "$5," & order & ")"
                Next
                workbook.SaveAs(input, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                ApplyChartDataEdit(workbook, operation)
                workbook.SaveAs(expected, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                workbook.Close()

                Using editor As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    ApplyChartDataEdit(editor, operation)
                    editor.SaveCopyAs(output, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using
                Assert.That(ReadClassicSeries(output), [Is].EqualTo(ReadClassicSeries(expected)))
                workbook = Me.CreateInstance(output, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                Dim reopened = GetFirstEmbeddedChart(workbook, ChartDataSheet)
                Assert.That(CInt(reopened.ChartType), [Is].EqualTo(chartType))
                Assert.That(DirectCast(reopened.SeriesCollection(), NativeExcel.SeriesCollection).Count, [Is].EqualTo(2))
            Finally
                If workbook IsNot Nothing AndAlso Not workbook.IsClosed Then workbook.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        Private Shared Function ReadClassicSeries(path As String) As String()
            Using document = SpreadsheetDocument.Open(path, False)
                Dim sheet = DirectCast(FindChartSourcePart(document.WorkbookPart, ChartDataSheet), WorksheetPart)
                Dim chart = sheet.DrawingsPart.ChartParts.Single().ChartSpace
                Dim series = chart.Descendants().Where(Function(item) item.LocalName = "ser").ToArray()
                Assert.That(series.Length, [Is].EqualTo(2))
                Return series.Select(Function(item)
                    Dim references = item.Descendants(Of ChartXml.Formula)().Select(Function(formula)
                        Dim reference = formula.Parent
                        Dim cache = If(DirectCast(reference.GetFirstChild(Of ChartXml.NumberingCache)(), OpenXmlElement),
                                       reference.GetFirstChild(Of ChartXml.StringCache)())
                        Dim points = If(cache?.ChildElements.Where(Function(child) child.LocalName = "pt").Select(
                            Function(point) point.GetAttribute("idx", "").Value & "=" & point.InnerText), Array.Empty(Of String)())
                        Return reference.Parent.LocalName & "=" & formula.Text & ":" &
                               If(cache?.GetFirstChild(Of ChartXml.PointCount)()?.Val?.Value.ToString(), "") & ":" & String.Join(",", points)
                    End Function)
                    Return item.Parent.LocalName & ":" & String.Join("|", references)
                End Function).ToArray()
            End Using
        End Function

        <TestCase(117, "Rename")>
        <TestCase(118, "Rename")>
        <TestCase(117, "InsideRows")>
        <TestCase(118, "InsideRows")>
        <TestCase(119, "InsideRows")>
        <TestCase(120, "InsideRows")>
        <TestCase(121, "InsideRows")>
        <TestCase(122, "InsideRows")>
        <TestCase(123, "InsideRows")>
        <TestCase(140, "InsideRows")>
        <TestCase(117, "DeleteRow")>
        <TestCase(118, "InsertColumn")>
        <TestCase(117, "DeleteCategories")>
        <TestCase(119, "DeleteCategories")>
        <TestCase(120, "DeleteCategories")>
        <TestCase(121, "DeleteCategories")>
        <TestCase(123, "DeleteCategories")>
        <TestCase(140, "DeleteCategories")>
        <TestCase(117, "DeleteTitleRow")>
        <TestCase(118, "DeleteTitleRow")>
        <TestCase(117, "DeleteValues")>
        <TestCase(118, "DeleteValues")>
        <TestCase(119, "DeleteValues")>
        <TestCase(120, "DeleteValues")>
        <TestCase(121, "DeleteValues")>
        <TestCase(122, "DeleteValues")>
        <TestCase(123, "DeleteValues")>
        <TestCase(140, "DeleteValues")>
        <TestCase(117, "DeleteDataRows")>
        <TestCase(118, "DeleteDataRows")>
        <TestCase(117, "DeleteSeriesRow")>
        <TestCase(119, "DeleteSeriesRow")>
        <TestCase(120, "DeleteSeriesRow")>
        <TestCase(121, "DeleteSeriesRow")>
        <TestCase(123, "DeleteSeriesRow")>
        <TestCase(117, "DeleteSeriesFirstRow")>
        <TestCase(119, "DeleteSeriesThenLast")>
        Public Sub ChartReferences_OpenXmlMatchesExcelChartExSourcesAndDrawingAnchors(chartType As Integer, operation As String)
            Dim caseDirectory = CreateChartTestDirectory()
            Dim input = Path.Combine(caseDirectory, "chartex.xlsx")
            Dim expected = Path.Combine(caseDirectory, "excel-edited.xlsx")
            Dim output = Path.Combine(caseDirectory, "openxml-edited.xlsx")
            Dim resultSheet = If(operation = "Rename", RenamedChartDataSheet, ChartDataSheet)
            Dim workbook As MsExcelDataOperations = Nothing
            Try
                workbook = CreateChartWorkbook()
                Dim data = DirectCast(workbook.Workbook.Worksheets.Item(ChartDataSheet), NativeExcel.Worksheet)
                data.Range("B2").Value2 = "Category"
                Dim numericOnly = chartType = 118 OrElse chartType = 122
                If Not numericOnly Then
                    Dim categories = If(chartType = 140, {"Germany", "France", "Canada"}, {"Category 3", "Category 4", "Category 5"})
                    For point = 3 To 5
                        data.Range("B" & point).Value2 = categories(point - 3)
                    Next
                End If
                Dim multipleSeries = operation.StartsWith("DeleteSeries", StringComparison.Ordinal)
                If multipleSeries Then
                    data.Range("D2").Value2 = "Profit"
                    data.Range("E2").Value2 = "Visits"
                    For point = 3 To 5
                        data.Range("D" & point).Value2 = point * 9
                        data.Range("E" & point).Value2 = point * 8
                    Next
                End If
                data.Range(If(numericOnly, "C2:C5", If(multipleSeries, "B2:E5", "B2:C5"))).Select()
                data.Shapes.AddChart2(-1, CType(chartType, NativeExcel.XlChartType), 240, 120, 220, 150)
                DirectCast(data.ChartObjects(1), NativeExcel.ChartObject).Placement = NativeExcel.XlPlacement.xlMoveAndSize
                workbook.SaveAs(input, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                ApplyChartDataEdit(workbook, operation)
                workbook.SaveAs(expected, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                workbook.Close()

                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
                Using editor As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    ApplyChartDataEdit(editor, operation)
                    editor.SaveCopyAs(output, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using
                Dim native = ChartExReferences(expected, resultSheet)
                Dim edited = ChartExReferences(output, resultSheet)
                Dim emptyChart = operation = "DeleteValues" OrElse operation = "DeleteDataRows" OrElse operation = "DeleteSeriesThenLast"
                Assert.That(native.HasChartData, [Is].EqualTo(Not emptyChart))
                If multipleSeries AndAlso Not emptyChart Then Assert.That(native.SeriesLayouts.Length, [Is].EqualTo(2))
                Assert.That(edited.References, [Is].EqualTo(native.References))
                Assert.That(edited.Anchors, [Is].EqualTo(native.Anchors))
                Assert.That(edited.HasChartData, [Is].EqualTo(native.HasChartData))
                Assert.That(edited.SeriesLayouts, [Is].EqualTo(native.SeriesLayouts))
                Assert.That(edited.AxisCount, [Is].EqualTo(native.AxisCount))

                workbook = Me.CreateInstance(expected, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                Dim nativeChart = GetFirstEmbeddedChart(workbook, resultSheet)
                Dim nativeType = CInt(nativeChart.ChartType)
                Dim nativeSeries = DirectCast(nativeChart.SeriesCollection(), NativeExcel.SeriesCollection).Count
                workbook.Close()
                workbook = Me.CreateInstance(output, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                Dim chart = GetFirstEmbeddedChart(workbook, resultSheet)
                Assert.That(CInt(chart.ChartType), [Is].EqualTo(nativeType))
                Assert.That(DirectCast(chart.SeriesCollection(), NativeExcel.SeriesCollection).Count, [Is].EqualTo(nativeSeries))
            Finally
                If workbook IsNot Nothing AndAlso Not workbook.IsClosed Then workbook.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        Private Shared Function GetFirstEmbeddedChart(workbook As MsExcelDataOperations, sheet As String) As NativeExcel.Chart
            Dim worksheet = DirectCast(workbook.Workbook.Worksheets.Item(sheet), NativeExcel.Worksheet)
            Return DirectCast(worksheet.ChartObjects(1), NativeExcel.ChartObject).Chart
        End Function

        Private Shared Function ChartExReferences(path As String, sheet As String) As (References As String(), Anchors As String(),
                                                                                      HasChartData As Boolean, SeriesLayouts As String(), AxisCount As Integer)
            Using document = SpreadsheetDocument.Open(path, False)
                Dim book = document.WorkbookPart
                Dim drawing = DirectCast(FindChartSourcePart(book, sheet), WorksheetPart).DrawingsPart
                Dim chart = drawing.Parts.Select(Function(part) part.OpenXmlPart).OfType(Of ExtendedChartPart)().Single()
                Dim names = book.Workbook.DefinedNames.Elements(Of DefinedName)().ToDictionary(Function(item) item.Name.Value, Function(item) item.Text)
                Dim namesUsed = chart.ChartSpace.Descendants(Of ExtendedChartXml.Formula)().Select(Function(formula) names(formula.Text)).ToArray()
                Dim anchors = drawing.WorksheetDrawing.Descendants(Of DrawingXml.MarkerType)().
                    Select(Function(marker) String.Join(",", marker.ChildElements.Select(Function(element) element.InnerText))).ToArray()
                Dim chartSpace = chart.ChartSpace
                Dim series = chartSpace.Descendants(Of ExtendedChartXml.Series)().Select(
                    Function(item) item.LayoutId?.Value.ToString() & ":" & item.GetFirstChild(Of ExtendedChartXml.DataId)()?.Val?.Value.ToString()
                ).ToArray()
                Return (namesUsed, anchors, chartSpace.ChartData IsNot Nothing, series, chartSpace.Descendants(Of ExtendedChartXml.Axis)().Count())
            End Using
        End Function

        ' Unsupported chart references must fail before even unrelated cell or anchor edits are applied.
        <TestCase(False)>
        <TestCase(True)>
        Public Sub ChartReferences_OpenXmlRejectsUnsupportedChartEditsBeforeMutation(chartSheet As Boolean)
            Dim caseDirectory = CreateChartTestDirectory()
            Dim input = Path.Combine(caseDirectory, "source.xlsx")
            Dim output = Path.Combine(caseDirectory, "copy.xlsx")
            Dim workbook As MsExcelDataOperations = Nothing
            Try
                workbook = CreateChartWorkbook()
                Dim data = DirectCast(workbook.Workbook.Worksheets.Item(ChartDataSheet), NativeExcel.Worksheet)
                Dim chart = If(chartSheet, AddReferenceChartSheet(workbook.Workbook, data), AddEmbeddedReferenceChart(data).Chart)
                ConfigureReferenceChart(chart, NativeExcel.XlChartType.xlColumnClustered)
                workbook.SaveAs(input, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                workbook.Close()
                Using document = SpreadsheetDocument.Open(input, True)
                    Dim partName = If(chartSheet, ChartViewSheet, ChartDataSheet)
                    Dim host = FindChartSourcePart(document.WorkbookPart, partName)
                    Dim drawing = host.Parts.Select(Function(item) item.OpenXmlPart).OfType(Of DrawingsPart)().Single()
                    Dim chartSpace = drawing.ChartParts.Single().ChartSpace
                    chartSpace.Descendants(Of ChartXml.Formula)().First().Text = "UnsupportedNamedRange"
                    chartSpace.Save()
                End Using

                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
                Using editor As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    Assert.That(editor.ChartSheetNames(), [Is].EqualTo(If(chartSheet, {ChartViewSheet}, Array.Empty(Of String)())))
                    Assert.Throws(Of NotSupportedException)(Sub() editor.AddRow(ChartDataSheet, 3, 1, True))
                    editor.SaveCopyAs(output, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using
                Assert.That(File.ReadAllBytes(output), [Is].EqualTo(File.ReadAllBytes(input)))
            Finally
                If workbook IsNot Nothing AndAlso Not workbook.IsClosed Then workbook.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        <TestCase(ExcelDataOperationsBase.CopyWorksheetFormulaMode.ExcelCompatible, 10)>
        <TestCase(ExcelDataOperationsBase.CopyWorksheetFormulaMode.PreserveFormulas, 70)>
        Public Sub WorksheetTransfer_ExcelReopensCopiedEmbeddedChartsAndChartSheets(mode As ExcelDataOperationsBase.CopyWorksheetFormulaMode,
                                                                                   expected As Double)
            Dim caseDirectory = CreateChartTestDirectory()
            Dim sourcePath = Path.Combine(caseDirectory, "source.xlsx")
            Dim targetPath = Path.Combine(caseDirectory, "target.xlsx")
            Dim output = Path.Combine(caseDirectory, "mixed-copy.xlsx")
            Dim workbook As MsExcelDataOperations = Nothing
            Try
                workbook = CreateChartWorkbook()
                Dim data = DirectCast(workbook.Workbook.Worksheets.Item(ChartDataSheet), NativeExcel.Worksheet)
                Dim embedded = AddEmbeddedReferenceChart(data)
                Dim sourceChart = AddReferenceChartSheet(workbook.Workbook, data)
                ConfigureReferenceChart(sourceChart, NativeExcel.XlChartType.xlColumnClustered)
                ConfigureReferenceChart(embedded.Chart, NativeExcel.XlChartType.xlColumnClustered)
                workbook.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                data.Range("C3").Value2 = 70
                workbook.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                workbook.Close()

                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
                Using source As New OpenXmlExcelDataOperations(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options),
                      target As New OpenXmlExcelDataOperations(targetPath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    Dim result = source.CopySheets({ChartDataSheet, ChartViewSheet}, target, 1, mode)
                    Assert.That(result.Select(Function(sheet) sheet.TargetSheetName), [Is].EqualTo({ChartViewSheet & " (2)", ChartDataSheet & " (2)"}))
                    Assert.That(result.Select(Function(sheet) sheet.TargetSheetIndex), [Is].EqualTo({1, 2}))
                    target.SaveCopyAs(output, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using

                workbook = Me.CreateInstance(output, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                Assert.That(workbook.SheetNames(), [Is].EqualTo({ChartViewSheet, ChartViewSheet & " (2)", ChartDataSheet & " (2)", ChartDataSheet}))
                workbook.RecalculateAll()
                Dim copiedData = DirectCast(workbook.Workbook.Worksheets.Item(ChartDataSheet & " (2)"), NativeExcel.Worksheet)
                Dim copiedEmbedded = DirectCast(copiedData.ChartObjects(1), NativeExcel.ChartObject)
                Dim copiedChartSheet = DirectCast(workbook.Workbook.Charts.Item(ChartViewSheet & " (2)"), NativeExcel.Chart)
                For Each chart In {copiedChartSheet, copiedEmbedded.Chart}
                    chart.Refresh()
                    Dim series = DirectCast(chart.SeriesCollection(1), NativeExcel.Series)
                    Dim points = DirectCast(series.Values, Array)
                    Assert.That(CDbl(points.GetValue(points.GetLowerBound(0))), [Is].EqualTo(expected))
                Next
            Finally
                If workbook IsNot Nothing AndAlso Not workbook.IsClosed Then workbook.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        Private Function CreateChartWorkbook() As MsExcelDataOperations
            Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile)
            Dim workbook = Me.CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile, options)
            Dim data = DirectCast(workbook.Workbook.Worksheets.Item(1), NativeExcel.Worksheet)
            data.Name = ChartDataSheet
            data.Range("C2").Value2 = "Revenue"
            data.Range("F3").Value2 = "Report"
            For point = 1 To 3
                DirectCast(data.Cells(point + 2, 2), NativeExcel.Range).Value2 = point
                DirectCast(data.Cells(point + 2, 3), NativeExcel.Range).Value2 = point * 10
            Next
            Return workbook
        End Function

        Private Shared Function AddEmbeddedReferenceChart(sheet As NativeExcel.Worksheet) As NativeExcel.ChartObject
            Dim anchor = sheet.Range("H10")
            Dim objects = DirectCast(sheet.ChartObjects(), NativeExcel.ChartObjects)
            Dim chartObject = objects.Add(CDbl(anchor.Left), CDbl(anchor.Top), 240, 160)
            chartObject.Placement = NativeExcel.XlPlacement.xlMoveAndSize
            Return chartObject
        End Function

        Private Shared Function AddReferenceChartSheet(workbook As NativeExcel.Workbook, before As NativeExcel.Worksheet) As NativeExcel.Chart
            Dim chart = DirectCast(workbook.Charts.Add(Before:=before), NativeExcel.Chart)
            chart.Name = ChartViewSheet
            Return chart
        End Function

        Private Shared Sub ConfigureReferenceChart(chart As NativeExcel.Chart, chartType As NativeExcel.XlChartType)
            chart.ChartType = chartType
            Dim series = DirectCast(chart.SeriesCollection(), NativeExcel.SeriesCollection)
            While series.Count > 0
                series.Item(1).Delete()
            End While
            Dim qualifier = ChartSheetQualifier(ChartDataSheet)
            Dim entry = series.NewSeries()
            entry.Formula = "=SERIES(" & qualifier & "$C$2," & qualifier & "$B$3:$B$5," & qualifier & "$C$3:$C$5,1)"
            chart.HasTitle = True
            chart.ChartTitle.Formula = "=" & qualifier & "$F$3"
        End Sub

        Private Shared Sub ApplyChartDataEdit(workbook As ExcelDataOperationsBase, operation As String)
            Select Case operation
                Case "RowAbove"
                    workbook.AddRow(ChartDataSheet, 1, 1, True)
                Case "FirstRow"
                    workbook.AddRow(ChartDataSheet, 2, 1, True)
                Case "InsideRows"
                    workbook.AddRow(ChartDataSheet, 3, 1, True)
                Case "AfterRows"
                    workbook.AddRow(ChartDataSheet, 5, 1, True)
                Case "DeleteRow"
                    workbook.RemoveRows(ChartDataSheet, 3, 1, True)
                Case "DeleteDataRows"
                    workbook.RemoveRows(ChartDataSheet, 2, 3, True)
                Case "DeleteTitleRow"
                    workbook.RemoveRows(ChartDataSheet, 1, 1, True)
                Case "DeleteSeriesRow"
                    workbook.RemoveRows(ChartDataSheet, 3, 1, True)
                Case "DeleteSeriesFirstRow"
                    workbook.RemoveRows(ChartDataSheet, 2, 1, True)
                Case "DeleteSeriesThenLast"
                    workbook.RemoveRows(ChartDataSheet, 3, 1, True)
                    workbook.RemoveRows(ChartDataSheet, 2, 2, True)
                Case "InsertColumn"
                    workbook.AddColumn(ChartDataSheet, 1, 1, True)
                Case "DeleteCategories"
                    workbook.RemoveColumns(ChartDataSheet, 1, 1, True)
                Case "DeleteValues"
                    workbook.RemoveColumns(ChartDataSheet, 2, 1, True)
                Case "ShiftValueCells"
                    workbook.AddCells(ChartDataSheet, 3, 2, 1, 1, ExcelDataOperationsBase.CellInsertShiftDirection.ShiftCellsDown, True)
                Case "Rename"
                    workbook.RenameWorksheet(ChartDataSheet, RenamedChartDataSheet, True)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(operation))
            End Select
        End Sub

        Private Shared Sub AssertSavedChartReferences(part As ChartPart, qualifier As String, seriesName As String, categories As String,
                                                      values As String, title As String, valueCache As String)
            Dim chart = part.ChartSpace.GetFirstChild(Of ChartXml.Chart)()
            Dim series As OpenXmlElement = chart.Descendants(Of ChartXml.ScatterChartSeries)().SingleOrDefault()
            Dim scatter = series IsNot Nothing
            If Not scatter Then series = chart.Descendants(Of ChartXml.BarChartSeries)().Single()
            Dim categorySource = If(scatter, DirectCast(series.GetFirstChild(Of ChartXml.XValues)(), OpenXmlElement),
                                    series.GetFirstChild(Of ChartXml.CategoryAxisData)())
            Dim valueSource = If(scatter, DirectCast(series.GetFirstChild(Of ChartXml.YValues)(), OpenXmlElement),
                                 series.GetFirstChild(Of ChartXml.Values)())
            Dim expectedCategory = If(scatter AndAlso categories Is Nothing, "#REF!", categories)
            Dim location = part.Uri.ToString()
            Assert.That(ChartReference(chart.GetFirstChild(Of ChartXml.Title)()), [Is].EqualTo(qualifier & title), location & " title")
            Assert.That(ChartReference(series.GetFirstChild(Of ChartXml.SeriesText)()),
                        [Is].EqualTo(If(seriesName Is Nothing, Nothing, qualifier & seriesName)), location & " series name")
            Assert.That(ChartReference(categorySource),
                        [Is].EqualTo(If(expectedCategory Is Nothing, Nothing, qualifier & expectedCategory)), location & " categories or X values")
            Assert.That(ChartReference(valueSource), [Is].EqualTo(qualifier & values), location & " values")
            Dim points = valueSource.Descendants(Of ChartXml.NumericPoint)().Select(
                Function(point) point.Index.Value.ToString(CultureInfo.InvariantCulture) & "=" & point.NumericValue.Text)
            Assert.That(String.Join(",", points), [Is].EqualTo(valueCache), location & " cached points")
        End Sub

        Private Shared Function ChartReference(element As OpenXmlElement) As String
            Return element?.Descendants(Of ChartXml.Formula)().SingleOrDefault()?.Text
        End Function

        Private Shared Function FindChartSourcePart(workbook As WorkbookPart, name As String) As OpenXmlPart
            Dim sheet = workbook.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = name)
            Return workbook.GetPartById(sheet.Id.Value)
        End Function

        Private Shared Function ChartSheetQualifier(name As String) As String
            Return "'" & name.Replace("'", "''") & "'!"
        End Function

        Private Shared Function CreateChartTestDirectory() As String
            Return Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ExcelOps-ChartReferences-" & Guid.NewGuid().ToString("N"))).FullName
        End Function

    End Class

End Namespace
