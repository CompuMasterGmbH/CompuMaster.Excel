Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports DocumentFormat.OpenXml.Packaging
Imports NUnit.Framework
Imports NativeExcel = Microsoft.Office.Interop.Excel

Namespace ExcelOpsTests.Engines
    Partial Public Class MsExcelOpsTest
        Private Shared Iterator Function EmbeddedChartPlacementCases() As IEnumerable(Of TestCaseData)
            Dim operations = {"DeleteBefore", "DeleteAfter", "DeleteInside", "DeleteFirst", "DeleteLast", "DeleteAll", "DeleteContaining",
                              "InsertBefore", "InsertAtStart", "InsertAfterFirst", "InsertInside", "InsertAtEnd", "InsertAfter"}
            For Each vertical In {True, False}
                For Each placement In {1, 2, 3}
                    For Each customSizes In {False, True}
                        For Each operation In operations
                            Yield New TestCaseData(vertical, placement, operation, customSizes, 51, False, Nothing)
                        Next
                    Next
                Next
                For Each placement In {1, 2, 3}
                    For Each operation In {"DeleteAll", "DeleteInside", "InsertInside"}
                        Yield New TestCaseData(vertical, placement, operation, True, 123, False, Nothing)
                    Next
                    For Each operation In {"InsertAtEnd", "DeleteLast", "DeleteUpToEdge"}
                        Yield New TestCaseData(vertical, placement, operation, False, 51, True, Nothing)
                    Next
                Next
                For Each mark In {"B3", "B4", "C3", "J20", "Format:J20"}
                    Yield New TestCaseData(vertical, 1, "DeleteAll", False, 51, False, mark)
                Next
                For Each chartType In {117, 118, 119, 120, 121, 122, 123, 140}
                    Yield New TestCaseData(vertical, 1, "DeleteAll", True, chartType, False, "J20")
                Next
            Next
        End Function

        ' Chart data is on a separate sheet so that a drawing deletion cannot be confused with a source deletion.
        <TestCaseSource(NameOf(EmbeddedChartPlacementCases))>
        Public Sub ChartReferences_EmbeddedChartPlacementMatchesExcel(vertical As Boolean, placement As Integer, operation As String,
                                                                      customSizes As Boolean, chartType As Integer, exactEnd As Boolean,
                                                                      mark As String)
            Dim caseDirectory = CreateChartTestDirectory()
            Dim input = Path.Combine(caseDirectory, "before.xlsx")
            Dim expected = Path.Combine(caseDirectory, "excel.xlsx")
            Dim actual = Path.Combine(caseDirectory, "openxml.xlsx")
            Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
            Dim workbook As MsExcelDataOperations = Nothing
            Dim mapApplication As CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper = Nothing
            Try
                If chartType = 140 Then mapApplication = NewMapChartApplication()
                workbook = CreateChartWorkbook(mapApplication)
                workbook.AddSheet(ChartDashboardSheet, Nothing)
                Dim sheet = DirectCast(workbook.Workbook.Worksheets.Item(ChartDashboardSheet), NativeExcel.Worksheet)
                Dim data = DirectCast(workbook.Workbook.Worksheets.Item(ChartDataSheet), NativeExcel.Worksheet)
                If customSizes Then
                    sheet.Range("B1").EntireColumn.ColumnWidth = 18
                    sheet.Range("D1").EntireColumn.ColumnWidth = 24
                    sheet.Range("G1").EntireColumn.ColumnWidth = 13
                    sheet.Range("A3").EntireRow.RowHeight = 28
                    sheet.Range("A5").EntireRow.RowHeight = 36
                    sheet.Range("A8").EntireRow.RowHeight = 24
                End If
                If mark IsNot Nothing Then
                    If mark.StartsWith("Format:", StringComparison.Ordinal) Then
                        sheet.Range(mark.Substring(7)).Interior.ColorIndex = 3
                    Else
                        sheet.Range(mark).Value2 = "Worksheet marker"
                    End If
                End If
                Dim start = sheet.Range("B3")
                Dim finish = sheet.Range("G8")
                Dim left = CDbl(start.Left) + 5
                Dim top = CDbl(start.Top) + 4
                Dim width = CDbl(finish.Left) + If(exactEnd, 0, 12) - left
                Dim height = CDbl(finish.Top) + If(exactEnd, 0, 8) - top
                Dim chart As NativeExcel.ChartObject
                If chartType <> 51 Then
                    data.Range("B2").Value2 = "Category"
                    Dim regions = {"Germany", "France", "Canada"}
                    For point = 3 To 5
                        data.Range("B" & point).Value2 = If(chartType = 140, regions(point - 3), "Item " & point)
                    Next
                    data.Activate()
                    data.Range(If(chartType = 118 OrElse chartType = 122, "C2:C5", "B2:C5")).Select()
                    Dim shape = data.Shapes.AddChart2(-1, CType(chartType, NativeExcel.XlChartType))
                    shape.Chart.Location(NativeExcel.XlChartLocation.xlLocationAsObject, ChartDashboardSheet)
                    chart = DirectCast(sheet.ChartObjects(1), NativeExcel.ChartObject)
                    chart.Left = left
                    chart.Top = top
                    chart.Width = width
                    chart.Height = height
                Else
                    chart = DirectCast(sheet.ChartObjects(), NativeExcel.ChartObjects).Add(left, top, width, height)
                    chart.Chart.ChartType = NativeExcel.XlChartType.xlColumnClustered
                    chart.Chart.SetSourceData(data.Range("B2:C5"))
                End If
                chart.Placement = CType(placement, NativeExcel.XlPlacement)
                workbook.SaveAs(input, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                workbook.Close()
                workbook = If(mapApplication Is Nothing, Me.CreateInstance(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options),
                    New MsExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, mapApplication, False, options))
                sheet = DirectCast(workbook.Workbook.Worksheets.Item(ChartDashboardSheet), NativeExcel.Worksheet)
                ApplyEmbeddedChartEdit(workbook, vertical, operation)
                Dim native = ReadChartPlacement(sheet)
                workbook.SaveAs(expected, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                workbook.Close()

                Using editor As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    ApplyEmbeddedChartEdit(editor, vertical, operation)
                    editor.SaveCopyAs(actual, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using
                If native.Length = 0 Then
                    Using document = SpreadsheetDocument.Open(actual, False), nativeDocument = SpreadsheetDocument.Open(expected, False)
                        Dim owner = DirectCast(FindChartSourcePart(document.WorkbookPart, ChartDashboardSheet), WorksheetPart)
                        Dim nativeOwner = DirectCast(FindChartSourcePart(nativeDocument.WorkbookPart, ChartDashboardSheet), WorksheetPart)
                        Assert.That(nativeOwner.DrawingsPart, [Is].Null, "Excel removes the drawing when the last embedded chart is deleted.")
                        Assert.That(owner.DrawingsPart, [Is].Null, "Open XML removes the last chart and its drawing part.")
                    End Using
                End If
                workbook = If(mapApplication Is Nothing, Me.CreateInstance(actual, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options),
                    New MsExcelDataOperations(actual, ExcelDataOperationsBase.OpenMode.OpenExistingFile, mapApplication, False, options))
                Dim reopened = DirectCast(workbook.Workbook.Worksheets.Item(ChartDashboardSheet), NativeExcel.Worksheet)
                Dim result = ReadChartPlacement(reopened)
                Assert.That(result.Length, [Is].EqualTo(native.Length), "Chart presence")
                For index = 0 To result.Length - 1
                    Assert.That(result(index), [Is].EqualTo(native(index)).Within(0.07), {"Left", "Top", "Width", "Height", "Placement"}(index))
                Next
            Finally
                Try
                    If workbook IsNot Nothing AndAlso Not workbook.IsClosed Then workbook.Close()
                Finally
                    mapApplication?.Dispose()
                    Directory.Delete(caseDirectory, True)
                End Try
            End Try
        End Sub

        Private Shared Function ReadChartPlacement(sheet As NativeExcel.Worksheet) As Double()
            Dim charts = DirectCast(sheet.ChartObjects(), NativeExcel.ChartObjects)
            If charts.Count = 0 Then Return Array.Empty(Of Double)()
            Assert.That(charts.Count, [Is].EqualTo(1))
            Dim chart = DirectCast(charts.Item(1), NativeExcel.ChartObject)
            Return {chart.Left, chart.Top, chart.Width, chart.Height, CDbl(chart.Placement)}
        End Function

        Private Shared Sub ApplyEmbeddedChartEdit(workbook As ExcelDataOperationsBase, vertical As Boolean, operation As String)
            Dim first = If(vertical, 2, 1)
            Dim last = If(vertical, 7, 6)
            Dim index As Integer
            Dim count = 1
            Select Case operation
                Case "DeleteBefore", "InsertBefore"
                    index = first - 1
                Case "DeleteFirst", "InsertAtStart"
                    index = first
                Case "InsertAfterFirst"
                    index = first + 1
                Case "DeleteInside", "InsertInside"
                    index = first + 2
                Case "DeleteLast", "InsertAtEnd"
                    index = last
                Case "DeleteAfter", "InsertAfter"
                    index = last + 1
                Case "DeleteAll"
                    index = first
                    count = last - first + 1
                Case "DeleteUpToEdge"
                    index = first
                    count = last - first
                Case "DeleteContaining"
                    index = first - 1
                    count = last - first + 3
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(operation))
            End Select
            If operation.StartsWith("Insert", StringComparison.Ordinal) Then
                If vertical Then
                    workbook.AddRow(ChartDashboardSheet, index, count, True)
                Else
                    workbook.AddColumn(ChartDashboardSheet, index, count, True)
                End If
            ElseIf vertical Then
                workbook.RemoveRows(ChartDashboardSheet, index, count, True)
            Else
                workbook.RemoveColumns(ChartDashboardSheet, index, count, True)
            End If
        End Sub
    End Class
End Namespace
