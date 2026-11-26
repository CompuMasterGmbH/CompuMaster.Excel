Imports System.IO
Imports System.Runtime.InteropServices
Imports CompuMaster.Excel.ExcelOps
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework
Imports NativeExcel = Microsoft.Office.Interop.Excel

Namespace ExcelOpsTests.Engines

    Partial Public Class MsExcelOpsTest

        ' Excel moves surviving filter fields and unhides the body when the final criterion is removed.
        <TestCase("InsertBeforeFilter")>
        <TestCase("DeleteBeforeFilter")>
        <TestCase("DeleteLastFilter")>
        <TestCase("InsertHeaderless")>
        <TestCase("DeleteHeaderless")>
        <TestCase("DeleteOneOfTwoFilters")>
        <TestCase("DeleteFilterKeepAnd")>
        <TestCase("DeleteFilterKeepOr")>
        <TestCase("DeleteFilterWithFormula")>
        <TestCase("InsertAtFirstColumn")>
        <TestCase("InsertBeforeTable")>
        <TestCase("InsertAfterHiddenColumn")>
        <TestCase("InsertAfterDefaultHiddenColumn")>
        <TestCase("InsertAfterTable")>
        <TestCase("InsertAtFarEdge")>
        <TestCase("DeleteOnlyCustomWidth")>
        <TestCase("DeleteHiddenColumnWithFilter")>
        <TestCase("DeleteHiddenColumnWithAllRowsVisible")>
        <TestCase("DeleteHiddenOutsideTableWithFilter")>
        <TestCase("DeleteMixedVisibilityWithFilter")>
        Public Sub StructuralReference_OpenXmlMatchesTableColumnsWithFiltersOrHiddenHeaders(operation As String)
            Dim directoryPath = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ExcelOps-TableColumns-" & Guid.NewGuid().ToString("N"))).FullName
            Dim input = Path.Combine(directoryPath, "before.xlsx")
            Dim nativeOutput = Path.Combine(directoryPath, "excel.xlsx")
            Dim openXmlOutput = Path.Combine(directoryPath, "openxml.xlsx")
            Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
            Dim workbook As MsExcelDataOperations = Nothing
            Try
                Dim createOptions As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile)
                workbook = Me.CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile, createOptions)
                Dim sheet = DirectCast(workbook.Workbook.Worksheets.Item(1), NativeExcel.Worksheet)
                sheet.Name = "Data"
                workbook.AddSheet("Checks", Nothing)
                For Each field In {(Column:=1, Header:="Id"), (Column:=2, Header:="Region"), (Column:=3, Header:="Sales")}
                    workbook.WriteCellValue("Data", 1, field.Column, field.Header)
                Next
                For row = 2 To 5
                    workbook.WriteCellValue("Data", row, 1, row - 1)
                    workbook.WriteCellValue("Data", row, 2, If(row Mod 2 = 1, "West", "East"))
                    workbook.WriteCellValue("Data", row, 3, If(row = 3, 5, If(row = 4, 10, 8)))
                Next
                If operation = "DeleteFilterWithFormula" Then workbook.WriteCellFormula("Data", 3, 3, "2+3", False)
                Dim table = sheet.ListObjects.Add(NativeExcel.XlListObjectSourceType.xlSrcRange, sheet.Range("B2:D6"),
                                                  Type.Missing, NativeExcel.XlYesNoGuess.xlYes)
                table.Name = "SalesTable"
                sheet.Range("A1").EntireColumn.ColumnWidth = 10
                If operation <> "DeleteOnlyCustomWidth" Then
                    If operation <> "InsertAfterDefaultHiddenColumn" Then sheet.Range("C1:D1").EntireColumn.ColumnWidth = 14
                    Select Case operation
                        Case "DeleteLastFilter", "DeleteOneOfTwoFilters", "DeleteFilterKeepAnd", "DeleteFilterKeepOr", "DeleteFilterWithFormula"
                            ' Keep the selection visible to exercise Excel's filter changes after a successful deletion.
                        Case "DeleteMixedVisibilityWithFilter"
                            sheet.Range("C1").EntireColumn.Hidden = True
                        Case Else
                            sheet.Range("C1:D1").EntireColumn.Hidden = True
                    End Select
                    sheet.Range("F1:G1").EntireColumn.ColumnWidth = 18
                End If
                If operation = "DeleteHiddenOutsideTableWithFilter" Then sheet.Range("F1").EntireColumn.Hidden = True
                If operation = "InsertAtFarEdge" Then sheet.Range("XFC1:XFD1").EntireColumn.ColumnWidth = 20
                Dim headerless = operation.EndsWith("Headerless", StringComparison.Ordinal)
                If headerless Then
                    table.ShowHeaders = False
                Else
                    table.Range.AutoFilter(2, If(operation = "DeleteHiddenColumnWithAllRowsVisible", "*", "West"))
                End If
                If operation = "DeleteLastFilter" Then
                    ' Excel also unhides manually hidden data rows and leaves surrounding rows hidden.
                    sheet.Range("A2").EntireRow.Hidden = True
                    sheet.Range("A6").EntireRow.Hidden = True
                    sheet.Range("A7").EntireRow.Hidden = True
                End If
                Select Case operation
                    Case "DeleteOneOfTwoFilters", "DeleteFilterWithFormula"
                        table.Range.AutoFilter(3, ">5")
                    Case "DeleteFilterKeepAnd"
                        table.Range.AutoFilter(3, ">5", NativeExcel.XlAutoFilterOperator.xlAnd, "<9")
                    Case "DeleteFilterKeepOr"
                        table.Range.AutoFilter(3, "<7", NativeExcel.XlAutoFilterOperator.xlOr, ">9")
                End Select
                workbook.WriteCellFormula("Checks", 0, 0, "COUNTIF(SalesTable[Region],""West"")", False)
                workbook.SaveAs(input, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)

                Dim inserting = operation.StartsWith("Insert", StringComparison.Ordinal)
                Dim columnIndex = 2
                Select Case operation
                    Case "InsertAtFirstColumn", "DeleteOnlyCustomWidth"
                        columnIndex = 0
                    Case "InsertBeforeTable", "DeleteBeforeFilter"
                        columnIndex = 1
                    Case "InsertAfterHiddenColumn", "InsertAfterDefaultHiddenColumn"
                        columnIndex = 3
                    Case "InsertAfterTable"
                        columnIndex = 4
                    Case "InsertAtFarEdge"
                        columnIndex = 16382
                    Case "DeleteHiddenOutsideTableWithFilter"
                        columnIndex = 5
                End Select
                Dim columns = If(operation = "InsertAfterTable" OrElse operation = "DeleteMixedVisibilityWithFilter", 2, 1)
                Dim edit As Action(Of ExcelDataOperationsBase) =
                    Sub(engine)
                        If inserting Then
                            engine.AddColumn("Data", columnIndex, columns, True)
                        Else
                            engine.RemoveColumns("Data", columnIndex, columns, True)
                        End If
                    End Sub
                If operation.StartsWith("DeleteHidden", StringComparison.Ordinal) Then
                    Assert.Throws(Of COMException)(Sub() edit(workbook))
                    workbook.Close()
                    Using editor As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                        Assert.Throws(Of StructuralReferenceCompatibilityException)(Sub() edit(editor))
                        editor.SaveCopyAs(openXmlOutput, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                    End Using
                    Assert.That(File.ReadAllBytes(openXmlOutput), [Is].EqualTo(File.ReadAllBytes(input)))
                    Dim optIn As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly) With {
                        .AllowKnownStructuralReferenceDifferences = True
                    }
                    Using editor As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, optIn)
                        edit(editor)
                        editor.SaveCopyAs(openXmlOutput, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                    End Using
                    Dim afterOptIn = ReadTableColumnState(openXmlOutput)
                    Assert.That(afterOptIn.Columns.Length, [Is].EqualTo(If(columnIndex = 2, 2, 3)))
                    Assert.That(afterOptIn.Formula, Does.Contain(If(columnIndex = 2, "#REF!", "SalesTable[Region]")))
                    Return
                End If
                edit(workbook)
                workbook.SaveAs(nativeOutput, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                workbook.Close()

                Using editor As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    If operation = "DeleteFilterWithFormula" Then
                        Assert.That(ReadTableColumnState(nativeOutput).HiddenRows, [Is].EqualTo({4UI}))
                        Assert.Throws(Of NotSupportedException)(Sub() edit(editor))
                    Else
                        edit(editor)
                    End If
                    editor.SaveCopyAs(openXmlOutput, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using
                If operation = "DeleteFilterWithFormula" Then
                    Assert.That(File.ReadAllBytes(openXmlOutput), [Is].EqualTo(File.ReadAllBytes(input)))
                    Return
                End If
                Dim expected = ReadTableColumnState(nativeOutput)
                Dim actual = ReadTableColumnState(openXmlOutput)
                Select Case operation
                    Case "DeleteOneOfTwoFilters"
                        Assert.That(expected.HiddenRows, [Is].EqualTo({4UI}))
                    Case "DeleteFilterKeepAnd"
                        Assert.That(expected.HiddenRows, [Is].EqualTo({4UI, 5UI}))
                    Case "DeleteFilterKeepOr"
                        Assert.That(expected.HiddenRows, [Is].EqualTo({3UI, 6UI}))
                End Select
                Assert.That(actual.TableRange, [Is].EqualTo(expected.TableRange))
                Assert.That(actual.Columns, [Is].EqualTo(expected.Columns))
                Assert.That(actual.FilterRange, [Is].EqualTo(expected.FilterRange))
                Assert.That(actual.Filters, [Is].EqualTo(expected.Filters))
                Assert.That(actual.HiddenRows, [Is].EqualTo(expected.HiddenRows))
                Assert.That(actual.Formula, [Is].EqualTo(expected.Formula))
                Assert.That(actual.Layout, [Is].EqualTo(expected.Layout))
                workbook = Me.CreateInstance(openXmlOutput, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                Dim reopened = DirectCast(workbook.Workbook.Worksheets.Item("Data"), NativeExcel.Worksheet).ListObjects.Item(1)
                Assert.That(reopened.ShowHeaders, [Is].EqualTo(Not headerless))
                Assert.That(reopened.ListColumns.Count, [Is].EqualTo(expected.Columns.Length))
            Finally
                If workbook IsNot Nothing AndAlso Not workbook.IsClosed Then workbook.Close()
                Directory.Delete(directoryPath, True)
            End Try
        End Sub

        Private Shared Function ReadTableColumnState(path As String) As (TableRange As String, Columns As String(), FilterRange As String,
                                                                         Filters As String(), HiddenRows As UInteger(), Formula As String,
                                                                         Layout As (First As UInteger, Last As UInteger, Width As Double?,
                                                                                    Custom As Boolean?, Hidden As Boolean?)())
            Using document = SpreadsheetDocument.Open(path, False)
                Dim data = StructuralReferenceCoverage.WorksheetFor(document.WorkbookPart, "Data")
                Dim table = data.TableDefinitionParts.Single().Table
                Dim columns = table.GetFirstChild(Of TableColumns)().Elements(Of TableColumn)().
                    Select(Function(column) column.Id.Value & ":" & column.Name.Value).ToArray()
                Dim filter = table.GetFirstChild(Of AutoFilter)()
                Dim criteria = If(filter?.Elements(Of FilterColumn)().Select(
                    Function(field) field.ColumnId.Value & ":" & String.Concat(field.ChildElements.Select(Function(rule) rule.OuterXml))
                ).ToArray(), Array.Empty(Of String)())
                Dim hidden = data.Worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)().
                    Where(Function(row) If(row.Hidden?.Value, False)).Select(Function(row) row.RowIndex.Value).ToArray()
                Dim checks = StructuralReferenceCoverage.WorksheetFor(document.WorkbookPart, "Checks")
                Dim formula = checks.Worksheet.Descendants(Of Cell)().Single().CellFormula.Text
                Dim layout = data.Worksheet.Elements(Of Columns)().SelectMany(Function(group) group.Elements(Of Column)()).
                    Select(Function(column) (column.Min.Value, column.Max.Value, column.Width?.Value, column.CustomWidth?.Value, column.Hidden?.Value)).ToArray()
                Return (table.Reference.Value, columns, filter?.Reference?.Value, criteria, hidden, formula, layout)
            End Using
        End Function

    End Class

End Namespace
