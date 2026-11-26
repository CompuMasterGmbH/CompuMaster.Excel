Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework
Imports NativeExcel = Microsoft.Office.Interop.Excel

Namespace ExcelOpsTests.Engines
    Partial Public Class MsExcelOpsTest

        <TestCase("ClearContents", False)>
        <TestCase("ClearContents", True)>
        <TestCase("ClearAll", False)>
        <TestCase("DeleteCell", False)>
        <TestCase("DeleteRow", False)>
        <TestCase("DeleteColumn", False)>
        Public Sub LookupLastCell_ExcelShrinksOnSaveOrWhenUsedRangeIsRead(operation As String, readUsedRange As Boolean)
            Dim directory = CreateChartTestDirectory()
            Dim path = IO.Path.Combine(directory, "excel-last-cell.xlsx")
            Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile)
            Dim workbook = CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile, options)
            Try
                Dim sheet = DirectCast(workbook.Workbook.Worksheets.Item(1), NativeExcel.Worksheet)
                Assert.That(NativeLastCell(sheet), [Is].EqualTo("A1"))
                sheet.Range("B2").Value2 = "base"
                sheet.Range("K30").Value2 = "edge"
                Assert.That(NativeLastCell(sheet), [Is].EqualTo("K30"))
                workbook.SaveAs(path, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)

                Select Case operation
                    Case "ClearContents"
                        sheet.Range("K30").ClearContents()
                    Case "ClearAll"
                        sheet.Range("K30").Clear()
                    Case "DeleteCell"
                        sheet.Range("K30").Delete(NativeExcel.XlDeleteShiftDirection.xlShiftUp)
                    Case "DeleteRow"
                        sheet.Range("K30").EntireRow.Delete()
                    Case "DeleteColumn"
                        sheet.Range("K30").EntireColumn.Delete()
                End Select
                Assert.That(NativeLastCell(sheet), [Is].EqualTo("K30"))
                If readUsedRange Then
                    ' LookupLastCell reads UsedRange first. Excel refreshes SpecialCells as a side effect.
                    Assert.That(workbook.LookupLastCell(sheet.Name).Address, [Is].EqualTo("B2"))
                    Assert.That(NativeLastCell(sheet), [Is].EqualTo("B2"))
                End If

                workbook.Save(ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(NativeLastCell(sheet), [Is].EqualTo("B2"))
                Assert.That(workbook.LookupLastCell(sheet.Name).Address, [Is].EqualTo("B2"))
            Finally
                workbook.Close()
                IO.Directory.Delete(directory, recursive:=True)
            End Try
        End Sub

        <TestCase("NumberFormat", "K30")>
        <TestCase("EmptyFormula", "K30")>
        <TestCase("Merge", "K30")>
        <TestCase("RowHeight", "B30")>
        <TestCase("HiddenRow", "B30")>
        <TestCase("RowFormat", "B30")>
        <TestCase("ColumnWidth", "B2")>
        <TestCase("HiddenColumn", "B2")>
        <TestCase("ColumnFormat", "K2")>
        <TestCase("TrailingColumnFormat", "J2")>
        <TestCase("LastColumnFormat", "XFD2")>
        <TestCase("DefaultSheetFormat", "B2")>
        <TestCase("Note", "K30")>
        <TestCase("Validation", "B2")>
        <TestCase("Hyperlink", "K30")>
        Public Sub LookupLastCell_MatchesExcelForSavedWorkbook(feature As String, expected As String)
            Dim directory = CreateChartTestDirectory()
            Dim path = IO.Path.Combine(directory, "excel-used-range.xlsx")
            Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile)
            Dim workbook = CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile, options)
            Try
                Dim sheetName As String
                Dim sheet = DirectCast(workbook.Workbook.Worksheets.Item(1), NativeExcel.Worksheet)
                sheetName = sheet.Name
                sheet.Range("B2").Value2 = "base"
                Select Case feature
                    Case "NumberFormat"
                        sheet.Range("K30").NumberFormat = "0.00"
                    Case "EmptyFormula"
                        sheet.Range("K30").Formula = "="""""
                    Case "Merge"
                        sheet.Range("J20:K30").Merge()
                    Case "RowHeight"
                        sheet.Range("K30").EntireRow.RowHeight = 27
                    Case "HiddenRow"
                        sheet.Range("K30").EntireRow.Hidden = True
                    Case "RowFormat"
                        sheet.Range("K30").EntireRow.Font.Bold = True
                    Case "ColumnWidth"
                        sheet.Range("K30").EntireColumn.ColumnWidth = 20
                    Case "HiddenColumn"
                        sheet.Range("K30").EntireColumn.Hidden = True
                    Case "ColumnFormat"
                        sheet.Range("K30").EntireColumn.Font.Bold = True
                    Case "TrailingColumnFormat"
                        sheet.Range("K:XFD").Font.Bold = True
                    Case "LastColumnFormat"
                        sheet.Range("XFD:XFD").Font.Bold = True
                    Case "DefaultSheetFormat"
                        sheet.Cells.Font.Bold = True
                    Case "Note"
                        sheet.Range("K30").AddComment("worksheet note")
                    Case "Validation"
                        sheet.Range("K29:K30").Validation.Add(NativeExcel.XlDVType.xlValidateList, Formula1:="one,two")
                    Case "Hyperlink"
                        sheet.Hyperlinks.Add(sheet.Range("K30"), "", "B2")
                End Select

                workbook.SaveAs(path, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(workbook.LookupLastCell(sheetName).Address, [Is].EqualTo(expected))
                workbook.Close()
                Using openXml As New OpenXmlExcelDataOperations(path, ExcelDataOperationsBase.OpenMode.OpenExistingFile,
                    New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                    Assert.That(openXml.LookupLastCell(sheetName).Address, [Is].EqualTo(expected))
                End Using
            Finally
                workbook.Close()
                IO.Directory.Delete(directory, recursive:=True)
            End Try
        End Sub

        Private Shared Function NativeLastCell(sheet As NativeExcel.Worksheet) As String
            Return sheet.Cells.SpecialCells(NativeExcel.XlCellType.xlCellTypeLastCell).Address(False, False)
        End Function

    End Class
End Namespace
