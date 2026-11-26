Imports System.Reflection
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    <TestCase("ClearContents")>
    <TestCase("RemoveCells")>
    <TestCase("RemoveRows")>
    <TestCase("RemoveColumns")>
    Public Sub LookupLastCell_TracksUnsavedWorksheetEdits(operation As String)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.DefaultBehaviourOnCreateFile)
        Dim path = IO.Path.Combine(_directory, "last-cell.xlsx")
        Using workbook As New OpenXmlExcelDataOperations(path, OpenMode.CreateFile, options)
            Dim sheetName = workbook.WorkSheetNames().Single()
            Assert.That(workbook.LookupLastCell(sheetName).Address, [Is].EqualTo("A1"))
            workbook.WriteCellValue(sheetName, 1, 1, 1)
            workbook.WriteCellValue(sheetName, 29, 10, 2)
            Assert.That(workbook.LookupLastCell(sheetName).Address, [Is].EqualTo("K30"))

            Select Case operation
                Case "ClearContents"
                    workbook.ClearCellContent(New ExcelCell(sheetName, "K30", ExcelCell.ValueTypes.All))
                Case "RemoveCells"
                    workbook.RemoveCells(sheetName, 29, 10, 1, 1, CellRemoveShiftDirection.ShiftCellsUp, True)
                Case "RemoveRows"
                    workbook.RemoveRows(sheetName, 29, 1, True)
                Case "RemoveColumns"
                    workbook.RemoveColumns(sheetName, 10, 1, True)
            End Select

            Assert.That(workbook.LookupLastCell(sheetName).Address, [Is].EqualTo("B2"))
            Assert.That(workbook.LookupLastRowIndex(sheetName), [Is].EqualTo(1))
            Assert.That(workbook.LookupLastColumnIndex(sheetName), [Is].EqualTo(1))
            workbook.SaveAs(path, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using workbook As New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.That(workbook.LookupLastCell(workbook.WorkSheetNames().Single()).Address, [Is].EqualTo("B2"))
        End Using
    End Sub

    ' A stale dimension and physical blank cells may remain in files produced by another engine.
    <Test>
    Public Sub LookupLastCell_IgnoresBlankRecordsAndCountsOrthogonalFormattedBounds()
        Dim path = IO.Path.Combine(_directory, "stale-dimension.xlsx")
        CreateTemplate(path)
        Using document = SpreadsheetDocument.Open(path, True)
            Dim worksheet = DirectCast(document.WorkbookPart.GetPartById("rId2"), WorksheetPart).Worksheet
            Dim data = worksheet.GetFirstChild(Of SheetData)()
            data.RemoveAllChildren(Of Row)()
            data.Append(
                New Row(New Cell With {.CellReference = "A1", .CellValue = New CellValue("8")}) With {.RowIndex = 1UI},
                New Row(New Cell With {.CellReference = "K2", .StyleIndex = 1UI}) With {.RowIndex = 2UI},
                New Row(New Cell With {.CellReference = "A8", .CellFormula = New CellFormula("""""")}) With {.RowIndex = 8UI},
                New Row(New Cell With {.CellReference = "Z200", .StyleIndex = 0UI}) With {.RowIndex = 200UI})
            If worksheet.SheetDimension IsNot Nothing Then
                worksheet.SheetDimension.Reference = "A1:Z200"
            Else
                worksheet.InsertBefore(New SheetDimension With {.Reference = "A1:Z200"}, worksheet.FirstChild)
            End If
        End Using

        Using workbook As New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.That(workbook.LookupLastCell("Notes").Address, [Is].EqualTo("K8"))
            workbook.WriteCellValue(Of String)("Notes", 7, 0, Nothing)
            Assert.That(workbook.LookupLastCell("Notes").Address, [Is].EqualTo("K2"))
            workbook.WriteCellValue("Notes", 1, 10, "temporary contents")
            workbook.WriteCellValue(Of String)("Notes", 1, 10, Nothing)
            Assert.That(workbook.LookupLastCell("Notes").Address, [Is].EqualTo("K2"))
        End Using
    End Sub

    ' Inject lookups during the actual recalculation scope without adding a production test callback.
    <TestCase(False)>
    <TestCase(True)>
    Public Sub LookupLastCell_ReusesBoundsPerSheetOnlyWithinARecalculation(fail As Boolean)
        Dim path = IO.Path.Combine(_directory, "recalculation-bounds.xlsx")
        CreateTemplate(path)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False}
        Using workbook As New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, options)
            workbook.WriteCellValue("Inputs", 29, 10, 1)
            workbook.WriteCellValue("Notes", 39, 11, 1)
            Dim recalculate = GetType(OpenXmlExcelDataOperations).GetMethod("RecalculateSheets", BindingFlags.Instance Or BindingFlags.NonPublic)
            Dim sheets = CheckLastCellsDuringRecalculation(workbook, fail)
            If fail Then
                Dim failure = Assert.Throws(Of TargetInvocationException)(Sub() recalculate.Invoke(workbook, {sheets}))
                Assert.That(failure.InnerException, [Is].TypeOf(Of InvalidOperationException)())
            Else
                recalculate.Invoke(workbook, {sheets})
            End If
            Assert.That(workbook.LookupLastCell("Inputs").Address, [Is].EqualTo("M50"))
            Assert.That(workbook.LookupLastCell("Notes").Address, [Is].EqualTo("N60"))
        End Using
    End Sub

    Private Shared Iterator Function CheckLastCellsDuringRecalculation(workbook As OpenXmlExcelDataOperations, fail As Boolean) As IEnumerable(Of String)
        Dim inputs = workbook.LookupLastCell("Inputs")
        Assert.That(inputs.Address, [Is].EqualTo("K30"))
        Assert.That(workbook.LookupLastCell("Notes").Address, [Is].EqualTo("L40"))
        inputs.SheetName = "Changed by caller"
        workbook.WriteCellValue("Inputs", 49, 12, 1)
        workbook.WriteCellValue("Notes", 59, 13, 1)
        Assert.That(workbook.LookupLastCell("inputs").Address, [Is].EqualTo("K30"))
        Assert.That(workbook.LookupLastCell("inputs").SheetName, [Is].EqualTo("inputs"))
        Assert.That(workbook.LookupLastCell("Notes").Address, [Is].EqualTo("L40"))
        Yield "Inputs"
        Assert.That(workbook.LookupLastCell("Notes").Address, [Is].EqualTo("L40"))
        If fail Then Throw New InvalidOperationException("Interrupted recalculation.")
        Yield "Notes"
    End Function

End Class
