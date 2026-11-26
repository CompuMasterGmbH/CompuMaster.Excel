Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    <Test>
    Public Sub SavesACellSelectionInAFrozenPaneWithoutActivatingItsSheet()
        Dim input = Path.Combine(_directory, "cell-selection.xlsx")
        Dim output = Path.Combine(_directory, "cell-selection-out.xlsx")
        CreateTemplate(input)
        Using document = SpreadsheetDocument.Open(input, True)
            Dim part = DirectCast(document.WorkbookPart.GetPartById("rId2"), WorksheetPart)
            part.Worksheet.SheetViews = New SheetViews(New SheetView(
                New Pane With {.State = PaneStateValues.Frozen, .HorizontalSplit = 2.0, .VerticalSplit = 2.0, .TopLeftCell = "C3",
                               .ActivePane = PaneValues.TopLeft},
                New Selection With {.Pane = PaneValues.TopRight, .ActiveCell = "C1"}) With {.WorkbookViewId = 0UI})
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = True}
        Using workbook As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            Assert.That(workbook.SelectedSheetName(), [Is].EqualTo("Inputs"))
            workbook.SelectCell(New ExcelCell("Notes", "D4", ExcelCell.ValueTypes.All))
            Assert.That(workbook.SelectedSheetName(), [Is].EqualTo("Inputs"))
            workbook.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using document = SpreadsheetDocument.Open(output, False)
            Dim part = DirectCast(document.WorkbookPart.GetPartById("rId2"), WorksheetPart)
            Dim view = part.Worksheet.SheetViews.Elements(Of SheetView)().Single()
            Assert.That(view.GetFirstChild(Of Pane)().ActivePane.Value, [Is].EqualTo(PaneValues.BottomRight))
            Dim bottomRight = view.Elements(Of Selection)().Single(Function(s) s.Pane.Value = PaneValues.BottomRight)
            Assert.That(bottomRight.ActiveCell.Value, [Is].EqualTo("D4"))
            Assert.That(bottomRight.SequenceOfReferences.InnerText, [Is].EqualTo("D4"))
            Assert.That(view.Elements(Of Selection)().Single(Function(s) s.Pane.Value = PaneValues.TopRight).ActiveCell.Value, [Is].EqualTo("C1"))
        End Using
    End Sub

    <Test>
    Public Sub CleansUpBrokenNamesAndPreservesUsableDefinitions()
        Dim input = Path.Combine(_directory, "defined-names.xlsx")
        Dim output = Path.Combine(_directory, "defined-names-out.xlsx")
        CreateTemplate(input)
        Using document = SpreadsheetDocument.Open(input, True)
            Dim names = document.WorkbookPart.Workbook.DefinedNames
            names.Append(New DefinedName("#REF!$C$1") With {.Name = "Broken"},
                         New DefinedName("$A$1") With {.Name = "InvalidScope", .LocalSheetId = 5UI},
                         New DefinedName("$A$1") With {.Name = "ValidScope", .LocalSheetId = 1UI},
                         New DefinedName("""#REF!""") With {.Name = "TextConstant"},
                         New DefinedName("'#REF! labels'!$A$1") With {.Name = "QuotedSheetName"})
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = True}
        Using workbook As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            workbook.CleanupRangeNames()
            workbook.CleanupRangeNames()
            workbook.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim names = document.WorkbookPart.Workbook.DefinedNames.Elements(Of DefinedName)()
            Assert.That(names.Select(Function(name) name.Name.Value),
                        [Is].EqualTo({"InputLabel", "ValidScope", "TextConstant", "QuotedSheetName"}))
        End Using
    End Sub

    <Test>
    Public Sub AddsAndRemovesSheetsWithExcelReferenceAndViewUpdates()
        Dim input = Path.Combine(_directory, "sheet-lifecycle.xlsx")
        Dim output = Path.Combine(_directory, "sheet-lifecycle-out.xlsx")
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(input))
            Dim kept = package.Workbook.Worksheets.Add("Kept")
            Dim removed = package.Workbook.Worksheets.Add("Removed")
            package.Workbook.Worksheets.Add("Tail")
            removed.Cells("A1").Value = 12
            kept.Cells("A1").Formula = "Removed!A1"
            kept.Cells("A2").Formula = "SUM(Removed!A1:B2)"
            kept.Cells("A3").Formula = "INDIRECT(""Removed!A1"")"
            package.Workbook.Names.Add("RemovedReference", removed.Cells("A1"))
            package.Save()
        End Using
        Using document = SpreadsheetDocument.Open(input, True)
            document.WorkbookPart.Workbook.DefinedNames.Append(New DefinedName("$A$1") With {.Name = "TailLocal", .LocalSheetId = 2UI})
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False, .DisableInitialCalculation = True}
        Using workbook As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            workbook.AddSheet("Inserted", "Removed")
            Assert.That(workbook.SheetNames(), [Is].EqualTo({"Kept", "Inserted", "Removed", "Tail"}))
            Assert.That(workbook.SelectedSheetName(), [Is].EqualTo("Inserted"))
            Assert.Throws(Of ArgumentException)(Sub() workbook.AddSheet("iNsErTeD", "Removed"))
            workbook.WriteCellValue("Inserted", 0, 0, 7)
            workbook.SelectSheet("Removed")
            workbook.RemoveSheet("Removed")
            Assert.That(workbook.SheetNames(), [Is].EqualTo({"Kept", "Inserted", "Tail"}))
            Assert.That(workbook.SelectedSheetName(), [Is].EqualTo("Tail"))
            Assert.That(workbook.LookupCellFormula(New ExcelCell("Kept", "A1", ExcelCell.ValueTypes.All)), [Is].EqualTo("#REF!A1"))
            Assert.That(workbook.LookupCellFormula(New ExcelCell("Kept", "A2", ExcelCell.ValueTypes.All)), [Is].EqualTo("SUM(#REF!A1:B2)"))
            Assert.That(workbook.LookupCellFormula(New ExcelCell("Kept", "A3", ExcelCell.ValueTypes.All)), [Is].EqualTo("INDIRECT(""Removed!A1"")"))
            Assert.That(workbook.LookupCellValue(Of Integer)("Inserted", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(7))
            workbook.RecalculateAll()
            Assert.That(workbook.LookupCellValueAsObject("Kept", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("#REF!"))
            workbook.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim book = document.WorkbookPart
            Dim names = book.Workbook.DefinedNames.Elements(Of DefinedName)().ToDictionary(Function(item) item.Name.Value)
            Assert.That(names("RemovedReference").Text, [Is].EqualTo("#REF!$A$1"))
            Assert.That(names("TailLocal").LocalSheetId.Value, [Is].EqualTo(2UI))
            Assert.That(book.WorksheetParts.Count(), [Is].EqualTo(3))
        End Using
    End Sub


End Class
