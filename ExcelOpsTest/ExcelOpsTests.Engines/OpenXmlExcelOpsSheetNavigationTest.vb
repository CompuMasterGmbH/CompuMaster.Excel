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

End Class
