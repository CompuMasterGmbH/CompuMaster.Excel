Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Validation
Imports ChartXml = DocumentFormat.OpenXml.Drawing.Charts
Imports CompuMaster.Excel.Test
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    <TestCase(CopyWorksheetFormulaMode.ExcelCompatible)>
    <TestCase(CopyWorksheetFormulaMode.PreserveFormulas)>
    Public Sub WorksheetTransfer_OpenXmlCopiesWithTheRequestedFormulaMode(mode As CopyWorksheetFormulaMode)
        Dim sourcePath = Path.Combine(_directory, "source.xlsx")
        Dim destinationPath = Path.Combine(_directory, "target.xlsx")
        Dim output = Path.Combine(_directory, "copied.xlsx")
        CreateStructuralWorkbook(sourcePath, {ValueCell("B2", 20)},
            {FormulaCell("A1", DataQualifier & "$B$2"), FormulaCell("A2", "Check!A1"), FormulaCell("A3", "A1+N(""Check!A1"")")})
        CreateStructuralWorkbook(destinationPath, {ValueCell("B2", 70)}, {ValueCell("A1", 100)})
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Dim expectedCrossSheet = If(mode = CopyWorksheetFormulaMode.ExcelCompatible, "'[1]O''Brien Data'!$B$2", DataQualifier & "$B$2")
        Dim expectedSelf = If(mode = CopyWorksheetFormulaMode.ExcelCompatible, "'Check (2)'!A1", "Check!A1")
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options)
            Using target As New OpenXmlExcelDataOperations(destinationPath, OpenMode.OpenExistingFile, options)
                target.WriteCellValue(Of Integer)(DataSheet, 0, 0, 17)
                Dim assigned = source.CopySheet("Check", target, mode)
                Assert.That(assigned.TargetSheetName, [Is].EqualTo("Check (2)"))
                Assert.That(assigned.TargetSheetIndex, [Is].EqualTo(3))
                Assert.That(target.LookupCellFormula(assigned.TargetSheetName, 0, 0), [Is].EqualTo(expectedCrossSheet))
                Assert.That(target.LookupCellFormula(assigned.TargetSheetName, 1, 0), [Is].EqualTo(expectedSelf))
                Assert.That(target.LookupCellFormula(assigned.TargetSheetName, 2, 0), [Is].EqualTo("A1+N(""Check!A1"")"))
                Assert.That(target.LookupCellValue(Of Integer)(DataSheet, 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(17))
                Assert.That(source.LookupCellFormula("Check", 0, 0), [Is].EqualTo(DataQualifier & "$B$2"))
                Dim second = source.CopySheet("Check", target, mode)
                Assert.That(second.TargetSheetName, [Is].EqualTo("Check (3)"))
                Assert.That(target.LookupCellFormula(second.TargetSheetName, 0, 0), [Is].EqualTo(expectedCrossSheet))
                target.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using
        End Using
        Using saved = SpreadsheetDocument.Open(output, False)
            Dim links = saved.WorkbookPart.ExternalWorkbookParts.ToArray()
            If mode = CopyWorksheetFormulaMode.ExcelCompatible Then
                Assert.That(links.Single().ExternalRelationships.Single().Uri.LocalPath, [Is].EqualTo(sourcePath))
            Else
                Assert.That(links, [Is].Empty)
            End If
            Assert.That(New OpenXmlValidator().Validate(saved).Select(Function(issue) issue.Description), [Is].Empty)
        End Using
        Using reopened As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            Assert.That(reopened.LookupCellFormula("Check (2)", 0, 0), [Is].EqualTo(expectedCrossSheet))
            Assert.That(reopened.LookupCellFormula("Check (2)", 1, 0), [Is].EqualTo(expectedSelf))
        End Using
    End Sub

    <Test>
    Public Sub WorksheetTransfer_OpenXmlCopiesGroupsWithLocalReferences()
        Dim sourcePath = Path.Combine(_directory, "source.xlsx")
        Dim destinationPath = Path.Combine(_directory, "target.xlsx")
        CreateStructuralWorkbook(sourcePath, {ValueCell("B2", 20)}, {FormulaCell("A1", DataQualifier & "$B$2")})
        CreateStructuralWorkbook(destinationPath, {ValueCell("B2", 70)}, {ValueCell("A1", 100)})
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options)
            Using target As New OpenXmlExcelDataOperations(destinationPath, OpenMode.OpenExistingFile, options)
                Assert.That(source.CopySheets({"Check", DataSheet}, target).Select(Function(item) item.TargetSheetName), [Is].EqualTo({DataSheet & " (2)", "Check (2)"}))
                Assert.That(target.LookupCellFormula("Check (2)", 0, 0), [Is].EqualTo("'O''Brien Data (2)'!$B$2"))
                Assert.That(target.LookupCellValue(Of Integer)(DataSheet & " (2)", 1, 1, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(20))
            End Using
        End Using
    End Sub

    <Test>
    Public Sub WorksheetTransfer_OpenXmlRequiresASavedSourceOnlyWhenCreatingAnExternalLink()
        Dim input = Path.Combine(_directory, "source.xlsx")
        CreateStructuralWorkbook(input, {ValueCell("B2", 20)}, {FormulaCell("A1", DataQualifier & "$B$2")})
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(File.ReadAllBytes(input), options),
              target As New OpenXmlExcelDataOperations(Nothing, OpenMode.CreateFile, options)
            Assert.Throws(Of NotSupportedException)(Sub() source.CopySheet("Check", target))
            Assert.That(target.SheetNames(), [Is].EqualTo({"Sheet1"}))
            Dim copied = source.CopySheet("Check", target, CopyWorksheetFormulaMode.PreserveFormulas)
            Assert.That(target.LookupCellFormula(copied.TargetSheetName, 0, 0), [Is].EqualTo(DataQualifier & "$B$2"))
        End Using
    End Sub

    <TestCase(CopyWorksheetFormulaMode.ExcelCompatible, 0)>
    <TestCase(CopyWorksheetFormulaMode.ExcelCompatible, 3)>
    <TestCase(CopyWorksheetFormulaMode.PreserveFormulas, 1)>
    Public Sub WorksheetTransfer_OpenXmlInsertsMixedGroupsAtTheRequestedIndex(mode As CopyWorksheetFormulaMode, index As Integer)
        Dim input = TestFiles.TestFileChartSheet01.FullName
        Dim targetPath = Path.Combine(_directory, "target-with-charts.xlsx")
        Dim output = Path.Combine(_directory, "mixed-copy.xlsx")
        File.Copy(input, targetPath)
        Using document = SpreadsheetDocument.Open(targetPath, True)
            Dim book = document.WorkbookPart
            Dim tail = book.AddNewPart(Of WorksheetPart)()
            tail.Worksheet = New Worksheet(New SheetData())
            book.Workbook.Sheets.Append(New Sheet With {.SheetId = 4UI, .Name = "Tail", .Id = book.GetIdOfPart(tail)})
            book.Workbook.DefinedNames = New DefinedNames(New DefinedName("Tail!$A$1") With {.Name = "LocalValue", .LocalSheetId = 2UI})
            Dim view = book.Workbook.BookViews.GetFirstChild(Of WorkbookView)()
            view.ActiveTab = 2UI
            view.FirstSheet = 1UI
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Dim expectedOrder As New List(Of String) From {"data", "chart", "Tail"}
        expectedOrder.InsertRange(index, {"data (2)", "chart (2)"})
        Using source As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options),
              target As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, options)
            Assert.Throws(Of ArgumentOutOfRangeException)(Sub() source.CopySheets({"data", "chart"}, target, -1, mode))
            Assert.Throws(Of ArgumentOutOfRangeException)(Sub() source.CopySheets({"data", "chart"}, target, 4, mode))
            Assert.That(target.SheetNames(), [Is].EqualTo({"data", "chart", "Tail"}))

            Dim copied = source.CopySheets({"chart", "data"}, target, index, mode)
            Assert.That(copied.Select(Function(item) item.TargetSheetName), [Is].EqualTo({"data (2)", "chart (2)"}))
            Assert.That(copied.Select(Function(item) item.TargetSheetIndex), [Is].EqualTo({index, index + 1}))
            Assert.That(target.SheetNames(), [Is].EqualTo(expectedOrder))
            Assert.That(target.ChartSheetNames(), [Is].EquivalentTo({"chart", "chart (2)"}))
            Assert.That(target.SheetNames().Count, [Is].EqualTo(5))
            Assert.That(target.WorkSheetNames().Count, [Is].EqualTo(3))
            Assert.That(target.SheetIndex("chart (2)"), [Is].EqualTo(copied(1).TargetSheetIndex))
            Assert.That(target.ChartSheetIndex("chart (2)"), [Is].EqualTo(If(index <= 1, 0, 1)))
            Assert.That(source.SheetNames(), [Is].EqualTo({"data", "chart"}))
            target.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using saved = SpreadsheetDocument.Open(output, False)
            Dim book = saved.WorkbookPart
            Dim expectedScope = If(index <= 2, 4UI, 2UI)
            Assert.That(book.Workbook.DefinedNames.Elements(Of DefinedName)().Single().LocalSheetId.Value, [Is].EqualTo(expectedScope))
            Dim view = book.Workbook.BookViews.GetFirstChild(Of WorkbookView)()
            Assert.That(view.ActiveTab.Value, [Is].EqualTo(expectedScope))
            Assert.That(view.FirstSheet.Value, [Is].EqualTo(If(index <= 1, 3UI, 1UI)))
            Assert.That(book.Workbook.Sheets.Elements(Of Sheet)().Select(Function(sheet) sheet.Name.Value), [Is].EqualTo(expectedOrder))

            Dim prefix = If(mode = CopyWorksheetFormulaMode.ExcelCompatible, "'data (2)'!", "data!")
            Assert.That(CopiedChartFormulas(book, "chart (2)"), [Is].EqualTo({prefix & "$A$6", prefix & "$B$5:$I$5", prefix & "$B$6:$I$6"}))
            Assert.That(CopiedChartFormulas(book, "chart"), [Is].EqualTo({"data!$A$6", "data!$B$5:$I$5", "data!$B$6:$I$6"}))
            Assert.That(book.ExternalWorkbookParts, [Is].Empty)
        End Using
    End Sub

    <Test>
    Public Sub WorksheetTransfer_OpenXmlChartCopiedAloneLinksBackToItsUncopiedData()
        Dim input = TestFiles.TestFileChartSheet01.FullName
        Dim output = Path.Combine(_directory, "chart-only-copy.xlsx")
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options),
              target As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            Dim copied = source.CopySheet("chart", target, 1)
            Assert.That(copied.TargetSheetName, [Is].EqualTo("chart (2)"))
            Assert.That(copied.TargetSheetIndex, [Is].EqualTo(1))
            target.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using saved = SpreadsheetDocument.Open(output, False)
            Dim book = saved.WorkbookPart
            Assert.That(CopiedChartFormulas(book, "chart (2)"), [Is].EqualTo({"[1]data!$A$6", "[1]data!$B$5:$I$5", "[1]data!$B$6:$I$6"}))
            Assert.That(book.ExternalWorkbookParts.Single().ExternalRelationships.Single().Uri.LocalPath, [Is].EqualTo(input))
        End Using
    End Sub

    Private Shared Function CopiedChartFormulas(book As WorkbookPart, name As String) As String()
        Dim sheet = book.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = name)
        Dim part = DirectCast(book.GetPartById(sheet.Id.Value), ChartsheetPart)
        Return part.DrawingsPart.ChartParts.Single().ChartSpace.Descendants(Of ChartXml.Formula)().
            Select(Function(formula) formula.Text).ToArray()
    End Function


    <TestCase(False)>
    <TestCase(True)>
    Public Sub WorksheetTransfer_OpenXmlRemapsDistinctWorkbookStyleTables(move As Boolean)
        Dim sourcePath = Path.Combine(_directory, "styled-source.xlsx")
        Dim targetPath = Path.Combine(_directory, "styled-target.xlsx")
        Dim output = Path.Combine(_directory, "styled-result.xlsx")
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(sourcePath))
            package.Workbook.Worksheets.Add("Retained")
            Dim sheet = package.Workbook.Worksheets.Add("Transfer")
            sheet.Cells("A1").Value = 1.25
            sheet.Cells("A1").Style.Numberformat.Format = "0.00 ""source"""
            sheet.Cells("A1").Style.Font.Bold = True
            sheet.Cells("A1").Style.Border.Bottom.Style = CompuMaster.Epplus4.Style.ExcelBorderStyle.Thin
            sheet.Row(2).Style.Numberformat.Format = "0.000 ""row"""
            sheet.Column(3).Style.Numberformat.Format = "0.0000 ""column"""
            package.Save()
        End Using
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(targetPath))
            Dim sheet = package.Workbook.Worksheets.Add("Original")
            sheet.Cells("A1").Value = 2.5
            sheet.Cells("A1").Style.Numberformat.Format = "0.0 ""target"""
            sheet.Cells("A1").Style.Font.Italic = True
            package.Save()
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = True}
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options),
              destination As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, options)
            If move Then
                source.MoveSheet("Transfer", destination)
                Assert.That(source.WorkSheetNames(), [Is].EqualTo({"Retained"}))
            Else
                source.CopySheet("Transfer", destination)
                Assert.That(source.WorkSheetNames(), [Is].EqualTo({"Retained", "Transfer"}))
            End If
            destination.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using destination As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            Assert.That(destination.LookupCellFormat("Original", 0, 0), [Is].EqualTo("0.0 ""target"""))
            Assert.That(destination.LookupCellFormat("Transfer", 0, 0), [Is].EqualTo("0.00 ""source"""))
            Assert.That(destination.LookupCellFormat("Transfer", 1, 1), [Is].EqualTo("0.000 ""row"""))
            Assert.That(destination.LookupCellFormat("Transfer", 4, 2), [Is].EqualTo("0.0000 ""column"""))
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim book = document.WorkbookPart
            Dim styles = book.WorkbookStylesPart.Stylesheet
            Dim sheet = book.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = "Transfer")
            Dim cell = DirectCast(book.GetPartById(sheet.Id.Value), WorksheetPart).Worksheet.Descendants(Of Cell)().Single(Function(item) item.CellReference.Value = "A1")
            Dim format = styles.CellFormats.Elements(Of CellFormat)().ElementAt(CInt(cell.StyleIndex.Value))
            Assert.That(styles.Fonts.Elements(Of Font)().ElementAt(CInt(format.FontId.Value)).Bold, [Is].Not.Null)
            Assert.That(styles.Borders.Elements(Of Border)().ElementAt(CInt(format.BorderId.Value)).BottomBorder.Style.Value,
                        [Is].EqualTo(BorderStyleValues.Thin))
        End Using
    End Sub


    <TestCase(CopySheetOption.TargetSheetMustExist)>
    <TestCase(CopySheetOption.TargetSheetMustNotExist)>
    Public Sub WorksheetTransfer_OpenXmlCopiesLegacyContentWithDistinctSheetNames(copyOption As CopySheetOption)
        Dim sourcePath = Path.Combine(_directory, "legacy-source.xlsx")
        Dim targetPath = Path.Combine(_directory, "legacy-target.xlsx")
        Dim output = Path.Combine(_directory, "legacy-result.xlsx")
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(sourcePath))
            Dim sheet = package.Workbook.Worksheets.Add("Inputs")
            sheet.Cells("A1").Value = "Transferred text"
            sheet.Cells("A1:B1").Merge = True
            sheet.Cells("A2").Value = 12
            sheet.Cells("A2").Style.Numberformat.Format = "0.000 ""source"""
            sheet.Cells("C3").Formula = "A2+1"
            package.Save()
        End Using
        Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(targetPath))
            Dim unrelated = package.Workbook.Worksheets.Add("Inputs")
            unrelated.Cells("A1").Value = "Unrelated text"
            unrelated.Cells("A2").Style.Numberformat.Format = "0.0 ""target"""
            If copyOption = CopySheetOption.TargetSheetMustExist Then
                package.Workbook.Worksheets.Add("Result").Cells("F20").Value = "Old content"
            End If
            package.Save()
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableInitialCalculation = True, .DisableCalculationEngine = False
        }
        Const cached As FormulaCalculationMode = FormulaCalculationMode.UseCachedValueOnly
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options),
              destination As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, options)
            source.CopySheetContent("Inputs", destination, "Result", copyOption)
            Assert.Throws(Of InvalidOperationException)(Sub() source.CopySheetContent("Inputs", destination, "Result", CopySheetOption.TargetSheetMustNotExist))
            destination.RecalculateAll()
            destination.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using destination As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            Assert.That(destination.LookupCellValue(Of String)("Inputs", 0, 0, cached), [Is].EqualTo("Unrelated text"))
            Assert.That(destination.LookupCellValue(Of String)("Result", 0, 0, cached), [Is].EqualTo("Transferred text"))
            Assert.That(destination.IsMergedCell("Result", 0, 1), [Is].True)
            Assert.That(destination.LookupCellFormat("Result", 1, 0), [Is].EqualTo("0.000 ""source"""))
            Assert.That(destination.LookupCellFormula("Result", 2, 2), [Is].EqualTo("A2+1"))
            Assert.That(destination.LookupCellValue(Of Integer)("Result", 2, 2, cached), [Is].EqualTo(13))
            Assert.That(destination.LookupCellValueAsObject("Result", 19, 5, cached), [Is].Null)
        End Using
    End Sub


    <Test>
    Public Sub WorksheetTransfer_OpenXmlPreservesThemeColorsAndWorksheetRelatedParts()
        Dim sourcePath = Path.Combine(_directory, "source-theme.xlsx")
        Dim targetPath = Path.Combine(_directory, "target-theme.xlsx")
        Dim output = Path.Combine(_directory, "theme-result.xlsx")
        Dim themeFixture = TestEnvironment.FullPathOfExistingTestFile("test_data", "HtmlExportThemeExcelOffice.xlsx")
        Dim settingsBytes As Byte() = {5, 12, 7, 31}
        For Each item In {(sourcePath, "Imported", "446699"), (targetPath, "Original", "AA2244")}
            Using package As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(item.Item1))
                Dim sheet = package.Workbook.Worksheets.Add(item.Item2)
                sheet.Cells("A1").Value = 9
                sheet.Cells("A1").Style.Font.Bold = True
                package.Save()
            End Using
            Using template = SpreadsheetDocument.Open(themeFixture, False), doc = SpreadsheetDocument.Open(item.Item1, True)
                Dim book = doc.WorkbookPart
                If book.ThemePart IsNot Nothing Then book.DeletePart(book.ThemePart)
                Dim theme = book.AddPart(template.WorkbookPart.ThemePart).Theme
                Dim accent = theme.ThemeElements.ColorScheme.Accent1Color
                accent.RemoveAllChildren()
                accent.AppendChild(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = item.Item3})
                Dim worksheetPart = book.WorksheetParts.Single()
                Dim cell = worksheetPart.Worksheet.Descendants(Of Cell)().Single()
                Dim styles = book.WorkbookStylesPart.Stylesheet
                Dim format = styles.CellFormats.Elements(Of CellFormat)().ElementAt(CInt(cell.StyleIndex.Value))
                styles.Fonts.Elements(Of Font)().ElementAt(CInt(format.FontId.Value)).Color = New Color() With {.Theme = 4UI}
                If item.Item2 = "Imported" Then
                    Dim settings = worksheetPart.AddNewPart(Of SpreadsheetPrinterSettingsPart)()
                    Using data As New MemoryStream(settingsBytes)
                        settings.FeedData(data)
                    End Using
                    worksheetPart.Worksheet.AddChild(New PageSetup With {.Id = worksheetPart.GetIdOfPart(settings)})
                    worksheetPart.Worksheet.AddChild(New PhoneticProperties With {.FontId = format.FontId.Value})
                End If
            End Using
        Next

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = True}
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options),
              destination As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, options)
            source.CopySheet("Imported", destination)
            destination.SaveCopyAs(output, False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using doc = SpreadsheetDocument.Open(output, False)
            Dim book = doc.WorkbookPart
            Dim styles = book.WorkbookStylesPart.Stylesheet
            Dim sheets = book.Workbook.Sheets.Elements(Of Sheet)().ToDictionary(Function(sheet) sheet.Name.Value)
            Dim copied = DirectCast(book.GetPartById(sheets("Imported").Id.Value), WorksheetPart)
            Dim original = DirectCast(book.GetPartById(sheets("Original").Id.Value), WorksheetPart)
            Dim format = styles.CellFormats.Elements(Of CellFormat)().ElementAt(CInt(copied.Worksheet.Descendants(Of Cell)().Single().StyleIndex.Value))
            Dim importedFont = styles.Fonts.Elements(Of Font)().ElementAt(CInt(format.FontId.Value))
            Assert.That(importedFont.Color.Rgb.Value, [Is].EqualTo("FF446699"))
            Assert.That(copied.Worksheet.GetFirstChild(Of PhoneticProperties)().FontId.Value, [Is].EqualTo(format.FontId.Value))
            Dim originalFormat = styles.CellFormats.Elements(Of CellFormat)().ElementAt(CInt(original.Worksheet.Descendants(Of Cell)().Single().StyleIndex.Value))
            Assert.That(styles.Fonts.Elements(Of Font)().ElementAt(CInt(originalFormat.FontId.Value)).Color.Theme.Value, [Is].EqualTo(4UI))
            Dim settings = DirectCast(copied.GetPartById(copied.Worksheet.GetFirstChild(Of PageSetup)().Id.Value), SpreadsheetPrinterSettingsPart)
            Assert.That(ReadBytes(settings.GetStream()), [Is].EqualTo(settingsBytes))
        End Using
    End Sub


    <Test>
    Public Sub WorksheetTransfer_OpenXmlRejectsMovingTheLastVisibleSheetAtomically()
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = True}
        Using source As New OpenXmlExcelDataOperations(Nothing, OpenMode.CreateFile, options),
              destination As New OpenXmlExcelDataOperations(Nothing, OpenMode.CreateFile, options)
            Dim sourceSheet = source.WorkSheetNames().Single()
            Dim targetSheet = destination.WorkSheetNames().Single()
            source.WriteCellValue(sourceSheet, 0, 0, 42)
            Assert.Throws(Of NotSupportedException)(Sub() source.MoveSheet(sourceSheet, destination))
            Assert.That(source.WorkSheetNames(), [Is].EqualTo({sourceSheet}))
            Assert.That(destination.WorkSheetNames(), [Is].EqualTo({targetSheet}))
            Assert.That(source.LookupCellValue(Of Integer)(sourceSheet, 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(42))
        End Using
    End Sub

End Class
