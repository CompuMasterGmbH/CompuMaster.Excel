Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Security.Cryptography
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Validation
Imports NUnit.Framework
Imports Spreadsheet = DocumentFormat.OpenXml.Spreadsheet
Imports Drawing = DocumentFormat.OpenXml.Drawing
Imports SpreadsheetDrawing = DocumentFormat.OpenXml.Drawing.Spreadsheet

<TestFixture>
Partial Public Class OpenXmlExcelOpsTest

    Private _directory As String

    <SetUp>
    Public Sub CreateTestDirectory()
        _directory = Path.Combine(Path.GetTempPath(), "ExcelOps-OpenXml-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_directory)
    End Sub

    <TearDown>
    Public Sub RemoveTestDirectory()
        If Directory.Exists(_directory) Then Directory.Delete(_directory, recursive:=True)
    End Sub

    ' Compare the actual package parts, including both links in the worksheet-to-image chain.
    <Test>
    Public Sub SaveCopyPreservesTemplatePartsAndSource()
        Dim template = Path.Combine(_directory, "template.xlsx")
        Dim output = Path.Combine(_directory, "edited.xlsx")
        CreateTemplate(template)
        Dim originalHash = HashFile(template)
        Dim text = "  =example _x0041_ " & ChrW(1) & vbCrLf & " text  "
        Using engine As New OpenXmlExcelDataOperations(template, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Dim workbook As ExcelDataOperationsBase = engine
            Assert.That(workbook.SheetNames(), [Is].EqualTo({"Inputs", "Notes", "Hidden"}))
            Assert.That(workbook.IsProtectedSheet("Inputs"), [Is].True)
            Assert.That(workbook.IsHiddenSheet("Hidden"), [Is].True)
            Assert.That(workbook.LookupCellValue(Of String)("inputs", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("Template text"))
            Assert.That(workbook.LookupCellValue(Of Integer)("Inputs", 0, 2, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(4))
            Assert.That(workbook.LookupCellIsLocked("Inputs", 0, 0), [Is].False)
            Assert.That(workbook.LookupCellIsLocked("Inputs", 0, 2), [Is].True)

            workbook.WriteCellValue("Inputs", 0, 0, text)
            workbook.WriteCellValue("Inputs", 0, 1, 12.5D)
            workbook.WriteCellValue(New ExcelCell("Inputs", "AA1", ExcelCell.ValueTypes.All), True)
            workbook.WriteCellValue("Inputs", 3, 2, 7)
            workbook.WriteCellValue(Of Object)("Inputs", 3, 2, Nothing)
            workbook.ProtectSheet("Inputs", ProtectionLevel.StandardWithInsertDeleteRows)
            workbook.UnhideSheet("Hidden")
            workbook.HideSheet("Notes", stronglyHide:=True)
            engine.RequestRecalculationOnOpen()
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
            Assert.That(workbook.FilePath, [Is].EqualTo(template))
            Assert.That(workbook.ReadOnly, [Is].True)
        End Using

        Assert.That(HashFile(template), [Is].EqualTo(originalHash))
        Using before As New ZipArchive(File.OpenRead(template), ZipArchiveMode.Read),
              after As New ZipArchive(File.OpenRead(output), ZipArchiveMode.Read)
            Assert.That(after.Entries.Select(Function(entry) entry.FullName),
                        [Is].EquivalentTo(before.Entries.Select(Function(entry) entry.FullName)))
            For Each entry In before.Entries
                If entry.FullName = "xl/workbook.xml" OrElse entry.FullName = "xl/worksheets/sheet1.xml" Then Continue For
                Assert.That(ReadBytes(after.GetEntry(entry.FullName).Open()), [Is].EqualTo(ReadBytes(entry.Open())), entry.FullName)
            Next
        End Using

        Using before = SpreadsheetDocument.Open(template, False), after = SpreadsheetDocument.Open(output, False)
            Dim originalWorkbook = before.WorkbookPart.Workbook
            Dim workbook = after.WorkbookPart.Workbook
            Assert.That(workbook.DefinedNames.OuterXml, [Is].EqualTo(originalWorkbook.DefinedNames.OuterXml))
            Assert.That(workbook.BookViews.OuterXml, [Is].EqualTo(originalWorkbook.BookViews.OuterXml))
            Assert.That(workbook.Sheets.Elements(Of Sheet)().Select(Function(sheet) sheet.State?.Value),
                        [Is].EqualTo(New SheetStateValues?() {Nothing, SheetStateValues.VeryHidden, SheetStateValues.Visible}))
            Assert.That(workbook.CalculationProperties.CalculationMode.Value, [Is].EqualTo(CalculateModeValues.Auto))
            Assert.That(workbook.CalculationProperties.FullCalculationOnLoad.Value, [Is].True)
            Dim part = DirectCast(after.WorkbookPart.GetPartById("rId1"), WorksheetPart)
            Dim originalPart = DirectCast(before.WorkbookPart.GetPartById("rId1"), WorksheetPart)
            Assert.That(part.Worksheet.SheetViews.OuterXml, [Is].EqualTo(originalPart.Worksheet.SheetViews.OuterXml))
            Assert.That(part.Worksheet.GetFirstChild(Of PageMargins)().OuterXml,
                        [Is].EqualTo(originalPart.Worksheet.GetFirstChild(Of PageMargins)().OuterXml))
            Assert.That(part.Worksheet.SheetDimension.Reference.Value, [Is].EqualTo("A1:AA4"))
            Dim protection = part.Worksheet.GetFirstChild(Of SheetProtection)()
            Assert.That(protection.Sheet.Value, [Is].True)
            Assert.That(protection.InsertRows.Value, [Is].False)
            Assert.That(protection.DeleteRows.Value, [Is].False)
            Dim cells = part.Worksheet.GetFirstChild(Of SheetData)().GetFirstChild(Of Row)().Elements(Of Cell)().ToArray()
            Assert.That(cells.Select(Function(cell) cell.CellReference.Value), [Is].EqualTo({"A1", "B1", "C1", "AA1"}))
            Assert.That(cells(0).StyleIndex.Value, [Is].EqualTo(1UI))
            Assert.That(cells(2).CellFormula.Text, [Is].EqualTo("B1*2"))
            Assert.That(cells(2).CellValue.Text, [Is].EqualTo("4"))
            Dim drawingId = part.Worksheet.GetFirstChild(Of Spreadsheet.Drawing)().Id.Value
            Dim drawings = DirectCast(part.GetPartById(drawingId), DrawingsPart)
            Dim imageId = drawings.WorksheetDrawing.Descendants(Of Drawing.Blip)().Single().Embed.Value
            Assert.That(ReadBytes(drawings.GetPartById(imageId).GetStream()), [Is].EqualTo(ImageBytes))
            Dim errors = New OpenXmlValidator().Validate(after).Select(Function(errorItem) errorItem.Description).ToArray()
            Assert.That(errors, [Is].Empty)
        End Using

        Using workbook As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.That(workbook.LookupCellValue(Of String)("Inputs", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(text))
            Assert.That(workbook.TryLookupCellValue(Of Decimal)("Inputs", 0, 1), [Is].EqualTo(12.5D))
            Assert.That(workbook.LookupCellValue(Of Boolean)("Inputs", 0, 26, FormulaCalculationMode.UseCachedValueOnly), [Is].True)
            Assert.That(workbook.IsEmptyCell("Inputs", 3, 2), [Is].True)
            Assert.That(workbook.TryLookupCellValue(Of Integer)("Inputs", 4, 3), [Is].Null)
        End Using
    End Sub

    ' Cached results are independently configurable for each copy of the same open workbook.
    <Test>
    Public Sub CacheResetOnlyChangesTheSavedCopy()
        Dim template = Path.Combine(_directory, "template.xlsx")
        Dim clean = Path.Combine(_directory, "clean.xlsx")
        Dim reset = Path.Combine(_directory, "reset.xlsx")
        Dim preserved = Path.Combine(_directory, "preserved.xlsx")
        CreateTemplate(template)
        Using stream = File.OpenRead(template)
            Using workbook As New OpenXmlExcelDataOperations(stream, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
                Assert.That(workbook.AutoCalculationEnabledWorkbookSetting, [Is].False)
                workbook.SaveCopyAs(clean, SaveOptionsForDisabledCalculationEngines.ResetCalculatedValuesForForcedCellRecalculationIfRecalculationRequired)
                Assert.That(HashFile(clean), [Is].EqualTo(HashFile(template)))
                workbook.WriteCellValue("Inputs", 0, 1, 11)
                workbook.SaveCopyAs(reset, SaveOptionsForDisabledCalculationEngines.ResetCalculatedValuesForForcedCellRecalculationIfRecalculationRequired)
                Assert.That(workbook.LookupCellValue(Of Integer)("Inputs", 0, 2, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(4))
                workbook.SaveCopyAs(preserved, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using
            Assert.That(stream.CanRead, [Is].True)
        End Using

        Using workbook As New OpenXmlExcelDataOperations(File.ReadAllBytes(reset), New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.That(workbook.LookupCellFormula("Inputs", 0, 2), [Is].EqualTo("B1*2"))
            Assert.Throws(Of MissingCalculatedCellValueException)(Sub() workbook.LookupCellValueAsObject("Inputs", 0, 2, FormulaCalculationMode.UseCachedValueOnly))
        End Using
        Using workbook As New OpenXmlExcelDataOperations(preserved, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.That(workbook.LookupCellValue(Of Integer)("Inputs", 0, 2, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(4))
        End Using
        Using document = SpreadsheetDocument.Open(reset, False)
            Assert.That(document.WorkbookPart.CalculationChainPart, [Is].Null)
            Assert.That(document.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad.Value, [Is].True)
        End Using
        Using document = SpreadsheetDocument.Open(preserved, False)
            Assert.That(document.WorkbookPart.CalculationChainPart, [Is].Not.Null)
        End Using
    End Sub

    ' In-place Save uses the tracked file, and opening a relative alias in read-only mode protects it.
    <Test>
    Public Sub SaveModesFollowFileOptions()
        Dim template = Path.Combine(_directory, "template.xlsx")
        Dim destination = Path.Combine(_directory, "destination.xlsx")
        CreateTemplate(template)
        Dim originalHash = HashFile(template)

        Using workbook As New OpenXmlExcelDataOperations(template, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.Throws(Of FileReadOnlyException)(Sub() workbook.Save())
            Assert.Throws(Of FileReadOnlyException)(
                Sub() workbook.SaveCopyAs(Path.Combine(_directory, ".", "template.xlsx"), SaveOptionsForDisabledCalculationEngines.NoReset))
            workbook.SaveAs(destination, SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Assert.That(workbook.ReadOnly, [Is].False)
            Assert.That(workbook.FilePath, [Is].EqualTo(destination))
            workbook.WriteCellValue("Inputs", 0, 0, 42)
            workbook.Save()
            workbook.ReloadFromFile()
            Assert.That(workbook.LookupCellValue(Of Integer)("Inputs", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(42))
        End Using
        Assert.That(HashFile(template), [Is].EqualTo(originalHash))
    End Sub

    <Test>
    Public Sub RejectsUnsupportedOperationsBeforeEditing()
        Dim template = Path.Combine(_directory, "template.xlsx")
        CreateTemplate(template)
        Assert.Throws(Of FeatureDisabledException)(Sub() OpenWithOptions(template, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = False}))
        Assert.Throws(Of NotSupportedException)(Sub() OpenWithOptions(template, New ExcelDataOperationsOptions("password")))
        Using workbook As New OpenXmlExcelDataOperations(template, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.Throws(Of FeatureDisabledException)(Sub() workbook.RecalculateAll())
            Assert.Throws(Of FeatureDisabledException)(Sub() workbook.RecalculateSheet("Inputs"))
            Assert.Throws(Of FeatureDisabledException)(Sub() workbook.RecalculateCell("Inputs", 0, 2))
            Assert.Throws(Of FeatureDisabledException)(Sub() workbook.WriteCellFormula("Inputs", 0, 2, "3*2", True))
            Assert.Throws(Of FeatureDisabledException)(
                Sub() workbook.SaveCopyAs(Path.Combine(_directory, "recalculated.xlsx"), True, SaveOptionsForDisabledCalculationEngines.NoReset))
            Assert.Throws(Of NotSupportedException)(Sub() workbook.AddRow("Inputs", 0, 1, True))
            Assert.Throws(Of NotSupportedException)(Sub() workbook.RemoveRows("Inputs", 0, 1, False))
            Assert.Throws(Of NotSupportedException)(Sub() workbook.ProtectSheet("Inputs", ProtectionLevel.SelectAndEditAllCellsButNoFurtherEditing))
            Assert.Throws(Of NotSupportedException)(Sub() workbook.WriteCellValue("Inputs", 0, 2, New Object()))
            Assert.Throws(Of ArgumentOutOfRangeException)(Sub() workbook.WriteCellValue("Inputs", 0, 2, Double.NaN))
            Assert.That(workbook.LookupCellFormula("Inputs", 0, 2), [Is].EqualTo("B1*2"))
            Assert.That(workbook.LookupCellValue(Of Integer)("Inputs", 0, 2, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(4))
        End Using
        Dim ex = Assert.Throws(Of FileCorruptedOrInvalidFileFormatException)(
            Sub()
                Using workbook As New OpenXmlExcelDataOperations(New Byte() {1, 2, 3}, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
                End Using
            End Sub)
        Assert.That(ex.InnerException, [Is].Not.Null)
    End Sub

    <Test>
    Public Sub RecalculatesSharedFormulaFollowersWithoutChangingTheirStoredGroups()
        Dim template = Path.Combine(_directory, "shared.xlsx")
        Dim output = Path.Combine(_directory, "calculated.xlsx")
        CreateTemplate(template)
        Using document = SpreadsheetDocument.Open(template, True)
            Dim source = DirectCast(document.WorkbookPart.GetPartById("rId1"), WorksheetPart).Worksheet
            Dim sheetData = source.GetFirstChild(Of SheetData)()
            Dim first = sheetData.GetFirstChild(Of Row)().Elements(Of Cell)().Last()
            first.CellFormula = New CellFormula("B1+$B$1") With {
                .FormulaType = CellFormulaValues.Shared, .SharedIndex = 0UI, .Reference = "C1:C3"
            }
            Dim second = New Cell With {
                .CellReference = "C2", .CellFormula = New CellFormula With {.FormulaType = CellFormulaValues.Shared, .SharedIndex = 0UI}
            }
            sheetData.Append(New Row(New Cell With {.CellReference = "B2", .CellValue = New CellValue("7")}, second) With {.RowIndex = 2UI})
            sheetData.Append(New Row(New Cell With {.CellReference = "C3", .CellValue = New CellValue("100")}) With {.RowIndex = 3UI})
            source.SheetDimension.Reference = "A1:C3"
            Dim other = DirectCast(document.WorkbookPart.GetPartById("rId2"), WorksheetPart).Worksheet.GetFirstChild(Of SheetData)()
            other.Append(New Row(New Cell With {.CellReference = "A1", .CellFormula = New CellFormula("SUM(Inputs!C1:C3)")}) With {.RowIndex = 1UI})
        End Using
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False}
        Using workbook As New OpenXmlExcelDataOperations(template, OpenMode.OpenExistingFile, options)
            workbook.RecalculateAll()
            Assert.That(workbook.LookupCellValue(Of Double)("Inputs", 1, 2, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(9.0))
            Assert.That(workbook.LookupCellValue(Of Double)("Notes", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(113.0))
            workbook.SaveCopyAs(output, recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim worksheet = DirectCast(document.WorkbookPart.GetPartById("rId1"), WorksheetPart).Worksheet
            Dim formulas = worksheet.Descendants(Of CellFormula)().ToArray()
            Assert.That(formulas.Select(Function(formula) formula.FormulaType.Value), [Is].All.EqualTo(CellFormulaValues.Shared))
            Assert.That(formulas.Select(Function(formula) formula.SharedIndex.Value), [Is].All.EqualTo(0UI))
            Assert.That(formulas(0).Reference.Value, [Is].EqualTo("C1:C3"))
            Assert.That(formulas(1).Text, [Is].Empty)
        End Using
    End Sub

    <Test>
    Public Sub RejectsADeepFormulaChainWithoutExhaustingTheThreadStack()
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False}
        Dim depthPath = Path.Combine(_directory, "depth.xlsx")
        Using document = SpreadsheetDocument.Create(depthPath, SpreadsheetDocumentType.Workbook)
            Dim part = document.AddWorkbookPart()
            Dim sheet = part.AddNewPart(Of WorksheetPart)("rId1")
            part.Workbook = New Workbook(New Sheets(New Sheet With {.Id = "rId1", .Name = "Depth", .SheetId = 1UI}))
            Dim data As New SheetData()
            For rowIndex = 1 To 600
                Dim nextCell = "A" & (rowIndex + 1).ToString(CultureInfo.InvariantCulture)
                Dim formula = New CellFormula("IF(" & nextCell & "=0,1," & nextCell & "+1)")
                data.Append(New Row(New Cell With {.CellReference = "A" & rowIndex.ToString(CultureInfo.InvariantCulture), .CellFormula = formula}) With {.RowIndex = CUInt(rowIndex)})
            Next
            sheet.Worksheet = New Worksheet(data)
        End Using
        Using workbook As New OpenXmlExcelDataOperations(depthPath, OpenMode.OpenExistingFile, options)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(
                Sub() workbook.RecalculateCell("Depth", 0, 0, throwExceptionOnCalculationError:=False))
            Assert.That(workbook.HasMissingCachedFormulaResult("Depth", 0, 0), [Is].True)
        End Using
    End Sub

    <Test>
    Public Sub WritesFormulaAndSwitchesActiveTabWhenHidingIt()
        Dim template = Path.Combine(_directory, "template.xlsx")
        Dim output = Path.Combine(_directory, "edited.xlsx")
        CreateTemplate(template)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableAutoCalculationInWorkbook = False}
        Using workbook As New OpenXmlExcelDataOperations(template, OpenMode.OpenExistingFile, options)
            Assert.That(workbook.AutoCalculationEnabledWorkbookSetting, [Is].True)
            workbook.HideSheet("Inputs", stronglyHide:=False)
            Assert.That(workbook.SelectedSheetName(), [Is].EqualTo("Notes"))
            Assert.Throws(Of InvalidOperationException)(Sub() workbook.HideSheet("Notes"))
            workbook.UnprotectSheet("Inputs")
            Assert.That(workbook.IsProtectedSheet("Inputs"), [Is].False)
            workbook.WriteCellFormula("Inputs", 0, 2, "=B1*3", False)
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim part = document.WorkbookPart
            Assert.That(part.Workbook.CalculationProperties.CalculationMode.Value, [Is].EqualTo(CalculateModeValues.Auto))
            Assert.That(part.Workbook.BookViews.GetFirstChild(Of WorkbookView)().ActiveTab.Value, [Is].EqualTo(1UI))
            Assert.That(part.Workbook.Sheets.GetFirstChild(Of Sheet)().State.Value, [Is].EqualTo(SheetStateValues.Hidden))
            Assert.That(part.CalculationChainPart, [Is].Null)
            Assert.That(DirectCast(part.GetPartById("rId1"), WorksheetPart).Worksheet.Descendants(Of CellFormula)().Single().Text,
                        [Is].EqualTo("B1*3"))
        End Using
    End Sub

    ' A non-anchor cell in an array formula is still part of that formula's range.
    <Test>
    Public Sub RejectsEditsInsideArrayFormulaRange()
        Dim template = Path.Combine(_directory, "template.xlsx")
        Dim output = Path.Combine(_directory, "edited.xlsx")
        CreateTemplate(template)
        Using document = SpreadsheetDocument.Open(template, True)
            Dim sheet = DirectCast(document.WorkbookPart.GetPartById("rId1"), WorksheetPart).Worksheet
            Dim row = sheet.GetFirstChild(Of SheetData)().GetFirstChild(Of Row)()
            Dim formula = row.Elements(Of Cell)().Last().CellFormula
            formula.FormulaType = CellFormulaValues.Array
            formula.Reference = "C1:D1"
            row.Append(New Cell With {.CellReference = "D1", .CellValue = New CellValue("4")})
            sheet.SheetDimension.Reference = "A1:D1"
        End Using
        Dim originalHash = HashFile(template)
        Using workbook As New OpenXmlExcelDataOperations(template, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.Throws(Of NotSupportedException)(Sub() workbook.WriteCellValue("Inputs", 0, 2, 12))
            Assert.Throws(Of NotSupportedException)(Sub() workbook.WriteCellValue("Inputs", 0, 3, 12))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Assert.That(HashFile(output), [Is].EqualTo(originalHash))
    End Sub

    <Test>
    Public Sub CreatesWorkbookThroughExcelOps()
        Dim path = IO.Path.Combine(_directory, "new.xlsx")
        Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile)
        Using workbook As New OpenXmlExcelDataOperations(path, OpenMode.CreateFile, options)
            Assert.That(workbook.SheetNames(), [Is].EqualTo({"Sheet1"}))
            workbook.WriteCellValue("Sheet1", 0, 0, "new")
            workbook.Save()
        End Using
        Using workbook As New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, options:=New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.That(workbook.LookupCellValue(Of String)("Sheet1", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("new"))
        End Using
    End Sub

    ' The same XLSX must produce the same table in both engines for common formatting.
    <Test>
    Public Sub HtmlMatchesEpplus4ForStyledGrid()
        Dim bytes As Byte()
        Using package As New CompuMaster.Epplus4.ExcelPackage()
            Dim sheet = package.Workbook.Worksheets.Add("Grid")
            sheet.Cells("A1:B1").Merge = True
            sheet.Cells("A1").Value = "Amounts & dates"
            sheet.Cells("A1").Style.Font.Bold = True
            sheet.Cells("A1").Style.Font.Italic = True
            sheet.Cells("A1").Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue)
            sheet.Cells("A1").Style.Fill.PatternType = CompuMaster.Epplus4.Style.ExcelFillStyle.Solid
            sheet.Cells("A1").Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray)
            sheet.Cells("A1").Style.HorizontalAlignment = CompuMaster.Epplus4.Style.ExcelHorizontalAlignment.Center
            sheet.Cells("A1").Style.WrapText = True
            sheet.Cells("A2").Value = 0.25
            sheet.Cells("A2").Style.Numberformat.Format = "0.00%"
            sheet.Cells("B2").Value = New DateTime(2026, 11, 25)
            sheet.Cells("B2").Style.Numberformat.Format = "yyyy-mm-dd"
            bytes = package.GetAsByteArray()
        End Using
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = True}
        Dim htmlOptions As New HtmlSheetExportOptions With {.TableHeaderRowIndexes = New List(Of Integer) From {0}}
        Dim reference As New EpplusFreeExcelDataOperations(bytes, options)
        Try
            Using workbook As New OpenXmlExcelDataOperations(bytes, options)
                Dim expected = reference.ExportSheetToHtml("Grid", htmlOptions, HtmlDocumentExportParts.TableOnly).ToString()
                Dim actual = workbook.ExportSheetToHtml("Grid", htmlOptions, HtmlDocumentExportParts.TableOnly).ToString()
                Assert.That(actual, [Is].EqualTo(expected))
                Assert.That(workbook.LookupCellFormat("Grid", 1, 0), [Is].EqualTo("0.00%"))
                Assert.That(workbook.IsMergedCell("Grid", 0, 1), [Is].True)
                Assert.That(workbook.LookupLastCell("Grid").Address, [Is].EqualTo("B2"))
            End Using
        Finally
            reference.Close()
        End Try
    End Sub

    ' Explicit cache resets affect the open workbook and the next saved copy.
    <Test>
    Public Sub ExplicitCacheResetsPreserveFormulasAndSource()
        Dim template = Path.Combine(_directory, "template.xlsx")
        Dim output = Path.Combine(_directory, "edited.xlsx")
        CreateTemplate(template)
        Using document = SpreadsheetDocument.Open(template, True)
            Dim worksheet = DirectCast(document.WorkbookPart.GetPartById("rId2"), WorksheetPart).Worksheet
            worksheet.GetFirstChild(Of SheetData)().Append(New Row(
                New Cell With {.CellReference = "A1", .CellFormula = New CellFormula("4+1"), .CellValue = New CellValue("5")},
                New Cell With {.CellReference = "B1", .CellFormula = New CellFormula("3+4"), .CellValue = New CellValue("7")}
            ) With {.RowIndex = 1UI})
            Dim hidden = DirectCast(document.WorkbookPart.GetPartById("rId3"), WorksheetPart).Worksheet
            hidden.GetFirstChild(Of SheetData)().Append(New Row(
                New Cell With {.CellReference = "A1", .CellFormula = New CellFormula("4+5"), .CellValue = New CellValue("9")}
            ) With {.RowIndex = 1UI})
        End Using
        Dim sourceHash = HashFile(template)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using workbook As New OpenXmlExcelDataOperations(template, OpenMode.OpenExistingFile, options)
            workbook.ResetCellValueFromFormulaCell(New ExcelCell("Inputs", "C1", ExcelCell.ValueTypes.All))
            Assert.That(workbook.HasMissingCachedFormulaResult("Inputs", 0, 2), [Is].True)
            Assert.That(workbook.LookupCellValue(Of Integer)("Notes", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(5))
            workbook.ResetCellValueFromFormulaCell("Notes", 0, 0)
            Assert.That(workbook.HasMissingCachedFormulaResult("Notes", 0, 0), [Is].True)
            Assert.That(workbook.LookupCellValue(Of Integer)("Notes", 0, 1, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(7))
            workbook.ResetCellValueFromFormulaCell("Notes")
            Assert.That(workbook.HasMissingCachedFormulaResult("Notes", 0, 1), [Is].True)
            Assert.That(workbook.LookupCellValue(Of Integer)("Hidden", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(9))
            workbook.ResetCellValueFromFormulaCellInWholeWorkbook()
            Assert.That(workbook.HasMissingCachedFormulaResult("Hidden", 0, 0), [Is].True)
            Assert.Throws(Of ArgumentException)(Sub() workbook.ResetCellValueFromFormulaCell("Inputs", 0, 1))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using workbook As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            Assert.That(workbook.HasMissingCachedFormulaResult("Inputs", 0, 2), [Is].True)
            Assert.That(workbook.LookupCellFormula("Inputs", 0, 2), [Is].EqualTo("B1*2"))
            Assert.That(workbook.LookupCellValue(Of Integer)("Inputs", 0, 1, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(2))
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Assert.That(document.WorkbookPart.CalculationChainPart, [Is].Null)
            Assert.That(document.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad.Value, [Is].True)
        End Using
        Assert.That(HashFile(template), [Is].EqualTo(sourceHash))
    End Sub

    ' The theme fixtures were written by Excel and exercise indexed colors, theme colors and tints.
    <TestCase("Office")>
    <TestCase("Office2013")>
    <TestCase("Legacy")>
    <TestCase("Ion")>
    <TestCase("Red")>
    Public Sub HtmlMatchesEpplus4ForExcelThemeFixtures(themeVariant As String)
        Dim template = TestEnvironment.FullPathOfExistingTestFile("test_data", "HtmlExportThemeExcel" & themeVariant & ".xlsx")
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableInitialCalculation = True}
        Dim reference As New EpplusFreeExcelDataOperations(template, OpenMode.OpenExistingFile, options)
        Try
            Using workbook As New OpenXmlExcelDataOperations(template, OpenMode.OpenExistingFile, options)
                Dim sheet = reference.WorkSheetNames().Single()
                Dim htmlOptions As New HtmlSheetExportOptions()
                Dim expected = reference.ExportSheetToHtml(sheet, htmlOptions, HtmlDocumentExportParts.TableOnly).ToString()
                Dim actual = workbook.ExportSheetToHtml(sheet, htmlOptions, HtmlDocumentExportParts.TableOnly).ToString()
                Assert.That(actual, [Is].EqualTo(expected))
            End Using
        Finally
            reference.Close()
        End Try
    End Sub

    Private Shared Sub OpenWithOptions(path As String, options As ExcelDataOperationsOptions)
        Using workbook As New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, options)
        End Using
    End Sub

    Private Shared Function ReadBytes(stream As Stream) As Byte()
        Using input = stream, result As New MemoryStream()
            input.CopyTo(result)
            Return result.ToArray()
        End Using
    End Function

    Private Shared Function HashFile(path As String) As Byte()
        Using sha = SHA256.Create(), stream = File.OpenRead(path)
            Return sha.ComputeHash(stream)
        End Using
    End Function

    Private Shared ReadOnly ImageBytes As Byte() =
        Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC")

    Private Shared Sub CreateTemplate(path As String)
        Using document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)
            Dim part = document.AddWorkbookPart()
            Dim inputs = part.AddNewPart(Of WorksheetPart)("rId1")
            Dim notes = part.AddNewPart(Of WorksheetPart)("rId2")
            Dim hidden = part.AddNewPart(Of WorksheetPart)("rId3")
            part.Workbook = New Workbook(
                New BookViews(New WorkbookView With {.ActiveTab = 0UI}),
                New Sheets(
                    New Sheet With {.Id = "rId1", .Name = "Inputs", .SheetId = 1UI},
                    New Sheet With {.Id = "rId2", .Name = "Notes", .SheetId = 2UI},
                    New Sheet With {.Id = "rId3", .Name = "Hidden", .SheetId = 3UI, .State = SheetStateValues.Hidden}),
                New DefinedNames(New DefinedName("'Inputs'!$A$1") With {.Name = "InputLabel"}),
                New CalculationProperties With {.CalculationMode = CalculateModeValues.Manual, .CalculationId = 191029UI})

            Dim style = part.AddNewPart(Of WorkbookStylesPart)()
            style.Stylesheet = New Stylesheet(
                New Fonts(New Font()),
                New Fills(New Fill(New PatternFill With {.PatternType = PatternValues.None}),
                          New Fill(New PatternFill With {.PatternType = PatternValues.Gray125})),
                New Borders(New Border()),
                New CellStyleFormats(New CellFormat()),
                New CellFormats(
                    New CellFormat With {.FormatId = 0UI},
                    New CellFormat(New Protection With {.Locked = False}) With {.FormatId = 0UI, .ApplyProtection = True}),
                New CellStyles(New CellStyle With {.Name = "Normal", .FormatId = 0UI, .BuiltinId = 0UI}))
            Dim strings = part.AddNewPart(Of SharedStringTablePart)()
            strings.SharedStringTable = New SharedStringTable(New SharedStringItem(New Spreadsheet.Text("Template text")))
            Dim chain = part.AddNewPart(Of CalculationChainPart)()
            chain.CalculationChain = New CalculationChain(New CalculationCell With {.CellReference = "C1", .SheetId = 1})

            inputs.Worksheet = New Worksheet(
                New SheetDimension With {.Reference = "A1:C1"},
                New SheetViews(New SheetView With {.WorkbookViewId = 0UI, .TabSelected = True, .ShowGridLines = False}),
                New SheetData(New Row(
                    New Cell With {.CellReference = "A1", .StyleIndex = 1UI, .DataType = CellValues.SharedString, .CellValue = New CellValue("0")},
                    New Cell With {.CellReference = "B1", .CellValue = New CellValue("2")},
                    New Cell With {.CellReference = "C1", .CellFormula = New CellFormula("B1*2"), .CellValue = New CellValue("4")}
                ) With {.RowIndex = 1UI}),
                New SheetProtection With {.Sheet = True},
                New PageMargins With {.Left = 0.7, .Right = 0.7, .Top = 0.75, .Bottom = 0.75, .Header = 0.3, .Footer = 0.3})
            notes.Worksheet = New Worksheet(New SheetViews(New SheetView With {.WorkbookViewId = 0UI}), New SheetData())
            hidden.Worksheet = New Worksheet(New SheetData(), New SheetProtection With {.Sheet = True})
            AddImage(inputs)
        End Using
    End Sub

    Private Shared Sub AddImage(part As WorksheetPart)
        Dim drawings = part.AddNewPart(Of DrawingsPart)("rIdDrawing")
        Dim image = drawings.AddImagePart(ImagePartType.Png, "rIdImage")
        Using input As New MemoryStream(ImageBytes, writable:=False)
            image.FeedData(input)
        End Using
        Dim picture As New SpreadsheetDrawing.Picture(
            New SpreadsheetDrawing.NonVisualPictureProperties(
                New SpreadsheetDrawing.NonVisualDrawingProperties With {.Id = 1UI, .Name = "Synthetic image"},
                New SpreadsheetDrawing.NonVisualPictureDrawingProperties()),
            New SpreadsheetDrawing.BlipFill(New Drawing.Blip With {.Embed = "rIdImage"}, New Drawing.Stretch(New Drawing.FillRectangle())),
            New SpreadsheetDrawing.ShapeProperties(
                New Drawing.Transform2D(New Drawing.Offset With {.X = 0L, .Y = 0L}, New Drawing.Extents With {.Cx = 9525L, .Cy = 9525L}),
                New Drawing.PresetGeometry(New Drawing.AdjustValueList()) With {.Preset = Drawing.ShapeTypeValues.Rectangle}))
        drawings.WorksheetDrawing = New SpreadsheetDrawing.WorksheetDrawing(
            New SpreadsheetDrawing.AbsoluteAnchor(
                New SpreadsheetDrawing.Position With {.X = 0L, .Y = 0L},
                New SpreadsheetDrawing.Extent With {.Cx = 9525L, .Cy = 9525L},
                picture, New SpreadsheetDrawing.ClientData()))
        part.Worksheet.Append(New Spreadsheet.Drawing With {.Id = "rIdDrawing"})
    End Sub

End Class
