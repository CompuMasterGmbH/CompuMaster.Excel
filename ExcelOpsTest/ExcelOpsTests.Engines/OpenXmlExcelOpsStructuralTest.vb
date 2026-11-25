Imports System.IO
Imports System.IO.Compression
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports CompuMaster.Excel.Test.ExcelOpsTests.Engines
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Validation
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    Private Const DataSheet As String = "O'Brien Data"
    Private Const DataQualifier As String = "'O''Brien Data'!"

    ' The reference strings cover whole-row and whole-column insertions and deletions as Excel stores them.
    <Test>
    Public Sub ShiftsRowsAndColumnsAndUpdatesReferencesAcrossSheets()
        Dim source = Path.Combine(_directory, "structure.xlsx")
        Dim output = Path.Combine(_directory, "structure-edited.xlsx")
        CreateStructuralWorkbook(source,
            {ValueCell("A1", 5), ValueCell("B2", 2), ValueCell("C3", 3), FormulaCell("E4", "SUM($B2:C$3)+A1"), ValueCell("A6", 9)},
            {FormulaCell("A1", "SUM(" & DataQualifier & "$B$2:$C$3)"),
             FormulaCell("A2", DataQualifier & "A1+" & DataQualifier & "$C$3"),
             FormulaCell("A3", "A1+SUM(" & DataQualifier & "$B:$C," & DataQualifier & "$2:$3)"),
             FormulaCell("A4", "LOG10 (100)+1E3+" & DataQualifier & "A1+N(""" & DataQualifier & "C3 """"literal"""""")")},
            New DefinedName(DataQualifier & "$B$2:$C$3") With {.Name = "DataBlock"},
            New DefinedName("$B$2") With {.Name = "LocalInput", .LocalSheetId = 0UI})

        Using workbook As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            workbook.AddRow(DataSheet, 2, 2, True)
            workbook.AddColumn(DataSheet, 2, 1, True)
            Assert.That(workbook.LookupCellFormula(DataSheet, 5, 5), [Is].EqualTo("SUM($B2:D$5)+A1"))
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM(" & DataQualifier & "$B$2:$D$5)"))
            Assert.That(workbook.LookupCellFormula("Check", 2, 0), [Is].EqualTo("A1+SUM(" & DataQualifier & "$B:$D," & DataQualifier & "$2:$5)"))

            workbook.RemoveRows(DataSheet, 0, 1, True)
            workbook.RemoveColumns(DataSheet, 3, 1, True)
            Assert.That(workbook.LookupCellFormula(DataSheet, 4, 4), [Is].EqualTo("SUM($B1:C$4)+#REF!"))
            Assert.That(workbook.LookupCellFormula("Check", 1, 0), [Is].EqualTo(DataQualifier & "#REF!+" & DataQualifier & "#REF!"))
            Assert.That(workbook.LookupCellFormula("Check", 2, 0), [Is].EqualTo("A1+SUM(" & DataQualifier & "$B:$C," & DataQualifier & "$1:$4)"))
            Assert.That(workbook.LookupCellFormula("Check", 3, 0), [Is].EqualTo(
                "LOG10 (100)+1E3+" & DataQualifier & "#REF!+N(""" & DataQualifier & "C3 """"literal"""""")"))
            Assert.That(workbook.LookupCellValue(Of Integer)(DataSheet, 0, 1, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(2))
            Assert.That(workbook.LookupCellValue(Of Integer)(DataSheet, 6, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(9))
            Assert.Throws(Of MissingCalculatedCellValueException)(Sub() workbook.LookupCellValueAsObject("Check", 0, 0, FormulaCalculationMode.UseCachedValueOnly))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using document = SpreadsheetDocument.Open(output, False)
            Dim workbook = document.WorkbookPart.Workbook
            Assert.That(workbook.DefinedNames.Elements(Of DefinedName)().Select(Function(item) item.Text),
                        [Is].EqualTo({DataQualifier & "$B$1:$C$4", "$B$1"}))
            Assert.That(workbook.CalculationProperties.CalculationMode.Value, [Is].EqualTo(CalculateModeValues.Manual))
            Assert.That(workbook.CalculationProperties.FullCalculationOnLoad.Value, [Is].True)
            Assert.That(document.WorkbookPart.CalculationChainPart, [Is].Null)
            Dim dataPart = DirectCast(document.WorkbookPart.GetPartById("rIdData"), WorksheetPart)
            Assert.That(dataPart.Worksheet.SheetDimension.Reference.Value, [Is].EqualTo("A1:E7"))
            Assert.That(dataPart.Worksheet.Descendants(Of Cell)().Where(Function(cell) cell.CellFormula IsNot Nothing).
                        Select(Function(cell) cell.CellValue), [Is].All.Null)
            Assert.That(New OpenXmlValidator().Validate(document).Select(Function(item) item.Description), [Is].Empty)
        End Using
        Using workbook As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.That(workbook.LookupCellFormula(DataSheet, 4, 4), [Is].EqualTo("SUM($B1:C$4)+#REF!"))
            Assert.Throws(Of MissingCalculatedCellValueException)(Sub() workbook.LookupCellValueAsObject("Check", 2, 0, FormulaCalculationMode.UseCachedValueOnly))
        End Using
        Using before = ZipFile.OpenRead(source), after = ZipFile.OpenRead(output)
            For Each entry In before.Entries.Where(Function(item) item.FullName.Contains("sheet3") OrElse item.FullName.StartsWith("docProps/", StringComparison.Ordinal) OrElse item.FullName.StartsWith("xl/drawings/", StringComparison.Ordinal) OrElse item.FullName.StartsWith("xl/media/", StringComparison.Ordinal))
                Assert.That(ReadBytes(after.GetEntry(entry.FullName).Open()), [Is].EqualTo(ReadBytes(entry.Open())), entry.FullName)
            Next
        End Using
    End Sub

    ' Excel accepts leading zeros in row references and clips ranges that cross the grid edge.
    <Test>
    Public Sub HandlesReferenceSpellingAndTheWorksheetBoundary()
        Dim source = Path.Combine(_directory, "reference-boundary.xlsx")
        CreateStructuralWorkbook(source, {ValueCell("A1", 5)},
            {FormulaCell("A1", DataQualifier & "$A$0000001"),
             FormulaCell("A2", DataQualifier & "B1048576"),
             FormulaCell("A3", "SUM(" & DataQualifier & "B1048575:B1048576)"),
             FormulaCell("A4", "SUM(" & DataQualifier & "B1048574:B1048576)")})
        Using workbook As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            workbook.AddRow(DataSheet, 0, 1, True)
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo(DataQualifier & "$A$2"))
            Assert.That(workbook.LookupCellFormula("Check", 1, 0), [Is].EqualTo(DataQualifier & "#REF!"))
            Assert.That(workbook.LookupCellFormula("Check", 2, 0), [Is].EqualTo("SUM(" & DataQualifier & "B1048576:B1048576)"))
            Assert.That(workbook.LookupCellFormula("Check", 3, 0), [Is].EqualTo("SUM(" & DataQualifier & "B1048575:B1048576)"))
        End Using
    End Sub

    ' Excel leaves references spanning outside a partial shift's strip intact.
    <TestCase(True)>
    <TestCase(False)>
    Public Sub UpdatesOnlyRangesEntirelyInsideTheShiftedStrip(vertical As Boolean)
        Dim source = Path.Combine(_directory, "cell-shift.xlsx")
        Dim output = Path.Combine(_directory, "cell-shift-edited.xlsx")
        Dim range = If(vertical, "$B2:B$4", "B$2:$D2")
        Dim movedRange = If(vertical, "$B3:B$5", "C$2:$E2")
        Dim shortenedRange = If(vertical, "$B3:B$4", "C$2:$D2")
        Dim firstMovedAddress = If(vertical, "$B3", "C$2")
        Dim removedAddress = If(vertical, "$B2", "B$2")
        Dim initialFormulaAddress = If(vertical, "B4", "D2")
        Dim movedFormulaAddress = If(vertical, "B5", "E2")
        Dim unaffectedAddress = If(vertical, "C2", "B3")
        CreateStructuralWorkbook(source, {ValueCell("B2", 2), FormulaCell(initialFormulaAddress, removedAddress & "+A1"), ValueCell(unaffectedAddress, 19)},
            {FormulaCell("A1", "SUM(" & DataQualifier & range & ")"),
             FormulaCell("A2", "SUM(" & DataQualifier & "A1:C3)"),
             FormulaCell("A3", DataQualifier & removedAddress),
             FormulaCell("A4", "SUM(" & DataQualifier & "B:B," & DataQualifier & "2:2)")})
        Using workbook As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            workbook.AddCells(DataSheet, 1, 1, 1, 1,
                If(vertical, CellInsertShiftDirection.ShiftCellsDown, CellInsertShiftDirection.ShiftCellsRight), True)
            Assert.That(workbook.LookupCellFormula(DataSheet, New ExcelCell(DataSheet, movedFormulaAddress, ExcelCell.ValueTypes.All).RowIndex,
                                                  New ExcelCell(DataSheet, movedFormulaAddress, ExcelCell.ValueTypes.All).ColumnIndex),
                        [Is].EqualTo(firstMovedAddress & "+A1"))
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM(" & DataQualifier & movedRange & ")"))
            Assert.That(workbook.LookupCellFormula("Check", 1, 0), [Is].EqualTo("SUM(" & DataQualifier & "A1:C3)"))
            Assert.That(workbook.LookupCellFormula("Check", 3, 0), [Is].EqualTo("SUM(" & DataQualifier & "B:B," & DataQualifier & "2:2)"))
            workbook.RemoveCells(DataSheet, If(vertical, 2, 1), If(vertical, 1, 2), 1, 1,
                If(vertical, CellRemoveShiftDirection.ShiftCellsUp, CellRemoveShiftDirection.ShiftCellsLeft), True)
            Assert.That(workbook.LookupCellFormula(DataSheet, If(vertical, 3, 1), If(vertical, 1, 3)), [Is].EqualTo("#REF!+A1"))
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM(" & DataQualifier & shortenedRange & ")"))
            Assert.That(workbook.LookupCellFormula("Check", 2, 0), [Is].EqualTo(DataQualifier & "#REF!"))
            Assert.That(workbook.LookupCellValue(Of Integer)(New ExcelCell(DataSheet, unaffectedAddress, ExcelCell.ValueTypes.All), FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(19))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using workbook As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM(" & DataQualifier & shortenedRange & ")"))
        End Using
    End Sub

    ' Sheet names can appear in references, error operands, defined names and string literals.
    <TestCase("NewData", "NewData!")>
    <TestCase("O'Brien's Numbers", "'O''Brien''s Numbers'!")>
    <TestCase("XFD1", "'XFD1'!")>
    <TestCase("A01", "'A01'!")>
    <TestCase(" ", "' '!")>
    Public Sub RenamesAWorksheetAndItsExplicitQualifiers(newName As String, newQualifier As String)
        Dim source = Path.Combine(_directory, "rename.xlsx")
        Dim output = Path.Combine(_directory, "renamed.xlsx")
        Dim reference = "SUM(" & DataQualifier & "B1:" & DataQualifier & "B2)+" & DataQualifier & "#REF!"
        Dim literal = "+N(""" & DataQualifier & "B1"")"
        CreateStructuralWorkbook(source, {FormulaCell("B2", "B1+Check!A2")},
            {FormulaCell("A1", reference & literal)},
            New DefinedName(DataQualifier & "$B1") With {.Name = "RelativeInput"},
            New DefinedName("$B$1") With {.Name = "LocalInput", .LocalSheetId = 0UI})

        Using engine As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Dim workbook As ExcelDataOperationsBase = engine
            Assert.Throws(Of NotSupportedException)(Sub() workbook.RenameWorksheet(DataSheet, newName, False))
            Assert.Throws(Of ArgumentException)(Sub() workbook.RenameWorksheet(DataSheet, "CHECK", True))
            Assert.Throws(Of ArgumentException)(Sub() workbook.RenameWorksheet(DataSheet, "Bad?Name", True))
            workbook.RenameWorksheet("o'brien data", newName, True)
            Assert.That(workbook.LookupCellFormula(newName, 1, 1), [Is].EqualTo("B1+Check!A2"))
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo(
                "SUM(" & newQualifier & "B1:" & newQualifier & "B2)+" & newQualifier & "#REF!" & literal))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using document = SpreadsheetDocument.Open(output, False)
            Dim workbook = document.WorkbookPart.Workbook
            Assert.That(workbook.Sheets.Elements(Of Sheet)().Select(Function(item) item.Name.Value), [Is].EqualTo({newName, "Check", "Unrelated"}))
            Assert.That(workbook.DefinedNames.Elements(Of DefinedName)().Select(Function(item) item.Text),
                        [Is].EqualTo({newQualifier & "$B1", "$B$1"}))
        End Using
    End Sub

    <Test>
    Public Sub WorkbookReferences_UpdatesLocalReferencesAndPreservesTheExternalLink()
        WorkbookReferenceCoverage.CopyPair(_directory)
        Dim source = Path.Combine(_directory, WorkbookReferenceCoverage.TargetName)
        Dim output = Path.Combine(_directory, "edited.xlsx")
        Dim linkedFile = Path.Combine(_directory, WorkbookReferenceCoverage.SourceName)
        Dim sourceHash = HashFile(linkedFile)
        Using workbook As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            WorkbookReferenceCoverage.EditTarget(workbook)
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Dim differences = WorkbookReferenceCoverage.CompareEditedTargetWithExcel(output)
        Assert.That(differences, [Is].Empty, String.Join(Environment.NewLine, differences))
        Assert.That(HashFile(linkedFile), [Is].EqualTo(sourceHash))
        Using before = ZipFile.OpenRead(source), after = ZipFile.OpenRead(output)
            For Each entry In before.Entries.Where(Function(item) item.FullName.StartsWith("xl/externalLinks/", StringComparison.Ordinal))
                Assert.That(ReadBytes(after.GetEntry(entry.FullName).Open()), [Is].EqualTo(ReadBytes(entry.Open())), entry.FullName)
            Next
        End Using
    End Sub

    <Test>
    Public Sub WorkbookReferences_RenamesAndDeletesLocallyWithTheExternalFileMissing()
        Dim target = TestEnvironment.FullPathOfExistingTestFile("test_data", WorkbookReferenceCoverage.TargetName)
        Dim localCopy = Path.Combine(_directory, WorkbookReferenceCoverage.TargetName)
        File.Copy(target, localCopy)
        Using workbook As New OpenXmlExcelDataOperations(localCopy, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            workbook.RenameWorksheet(WorkbookReferenceCoverage.LocalSheet, WorkbookReferenceCoverage.RemoteSheet, True)
            workbook.RemoveRows(WorkbookReferenceCoverage.RemoteSheet, 0, 1, True)
            workbook.RemoveColumns(WorkbookReferenceCoverage.RemoteSheet, 0, 1, True)
            Assert.That(workbook.LookupCellFormula(WorkbookReferenceCoverage.ChecksSheet, 1, 0),
                [Is].EqualTo("SUM('Remote Data'!$A1:A$2,'[1]Remote Data'!$B$2:$B$3)"))
        End Using
    End Sub

    ' The link follows the workbook identity even when a local worksheet has the same name.
    <Test>
    Public Sub WorkbookSession_UpdatesAllAttachedDependentsAndTheirExternalCaches()
        WorkbookReferenceCoverage.CopyPair(_directory)
        Dim sourcePath = Path.Combine(_directory, WorkbookReferenceCoverage.SourceName)
        Dim firstPath = Path.Combine(_directory, WorkbookReferenceCoverage.TargetName)
        Dim secondPath = Path.Combine(_directory, "second.xlsx")
        Dim output = Path.Combine(_directory, "saved-links.xlsx")
        File.Copy(firstPath, secondPath)
        Dim sourceName = WorkbookReferenceCoverage.RemoteSheet
        Dim updatedName = "O'Brien source"
        Dim checks = WorkbookReferenceCoverage.ChecksSheet
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using session As New OpenXmlWorkbookSession(),
              source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options),
              first As New OpenXmlExcelDataOperations(firstPath, OpenMode.OpenExistingFile, options),
              second As New OpenXmlExcelDataOperations(secondPath, OpenMode.OpenExistingFile, options)
            first.RenameWorksheet(WorkbookReferenceCoverage.LocalSheet, sourceName, True)
            session.Attach(source)
            session.Attach(first)
            session.Attach(second)
            source.AddRow(sourceName, 1, 1, True)
            Assert.That(first.LookupCellFormula(checks, 0, 0), [Is].EqualTo("'[1]Remote Data'!$B$3"))
            Assert.That(second.LookupCellFormula(checks, 0, 0), [Is].EqualTo("'[1]Remote Data'!$B$3"))
            Assert.That(session.Detach(second), [Is].True)
            source.RenameWorksheet(sourceName, updatedName, True)
            Assert.That(first.LookupCellFormula(checks, 1, 0),
                        [Is].EqualTo("SUM('Remote Data'!$B2:B$3,'[1]O''Brien source'!$B$3:$B$4)"))
            Assert.That(second.LookupCellFormula(checks, 0, 0), [Is].EqualTo("'[1]Remote Data'!$B$3"))
            Assert.That(first.HasMissingCachedFormulaResult(checks, 0, 0), [Is].True)
            first.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim external = document.WorkbookPart.ExternalWorkbookParts.Single().ExternalLink.GetFirstChild(Of ExternalBook)()
            Assert.That(external.SheetNames.GetFirstChild(Of SheetName)().Val.Value, [Is].EqualTo("O'Brien source"))
            Dim cached = external.SheetDataSet.Descendants(Of ExternalCell)()
            Assert.That(cached.Select(Function(cell) cell.CellReference.Value), [Is].EqualTo({"B3", "B4"}))
            Assert.That(cached.Select(Function(cell) cell.Xstring.Text), [Is].EqualTo({"11", "13"}))
        End Using
    End Sub

    ' The second edit sees the moved filter in the current unsaved table definition.
    <Test>
    Public Sub StructuralReference_UpdatesTableFiltersAcrossSuccessiveColumnEdits()
        Dim source = Path.Combine(_directory, "filtered-table.xlsx")
        Dim output = Path.Combine(_directory, "filtered-table-edited.xlsx")
        File.Copy(TestEnvironment.FullPathOfExistingTestFile("test_data", StructuralReferenceCoverage.FixtureName), source)
        Using document = SpreadsheetDocument.Open(source, True)
            Dim data = StructuralReferenceCoverage.WorksheetFor(document.WorkbookPart, "Data")
            Dim table = data.TableDefinitionParts.Single().Table
            table.GetFirstChild(Of AutoFilter)().Append(New FilterColumn(New Filters(New Filter With {.Val = "2"})) With {.ColumnId = 1UI})
            For Each row In data.Worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)().Where(
                Function(item) item.RowIndex.Value = 4UI OrElse item.RowIndex.Value = 7UI)
                row.Hidden = True
            Next
        End Using
        Using editor As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            editor.AddColumn("Data", 2, 1, True)
            editor.RemoveColumns("Data", 3, 1, True)
            Assert.That(editor.LookupCellFormula("Checks", 13, 0), [Is].EqualTo("SUM(#REF!)"))
            editor.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim data = StructuralReferenceCoverage.WorksheetFor(document.WorkbookPart, "Data")
            Dim table = data.TableDefinitionParts.Single().Table
            Dim filter = table.GetFirstChild(Of AutoFilter)()
            Assert.That(filter.Reference.Value, [Is].EqualTo("B3:D6"))
            Assert.That(filter.Elements(Of FilterColumn)(), [Is].Empty)
            Dim hidden = data.Worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)().
                Where(Function(row) If(row.Hidden?.Value, False)).Select(Function(row) row.RowIndex.Value)
            Assert.That(hidden, [Is].EqualTo({7UI}))
        End Using
    End Sub

    <Test>
    Public Sub WorkbookSession_UpdatesAnExternalReferenceInACalculatedTableColumn()
        Dim sourcePath = Path.Combine(_directory, WorkbookReferenceCoverage.SourceName)
        Dim targetPath = Path.Combine(_directory, "linked-table.xlsx")
        Dim output = Path.Combine(_directory, "saved-table.xlsx")
        File.Copy(TestEnvironment.FullPathOfExistingTestFile("test_data", WorkbookReferenceCoverage.SourceName), sourcePath)
        File.Copy(TestEnvironment.FullPathOfExistingTestFile("test_data", StructuralReferenceCoverage.FixtureName), targetPath)
        Using document = SpreadsheetDocument.Open(targetPath, True)
            Dim book = document.WorkbookPart
            book.Workbook.ExternalReferences = New ExternalReferences()
            AddMoveTestLink(book, sourcePath, WorkbookReferenceCoverage.RemoteSheet)
            Dim sheet = StructuralReferenceCoverage.WorksheetFor(book, "PlainTable")
            Dim table = sheet.TableDefinitionParts.Single().Table
            Dim column = table.GetFirstChild(Of TableColumns)().Elements(Of TableColumn)().Last()
            column.CalculatedColumnFormula.Text = "C4*'[1]Remote Data'!$B$2"
            For row = 4 To 6
                Dim address = "D" & row
                Dim cell = sheet.Worksheet.Descendants(Of Cell)().Single(Function(item) item.CellReference.Value = address)
                cell.CellFormula = New CellFormula("C" & row & "*'[1]Remote Data'!$B$2")
            Next
        End Using
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using session As New OpenXmlWorkbookSession(),
              source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options),
              dependent As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, options)
            session.Attach(source)
            session.Attach(dependent)
            source.AddRow(WorkbookReferenceCoverage.RemoteSheet, 1, 1, True)
            Assert.That(dependent.LookupCellFormula("PlainTable", 3, 3), [Is].EqualTo("C4*'[1]Remote Data'!$B$3"))
            dependent.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim table = StructuralReferenceCoverage.WorksheetFor(document.WorkbookPart, "PlainTable").TableDefinitionParts.Single().Table
            Dim column = table.GetFirstChild(Of TableColumns)().Elements(Of TableColumn)().Last()
            Assert.That(table.Reference.Value, [Is].EqualTo("B3:D6"))
            Assert.That(column.CalculatedColumnFormula.Text, [Is].EqualTo("C4*'[1]Remote Data'!$B$3"))
        End Using
    End Sub

    <Test>
    Public Sub WorkbookSession_RejectsAnUnsupportedDependentBeforeChangingAnyWorkbook()
        WorkbookReferenceCoverage.CopyPair(_directory)
        Dim sourcePath = Path.Combine(_directory, WorkbookReferenceCoverage.SourceName)
        Dim targetPath = Path.Combine(_directory, WorkbookReferenceCoverage.TargetName)
        Dim output = Path.Combine(_directory, "source-unchanged.xlsx")
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using session As New OpenXmlWorkbookSession(),
              source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options),
              target As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, options)
            session.Attach(source)
            session.Attach(target)
            target.WriteCellFormula(WorkbookReferenceCoverage.ChecksSheet, 3, 0, "A1#", False)
            Assert.Throws(Of NotSupportedException)(Sub() source.AddRow(WorkbookReferenceCoverage.RemoteSheet, 1, 1, True))
            Assert.That(target.LookupCellFormula(WorkbookReferenceCoverage.ChecksSheet, 0, 0), [Is].EqualTo("'[1]Remote Data'!$B$2"))
            source.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Assert.That(HashFile(output), [Is].EqualTo(HashFile(sourcePath)))
    End Sub

    <Test>
    Public Sub ThreeDimensionalReferencesKeepTheirCoordinatesAndFollowEndpointRenames()
        Dim source = Path.Combine(_directory, "sheet-spans.xlsx")
        CreateStructuralWorkbook(source, {ValueCell("B2", 5)},
            {FormulaCell("A1", "SUM('O''Brien Data:Unrelated'!$B$2:$C$4)"),
             FormulaCell("A2", "SUM(Check:Unrelated!B2)"),
             FormulaCell("A3", "SUM('O''Brien Data:Unrelated'!$B:$C)+B2")})
        Using workbook = OpenStructuralWorkbook(source)
            workbook.AddRow(DataSheet, 0, 1, True)
            workbook.AddColumn("Check", 1, 1, True)
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM('O''Brien Data:Unrelated'!$B$2:$C$4)"))
            Assert.That(workbook.LookupCellFormula("Check", 2, 0), [Is].EqualTo("SUM('O''Brien Data:Unrelated'!$B:$C)+C2"))
            workbook.RenameWorksheet(DataSheet, "Beginning", True)
            workbook.RenameWorksheet("Unrelated", "O'Brien end", True)
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM('Beginning:O''Brien end'!$B$2:$C$4)"))
            Assert.That(workbook.LookupCellFormula("Check", 1, 0), [Is].EqualTo("SUM('Check:O''Brien end'!B2)"))
        End Using
    End Sub

    <TestCase("SUM('O''Brien Data:Missing'!B2)")>
    <TestCase("SUM(Check:Check!B2)")>
    <TestCase("'C:\Temp\[External.xlsx]Data'!B2")>
    <TestCase("[1]Remote!B2")>
    <TestCase("'[1]Remote Data'!B2")>
    <TestCase("SUM(Table1[Amount])")>
    <TestCase(DataQualifier & "B2#")>
    <TestCase("_xlfn.ANCHORARRAY(" & DataQualifier & "B2)")>
    Public Sub UnsupportedFormulaReferencesLeaveTheWorkbookByteIdentical(formula As String)
        Dim source = Path.Combine(_directory, "unsupported.xlsx")
        Dim output = Path.Combine(_directory, "unchanged.xlsx")
        CreateStructuralWorkbook(source, {FormulaCell("B2", "A1")}, {FormulaCell("A1", formula)})
        Using workbook As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            Assert.Throws(Of NotSupportedException)(Sub() workbook.AddRow(DataSheet, 0, 1, True))
            Assert.Throws(Of NotSupportedException)(Sub() workbook.RenameWorksheet(DataSheet, "Data", True))
            Assert.That(workbook.LookupCellFormula(DataSheet, 1, 1), [Is].EqualTo("A1"))
            Assert.That(workbook.LookupCellValue(Of Integer)(DataSheet, 1, 1, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(42))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Assert.That(HashFile(output), [Is].EqualTo(HashFile(source)))
    End Sub

    <TestCase("DefinedName")>
    <TestCase("ConditionalFormatting")>
    <TestCase("GroupedFormula")>
    <TestCase("RowExtension")>
    <TestCase("GridOverflow")>
    Public Sub UnsupportedStructureLeavesPreviouslyStagedChangesUnapplied(feature As String)
        Dim source = Path.Combine(_directory, "unsupported-structure.xlsx")
        Dim output = Path.Combine(_directory, "unchanged-structure.xlsx")
        Dim address = If(feature = "GridOverflow", "B1048576", "B2")
        CreateStructuralWorkbook(source, {FormulaCell(address, "A1")}, {FormulaCell("A1", DataQualifier & address)})
        Using document = SpreadsheetDocument.Open(source, True)
            Dim part = document.WorkbookPart
            Dim worksheet = DirectCast(part.GetPartById("rIdData"), WorksheetPart).Worksheet
            Select Case feature
                Case "DefinedName"
                    part.Workbook.DefinedNames = New DefinedNames(New DefinedName(DataQualifier & "B2") With {.Name = "RelativeInput"})
                Case "ConditionalFormatting"
                    Dim other = DirectCast(part.GetPartById("rIdCheck"), WorksheetPart).Worksheet
                    other.AddChild(New ConditionalFormatting())
                Case "RowExtension"
                    Dim other = DirectCast(part.GetPartById("rIdCheck"), WorksheetPart).Worksheet
                    other.GetFirstChild(Of SheetData)().GetFirstChild(Of Row)().Append(New ExtensionList())
                Case "GroupedFormula"
                    Dim formula = worksheet.Descendants(Of CellFormula)().Single()
                    formula.FormulaType = CellFormulaValues.Array
                    formula.Reference = "B2:B3"
            End Select
        End Using

        Using workbook As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly))
            If feature = "GridOverflow" Then
                Assert.Throws(Of InvalidOperationException)(Sub() workbook.AddRow(DataSheet, 0, 1, True))
            Else
                Assert.Throws(Of NotSupportedException)(Sub() workbook.AddRow(DataSheet, 0, 1, True))
            End If
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo(DataQualifier & address))
            Assert.That(workbook.LookupCellValue(Of Integer)("Check", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(42))
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Assert.That(HashFile(output), [Is].EqualTo(HashFile(source)))
    End Sub

    <Test>
    Public Sub DeletingASpillAnchorClearsItsRemainingGeneratedCells()
        Dim source = Path.Combine(_directory, "dynamic-array.xlsx")
        Dim anchor = FormulaCell("B2", "_xlfn.SEQUENCE(3)")
        anchor.CellFormula.FormulaType = CellFormulaValues.Array
        anchor.CellFormula.Reference = "B2:B4"
        anchor.CellMetaIndex = 1UI
        CreateStructuralWorkbook(source, {anchor, ValueCell("B3", 2), ValueCell("B4", 3), ValueCell("B5", 7)},
            {FormulaCell("A1", "SUM(_xlfn.ANCHORARRAY(" & DataQualifier & "B2))")})
        Using document = SpreadsheetDocument.Open(source, True)
            Dim part = document.WorkbookPart.AddNewPart(Of CellMetadataPart)()
            Using output As New StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write))
                output.Write(<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                                      xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
                                 <metadataTypes count="1"><metadataType name="XLDAPR" cellMeta="1"/></metadataTypes>
                                 <futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">
                                     <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>
                                 </ext></extLst></bk></futureMetadata>
                                 <cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>
                             </metadata>.ToString())
            End Using
        End Using
        Using workbook = OpenStructuralWorkbook(source)
            workbook.RemoveRows(DataSheet, 1, 1, True)
            Assert.That(workbook.IsEmptyCell(DataSheet, 1, 1), [Is].True)
            Assert.That(workbook.IsEmptyCell(DataSheet, 2, 1), [Is].True)
            Assert.That(workbook.LookupCellValue(Of Integer)(DataSheet, 3, 1, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(7))
            Assert.That(workbook.LookupCellFormula("Check", 0, 0), [Is].EqualTo("SUM(_xlfn.ANCHORARRAY(" & DataQualifier & "#REF!))"))
        End Using
    End Sub

    Private Shared Function FormulaCell(address As String, formula As String) As Cell
        Return New Cell With {.CellReference = address, .CellFormula = New CellFormula(formula), .CellValue = New CellValue("42")}
    End Function

    Private Shared Function ValueCell(address As String, value As Integer) As Cell
        Return New Cell With {.CellReference = address, .CellValue = New CellValue(value)}
    End Function

    Private Shared Function StructuralSheet(cells As IEnumerable(Of Cell)) As Worksheet
        Dim data As New SheetData()
        For Each group In cells.GroupBy(Function(cell) New ExcelCell(Nothing, cell.CellReference.Value, ExcelCell.ValueTypes.All).RowNumber).OrderBy(Function(item) item.Key)
            data.Append(New Row(group.OrderBy(Function(cell) New ExcelCell(Nothing, cell.CellReference.Value, ExcelCell.ValueTypes.All).ColumnIndex)) With {.RowIndex = CUInt(group.Key)})
        Next
        Return New Worksheet(New SheetDimension With {.Reference = "A1"}, data)
    End Function

    Private Shared Sub CreateStructuralWorkbook(path As String, data As Cell(), check As Cell(), ParamArray names As DefinedName())
        Using document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)
            Dim workbookPart = document.AddWorkbookPart()
            workbookPart.AddNewPart(Of WorksheetPart)("rIdData").Worksheet = StructuralSheet(data)
            workbookPart.AddNewPart(Of WorksheetPart)("rIdCheck").Worksheet = StructuralSheet(check)
            Dim unrelatedPart = workbookPart.AddNewPart(Of WorksheetPart)("rIdUnrelated")
            unrelatedPart.Worksheet = StructuralSheet({ValueCell("A1", 17)})
            AddImage(unrelatedPart)
            document.PackageProperties.Title = "Synthetic structural edit"
            workbookPart.Workbook = New Workbook(
                New Sheets(New Sheet With {.Id = "rIdData", .SheetId = 1UI, .Name = DataSheet},
                           New Sheet With {.Id = "rIdCheck", .SheetId = 2UI, .Name = "Check"},
                           New Sheet With {.Id = "rIdUnrelated", .SheetId = 3UI, .Name = "Unrelated"}),
                New CalculationProperties With {.CalculationId = 191029UI, .CalculationMode = CalculateModeValues.Manual})
            If names.Length > 0 Then workbookPart.Workbook.DefinedNames = New DefinedNames(names)
            workbookPart.AddNewPart(Of CalculationChainPart)().CalculationChain =
                New CalculationChain(New CalculationCell With {.CellReference = "A1", .SheetId = 2})
        End Using
    End Sub

End Class
