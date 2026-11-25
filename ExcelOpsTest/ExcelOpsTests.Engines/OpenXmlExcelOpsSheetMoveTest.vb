Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    <Test>
    Public Sub WorksheetTransfer_OpenXmlLeavesBothWorkbooksUntouchedWhenARequiredLinkCannotBeSaved()
        Dim input = Path.Combine(_directory, "move-source.xlsx")
        CreateSimpleMoveWorkbook(input, withIncomingReference:=True)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options),
              target As New OpenXmlExcelDataOperations(Nothing, OpenMode.CreateFile, options)
            target.WriteCellValue(Of Integer)("Sheet1", 0, 0, 42)
            Assert.Throws(Of NotSupportedException)(Sub() source.MoveSheet("Data", target, 0))
            Assert.That(source.SheetNames(), [Is].EqualTo({"Data", "Remain"}))
            Assert.That(source.LookupCellFormula("Remain", 0, 0), [Is].EqualTo("Data!A1"))
            Assert.That(target.SheetNames(), [Is].EqualTo({"Sheet1"}))
            Assert.That(target.LookupCellValue(Of Integer)("Sheet1", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(42))
        End Using
    End Sub

    <Test>
    Public Sub WorksheetTransfer_OpenXmlCanMoveBetweenUnsavedWorkbooksWithoutCreatingALink()
        Dim input = Path.Combine(_directory, "move-source.xlsx")
        CreateSimpleMoveWorkbook(input, withIncomingReference:=False)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(File.ReadAllBytes(input), options),
              target As New OpenXmlExcelDataOperations(Nothing, OpenMode.CreateFile, options)
            Assert.Throws(Of NotSupportedException)(Sub() source.MoveSheets({"Data", "Remain"}, target))
            Assert.That(source.SheetNames(), [Is].EqualTo({"Data", "Remain"}))
            Assert.That(target.SheetNames(), [Is].EqualTo({"Sheet1"}))
            Dim moved = source.MoveSheet("Data", target, 0)
            Assert.That(moved.TargetSheetName, [Is].EqualTo("Data"))
            Assert.That(moved.TargetSheetIndex, [Is].Zero)
            Assert.That(source.SheetNames(), [Is].EqualTo({"Remain"}))
            Assert.That(target.SheetNames(), [Is].EqualTo({"Data", "Sheet1"}))
            Assert.That(target.LookupCellValue(Of Integer)("Data", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(12))
        End Using
    End Sub

    <Test>
    Public Sub WorksheetTransfer_OpenXmlRemovesAnUnusedOriginLinkAndRenumbersRemainingLinks()
        Dim sourcePath = Path.Combine(_directory, "origin.xlsx")
        Dim targetPath = Path.Combine(_directory, "target.xlsx")
        Dim thirdPath = Path.Combine(_directory, "other.xlsx")
        Dim output = Path.Combine(_directory, "moved.xlsx")
        CreateSimpleMoveWorkbook(sourcePath, withIncomingReference:=False)
        CreateSimpleMoveWorkbook(targetPath, withIncomingReference:=False)
        Using document = SpreadsheetDocument.Open(targetPath, True)
            Dim book = document.WorkbookPart
            Dim sheet = book.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = "Remain")
            Dim part = DirectCast(book.GetPartById(sheet.Id.Value), WorksheetPart)
            part.Worksheet.GetFirstChild(Of SheetData)().Append(
                New Row(New Cell With {.CellReference = "A1", .CellFormula = New CellFormula("[1]Data!A1")}) With {.RowIndex = 1UI},
                New Row(New Cell With {.CellReference = "A2", .CellFormula = New CellFormula("[2]Other!A1")}) With {.RowIndex = 2UI})
            book.Workbook.ExternalReferences = New ExternalReferences()
            AddMoveTestLink(book, sourcePath, "Data")
            AddMoveTestLink(book, thirdPath, "Other")
        End Using

        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options),
              target As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, options)
            Dim moved = source.MoveSheet("Data", target, 1)
            Assert.That(moved.TargetSheetName, [Is].EqualTo("Data (2)"))
            Assert.That(target.LookupCellFormula("Remain", 0, 0), [Is].EqualTo("'Data (2)'!A1"))
            Assert.That(target.LookupCellFormula("Remain", 1, 0), [Is].EqualTo("[1]Other!A1"))
            target.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using saved = SpreadsheetDocument.Open(output, False)
            Dim link = saved.WorkbookPart.ExternalWorkbookParts.Single()
            Assert.That(link.ExternalRelationships.Single().Uri.LocalPath, [Is].EqualTo(thirdPath))
            Assert.That(saved.WorkbookPart.Workbook.ExternalReferences.Elements(Of ExternalReference)().Count(), [Is].EqualTo(1))
        End Using
    End Sub

    <Test>
    Public Sub WorksheetTransfer_OpenXmlSessionRedirectsAThirdWorkbookToTheMovedSheet()
        Dim sourcePath = Path.Combine(_directory, "origin.xlsx")
        Dim targetPath = Path.Combine(_directory, "destination.xlsx")
        Dim dependentPath = Path.Combine(_directory, "dependent.xlsx")
        Dim output = Path.Combine(_directory, "saved-dependent.xlsx")
        CreateSimpleMoveWorkbook(sourcePath, withIncomingReference:=False)
        CreateSimpleMoveWorkbook(targetPath, withIncomingReference:=False)
        CreateSimpleMoveWorkbook(dependentPath, withIncomingReference:=False)
        Using document = SpreadsheetDocument.Open(dependentPath, True)
            Dim book = document.WorkbookPart
            book.Workbook.ExternalReferences = New ExternalReferences()
            AddMoveTestLink(book, sourcePath, "Data")
        End Using
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly)
        Using session As New OpenXmlWorkbookSession(),
              source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, options),
              target As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, options),
              dependent As New OpenXmlExcelDataOperations(dependentPath, OpenMode.OpenExistingFile, options)
            dependent.WriteCellFormula("Remain", 0, 0, "[1]Data!$A$1", False)
            session.Attach(source)
            session.Attach(target)
            session.Attach(dependent)
            Dim moved = source.MoveSheet("Data", target, 0)
            Assert.That(moved.TargetSheetName, [Is].EqualTo("Data (2)"))
            Assert.That(dependent.LookupCellFormula("Remain", 0, 0), [Is].EqualTo("'[1]Data (2)'!$A$1"))
            target.RenameWorksheet("Data (2)", "Moved", True)
            Assert.That(dependent.LookupCellFormula("Remain", 0, 0), [Is].EqualTo("'[1]Moved'!$A$1"))
            dependent.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Dim link = document.WorkbookPart.ExternalWorkbookParts.Single()
            Assert.That(link.ExternalRelationships.Single().Uri.LocalPath, [Is].EqualTo(targetPath))
        End Using
    End Sub

    ' Both save orders retain the temporary workbook identity until the dependent is saved again.
    <TestCase(False)>
    <TestCase(True)>
    Public Sub WorksheetTransfer_OpenXmlSessionResolvesNewWorkbookLinksOnFirstSave(targetFirst As Boolean)
        Dim input = Path.Combine(_directory, "template.xlsx")
        Dim sourcePath = Path.Combine(_directory, "source.xlsx")
        Dim targetPath = Path.Combine(_directory, "target.xlsx")
        CreateSimpleMoveWorkbook(input, withIncomingReference:=True)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.DefaultBehaviourOnCreateFile)
        Using session As New OpenXmlWorkbookSession(),
              source As New OpenXmlExcelDataOperations(File.ReadAllBytes(input), options),
              target As New OpenXmlExcelDataOperations(Nothing, OpenMode.CreateFile, options)
            session.Attach(source)
            session.Attach(target)
            source.WriteCellFormula("Data", 0, 1, "Remain!B1", False)
            source.MoveSheet("Data", target, 0)
            target.AddRow("Data", 0, 1, True)
            Assert.That(source.LookupCellFormula("Remain", 0, 0), [Is].EqualTo("[1]Data!A2"))
            Assert.That(target.LookupCellFormula("Data", 1, 1), [Is].EqualTo("[1]Remain!B1"))
            If targetFirst Then
                target.SaveAs(targetPath, SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(SavedLink(targetPath).OriginalString, Does.StartWith("Book-"))
                source.SaveAs(sourcePath, SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(SavedLink(targetPath).OriginalString, Does.StartWith("Book-"))
            Else
                source.SaveAs(sourcePath, SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(SavedLink(sourcePath).OriginalString, Does.StartWith("Book-"))
                target.SaveAs(targetPath, SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(SavedLink(sourcePath).OriginalString, Does.StartWith("Book-"))
            End If
            source.Save()
            target.Save()
            Assert.That(SavedLink(sourcePath).LocalPath, [Is].EqualTo(targetPath))
            Assert.That(SavedLink(targetPath).LocalPath, [Is].EqualTo(sourcePath))
        End Using
    End Sub

    Private Shared Function SavedLink(path As String) As Uri
        Using document = SpreadsheetDocument.Open(path, False)
            Return document.WorkbookPart.ExternalWorkbookParts.Single().ExternalRelationships.Single().Uri
        End Using
    End Function

    Private Shared Sub AddMoveTestLink(book As WorkbookPart, path As String, sheet As String)
        Dim part = book.AddNewPart(Of ExternalWorkbookPart)()
        Dim link = part.AddExternalRelationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath", New Uri(path))
        part.ExternalLink = New ExternalLink(New ExternalBook(New SheetNames(New SheetName With {.Val = sheet})) With {.Id = link.Id})
        book.Workbook.ExternalReferences.Append(New ExternalReference With {.Id = book.GetIdOfPart(part)})
    End Sub

    Private Shared Sub CreateSimpleMoveWorkbook(path As String, withIncomingReference As Boolean)
        Using document = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
            Dim book = document.AddWorkbookPart()
            Dim source = book.AddNewPart(Of WorksheetPart)()
            Dim remaining = book.AddNewPart(Of WorksheetPart)()
            source.Worksheet = New Worksheet(New SheetData(New Row(New Cell With {.CellReference = "A1", .CellValue = New CellValue("12")}) With {.RowIndex = 1UI}))
            remaining.Worksheet = New Worksheet(New SheetData())
            If withIncomingReference Then
                Dim formula As New Cell With {.CellReference = "A1", .CellFormula = New CellFormula("Data!A1")}
                remaining.Worksheet.GetFirstChild(Of SheetData)().Append(New Row(formula) With {.RowIndex = 1UI})
            End If
            book.Workbook = New Workbook(New Sheets(
                New Sheet With {.Name = "Data", .SheetId = 1UI, .Id = book.GetIdOfPart(source)},
                New Sheet With {.Name = "Remain", .SheetId = 2UI, .Id = book.GetIdOfPart(remaining)}))
        End Using
    End Sub

End Class
