Imports System.IO
Imports System.Text
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework
Imports OfficeExcel = DocumentFormat.OpenXml.Office.Excel
Imports Excel2010 = DocumentFormat.OpenXml.Office2010.Excel

Partial Public Class OpenXmlExcelOpsTest

    <Test>
    Public Sub AllowsRemovingAnotherSheetAndCopiesWorksheetWithLegacyNotesAndHeaderImage()
        Dim input = Path.Combine(_directory, "notes.xlsx")
        Dim output = Path.Combine(_directory, "notes-copy.xlsx")
        CreateStructuralWorkbook(input, {ValueCell("A1", 1)}, {ValueCell("A1", 2)})
        Dim noteDrawing = LegacyNotesVml("Note")
        Dim headerDrawing = Encoding.UTF8.GetBytes(
            "<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office""><v:shape id=""LH""><v:imagedata o:relid=""rId1""/></v:shape></xml>")
        Using document = SpreadsheetDocument.Open(input, True)
            Dim sheet = document.WorkbookPart.WorksheetParts.First()
            Dim notes = sheet.AddNewPart(Of WorksheetCommentsPart)()
            notes.Comments = New Comments(New Authors(New Author("Test author")),
                New CommentList(New Comment(New CommentText(New Text("A note"))) With {.Reference = "A1", .AuthorId = 0UI}))
            Dim vml = sheet.AddNewPart(Of VmlDrawingPart)()
            vml.FeedData(New MemoryStream(noteDrawing))
            sheet.Worksheet.AddChild(New LegacyDrawing With {.Id = sheet.GetIdOfPart(vml)})

            Dim header = sheet.AddNewPart(Of VmlDrawingPart)()
            header.FeedData(New MemoryStream(headerDrawing))
            Dim image = header.AddImagePart(ImagePartType.Png, "rId1")
            image.FeedData(New MemoryStream(ImageBytes))
            sheet.Worksheet.AddChild(New HeaderFooter(New OddHeader("&L&G")))
            sheet.Worksheet.AddChild(New LegacyDrawingHeaderFooter With {.Id = sheet.GetIdOfPart(header)})
        End Using

        Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options),
            target As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, options)
            Assert.Throws(Of NotSupportedException)(Sub() source.AddRow(DataSheet, 0, 1, True))
            source.RemoveSheet("Check")
            source.RenameWorksheet(DataSheet, "Report", True)
            source.CopySheet("Report", target, CopyWorksheetFormulaMode.PreserveFormulas)
            target.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using saved = SpreadsheetDocument.Open(output, False)
            Dim item = saved.WorkbookPart.Workbook.Sheets.Elements(Of Sheet)().Single(Function(itemToCheck) itemToCheck.Name.Value = "Report")
            Dim sheet = DirectCast(saved.WorkbookPart.GetPartById(item.Id.Value), WorksheetPart)
            Assert.That(sheet.WorksheetCommentsPart.Comments.GetFirstChild(Of CommentList)().GetFirstChild(Of Comment)().InnerText, [Is].EqualTo("A note"))
            Dim vml = sheet.GetPartById(sheet.Worksheet.GetFirstChild(Of LegacyDrawing)().Id.Value)
            Dim header = sheet.GetPartById(sheet.Worksheet.GetFirstChild(Of LegacyDrawingHeaderFooter)().Id.Value)
            Assert.That(ReadBytes(vml.GetStream()), [Is].EqualTo(noteDrawing))
            Assert.That(ReadBytes(header.GetStream()), [Is].EqualTo(headerDrawing))
            Assert.That(ReadBytes(header.GetPartById("rId1").GetStream()), [Is].EqualTo(ImageBytes))
        End Using
    End Sub

    <Test>
    Public Sub ContinuesToRejectLegacyControlsDuringStructuralEdits()
        Dim input = Path.Combine(_directory, "control.xlsx")
        CreateStructuralWorkbook(input, {ValueCell("A1", 1)}, {ValueCell("A1", 2)})
        Using document = SpreadsheetDocument.Open(input, True)
            Dim sheet = document.WorkbookPart.WorksheetParts.First()
            Dim part = sheet.AddNewPart(Of VmlDrawingPart)()
            part.FeedData(New MemoryStream(LegacyNotesVml("Button")))
            sheet.Worksheet.AddChild(New LegacyDrawing With {.Id = sheet.GetIdOfPart(part)})
        End Using
        Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
        Using source As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            Assert.Throws(Of NotSupportedException)(Sub() source.RemoveSheet("Check"))
        End Using
    End Sub

    <Test>
    Public Sub RemovingSheetPreservesWorksheetConditionalFormattingRulesAndUpdatesTheirReferences()
        Dim input = Path.Combine(_directory, "conditional-format.xlsx")
        Dim output = Path.Combine(_directory, "conditional-format-copied.xlsx")
        CreateStructuralWorkbook(input, {ValueCell("A1", 1)}, {ValueCell("A1", 2)})
        Using document = SpreadsheetDocument.Open(input, True)
            Dim data = DirectCast(document.WorkbookPart.GetPartById("rIdData"), WorksheetPart)
            Dim rules As New ConditionalFormatting(
                New ConditionalFormattingRule(New Formula("A1>0")) With {.Type = ConditionalFormatValues.Expression, .Priority = 1},
                New ConditionalFormattingRule(New Formula("SUM(Check!A1)>0")) With {.Type = ConditionalFormatValues.Expression, .Priority = 2})
            rules.SequenceOfReferences = New DocumentFormat.OpenXml.ListValue(Of DocumentFormat.OpenXml.StringValue) With {.InnerText = "A1:A3"}
            data.Worksheet.AddChild(rules)
            Dim extendedRule As New Excel2010.ConditionalFormattingRule(
                New OfficeExcel.Formula("SUM(Check!A1)>0"), New Excel2010.DifferentialType()) With {
                .Priority = 3, .Id = "{EDAB42D2-B167-4F68-8101-A4FC219DF43B}"
            }
            extendedRule.SetAttribute(New DocumentFormat.OpenXml.OpenXmlAttribute("type", "", "expression"))
            Dim extended As New Excel2010.ConditionalFormatting(extendedRule, New OfficeExcel.ReferenceSequence("A1:A3"))
            Dim entry As New WorksheetExtension(New Excel2010.ConditionalFormattings(extended)) With {
                .Uri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"
            }
            data.Worksheet.AddChild(New WorksheetExtensionList(entry))
        End Using

        Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
        Using workbook As New OpenXmlExcelDataOperations(input, OpenMode.OpenExistingFile, options)
            workbook.RemoveSheet("Check")
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using saved = SpreadsheetDocument.Open(output, False)
            Dim data = DirectCast(saved.WorkbookPart.GetPartById("rIdData"), WorksheetPart)
            Dim rules = data.Worksheet.GetFirstChild(Of ConditionalFormatting)()
            Assert.That(rules.SequenceOfReferences.InnerText, [Is].EqualTo("A1:A3"))
            Assert.That(rules.Descendants(Of Formula)().Select(Function(rule) rule.Text), [Is].EqualTo({"A1>0", "SUM(#REF!A1)>0"}))
            Assert.That(data.Worksheet.Descendants(Of OfficeExcel.Formula)().Single().Text, [Is].EqualTo("SUM(#REF!A1)>0"))
        End Using
    End Sub

    Private Shared Function LegacyNotesVml(objectType As String) As Byte()
        Return Encoding.UTF8.GetBytes(
            "<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">" &
            "<v:shape id=""Note1""><x:ClientData ObjectType=""" & objectType &
            """><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>")
    End Function

End Class
