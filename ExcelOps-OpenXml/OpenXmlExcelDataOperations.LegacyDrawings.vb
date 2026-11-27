Imports System.Xml
Imports System.Xml.Linq
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations
        Private Sub ValidateLegacyWorksheetDrawings(part As WorksheetPart, shiftsCells As Boolean)
            Dim worksheet = part.Worksheet
            Dim notes = worksheet.GetFirstChild(Of LegacyDrawing)()
            If notes IsNot Nothing Then
                If shiftsCells OrElse part.WorksheetCommentsPart Is Nothing OrElse
                   Not IsNotesDrawing(GetVmlPart(part, notes.Id?.Value)) Then
                    Throw UnsupportedOperation("Updating references in a worksheet with legacyDrawing")
                End If
            End If

            Dim header = worksheet.GetFirstChild(Of LegacyDrawingHeaderFooter)()
            If header Is Nothing Then Return
            Dim drawing = GetVmlPart(part, header.Id?.Value)
            If drawing Is Nothing OrElse drawing.ExternalRelationships.Any() OrElse drawing.HyperlinkRelationships.Any() OrElse
               drawing.Parts.Any(Function(item) Not TypeOf item.OpenXmlPart Is ImagePart) Then
                Throw UnsupportedOperation("Updating references in a worksheet with unsupported header or footer drawings")
            End If
        End Sub

        Private Shared Function GetVmlPart(part As WorksheetPart, id As String) As VmlDrawingPart
            Return part.Parts.Where(Function(item) String.Equals(item.RelationshipId, id, StringComparison.Ordinal)).
                Select(Function(item) TryCast(item.OpenXmlPart, VmlDrawingPart)).FirstOrDefault()
        End Function

        Private Shared Function IsNotesDrawing(part As VmlDrawingPart) As Boolean
            If part Is Nothing OrElse part.Parts.Any() OrElse part.ExternalRelationships.Any() OrElse part.HyperlinkRelationships.Any() Then Return False
            Dim vml As XNamespace = "urn:schemas-microsoft-com:vml"
            Dim excel As XNamespace = "urn:schemas-microsoft-com:office:excel"
            Try
                Using input = part.GetStream(), reader = XmlReader.Create(input, New XmlReaderSettings With {.DtdProcessing = DtdProcessing.Prohibit})
                    Dim root = XDocument.Load(reader).Root
                    Dim shapes = root.Descendants(vml + "shape").ToArray()
                    Return shapes.Length > 0 AndAlso Not root.Descendants(vml + "group").Any() AndAlso
                        shapes.All(Function(shape)
                            Dim clients = shape.Elements(excel + "ClientData").ToArray()
                            Return clients.Length = 1 AndAlso CStr(clients(0).Attribute("ObjectType")) = "Note"
                        End Function)
                End Using
            Catch ex As XmlException
                Return False
            End Try
        End Function

        Private Shared Function IsSupportedLegacyWorksheetPart(part As WorksheetPart, relationship As IdPartPair) As Boolean
            If TypeOf relationship.OpenXmlPart Is WorksheetCommentsPart Then Return part.Worksheet.GetFirstChild(Of LegacyDrawing)() IsNot Nothing
            If Not TypeOf relationship.OpenXmlPart Is VmlDrawingPart Then Return False
            Return String.Equals(relationship.RelationshipId, part.Worksheet.GetFirstChild(Of LegacyDrawing)()?.Id?.Value, StringComparison.Ordinal) OrElse
                   String.Equals(relationship.RelationshipId, part.Worksheet.GetFirstChild(Of LegacyDrawingHeaderFooter)()?.Id?.Value, StringComparison.Ordinal)
        End Function
    End Class

End Namespace
