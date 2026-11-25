Imports System.Globalization
Imports System.IO
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Function ExternalWorkbookUri(index As Integer) As Uri
            Dim reference = WorkbookPart.Workbook.ExternalReferences?.Elements(Of ExternalReference)().ElementAtOrDefault(index - 1)
            Dim part As OpenXmlPart = Nothing
            If reference?.Id Is Nothing OrElse Not WorkbookPart.TryGetPartById(reference.Id.Value, part) Then Return Nothing
            Dim link = TryCast(part, ExternalWorkbookPart)
            Dim relationshipId As String = link?.ExternalLink?.GetFirstChild(Of ExternalBook)()?.Id?.Value
            If relationshipId Is Nothing Then Return Nothing
            Return link.ExternalRelationships.FirstOrDefault(Function(item) item.Id = relationshipId)?.Uri
        End Function

        Private NotInheritable Partial Class FormulaCalculation
            Private ReadOnly _externalWorkbooks As New Dictionary(Of Integer, FormulaCalculation)

            Private Function ResolveReference(area As FormulaArea) As FormulaCalculation
                If Not area.WorkbookIndex.HasValue Then Return Me
                Dim index = area.WorkbookIndex.Value
                Dim source As FormulaCalculation = Nothing
                If _externalWorkbooks.TryGetValue(index, source) Then Return source

                Dim uri = _owner.ExternalWorkbookUri(index)
                Dim workbook = FindExternalWorkbook(uri)
                If workbook Is Nothing Then
                    Throw New NotSupportedException("External workbook [" & index.ToString(CultureInfo.InvariantCulture) &
                        "] is unavailable. Open its source workbook in the same calculation session.")
                End If
                source = ForWorkbook(workbook)
                _externalWorkbooks.Add(index, source)
                Return source
            End Function

            Private Function FindExternalWorkbook(uri As Uri) As OpenXmlExcelDataOperations
                If uri Is Nothing OrElse _run.Session Is Nothing Then Return Nothing
                uri = If(_owner.WorkbookSession?.ResolveTemporaryLink(uri), uri)
                If uri.IsAbsoluteUri AndAlso Not uri.IsFile Then Return Nothing
                Dim relativePath As String = Nothing
                Dim linkedPath As String = Nothing
                If uri.IsAbsoluteUri Then
                    linkedPath = uri.LocalPath
                Else
                    relativePath = Uri.UnescapeDataString(uri.OriginalString).Replace("/"c, Path.DirectorySeparatorChar).Replace("\"c, Path.DirectorySeparatorChar)
                    If Path.IsPathRooted(relativePath) OrElse relativePath.IndexOf(":"c) >= 0 Then Return Nothing
                    Dim ownerPath = ReferencePath(_owner)
                    If Not String.IsNullOrEmpty(ownerPath) Then
                        linkedPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(ownerPath), relativePath))
                    End If
                End If

                If linkedPath IsNot Nothing Then
                    Dim comparison = If(Path.DirectorySeparatorChar = "\"c, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
                    Dim pathMatch = _run.Session.Workbooks.Values.OfType(Of OpenXmlExcelDataOperations)().FirstOrDefault(
                        Function(workbook) Not workbook.IsClosed AndAlso String.Equals(ReferencePath(workbook), linkedPath, comparison))
                    If pathMatch IsNot Nothing Then Return pathMatch
                End If

                If relativePath Is Nothing OrElse Path.GetFileName(relativePath) <> relativePath Then Return Nothing
                Dim temporaryMatch = _run.Session.Workbooks.Values.OfType(Of OpenXmlExcelDataOperations)().FirstOrDefault(
                    Function(workbook) Not workbook.IsClosed AndAlso String.Equals(workbook.TemporaryWorkbookName, relativePath, StringComparison.Ordinal))
                If temporaryMatch IsNot Nothing Then Return temporaryMatch
                Dim workbookEntry As ExcelDataOperationsBase = Nothing
                If Not _run.Session.Workbooks.TryGetValue(relativePath, workbookEntry) OrElse
                   workbookEntry.IsClosed Then Return Nothing
                Dim source = TryCast(workbookEntry, OpenXmlExcelDataOperations)
                If source Is Nothing OrElse Not String.IsNullOrEmpty(ReferencePath(source)) Then Return Nothing
                Return source
            End Function

            Private Function ReferencePath(workbook As OpenXmlExcelDataOperations) As String
                Return If(workbook Is _run.CopyWorkbook, _run.CopySourcePath, workbook.FilePath)
            End Function
        End Class

    End Class

End Namespace
