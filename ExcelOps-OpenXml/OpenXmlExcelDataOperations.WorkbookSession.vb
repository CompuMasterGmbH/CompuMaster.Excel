Imports System.IO
Imports System.Globalization
Imports System.Threading
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Friend Shared ReadOnly WorkbookSessionLock As New Object()
        Friend Property WorkbookSession As OpenXmlWorkbookSession
        Friend ReadOnly Property TemporaryWorkbookName As String
            Get
                Return _temporaryWorkbookName
            End Get
        End Property

        Private _temporaryWorkbookName As String

        Private Sub BeginWorkbookIdentity()
            _temporaryWorkbookName = "Book-" & Guid.NewGuid().ToString("N")
            WorkbookSession?.RegisterTemporaryLink(_temporaryWorkbookName)
        End Sub

        Private Sub ApplyWorkbookSessionEdit(sheetName As String, shift As StructuralShift, newSheetName As String)
            SyncLock WorkbookSessionLock
                Dim workbooks = If(WorkbookSession?.Snapshot(), {Me})
                WithLockedWorkbooks(workbooks,
                    Function()
                        ApplyEditToSession(workbooks, sheetName, shift, newSheetName)
                        Return True
                    End Function)
            End SyncLock
        End Sub

        Private Shared Function WithLockedWorkbooks(Of T)(workbooks As OpenXmlExcelDataOperations(), action As Func(Of T)) As T
            Dim locked As Integer
            Try
                For Each workbook In workbooks
                    Dim lockTaken = False
                    Try
                        Monitor.Enter(workbook.WorkbookSaveLock, lockTaken)
                    Finally
                        If lockTaken Then locked += 1
                    End Try
                Next
                Return action()
            Finally
                For index = locked - 1 To 0 Step -1
                    Monitor.Exit(workbooks(index).WorkbookSaveLock)
                Next
            End Try
        End Function

        Private Sub ApplyEditToSession(workbooks As OpenXmlExcelDataOperations(), sheetName As String,
                                       shift As StructuralShift, newSheetName As String)
            Dim name = FindSheet(sheetName).Name.Value
            If newSheetName IsNot Nothing AndAlso name = newSheetName Then Return
            Dim savedPath = WorkbookFilePath
            Using changes As New StagedWorkbookChanges()
                For Each dependent In workbooks
                    If dependent Is Me OrElse dependent.IsClosed Then Continue For
                    If Not String.IsNullOrEmpty(savedPath) AndAlso dependent.MatchesTransferPath(New Uri(Path.GetFullPath(savedPath))) Then
                        Throw UnsupportedOperation("Updating a session containing the same saved workbook twice")
                    End If
                    Dim links = New HashSet(Of Integer)(dependent.FindLinksTo(Me))
                    If links.Count = 0 Then Continue For
                    changes.Add(dependent,
                        Sub(book) dependent.RewriteIncomingSessionLinks(book, links, name, shift, newSheetName))
                Next

                ApplyStructuralEditCore(name, shift, newSheetName)
                changes.Apply()
            End Using
        End Sub

        Private Function ResolvedSessionLinks(book As WorkbookPart) As List(Of (Id As String, Link As ExternalRelationship, Path As Uri))
            Dim changes As New List(Of (Id As String, Link As ExternalRelationship, Path As Uri))
            Dim session = WorkbookSession
            If session Is Nothing Then Return changes
            For Each part In book.ExternalWorkbookParts
                Dim bookId = part.ExternalLink?.GetFirstChild(Of ExternalBook)()?.Id?.Value
                Dim link = part.ExternalRelationships.FirstOrDefault(Function(item) item.Id = bookId)
                If link Is Nothing Then Continue For
                Dim resolved = session.ResolveTemporaryLink(link.Uri)
                If resolved IsNot Nothing AndAlso resolved.IsAbsoluteUri Then
                    changes.Add((book.GetIdOfPart(part), link, resolved))
                End If
            Next
            Return changes
        End Function

        Private Sub RewriteIncomingSessionLinks(book As WorkbookPart, links As HashSet(Of Integer), editedSheet As String,
                                                shift As StructuralShift, newSheetName As String)
            ValidateWorkbookStructure()
            For Each sheet In book.Workbook.Sheets.Elements(Of Sheet)()
                Dim part = DirectCast(book.GetPartById(sheet.Id.Value), WorksheetPart)
                If part.TableDefinitionParts.Any() OrElse If(part.DrawingsPart?.ChartParts.Any(), False) Then
                    Throw UnsupportedOperation("Updating incoming workbook links in a workbook with tables or charts")
                End If
                ValidateWorksheetStructure(part, Nothing)
                Dim original = GetSheetData(part)
                Dim data = ExpandSharedFormulas(original, sheet.Name.Value, Nothing)
                Dim changed = data IsNot original
                If changed Then part.Worksheet.ReplaceChild(data, original)
                For Each formula In data.Descendants(Of CellFormula)()
                    Dim updated = RewriteFormula(formula.Text, Nothing, editedSheet, shift, newSheetName, False,
                                                 formulaWorkbook:=book, structuralExternalLinks:=links)
                    If updated = formula.Text Then Continue For
                    formula.Text = updated
                    changed = True
                Next
                If changed Then part.Worksheet.Save()
            Next

            If book.Workbook.DefinedNames IsNot Nothing Then
                For Each definition In book.Workbook.DefinedNames.Elements(Of DefinedName)()
                    ' An incoming update must never bind unqualified references to the source worksheet.
                    definition.Text = RewriteFormula(definition.Text, Nothing, editedSheet, shift, newSheetName, False,
                                                     formulaWorkbook:=book, structuralExternalLinks:=links)
                Next
            End If

            For Each index In links
                Dim external = TransferExternalBook(book, index)
                If external.ExternalDefinedNames?.HasChildren = True Then
                    Throw UnsupportedOperation("Updating incoming workbook links with cached external defined names")
                End If
                Dim names = external.SheetNames?.Elements(Of SheetName)().ToArray()
                If names Is Nothing Then Continue For
                For workbookSheetIndex As Integer = 0 To names.Length - 1
                    If Not String.Equals(names(workbookSheetIndex).Val?.Value, editedSheet, StringComparison.OrdinalIgnoreCase) Then Continue For
                    If newSheetName IsNot Nothing Then names(workbookSheetIndex).Val = newSheetName
                    If shift IsNot Nothing AndAlso external.SheetDataSet IsNot Nothing Then
                        Dim cachedSheetId = CUInt(workbookSheetIndex)
                        Dim cached = external.SheetDataSet.Elements(Of ExternalSheetData)().SingleOrDefault(
                            Function(sheetData) If(sheetData.SheetId?.Value, UInteger.MaxValue) = cachedSheetId)
                        If cached IsNot Nothing Then ShiftExternalCache(cached, shift)
                    End If
                Next
                DirectCast(external.Parent, ExternalLink).Save()
            Next
        End Sub

        Private Shared Sub ShiftExternalCache(data As ExternalSheetData, shift As StructuralShift)
            Dim rows As New SortedDictionary(Of Integer, SortedDictionary(Of Integer, ExternalCell))
            For Each sourceRow In data.Elements(Of ExternalRow)()
                For Each cell In sourceRow.Elements(Of ExternalCell)()
                    Dim address = cell.CellReference?.Value
                    Dim addressEnd As Integer
                    Dim reference = ReadEndpoint(If(address, ""), 0, addressEnd)
                    If reference Is Nothing OrElse Not reference.IsCell OrElse addressEnd <> address.Length Then
                        Throw New InvalidDataException("An external link cache contains an invalid cell reference.")
                    End If
                    Dim row = reference.Row
                    Dim column = reference.Column
                    If shift.Vertical AndAlso shift.Includes(column) Then
                        row = shift.MoveIndex(row)
                    ElseIf Not shift.Vertical AndAlso shift.Includes(row) Then
                        column = shift.MoveIndex(column)
                    End If
                    If row < 0 OrElse column < 0 Then Continue For
                    Dim rowCells As SortedDictionary(Of Integer, ExternalCell) = Nothing
                    If Not rows.TryGetValue(row, rowCells) Then
                        rowCells = New SortedDictionary(Of Integer, ExternalCell)
                        rows.Add(row, rowCells)
                    End If
                    Dim moved = DirectCast(cell.CloneNode(True), ExternalCell)
                    moved.CellReference = ExcelCell.LocalCellAddress(row, column)
                    rowCells.Add(column, moved)
                Next
            Next

            data.RemoveAllChildren(Of ExternalRow)()
            For Each row In rows
                data.Append(New ExternalRow(row.Value.Values) With {.RowIndex = CUInt(row.Key + 1)})
            Next
        End Sub

    End Class

End Namespace
