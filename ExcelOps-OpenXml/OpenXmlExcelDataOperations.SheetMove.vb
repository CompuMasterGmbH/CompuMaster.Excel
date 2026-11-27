Imports System.Globalization
Imports System.IO
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Sub ValidateMoveWorkbook()
            ValidateWorkbookStructure(allowChartSheets:=True)
            For Each sheet In WorkbookSheets()
                Dim part = WorkbookPart.GetPartById(sheet.Id.Value)
                If Not (TypeOf part Is WorksheetPart OrElse TypeOf part Is ChartsheetPart) Then
                    Throw UnsupportedOperation("Updating references in a workbook with unsupported sheet types")
                End If
            Next
        End Sub

        Private Sub RewriteMoveDependents(book As WorkbookPart, excluded As Dictionary(Of String, String),
                                          rewriteQualifier As Func(Of String, String, String),
                                          Optional formulaWorkbook As WorkbookPart = Nothing,
                                          Optional observeName As Action(Of String, String) = Nothing)
            For Each sheet In book.Workbook.Sheets.Elements(Of Sheet)()
                Dim name = sheet.Name.Value
                If excluded IsNot Nothing AndAlso excluded.ContainsKey(name) Then Continue For
                Dim part = book.GetPartById(sheet.Id.Value)
                Dim worksheet = TryCast(part, WorksheetPart)
                If worksheet IsNot Nothing Then
                    If worksheet.TableDefinitionParts.Any() Then
                        Throw UnsupportedOperation("Updating move references in a workbook with tables")
                    End If
                    ValidateWorksheetStructure(worksheet, Nothing)
                    Dim original = GetSheetData(worksheet)
                    Dim data = ExpandSharedFormulas(original, name, Nothing)
                    Dim changed = data IsNot original
                    If changed Then worksheet.Worksheet.ReplaceChild(data, original)
                    For Each formula In data.Descendants(Of CellFormula)()
                        Dim revised = RewriteFormula(formula.Text, name, name, Nothing, Nothing, False,
                                                     rewriteSheetQualifier:=rewriteQualifier, formulaWorkbook:=formulaWorkbook, observeName:=observeName)
                        If revised <> formula.Text Then
                            formula.Text = revised
                            changed = True
                        End If
                    Next
                    If changed Then SaveWorksheet(worksheet.Worksheet, worksheet)
                    Dim drawing = worksheet.DrawingsPart
                    If drawing IsNot Nothing Then
                        Dim charts = PrepareDrawingChartCopy(drawing, name, CopyWorksheetFormulaMode.ExcelCompatible, rewriteQualifier,
                                                            move:=True, retained:=True, formulaWorkbook:=formulaWorkbook)
                        SaveTransferredCharts(drawing, charts)
                    End If
                Else
                    Dim chartSheet = DirectCast(part, ChartsheetPart)
                    Dim charts = PrepareChartSheetCopy(chartSheet, name, CopyWorksheetFormulaMode.ExcelCompatible, rewriteQualifier,
                                                       move:=True, retained:=True, formulaWorkbook:=formulaWorkbook)
                    SaveTransferredCharts(chartSheet.DrawingsPart, charts)
                End If
            Next
            If book.Workbook.DefinedNames Is Nothing Then Return
            For Each definition In book.Workbook.DefinedNames.Elements(Of DefinedName)()
                If excluded IsNot Nothing AndAlso definition.LocalSheetId IsNot Nothing Then
                    Dim scope = book.Workbook.Sheets.Elements(Of Sheet)().ElementAt(CInt(definition.LocalSheetId.Value))
                    If excluded.ContainsKey(scope.Name.Value) Then Continue For
                End If
                definition.Text = RewriteFormula(definition.Text, Nothing, Nothing, Nothing, Nothing, True,
                                                 rewriteSheetQualifier:=rewriteQualifier, formulaWorkbook:=formulaWorkbook, observeName:=observeName)
            Next
        End Sub

        Private Function FindLinksTo(source As OpenXmlExcelDataOperations) As Integer()
            Dim count = If(WorkbookPart.Workbook.ExternalReferences?.ChildElements.Count, 0)
            Dim matches As New List(Of Integer)
            For index = 1 To count
                Dim external = TransferExternalBook(WorkbookPart, index)
                Dim linkedPath = ResolveTransferLink(index, external)
                If linkedPath IsNot Nothing AndAlso source.MatchesTransferPath(linkedPath) Then matches.Add(index)
            Next
            Return matches.ToArray()
        End Function

        Private Sub StageSessionMoveDependents(changes As StagedWorkbookChanges, workbooks As OpenXmlExcelDataOperations(),
                                               destination As OpenXmlExcelDataOperations, assigned As Dictionary(Of String, String),
                                               targetSheetNames As String())
            For Each dependent In workbooks
                If dependent Is Me OrElse dependent Is destination OrElse dependent.IsClosed Then Continue For
                Dim oldLinks = dependent.FindLinksTo(Me)
                If oldLinks.Length = 0 Then Continue For
                dependent.ValidateMoveWorkbook()
                Dim links As New SheetTransferLinks(dependent)
                Dim keep As New HashSet(Of Integer)
                changes.Add(dependent,
                    Sub(book)
                        dependent.RewriteMoveDependents(book, Nothing,
                            Function(name, qualifier)
                                If Not name.StartsWith("[", StringComparison.Ordinal) Then Return qualifier
                                Dim external = dependent.ReadTransferExternalReference(name)
                                Dim renamed As String = Nothing
                                If oldLinks.Contains(external.Index) AndAlso assigned.TryGetValue(external.Sheet, renamed) Then
                                    Return links.ReferenceTo(destination.TransferWorkbookPath(allowUnsaved:=True), targetSheetNames, renamed)
                                End If
                                keep.Add(external.Index)
                                Return qualifier
                            End Function)
                        links.ApplyTo(book)
                        dependent.RemoveUnusedMoveLinks(book, oldLinks.Where(Function(index) Not keep.Contains(index)).ToArray())
                    End Sub)
            Next
        End Sub

        Private Sub RemoveUnusedMoveLinks(book As WorkbookPart, obsolete As Integer())
            If obsolete.Length = 0 Then Return
            Dim references = book.Workbook.ExternalReferences
            Dim records = references.Elements(Of ExternalReference)().ToArray()
            Dim indexes(records.Length) As Integer
            Dim newIndex As Integer
            For index = 1 To records.Length
                If obsolete.Contains(index) Then Continue For
                newIndex += 1
                indexes(index) = newIndex
            Next

            ' Rewrite while the original association between formula indexes and workbook links still exists.
            If obsolete.First() <= records.Length - obsolete.Length Then
                RewriteMoveDependents(book, Nothing,
                    Function(name, originalQualifier)
                        If Not name.StartsWith("[", StringComparison.Ordinal) Then Return originalQualifier
                        Dim close = name.IndexOf("]"c)
                        Dim oldIndex = Integer.Parse(name.Substring(1, close - 1), CultureInfo.InvariantCulture)
                        Dim replacement = indexes(oldIndex)
                        If replacement = 0 Then Throw New InvalidDataException("A workbook link marked unused still has a reference.")
                        If replacement = oldIndex Then Return originalQualifier
                        Dim open = originalQualifier.IndexOf("["c)
                        Dim bracket = originalQualifier.IndexOf("]"c, open + 1)
                        Return originalQualifier.Substring(0, open + 1) &
                            replacement.ToString(CultureInfo.InvariantCulture) & originalQualifier.Substring(bracket)
                    End Function, formulaWorkbook:=book)
            End If

            For Each index In obsolete
                Dim reference = records(index - 1)
                book.DeletePart(reference.Id.Value)
                reference.Remove()
            Next
            If Not references.HasChildren Then references.Remove()
        End Sub

        Private Shared Sub RemoveTransferredSheets(book As WorkbookPart, assigned As Dictionary(Of String, String))
            Dim workbook = book.Workbook
            Dim oldOrder = workbook.Sheets.Elements(Of Sheet)().ToArray()
            Dim remaining As New List(Of Sheet)
            Dim positions(oldOrder.Length - 1) As Integer
            For index = 0 To oldOrder.Length - 1
                Dim sheet = oldOrder(index)
                If assigned.ContainsKey(sheet.Name.Value) Then
                    positions(index) = -1
                    book.DeletePart(sheet.Id.Value)
                    sheet.Remove()
                Else
                    positions(index) = remaining.Count
                    remaining.Add(sheet)
                End If
            Next
            If workbook.DefinedNames IsNot Nothing Then
                For Each definition In workbook.DefinedNames.Elements(Of DefinedName)().ToArray()
                    If definition.LocalSheetId IsNot Nothing Then
                        Dim index = positions(CInt(definition.LocalSheetId.Value))
                        If index < 0 Then
                            definition.Remove()
                        Else
                            definition.LocalSheetId = CUInt(index)
                        End If
                    End If
                Next
                If Not workbook.DefinedNames.HasChildren Then workbook.DefinedNames.Remove()
            End If
            If workbook.BookViews Is Nothing Then Return
            Dim hasVisibleSheets = remaining.Any(AddressOf IsVisible)
            For Each view In workbook.BookViews.Elements(Of WorkbookView)()
                If Not hasVisibleSheets Then
                    view.ActiveTab = Nothing
                    view.FirstSheet = Nothing
                    Continue For
                End If
                view.ActiveTab = RemapRemovedSheetPosition(If(view.ActiveTab?.Value, 0UI), positions, remaining)
                If view.FirstSheet IsNot Nothing Then
                    view.FirstSheet = RemapRemovedSheetPosition(view.FirstSheet.Value, positions, remaining)
                End If
            Next
        End Sub

        Private Shared Function RemapRemovedSheetPosition(previous As UInteger, positions As Integer(), remaining As List(Of Sheet)) As UInteger
            For index = CInt(System.Math.Min(previous, CUInt(positions.Length))) To positions.Length - 1
                Dim position = positions(index)
                If position >= 0 AndAlso IsVisible(remaining(position)) Then Return CUInt(position)
            Next
            Return CUInt(remaining.FindLastIndex(AddressOf IsVisible))
        End Function

        Private Function StageSheetTransfer(edit As Action(Of WorkbookPart)) As MemoryStream
            Dim snapshot As New MemoryStream()
            Try
                _source.WriteTo(snapshot)
                snapshot.Position = 0
                Using staged = SpreadsheetDocument.Open(snapshot, True, New OpenSettings With {.AutoSave = False})
                    Dim book = staged.WorkbookPart
                    SaveChangedParts(WorkbookPart, book, clearCache:=False)
                    edit(book)
                    For Each part In book.WorksheetParts
                        If ClearFormulaResults(GetSheetData(part), book.CellMetadataPart?.Metadata) Then SaveWorksheet(part.Worksheet, part)
                    Next
                    If book.CalculationChainPart IsNot Nothing Then book.DeletePart(book.CalculationChainPart)
                    If book.Workbook.CalculationProperties Is Nothing Then book.Workbook.CalculationProperties = New CalculationProperties()
                    book.Workbook.CalculationProperties.FullCalculationOnLoad = True
                    book.Workbook.Save()
                End Using
                Return snapshot
            Catch
                snapshot.Dispose()
                Throw
            End Try
        End Function

        Private NotInheritable Class StagedWorkbookChanges
            Implements IDisposable

            Private ReadOnly _changes As New List(Of (Workbook As OpenXmlExcelDataOperations, Snapshot As MemoryStream, Document As SpreadsheetDocument))

            Friend Sub Add(workbook As OpenXmlExcelDataOperations, edit As Action(Of WorkbookPart))
                Dim snapshot = workbook.StageSheetTransfer(edit)
                Try
                    snapshot.Position = 0
                    Dim document = SpreadsheetDocument.Open(snapshot, False, New OpenSettings With {.AutoSave = False})
                    _changes.Add((workbook, snapshot, document))
                Catch
                    snapshot.Dispose()
                    Throw
                End Try
            End Sub

            Friend Sub Apply()
                Dim applied As Integer
                Try
                    For Each change In _changes
                        change.Workbook.InstallTransferSnapshot(change.Snapshot, change.Document)
                        applied += 1
                    Next
                Finally
                    _changes.RemoveRange(0, applied)
                End Try
            End Sub

            Private Sub Dispose() Implements IDisposable.Dispose
                For Each change In _changes
                    change.Document.Dispose()
                    change.Snapshot.Dispose()
                Next
                _changes.Clear()
            End Sub
        End Class

        Private Function TransferWorkbookPath(allowUnsaved As Boolean) As Uri
            Dim path = WorkbookFilePath
            If String.IsNullOrEmpty(path) OrElse Not File.Exists(path) Then
                If allowUnsaved Then Return New Uri(TemporaryWorkbookName, UriKind.Relative)
                Throw UnsupportedOperation("Creating a workbook link to a workbook that has not been saved")
            End If
            Return WorkbookFileUri(path)
        End Function

        Friend Shared Function WorkbookFileUri(path As String) As Uri
            ' An explicit file URI survives package serialization on Windows and Unix.
            Return New Uri(New Uri(System.IO.Path.GetFullPath(path)).AbsoluteUri, UriKind.Absolute)
        End Function

        Private Function MatchesTransferPath(path As Uri) As Boolean
            If Not path.IsAbsoluteUri Then Return path.OriginalString = TemporaryWorkbookName
            Dim savedPath = WorkbookFilePath
            Return Not String.IsNullOrEmpty(savedPath) AndAlso
                SameTransferPath(WorkbookFileUri(savedPath), path)
        End Function

        Private Shared Function SameTransferPath(first As Uri, second As Uri) As Boolean
            If Not first.IsAbsoluteUri OrElse Not second.IsAbsoluteUri Then
                Return Not first.IsAbsoluteUri AndAlso Not second.IsAbsoluteUri AndAlso first.OriginalString = second.OriginalString
            End If
            If first.IsFile AndAlso second.IsFile Then
                Dim comparison = If(Path.DirectorySeparatorChar = "\"c, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
                Return String.Equals(first.LocalPath, second.LocalPath, comparison)
            End If
            Return first.Equals(second)
        End Function

        Private Function ReadTransferExternalReference(qualifier As String) As (Index As Integer, Path As Uri, Sheet As String, SheetNames As String())
            Dim close = qualifier.IndexOf("]"c)
            Dim index = Integer.Parse(qualifier.Substring(1, close - 1), CultureInfo.InvariantCulture)
            Dim sheet = qualifier.Substring(close + 1)
            Dim book = TransferExternalBook(WorkbookPart, index)
            Dim linkedPath = ResolveTransferLink(index, book)
            If linkedPath Is Nothing Then
                Throw UnsupportedOperation("Resolving a workbook link during a sheet transfer")
            End If
            Dim names = book.SheetNames?.Elements(Of SheetName)().Select(Function(item) item.Val.Value).ToArray()
            Return (index, linkedPath, sheet, If(names, {sheet}))
        End Function

        Private Shared Function TransferExternalBook(book As WorkbookPart, index As Integer) As ExternalBook
            Dim reference = TryCast(book.Workbook.ExternalReferences?.ChildElements(index - 1), ExternalReference)
            Dim part As OpenXmlPart = Nothing
            If reference?.Id Is Nothing OrElse Not book.TryGetPartById(reference.Id.Value, part) Then Return Nothing
            Return TryCast(part, ExternalWorkbookPart)?.ExternalLink?.GetFirstChild(Of ExternalBook)()
        End Function

        Private Function ResolveTransferLink(index As Integer, book As ExternalBook) As Uri
            If book?.Id Is Nothing Then Return Nothing
            Dim reference = DirectCast(WorkbookPart.Workbook.ExternalReferences.ChildElements(index - 1), ExternalReference)
            Dim part = DirectCast(WorkbookPart.GetPartById(reference.Id.Value), ExternalWorkbookPart)
            Dim relationship = part.ExternalRelationships.FirstOrDefault(Function(item) item.Id = book.Id.Value)
            If relationship Is Nothing Then Return Nothing
            If relationship.Uri.IsAbsoluteUri Then Return relationship.Uri
            Dim resolved = WorkbookSession?.ResolveTemporaryLink(relationship.Uri)
            If resolved IsNot Nothing Then Return resolved
            If String.IsNullOrEmpty(FilePath) Then Return Nothing
            Return New Uri(WorkbookFileUri(FilePath), relationship.Uri)
        End Function

        ' Collects required links without modifying the open workbook or changing existing link indexes.
        Private NotInheritable Class SheetTransferLinks
            Private ReadOnly _recipient As OpenXmlExcelDataOperations
            Private ReadOnly _existingCount As Integer
            Private _newCount As Integer
            Private ReadOnly _links As New List(Of (Path As Uri, Index As Integer, Sheets As List(Of String)))

            Friend Sub New(recipient As OpenXmlExcelDataOperations)
                _recipient = recipient
                _existingCount = If(recipient.WorkbookPart.Workbook.ExternalReferences?.ChildElements.Count, 0)
            End Sub

            Friend Function ReferenceTo(path As Uri, sheetNames As IEnumerable(Of String), sheet As String) As String
                Dim link = _links.Find(Function(item) SameTransferPath(item.Path, path))
                If link.Path Is Nothing Then
                    Dim index = FindIndex(path)
                    If index = 0 Then
                        _newCount += 1
                        index = _existingCount + _newCount
                    End If
                    link = (path, index, New List(Of String)(sheetNames))
                    _links.Add(link)
                End If
                If Not link.Sheets.Contains(sheet, StringComparer.OrdinalIgnoreCase) Then link.Sheets.Add(sheet)
                Dim qualifier = FormatSheetName(sheet)
                Dim prefix = "[" & link.Index.ToString(CultureInfo.InvariantCulture) & "]"
                Return If(qualifier.StartsWith("'", StringComparison.Ordinal), qualifier.Insert(1, prefix), prefix & qualifier)
            End Function

            Friend Function UsesIndex(index As Integer) As Boolean
                Return _links.Any(Function(link) link.Index = index)
            End Function

            Friend Sub ApplyTo(book As WorkbookPart)
                For Each link In _links
                    If link.Index > _existingCount Then
                        AddTransferLink(book, link.Path, link.Sheets)
                    Else
                        Dim external = TransferExternalBook(book, link.Index)
                        If external.SheetNames Is Nothing Then external.SheetNames = New SheetNames()
                        Dim known As New HashSet(Of String)(
                            external.SheetNames.Elements(Of SheetName)().Select(Function(sheet) sheet.Val.Value), StringComparer.OrdinalIgnoreCase)
                        For Each name In link.Sheets
                            If known.Add(name) Then external.SheetNames.Append(New SheetName With {.Val = name})
                        Next
                        DirectCast(external.Parent, ExternalLink).Save()
                    End If
                Next
            End Sub

            Private Function FindIndex(path As Uri) As Integer
                For index = 1 To _existingCount
                    Dim external = TransferExternalBook(_recipient.WorkbookPart, index)
                    Dim existing = _recipient.ResolveTransferLink(index, external)
                    If existing IsNot Nothing AndAlso SameTransferPath(existing, path) Then Return index
                Next
                Return 0
            End Function
        End Class

    End Class

End Namespace
