Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private ReadOnly _changedTables As New Dictionary(Of TableDefinitionPart, Table)

        Private NotInheritable Class StructuralTable
            Friend ReadOnly Worksheet As WorksheetPart
            Friend ReadOnly SheetName As String
            Friend ReadOnly Part As TableDefinitionPart
            Friend ReadOnly Source As Table
            Friend ReadOnly First As ReferenceEndpoint
            Friend ReadOnly Last As ReferenceEndpoint
            Friend ReadOnly HeaderCount As Integer
            Friend ReadOnly TotalsCount As Integer

            Friend TargetFirstRow As Integer
            Friend TargetLastRow As Integer
            Friend TargetFirstColumn As Integer
            Friend TargetLastColumn As Integer
            Friend AddsDataRows As Boolean

            Friend Sub New(worksheet As WorksheetPart, sheetName As String, part As TableDefinitionPart, source As Table)
                Me.Worksheet = worksheet
                Me.SheetName = sheetName
                Me.Part = part
                Me.Source = source
                Dim area = ReadStructuralRange(source.Reference?.Value)
                First = area.First
                Last = area.Last
                TargetFirstRow = First.Row
                TargetLastRow = Last.Row
                TargetFirstColumn = First.Column
                TargetLastColumn = Last.Column
                Dim headers = If(source.HeaderRowCount?.Value, 1UI)
                Dim totals = If(source.TotalsRowCount?.Value, 0UI)
                If headers > 1UI OrElse totals > 1UI OrElse headers + totals > CUInt(Last.Row - First.Row + 1) Then
                    Throw New NotSupportedException("Updating references in a table with unsupported header or totals rows.")
                End If
                HeaderCount = CInt(headers)
                TotalsCount = CInt(totals)
            End Sub

            Friend ReadOnly Property FirstDataRow As Integer
                Get
                    Return First.Row + HeaderCount
                End Get
            End Property

            Friend ReadOnly Property LastDataRow As Integer
                Get
                    Return Last.Row - TotalsCount
                End Get
            End Property

            Friend Function FormatRange(lastRow As Integer) As String
                Return ExcelCell.LocalCellAddress(TargetFirstRow, TargetFirstColumn) & ":" &
                       ExcelCell.LocalCellAddress(lastRow, TargetLastColumn)
            End Function
        End Class

        Private Function GetStructuralTables(sheets As IEnumerable(Of Sheet), Optional book As WorkbookPart = Nothing) As List(Of StructuralTable)
            Dim result As New List(Of StructuralTable)
            book = If(book, WorkbookPart)
            For Each sheet In sheets
                Dim worksheet = TryCast(book.GetPartById(sheet.Id.Value), WorksheetPart)
                If worksheet Is Nothing Then Continue For
                For Each part In worksheet.TableDefinitionParts
                    Dim source As Table = Nothing
                    If Not _changedTables.TryGetValue(part, source) Then source = part.Table
                    If source Is Nothing OrElse part.Parts.Any() OrElse part.ExternalRelationships.Any() OrElse
                       source.ChildElements.Any(Function(child) Not (TypeOf child Is AutoFilter OrElse TypeOf child Is SortState OrElse
                                                         TypeOf child Is TableColumns OrElse TypeOf child Is TableStyleInfo)) Then
                        Throw UnsupportedOperation("Updating references in an extended or externally linked table")
                    End If
                    result.Add(New StructuralTable(worksheet, sheet.Name.Value, part, source))
                Next
            Next
            Return result
        End Function

        Private Shared Function GetStructuralTableNames(tables As IEnumerable(Of StructuralTable)) As HashSet(Of String)
            Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each table In tables
                If table.Source.Name IsNot Nothing Then names.Add(table.Source.Name.Value)
                If table.Source.DisplayName IsNot Nothing Then names.Add(table.Source.DisplayName.Value)
            Next
            Return names
        End Function

        Private Function StageTableChanges(tables As List(Of StructuralTable), editedPart As WorksheetPart, editedName As String,
                                           shift As StructuralShift, newSheetName As String, tableNames As HashSet(Of String),
                                           fills As List(Of (Row As Integer, Column As Integer, Formula As String)),
                                           Optional book As WorkbookPart = Nothing, Optional links As HashSet(Of Integer) = Nothing) As Dictionary(Of TableDefinitionPart, Table)
            If shift IsNot Nothing Then
                For Each table In tables
                    If table.Worksheet Is editedPart Then shift.PlanTableShift(table)
                Next
            End If

            Dim changes As New Dictionary(Of TableDefinitionPart, Table)
            For Each table In tables
                Dim replacement = DirectCast(table.Source.CloneNode(True), Table)
                Dim changed = table.TargetFirstRow <> table.First.Row OrElse table.TargetLastRow <> table.Last.Row OrElse
                              table.TargetFirstColumn <> table.First.Column OrElse table.TargetLastColumn <> table.Last.Column
                If changed Then
                    replacement.Reference = table.FormatRange(table.TargetLastRow)
                    Dim filter = replacement.GetFirstChild(Of AutoFilter)()
                    If replacement.GetFirstChild(Of SortState)() IsNot Nothing OrElse filter?.GetFirstChild(Of SortState)() IsNot Nothing Then
                        Throw UnsupportedOperation("Shifting a table with a saved sort state")
                    End If
                    If filter IsNot Nothing Then
                        filter.Reference = table.FormatRange(table.TargetLastRow - table.TotalsCount)
                    End If
                End If

                Dim columns = replacement.GetFirstChild(Of TableColumns)()?.Elements(Of TableColumn)().ToArray()
                If columns Is Nothing OrElse columns.Length <> table.Last.Column - table.First.Column + 1 Then
                    Throw New InvalidDataException("The table's columns do not match its range.")
                End If
                For columnIndex = 0 To columns.Length - 1
                    Dim column = columns(columnIndex)
                    Dim calculation = column.CalculatedColumnFormula
                    If calculation IsNot Nothing Then
                        Dim original = calculation.Text
                        Dim rewritten = RewriteFormula(original, table.SheetName, editedName, shift, newSheetName, False, tableNames,
                                                       formulaWorkbook:=book, structuralExternalLinks:=links)
                        If table.AddsDataRows Then
                            If calculation.GetAttributes().Any(Function(attribute) attribute.LocalName = "array" AndAlso
                                                               (attribute.Value = "1" OrElse attribute.Value = "true")) Then
                                Throw UnsupportedOperation("Extending an array calculated column")
                            End If
                            ' Excel takes the template at the leading edge and copies relative rows for the added cells.
                            Dim fillTemplate = If(shift.Start = table.FirstDataRow, original, rewritten)
                            For row = shift.Start To shift.Start + shift.Count - 1
                                Dim formula = RewriteFormula(fillTemplate, table.SheetName, editedName, Nothing, Nothing, False,
                                                             tableNames, copyRowOffset:=row - table.FirstDataRow)
                                fills.Add((row, table.First.Column + columnIndex, formula))
                                If row = table.FirstDataRow Then rewritten = formula
                            Next
                        End If
                        If rewritten <> original Then
                            calculation.Text = rewritten
                            changed = True
                        End If
                    End If

                    Dim totals = column.TotalsRowFormula
                    If totals IsNot Nothing Then
                        Dim text = RewriteFormula(totals.Text, table.SheetName, editedName, shift, newSheetName, False, tableNames,
                                                  formulaWorkbook:=book, structuralExternalLinks:=links)
                        If text <> totals.Text Then
                            totals.Text = text
                            changed = True
                        End If
                    End If
                Next
                If changed Then changes.Add(table.Part, replacement)
            Next
            Return changes
        End Function

        Private Sub SaveChangedTables(source As WorkbookPart, target As WorkbookPart)
            For Each change In _changedTables
                Dim owner = change.Key.GetParentParts().OfType(Of WorksheetPart)().Single()
                Dim targetSheet = DirectCast(target.GetPartById(source.GetIdOfPart(owner)), WorksheetPart)
                Dim targetTable = DirectCast(targetSheet.GetPartById(owner.GetIdOfPart(change.Key)), TableDefinitionPart)
                Using stream = targetTable.GetStream(FileMode.Create, FileAccess.Write)
                    change.Value.Save(stream)
                End Using
            Next
        End Sub

    End Class

End Namespace
