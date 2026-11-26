Imports System.IO
Imports DocumentFormat.OpenXml
Imports System.Globalization
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
            Friend RemovesFirstDataRow As Boolean
            Friend RemovesTotalsRow As Boolean
            Friend ClearsDataBody As Boolean
            Friend ClearsFilterCriteria As Boolean
            Friend ReadOnly ReappliedFilters As New List(Of (Column As Integer, Match As Func(Of Double, Boolean)))
            Friend ColumnEditOffset As Integer = -1
            Friend ReadOnly DeletedColumnNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

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
                                           fills As List(Of (Row As Integer, Column As Integer, Formula As String, Header As String)),
                                           Optional book As WorkbookPart = Nothing, Optional links As HashSet(Of Integer) = Nothing,
                                           Optional allowSpillReferences As Boolean = False) As Dictionary(Of TableDefinitionPart, Table)
            If shift IsNot Nothing Then
                For Each table In tables
                    If table.Worksheet Is editedPart Then shift.PlanTableShift(table)
                Next
            End If

            Dim changes As New Dictionary(Of TableDefinitionPart, Table)
            For Each table In tables
                Dim replacement = DirectCast(table.Source.CloneNode(True), Table)
                Dim changed = table.TargetFirstRow <> table.First.Row OrElse table.TargetLastRow <> table.Last.Row OrElse
                              table.TargetFirstColumn <> table.First.Column OrElse table.TargetLastColumn <> table.Last.Column OrElse
                              table.ClearsDataBody
                If changed Then
                    replacement.Reference = table.FormatRange(table.TargetLastRow)
                    If table.ClearsDataBody Then replacement.InsertRow = True
                    If table.RemovesTotalsRow Then
                        replacement.TotalsRowShown = False
                        replacement.TotalsRowCount = Nothing
                    End If
                    Dim filter = replacement.GetFirstChild(Of AutoFilter)()
                    If replacement.GetFirstChild(Of SortState)() IsNot Nothing OrElse filter?.GetFirstChild(Of SortState)() IsNot Nothing Then
                        Throw UnsupportedOperation("Shifting a table with a saved sort state")
                    End If
                    If filter IsNot Nothing Then
                        filter.Reference = table.FormatRange(table.TargetLastRow - If(table.RemovesTotalsRow, 0, table.TotalsCount))
                    End If
                End If

                Dim columns = replacement.GetFirstChild(Of TableColumns)()?.Elements(Of TableColumn)().ToArray()
                If columns Is Nothing OrElse columns.Length <> table.Last.Column - table.First.Column + 1 Then
                    Throw New InvalidDataException("The table's columns do not match its range.")
                End If
                If table.ColumnEditOffset >= 0 Then
                    EditTableColumnFilters(replacement.GetFirstChild(Of AutoFilter)(), table, shift)
                    EditTableColumns(replacement.GetFirstChild(Of TableColumns)(), columns, table, shift, fills)
                    columns = replacement.GetFirstChild(Of TableColumns)().Elements(Of TableColumn)().ToArray()
                End If
                For columnIndex = 0 To columns.Length - 1
                    Dim column = columns(columnIndex)
                    Dim calculation = column.CalculatedColumnFormula
                    ' Excel keeps the template when deleting its first data row. The remaining cell formulas still move.
                    If calculation IsNot Nothing AndAlso Not table.RemovesFirstDataRow Then
                        Dim original = calculation.Text
                        Dim rewritten = RewriteFormula(original, table.SheetName, editedName, shift, newSheetName, False, tableNames,
                                                       formulaWorkbook:=book, structuralExternalLinks:=links, contextualTable:=table,
                                                       allowSpillReferences:=allowSpillReferences)
                        If table.AddsDataRows Then
                            If calculation.GetAttributes().Any(Function(attribute) attribute.LocalName = "array" AndAlso
                                                               (attribute.Value = "1" OrElse attribute.Value = "true")) Then
                                Throw UnsupportedOperation("Extending an array calculated column")
                            End If
                            ' Excel takes the template at the leading edge and copies relative rows for the added cells.
                            Dim fillTemplate = If(shift.Start = table.FirstDataRow, original, rewritten)
                            For row = shift.Start To shift.Start + shift.Count - 1
                                Dim formula = RewriteFormula(fillTemplate, table.SheetName, editedName, Nothing, Nothing, False,
                                                             tableNames, copyRowOffset:=row - table.FirstDataRow, allowSpillReferences:=allowSpillReferences)
                                fills.Add((row, table.First.Column + columnIndex, formula, Nothing))
                                If row = table.FirstDataRow Then rewritten = formula
                            Next
                        End If
                        If rewritten <> original Then
                            calculation.Text = rewritten
                            changed = True
                        End If
                    End If

                    Dim totals = column.TotalsRowFormula
                    If table.RemovesTotalsRow Then
                        column.TotalsRowLabel = Nothing
                        column.TotalsRowFunction = Nothing
                        column.TotalsRowFormula = Nothing
                    ElseIf totals IsNot Nothing Then
                        Dim text = RewriteFormula(totals.Text, table.SheetName, editedName, shift, newSheetName, False, tableNames,
                                                  formulaWorkbook:=book, structuralExternalLinks:=links, contextualTable:=table,
                                                  allowSpillReferences:=allowSpillReferences)
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

        Private Shared Sub EditTableColumns(definition As TableColumns, columns As TableColumn(), table As StructuralTable,
                                            shift As StructuralShift, fills As List(Of (Row As Integer, Column As Integer, Formula As String, Header As String)))
            If shift.Inserting Then
                Dim names As New HashSet(Of String)(columns.Select(Function(column) column.Name.Value), StringComparer.OrdinalIgnoreCase)
                Dim anchor = columns(table.ColumnEditOffset)
                Dim id = columns.Max(Function(column) column.Id.Value)
                Dim number = 1
                Dim prefix = If(CultureInfo.CurrentUICulture.TwoLetterISOLanguageName = "de", "Spalte", "Column")
                For index = 0 To shift.Count - 1
                    Dim name As String
                    Do
                        name = prefix & number.ToString(CultureInfo.InvariantCulture)
                        number += 1
                    Loop Until names.Add(name)
                    id += 1UI
                    definition.InsertBefore(New TableColumn With {.Id = id, .Name = name}, anchor)
                    If table.HeaderCount > 0 Then fills.Add((table.First.Row, shift.Start + index, Nothing, name))
                Next
            Else
                For Each removed In columns.Skip(table.ColumnEditOffset).Take(shift.Count)
                    removed.Remove()
                Next
            End If
            definition.Count = CUInt(definition.ChildElements.Count)
        End Sub

        Private Sub EditTableColumnFilters(filter As AutoFilter, table As StructuralTable, shift As StructuralShift)
            If filter Is Nothing Then Return
            If filter.ChildElements.Any(Function(child) Not TypeOf child Is FilterColumn) Then
                Throw UnsupportedOperation("Shifting Excel table columns with extended filter settings")
            End If
            Dim columns = filter.Elements(Of FilterColumn)().ToArray()
            Dim deletedCriteria = False
            For Each column In columns
                If column.ColumnId Is Nothing OrElse column.ColumnId.Value > CUInt(table.Last.Column - table.First.Column) Then
                    Throw New InvalidDataException("A table filter references an invalid column.")
                End If
                If column.ChildElements.Any(Function(child) TypeOf child Is ExtensionList OrElse TypeOf child Is OpenXmlUnknownElement) Then
                    Throw UnsupportedOperation("Shifting Excel table columns with extended filter criteria")
                End If
                Dim offset = CInt(column.ColumnId.Value)
                If offset < table.ColumnEditOffset Then Continue For
                If shift.Inserting Then
                    column.ColumnId = CUInt(offset + shift.Count)
                ElseIf offset >= table.ColumnEditOffset + shift.Count Then
                    column.ColumnId = CUInt(offset - shift.Count)
                Else
                    deletedCriteria = deletedCriteria OrElse column.HasChildren
                    column.Remove()
                End If
            Next
            If deletedCriteria Then
                For Each column In filter.Elements(Of FilterColumn)().Where(Function(item) item.HasChildren)
                    table.ReappliedFilters.Add((table.TargetFirstColumn + CInt(column.ColumnId.Value), ReadNumericFilter(column)))
                Next
            End If
            table.ClearsFilterCriteria = deletedCriteria
        End Sub

        Private Function ReadNumericFilter(column As FilterColumn) As Func(Of Double, Boolean)
            Dim filters = column.GetFirstChild(Of CustomFilters)()
            Dim conditions = filters?.Elements(Of CustomFilter)().ToArray()
            If column.ChildElements.Count <> 1 OrElse conditions Is Nothing OrElse conditions.Length = 0 OrElse
               conditions.Length > 2 OrElse filters.ChildElements.Count <> conditions.Length Then
                Throw UnsupportedOperation("Reapplying a table filter other than a simple numeric custom filter")
            End If
            Dim first = ReadNumericFilterCondition(conditions(0))
            If conditions.Length = 1 Then Return first
            Dim second = ReadNumericFilterCondition(conditions(1))
            If If(filters.[And]?.Value, False) Then Return Function(value) first(value) AndAlso second(value)
            Return Function(value) first(value) OrElse second(value)
        End Function

        Private Function ReadNumericFilterCondition(condition As CustomFilter) As Func(Of Double, Boolean)
            Dim threshold As Double
            If Not Double.TryParse(condition.Val?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, threshold) OrElse
               Double.IsNaN(threshold) OrElse Double.IsInfinity(threshold) Then
                Throw UnsupportedOperation("Reapplying a non-numeric custom table filter")
            End If
            Select Case If(condition.Operator?.Value, FilterOperatorValues.Equal)
                Case FilterOperatorValues.Equal
                    Return Function(value) value = threshold
                Case FilterOperatorValues.NotEqual
                    Return Function(value) value <> threshold
                Case FilterOperatorValues.LessThan
                    Return Function(value) value < threshold
                Case FilterOperatorValues.LessThanOrEqual
                    Return Function(value) value <= threshold
                Case FilterOperatorValues.GreaterThan
                    Return Function(value) value > threshold
                Case FilterOperatorValues.GreaterThanOrEqual
                    Return Function(value) value >= threshold
                Case Else
                    Throw UnsupportedOperation("Reapplying a custom table filter with an unsupported operator")
            End Select
        End Function

        Private Sub ApplyTableFilterVisibility(data As SheetData, tables As IEnumerable(Of StructuralTable))
            Dim affected = tables.Where(Function(item) item.ClearsFilterCriteria).ToArray()
            For Each table In affected
                If table.ReappliedFilters.Count > 0 AndAlso affected.Any(
                    Function(other) other IsNot table AndAlso other.FirstDataRow <= table.LastDataRow AndAlso table.FirstDataRow <= other.LastDataRow) Then
                    Throw UnsupportedOperation("Reapplying table filters over overlapping worksheet rows in the same edit")
                End If
                Dim dataRows = 0
                For Each row In data.Elements(Of Row)()
                    Dim index = CInt(row.RowIndex.Value) - 1
                    If index < table.FirstDataRow OrElse index > table.LastDataRow Then Continue For
                    dataRows += 1
                    Dim visible = table.ReappliedFilters.All(
                        Function(filter) filter.Match(ReadNumericFilterOperand(row, index, filter.Column)))
                    row.Hidden = If(visible, Nothing, New BooleanValue(True))
                Next
                If table.ReappliedFilters.Count > 0 AndAlso dataRows <> table.LastDataRow - table.FirstDataRow + 1 Then
                    Throw UnsupportedOperation("Reapplying a numeric table filter to a table with missing data rows")
                End If
            Next
        End Sub

        Private Function ReadNumericFilterOperand(row As Row, index As Integer, column As Integer) As Double
            Dim address = ExcelCell.LocalCellAddress(index, column)
            Dim cell = row.Elements(Of Cell)().FirstOrDefault(Function(item) item.CellReference.Value = address)
            Dim value As Double
            If cell Is Nothing OrElse cell.CellFormula IsNot Nothing OrElse
               If(row.StyleIndex?.Value, 0UI) <> 0UI OrElse If(cell.StyleIndex?.Value, 0UI) <> 0UI OrElse
               If(cell.DataType?.Value, CellValues.Number) <> CellValues.Number OrElse
               Not Double.TryParse(cell.CellValue?.Text, NumberStyles.Float, CultureInfo.InvariantCulture, value) OrElse
               Double.IsNaN(value) OrElse Double.IsInfinity(value) Then
                Throw UnsupportedOperation("Reapplying a numeric table filter to anything other than unformatted numeric constants")
            End If
            Return value
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
