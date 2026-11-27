Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports MergeCollection = DocumentFormat.OpenXml.Spreadsheet.MergeCells

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <inheritdoc/>
        ''' <remarks>Clears values, formulas and cell formatting on supported worksheets. A merge or formula group must be wholly inside the range.</remarks>
        Public Overrides Sub ClearCells(overrideSheetName As String, rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)
            If rangeFirstCell Is Nothing Then Throw New ArgumentNullException(NameOf(rangeFirstCell))
            If rangeLastCell Is Nothing Then Throw New ArgumentNullException(NameOf(rangeLastCell))
            If Not String.Equals(rangeFirstCell.SheetName, rangeLastCell.SheetName, StringComparison.OrdinalIgnoreCase) Then
                Throw New ArgumentException("The range addresses must belong to the same worksheet.", NameOf(rangeLastCell))
            End If
            Dim sheetName = If(String.IsNullOrEmpty(overrideSheetName), rangeFirstCell.SheetName, overrideSheetName)
            Dim range = CellRectangle(sheetName, rangeFirstCell.RowIndex, rangeFirstCell.ColumnIndex, rangeLastCell.RowIndex, rangeLastCell.ColumnIndex)
            ClearCellRectangle(sheetName, range)
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Clears cells and formatting and preserves worksheet row heights, column widths and layout.</remarks>
        Public Overrides Sub ClearSheet(sheetName As String)
            ClearCellRectangle(sheetName, Nothing)
        End Sub

        Private Sub ClearCellRectangle(sheetName As String, range As ExcelRange)
            SyncLock WorkbookSaveLock
                Dim part = GetWorksheetPart(sheetName)
                ValidateCellLayoutEdit(part, range, "Clearing cells")
                CheckFormulaGroups(part, sheetName, range, allowEntireGroup:=True)
                Dim merges = part.Worksheet.GetFirstChild(Of MergeCollection)()
                Dim selectedMerges = If(merges?.Elements(Of MergeCell)(), Enumerable.Empty(Of MergeCell)()).
                    Where(Function(merge) range Is Nothing OrElse RectanglesOverlap(range, New ExcelRange(sheetName, merge.Reference.Value))).ToArray()
                If range IsNot Nothing AndAlso selectedMerges.Any(Function(merge) Not ContainsRectangle(range, New ExcelRange(sheetName, merge.Reference.Value))) Then
                    Throw New InvalidOperationException("A merged range must be cleared as a whole.")
                End If

                Dim data = GetSheetData(part)
                Dim columns = part.Worksheet.Elements(Of Columns)().SelectMany(Function(group) group.Elements(Of Column)()).ToArray()
                Dim rows = EnumerateWorksheetRows(data).ToArray()
                If range IsNot Nothing AndAlso (
                    columns.Any(Function(column) If(column.Style?.Value, 0UI) > 0 AndAlso
                        column.Max.Value > range.AddressStart.ColumnIndex AndAlso column.Min.Value <= range.AddressEnd.ColumnIndex + 1) OrElse
                    rows.Any(Function(row) If(row.Value.CustomFormat?.Value, False) AndAlso If(row.Value.StyleIndex?.Value, 0UI) > 0 AndAlso
                        row.Index >= range.AddressStart.RowIndex AndAlso row.Index <= range.AddressEnd.RowIndex)) Then
                    Throw UnsupportedOperation("Clearing part of a worksheet with inherited row or column formatting")
                End If

                For Each merge In selectedMerges
                    merge.Remove()
                Next
                If selectedMerges.Length > 0 Then UpdateMergeCount(merges)
                For Each entry In EnumerateWorksheetCells(data).ToArray()
                    If range IsNot Nothing AndAlso Not CellIsInRange(entry.RowIndex, entry.ColumnIndex, range) Then
                        If entry.Value.CellReference Is Nothing Then entry.Value.CellReference = ExcelCell.LocalCellAddress(entry.RowIndex, entry.ColumnIndex)
                        Continue For
                    End If
                    If entry.Value.CellFormula IsNot Nothing Then _calculationChainInvalid = True
                    entry.Value.Remove()
                    RecordCellEdit(sheetName, entry.RowIndex, entry.ColumnIndex, calculateOnEdit:=False)
                Next

                If range Is Nothing Then
                    For Each row In rows
                        row.Value.StyleIndex = Nothing
                        row.Value.CustomFormat = Nothing
                    Next
                    For Each column In columns
                        column.Style = Nothing
                    Next
                    If part.Worksheet.SheetDimension IsNot Nothing Then part.Worksheet.SheetDimension.Reference = "A1"
                End If
                FinishCellLayoutEdit(part)
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Keeps the upper-left cell value and clears other values in the merged range. Existing overlapping merges are rejected.</remarks>
        Public Overrides Sub MergeCells(sheetName As String, fromRowIndex As Integer, fromColumnIndex As Integer, toRowIndex As Integer, toColumnIndex As Integer)
            Dim range = CellRectangle(sheetName, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex)
            If range.AddressStart.Address = range.AddressEnd.Address Then Return
            SyncLock WorkbookSaveLock
                Dim part = GetWorksheetPart(sheetName)
                ValidateCellLayoutEdit(part, range, "Merging cells")
                CheckFormulaGroups(part, sheetName, range, allowEntireGroup:=False)
                Dim reference = range.AddressStart.Address & ":" & range.AddressEnd.Address
                Dim merges = part.Worksheet.GetFirstChild(Of MergeCollection)()
                If merges IsNot Nothing Then
                    For Each existing In merges.Elements(Of MergeCell)()
                        If String.Equals(reference, existing.Reference.Value, StringComparison.OrdinalIgnoreCase) Then Return
                        If RectanglesOverlap(range, New ExcelRange(sheetName, existing.Reference.Value)) Then
                            Throw New InvalidOperationException("The requested range overlaps an existing merge.")
                        End If
                    Next
                Else
                    merges = New MergeCollection()
                    part.Worksheet.AddChild(merges)
                End If

                For Each entry In EnumerateWorksheetCells(GetSheetData(part))
                    If entry.RowIndex = range.AddressStart.RowIndex AndAlso entry.ColumnIndex = range.AddressStart.ColumnIndex Then Continue For
                    If Not CellIsInRange(entry.RowIndex, entry.ColumnIndex, range) Then Continue For
                    Dim cell = entry.Value
                    If cell.CellFormula IsNot Nothing Then _calculationChainInvalid = True
                    cell.CellFormula = Nothing
                    cell.CellValue = Nothing
                    cell.InlineString = Nothing
                    cell.DataType = Nothing
                    RecordCellEdit(sheetName, entry.RowIndex, entry.ColumnIndex, calculateOnEdit:=False)
                Next
                merges.AppendChild(New MergeCell With {.Reference = reference})
                UpdateMergeCount(merges)
                UpdateDimension(part.Worksheet, sheetName, range.AddressEnd.RowIndex, range.AddressEnd.ColumnIndex)
                FinishCellLayoutEdit(part)
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub UnMergeCells(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            ValidateCoordinates(rowIndex, columnIndex)
            SyncLock WorkbookSaveLock
                Dim part = GetWorksheetPart(sheetName)
                Dim merges = part.Worksheet.GetFirstChild(Of MergeCollection)()
                Dim merged = merges?.Elements(Of MergeCell)().FirstOrDefault(
                    Function(item) CellIsInRange(rowIndex, columnIndex, New ExcelRange(sheetName, item.Reference.Value)))
                If merged Is Nothing Then Return
                ValidateCellLayoutEdit(part, New ExcelRange(sheetName, merged.Reference.Value), "Unmerging cells")
                merged.Remove()
                UpdateMergeCount(merges)
                _changedSheets.Add(part)
            End SyncLock
        End Sub

        Private Sub ValidateCellLayoutEdit(part As WorksheetPart, range As ExcelRange, operation As String)
            If If(part.Worksheet.GetFirstChild(Of SheetProtection)()?.Sheet?.Value, False) Then Throw UnsupportedOperation(operation & " in a protected worksheet")
            ValidateWorksheetStructure(part, Nothing, allowSpills:=HasSupportedSpillMetadata(FormulaCellMetadata))
            If part.TableDefinitionParts.Any(Function(table) range Is Nothing OrElse RectanglesOverlap(range, New ExcelRange(
                range.AddressStart.SheetName, table.Table.Reference.Value))) Then
                Throw UnsupportedOperation(operation & " inside an Excel table")
            End If
        End Sub

        Private Sub CheckFormulaGroups(part As WorksheetPart, sheetName As String, range As ExcelRange, allowEntireGroup As Boolean)
            If range Is Nothing Then Return
            For Each formula In GetSheetData(part).Descendants(Of CellFormula)()
                If formula.Reference Is Nothing OrElse formula.FormulaType Is Nothing OrElse formula.FormulaType.Value = CellFormulaValues.Normal Then Continue For
                Dim group As New ExcelRange(sheetName, formula.Reference.Value)
                If RectanglesOverlap(range, group) AndAlso Not (allowEntireGroup AndAlso ContainsRectangle(range, group)) Then
                    Throw UnsupportedOperation("Editing part of a shared, array or data-table formula range")
                End If
            Next
        End Sub

        Private Sub FinishCellLayoutEdit(part As WorksheetPart)
            InvalidateGroupedFormulaRanges(part)
            _changedSheets.Add(part)
            RecalculationRequired = True
            If _pendingCellEdits.Count > 0 Then
                Dim calculation As New FormulaCalculation(Me)
                calculation.RecalculateAfterEdit()
            End If
        End Sub

        Private Shared Sub UpdateMergeCount(merges As MergeCollection)
            Dim count = merges.Elements(Of MergeCell)().Count()
            If count = 0 Then
                merges.Remove()
            Else
                merges.Count = CUInt(count)
            End If
        End Sub

        Private Shared Function CellRectangle(sheetName As String, fromRow As Integer, fromColumn As Integer, toRow As Integer, toColumn As Integer) As ExcelRange
            ValidateCoordinates(fromRow, fromColumn)
            ValidateCoordinates(toRow, toColumn)
            Dim first = ExcelCell.LocalCellAddress(Math.Min(fromRow, toRow), Math.Min(fromColumn, toColumn))
            Dim last = ExcelCell.LocalCellAddress(Math.Max(fromRow, toRow), Math.Max(fromColumn, toColumn))
            Return New ExcelRange(sheetName, first & ":" & last)
        End Function

        Private Shared Function RectanglesOverlap(first As ExcelRange, second As ExcelRange) As Boolean
            Return first.AddressStart.RowIndex <= second.AddressEnd.RowIndex AndAlso second.AddressStart.RowIndex <= first.AddressEnd.RowIndex AndAlso
                first.AddressStart.ColumnIndex <= second.AddressEnd.ColumnIndex AndAlso second.AddressStart.ColumnIndex <= first.AddressEnd.ColumnIndex
        End Function

        Private Shared Function ContainsRectangle(outer As ExcelRange, inner As ExcelRange) As Boolean
            Return CellIsInRange(inner.AddressStart.RowIndex, inner.AddressStart.ColumnIndex, outer) AndAlso
                CellIsInRange(inner.AddressEnd.RowIndex, inner.AddressEnd.ColumnIndex, outer)
        End Function

    End Class

End Namespace
