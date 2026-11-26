Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports ThreadedComment = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedComment

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private _recalculationLastCells As Dictionary(Of WorksheetPart, (RowIndex As Integer, ColumnIndex As Integer))

        ''' <inheritdoc/>
        ''' <remarks>Uses the current worksheet data and formatting. Changes are visible before saving.
        ''' During a workbook or worksheet recalculation, the bounds are reused per worksheet until the run ends.</remarks>
        Public Overrides Function LookupLastCell(sheetName As String) As ExcelCell
            SyncLock WorkbookSaveLock
                Dim part = GetWorksheetPart(sheetName)
                Dim bounds As (RowIndex As Integer, ColumnIndex As Integer)
                If _recalculationLastCells Is Nothing OrElse Not _recalculationLastCells.TryGetValue(part, bounds) Then
                    bounds = FindLastUsedCoordinates(part)
                    _recalculationLastCells?.Add(part, bounds)
                End If
                Return New ExcelCell(sheetName, bounds.RowIndex, bounds.ColumnIndex, ExcelCell.ValueTypes.All)
            End SyncLock
        End Function

        Private Shared Function FindLastUsedCoordinates(part As WorksheetPart) As (RowIndex As Integer, ColumnIndex As Integer)
            Dim lastRow = 0
            Dim lastColumn = 0
            Dim data = GetSheetData(part)
            For Each item In EnumerateWorksheetCells(data)
                Dim cell = item.Value
                If cell.CellFormula Is Nothing AndAlso cell.CellValue Is Nothing AndAlso cell.InlineString Is Nothing AndAlso
                   If(cell.StyleIndex?.Value, 0UI) = 0UI Then Continue For
                lastRow = Math.Max(lastRow, item.RowIndex)
                lastColumn = Math.Max(lastColumn, item.ColumnIndex)
            Next

            For Each item In EnumerateWorksheetRows(data)
                Dim row = item.Value
                If row.Height IsNot Nothing OrElse If(row.Hidden?.Value, False) OrElse
                   (If(row.CustomFormat?.Value, False) AndAlso If(row.StyleIndex?.Value, 0UI) <> 0UI) Then
                    lastRow = Math.Max(lastRow, item.Index)
                End If
            Next

            For Each group In part.Worksheet.Elements(Of Columns)()
                For Each column In group.Elements(Of Column)()
                    If column.Min Is Nothing OrElse column.Max Is Nothing OrElse If(column.Style?.Value, 0UI) = 0UI Then Continue For
                    ' Excel stores a multi-column style reaching XFD as a default for the remaining columns.
                    Dim last = If(column.Max.Value = ColumnCount AndAlso column.Min.Value < ColumnCount, column.Min.Value - 1UI, column.Max.Value)
                    lastColumn = Math.Max(lastColumn, CInt(last) - 1)
                Next
            Next

            For Each reference In UsedRangeReferences(part)
                Dim address As New ExcelCell(Nothing, Tools.LookupCellAddresFromRange(reference, Tools.LookupCellAddresFromRangeMode.LastCell), ExcelCell.ValueTypes.All)
                lastRow = Math.Max(lastRow, address.RowIndex)
                lastColumn = Math.Max(lastColumn, address.ColumnIndex)
            Next
            Return (lastRow, lastColumn)
        End Function

        Private Shared Iterator Function UsedRangeReferences(part As WorksheetPart) As IEnumerable(Of String)
            For Each group In part.Worksheet.Elements(Of MergeCells)()
                For Each merge In group.Elements(Of MergeCell)()
                    Yield merge.Reference.Value
                Next
            Next
            For Each group In part.Worksheet.Elements(Of Hyperlinks)()
                For Each link In group.Elements(Of Hyperlink)()
                    Yield link.Reference.Value
                Next
            Next
            Dim notes = part.WorksheetCommentsPart?.Comments?.GetFirstChild(Of CommentList)()
            If notes IsNot Nothing Then
                For Each note In notes.Elements(Of Comment)()
                    Yield note.Reference.Value
                Next
            End If
            For Each commentsPart In part.WorksheetThreadedCommentsParts
                For Each comment In commentsPart.ThreadedComments.Elements(Of ThreadedComment)()
                    Yield comment.Ref.Value
                Next
            Next
        End Function

    End Class

End Namespace
