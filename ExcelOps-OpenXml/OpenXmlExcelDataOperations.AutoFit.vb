Imports System.Globalization
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NumberFormat = ExcelNumberFormat.NumberFormat

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <inheritdoc/>
        ''' <remarks>Estimates column widths from formatted cell text using a portable text-width model.</remarks>
        Public Overrides Sub AutoFitColumns(sheetName As String)
            AutoFitColumnsCore(sheetName, Nothing, Nothing)
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub AutoFitColumns(sheetName As String, minimumWidth As Double)
            AutoFitColumnsCore(sheetName, Nothing, minimumWidth)
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub AutoFitColumns(sheetName As String, columnIndex As Integer)
            ValidateCoordinates(0, columnIndex)
            AutoFitColumnsCore(sheetName, columnIndex, Nothing)
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub AutoFitColumns(sheetName As String, columnIndex As Integer, minimumWidth As Double)
            ValidateCoordinates(0, columnIndex)
            AutoFitColumnsCore(sheetName, columnIndex, minimumWidth)
        End Sub

        Private Sub AutoFitColumnsCore(sheetName As String, selectedColumn As Integer?, minimumWidth As Double?)
            If minimumWidth.HasValue AndAlso (Double.IsNaN(minimumWidth.Value) OrElse minimumWidth.Value < 0 OrElse minimumWidth.Value > 255) Then
                Throw New ArgumentOutOfRangeException(NameOf(minimumWidth))
            End If

            SyncLock WorkbookSaveLock
                Dim part = GetWorksheetPart(sheetName)
                Dim worksheet = part.Worksheet
                Dim columns = worksheet.Elements(Of Columns)().ToArray()
                If columns.Length > 1 Then Throw UnsupportedOperation("Fitting worksheets with multiple column layout groups")
                Dim original = columns.FirstOrDefault()
                Dim layouts = ReadColumnLayouts(original)
                Dim styles = WorkbookPart.WorkbookStylesPart?.Stylesheet
                Dim fonts = If(styles?.Fonts?.Elements(Of Font)().ToArray(), Array.Empty(Of Font)())
                Dim normalStyle = GetCellFormat(styles, 0UI)
                Dim normalSize = If(fonts.ElementAtOrDefault(CInt(If(normalStyle?.FontId?.Value, 0UI)))?.FontSize?.Val?.Value, 11.0)
                Dim merges = MergedCells(sheetName)
                Dim rowElements = EnumerateWorksheetRows(GetSheetData(part)).ToDictionary(Function(row) row.Index, Function(row) row.Value)
                Dim formats As New Dictionary(Of UInteger, NumberFormat)
                Dim fitted As New Dictionary(Of Integer, Double)
                Dim minimum = If(minimumWidth, If(worksheet.SheetFormatProperties?.DefaultColumnWidth?.Value, 8.43))

                For Each entry In EnumerateWorksheetCells(GetSheetData(part))
                    Dim index = entry.ColumnIndex
                    If (selectedColumn.HasValue AndAlso index <> selectedColumn.Value) OrElse If(layouts(index)?.Hidden?.Value, False) Then Continue For
                    If Not fitted.ContainsKey(index) Then fitted.Add(index, minimum)
                    Dim styleIndex = WorksheetCellStyleIndex(worksheet, rowElements(entry.RowIndex), entry.Value, index)
                    Dim style = GetCellFormat(styles, styleIndex)
                    If If(style?.Alignment?.WrapText?.Value, False) OrElse
                        merges.Any(Function(range) CellIsInRange(entry.RowIndex, index, range)) Then Continue For
                    Dim value = ReadCellValue(sheetName, entry.RowIndex, index, entry.Value)
                    Dim display = FormatCellValue(value, GetCachedNumberFormat(styles, styleIndex, formats))
                    If display.Length = 0 Then Continue For
                    Dim font = fonts.ElementAtOrDefault(CInt(If(style?.FontId?.Value, 0UI)))
                    Dim fontRatio = If(font?.FontSize?.Val?.Value, normalSize) / If(normalSize > 0.0, normalSize, 11.0)
                    Dim bold = If(font?.Bold IsNot Nothing AndAlso If(font.Bold.Val?.Value, True), 1.1, 1.0)
                    Dim indent = If(style?.Alignment?.Indent?.Value, 0UI) * 2.0
                    Dim width = PortableTextWidth(display) * fontRatio * bold + indent + 0.75
                    fitted(index) = Math.Min(255.0, Math.Max(fitted(index), Math.Ceiling(width * 256.0) / 256.0))
                Next

                If fitted.Count = 0 Then Return
                For Each width In fitted
                    Dim column = If(TryCast(layouts(width.Key)?.CloneNode(False), Column), New Column())
                    column.Width = width.Value
                    column.CustomWidth = True
                    column.BestFit = True
                    layouts(width.Key) = column
                Next

                Dim updated = BuildColumnLayouts(original, layouts)
                If original Is Nothing Then
                    worksheet.AddChild(updated)
                Else
                    worksheet.ReplaceChild(updated, original)
                End If
                _changedSheets.Add(part)
            End SyncLock
        End Sub

        Private Shared Function BuildColumnLayouts(original As Columns, layouts As Column()) As Columns
            Dim updated = If(TryCast(original?.CloneNode(False), Columns), New Columns())
            Dim previous As Column = Nothing
            For index = 0 To layouts.Length - 1
                Dim column = layouts(index)
                If column Is Nothing Then
                    previous = Nothing
                ElseIf previous IsNot Nothing AndAlso SameColumnLayout(previous, column) Then
                    previous.Max = CUInt(index + 1)
                Else
                    previous = DirectCast(column.CloneNode(False), Column)
                    previous.Min = CUInt(index + 1)
                    previous.Max = previous.Min.Value
                    updated.AppendChild(previous)
                End If
            Next
            Return updated
        End Function

        Private Shared Function PortableTextWidth(text As String) As Double
            Dim longest As Double
            For Each line In text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(ChrW(10))
                Dim width As Double
                For Each character In line
                    If Char.IsLowSurrogate(character) OrElse Char.GetUnicodeCategory(character) = UnicodeCategory.NonSpacingMark Then Continue For
                    If Char.IsHighSurrogate(character) OrElse CUInt(AscW(character) And &HFFFF) >= &H2E80UI Then
                        width += 2.0
                    ElseIf "ilI.,':;!| ".IndexOf(character) >= 0 Then
                        width += 0.5
                    ElseIf "mMwW@%".IndexOf(character) >= 0 Then
                        width += 1.4
                    Else
                        width += 1.0
                    End If
                Next
                longest = Math.Max(longest, width)
            Next
            Return longest
        End Function

    End Class

End Namespace
