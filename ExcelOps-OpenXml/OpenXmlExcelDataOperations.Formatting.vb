Imports System.Globalization
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations
        Private ReadOnly Property CurrentStyles As Stylesheet
            Get
                Return If(_editedStyles, WorkbookPart.WorkbookStylesPart?.Stylesheet)
            End Get
        End Property

        <Flags>
        Private Enum RangeEdges
            None = 0
            Left = 1
            Right = 2
            Top = 4
            Bottom = 8
        End Enum

        ''' <inheritdoc/>
        Public Sub ApplyCellFormatting(range As ExcelRange, formatting As ExcelCellFormatting) Implements IExcelWorksheetFormatting.ApplyCellFormatting
            If range Is Nothing Then Throw New ArgumentNullException(NameOf(range))
            If formatting Is Nothing Then Throw New ArgumentNullException(NameOf(formatting))
            Dim first = range.AddressStart
            Dim last = range.AddressEnd
            ValidateAddress(first)
            ValidateAddress(last)
            ValidateCoordinates(first.RowIndex, first.ColumnIndex)
            ValidateCoordinates(last.RowIndex, last.ColumnIndex)
            If Not String.Equals(first.SheetName, last.SheetName, StringComparison.OrdinalIgnoreCase) OrElse
               first.RowIndex > last.RowIndex OrElse first.ColumnIndex > last.ColumnIndex Then
                Throw New ArgumentException("Expected a rectangular range on one worksheet.", NameOf(range))
            End If
            ValidateCellFormatting(formatting)
            Dim changesFont = formatting.FontName IsNot Nothing OrElse formatting.FontSize.HasValue OrElse
                formatting.Bold.HasValue OrElse formatting.Italic.HasValue OrElse formatting.FontColor.HasValue
            Dim changesAlignment = formatting.HorizontalAlignment.HasValue OrElse formatting.VerticalAlignment.HasValue OrElse formatting.WrapText.HasValue
            Dim changesBorders = formatting.OutsideBorder IsNot Nothing OrElse formatting.InsideBorder IsNot Nothing

            SyncLock WorkbookSaveLock
                Dim part = GetWorksheetPart(first.SheetName)
                If Not (changesFont OrElse changesAlignment OrElse changesBorders OrElse
                        formatting.FillColor.HasValue OrElse formatting.NumberFormat IsNot Nothing) Then Return

                Dim styles = If(DirectCast(CurrentStyles?.CloneNode(True), Stylesheet), New Stylesheet())
                InitializeCellFormats(styles)
                Dim changedFormats As New Dictionary(Of (Style As UInteger, Edges As RangeEdges), UInteger)
                Dim data = GetSheetData(part)
                Dim rows = EnumerateWorksheetRows(data).GetEnumerator()
                Dim hasRow = rows.MoveNext()
                For rowIndex = first.RowIndex To last.RowIndex
                    While hasRow AndAlso rows.Current.Index < rowIndex
                        hasRow = rows.MoveNext()
                    End While
                    Dim row As Row
                    If hasRow AndAlso rows.Current.Index = rowIndex Then
                        row = rows.Current.Value
                    Else
                        row = data.InsertBefore(New Row With {.RowIndex = CUInt(rowIndex + 1)}, If(hasRow, rows.Current.Value, Nothing))
                    End If
                    For columnIndex = first.ColumnIndex To last.ColumnIndex
                        Dim cell = FindCell(row, rowIndex, columnIndex, True)
                        Dim oldStyle = WorksheetCellStyleIndex(part.Worksheet, row, cell, columnIndex)
                        Dim edges = RangeEdges.None
                        If changesBorders Then
                            If columnIndex = first.ColumnIndex Then edges = edges Or RangeEdges.Left
                            If columnIndex = last.ColumnIndex Then edges = edges Or RangeEdges.Right
                            If rowIndex = first.RowIndex Then edges = edges Or RangeEdges.Top
                            If rowIndex = last.RowIndex Then edges = edges Or RangeEdges.Bottom
                        End If
                        Dim index As UInteger
                        If Not changedFormats.TryGetValue((oldStyle, edges), index) Then
                            Dim template = If(GetCellFormat(styles, oldStyle), GetCellFormat(styles, 0UI))
                            Dim updated = DirectCast(template.CloneNode(True), CellFormat)
                            ApplyFormatChanges(styles, updated, formatting, changesFont, changesAlignment, changesBorders, edges)
                            index = WorksheetStyleTransfer.AppendStyleRecord(styles.CellFormats, updated)
                            changedFormats.Add((oldStyle, edges), index)
                        End If
                        cell.StyleIndex = index
                    Next
                Next
                UpdateDimension(part.Worksheet, first.SheetName, first.RowIndex, first.ColumnIndex)
                UpdateDimension(part.Worksheet, first.SheetName, last.RowIndex, last.ColumnIndex)
                _editedStyles = styles
                _changedSheets.Add(part)
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        Public Sub SetColumnWidth(sheetName As String, columnIndex As Integer, width As Double) Implements IExcelWorksheetFormatting.SetColumnWidth
            ValidateCoordinates(0, columnIndex)
            If Double.IsNaN(width) OrElse width < 0 OrElse width > 255 Then Throw New ArgumentOutOfRangeException(NameOf(width))
            SyncLock WorkbookSaveLock
                Dim part = GetWorksheetPart(sheetName)
                Dim columns = part.Worksheet.Elements(Of Columns)().ToArray()
                If columns.Length > 1 Then Throw UnsupportedOperation("Setting widths in worksheets with multiple column layout groups")
                Dim original = columns.SingleOrDefault()
                Dim layouts = ReadColumnLayouts(original)
                Dim column = If(DirectCast(layouts(columnIndex)?.CloneNode(False), Column), New Column())
                column.Width = width
                column.CustomWidth = True
                column.BestFit = Nothing
                layouts(columnIndex) = column
                Dim updated = BuildColumnLayouts(original, layouts)
                If original Is Nothing Then
                    part.Worksheet.AddChild(updated)
                Else
                    part.Worksheet.ReplaceChild(updated, original)
                End If
                _changedSheets.Add(part)
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        Public Sub SetRowHeight(sheetName As String, rowIndex As Integer, height As Double) Implements IExcelWorksheetFormatting.SetRowHeight
            ValidateCoordinates(rowIndex, 0)
            If Double.IsNaN(height) OrElse height < 0 OrElse height > 409 Then Throw New ArgumentOutOfRangeException(NameOf(height))
            SyncLock WorkbookSaveLock
                Dim part = GetWorksheetPart(sheetName)
                Dim row = FindRow(GetSheetData(part), rowIndex, True)
                row.Height = height
                row.CustomHeight = True
                _changedSheets.Add(part)
            End SyncLock
        End Sub

        Private Shared Sub ValidateCellFormatting(formatting As ExcelCellFormatting)
            If formatting.NumberFormat IsNot Nothing Then
                If Not New ExcelNumberFormat.NumberFormat(formatting.NumberFormat).IsValid Then
                    Throw New ArgumentException("Expected an Excel number format code.", NameOf(formatting))
                End If
                Xml.XmlConvert.VerifyXmlChars(formatting.NumberFormat)
            End If
            If formatting.FontName IsNot Nothing Then
                If String.IsNullOrWhiteSpace(formatting.FontName) Then Throw New ArgumentException("Expected a font name.", NameOf(formatting))
                Xml.XmlConvert.VerifyXmlChars(formatting.FontName)
            End If
            If formatting.FontSize.HasValue AndAlso
               (Double.IsNaN(formatting.FontSize.Value) OrElse formatting.FontSize.Value < 1 OrElse formatting.FontSize.Value > 409) Then
                Throw New ArgumentOutOfRangeException(NameOf(formatting), "Font size must be between 1 and 409 points.")
            End If
            If (formatting.HorizontalAlignment.HasValue AndAlso Not [Enum].IsDefined(GetType(ExcelHorizontalAlignment), formatting.HorizontalAlignment.Value)) OrElse
               (formatting.VerticalAlignment.HasValue AndAlso Not [Enum].IsDefined(GetType(ExcelVerticalAlignment), formatting.VerticalAlignment.Value)) OrElse
               (formatting.OutsideBorder IsNot Nothing AndAlso Not [Enum].IsDefined(GetType(ExcelBorderLineStyle), formatting.OutsideBorder.Style)) OrElse
               (formatting.InsideBorder IsNot Nothing AndAlso Not [Enum].IsDefined(GetType(ExcelBorderLineStyle), formatting.InsideBorder.Style)) Then
                Throw New ArgumentException("The requested alignment or border style is invalid.", NameOf(formatting))
            End If
        End Sub

        Private Shared Sub InitializeCellFormats(styles As Stylesheet)
            If styles.Fonts Is Nothing Then styles.Fonts = New Fonts()
            If Not styles.Fonts.HasChildren Then WorksheetStyleTransfer.AppendStyleRecord(styles.Fonts, New Font())
            If styles.Fills Is Nothing Then styles.Fills = New Fills()
            If Not styles.Fills.HasChildren Then
                WorksheetStyleTransfer.AppendStyleRecord(styles.Fills, New Fill(New PatternFill With {.PatternType = PatternValues.None}))
                WorksheetStyleTransfer.AppendStyleRecord(styles.Fills, New Fill(New PatternFill With {.PatternType = PatternValues.Gray125}))
            End If
            If styles.Borders Is Nothing Then styles.Borders = New Borders()
            If Not styles.Borders.HasChildren Then WorksheetStyleTransfer.AppendStyleRecord(styles.Borders, New Border())
            If styles.CellStyleFormats Is Nothing Then styles.CellStyleFormats = New CellStyleFormats()
            If Not styles.CellStyleFormats.HasChildren Then WorksheetStyleTransfer.AppendStyleRecord(styles.CellStyleFormats, New CellFormat())
            If styles.CellFormats Is Nothing Then styles.CellFormats = New CellFormats()
            If Not styles.CellFormats.HasChildren Then WorksheetStyleTransfer.AppendStyleRecord(styles.CellFormats, New CellFormat With {.FormatId = 0UI})
        End Sub

        Private Shared Sub ApplyFormatChanges(styles As Stylesheet, format As CellFormat, formatting As ExcelCellFormatting,
                                              changesFont As Boolean, changesAlignment As Boolean, changesBorders As Boolean, edges As RangeEdges)
            Dim baseFormat = styles.CellStyleFormats.Elements(Of CellFormat)().ElementAtOrDefault(CInt(If(format.FormatId?.Value, 0UI)))
            If formatting.NumberFormat IsNot Nothing Then
                format.NumberFormatId = FindOrAddNumberFormat(styles, formatting.NumberFormat)
                format.ApplyNumberFormat = True
            End If
            If changesFont Then
                Dim fontId = If(If(format.ApplyFont?.Value, True), format.FontId?.Value, Nothing)
                Dim original = styles.Fonts.Elements(Of Font)().ElementAtOrDefault(CInt(If(fontId, If(baseFormat?.FontId?.Value, 0UI))))
                Dim font = If(DirectCast(original?.CloneNode(True), Font), New Font())
                If formatting.FontName IsNot Nothing Then
                    font.FontName = New FontName With {.Val = formatting.FontName}
                    font.FontScheme = Nothing
                End If
                If formatting.FontSize.HasValue Then font.FontSize = New FontSize With {.Val = formatting.FontSize.Value}
                If formatting.Bold.HasValue Then font.Bold = New Bold With {.Val = formatting.Bold.Value}
                If formatting.Italic.HasValue Then font.Italic = New Italic With {.Val = formatting.Italic.Value}
                If formatting.FontColor.HasValue Then
                    Dim color = formatting.FontColor.Value
                    font.Color = If(color.IsEmpty, New Color With {.Auto = True}, New Color With {.Rgb = ToSpreadsheetRgb(color)})
                End If
                format.FontId = WorksheetStyleTransfer.AppendStyleRecord(styles.Fonts, font)
                format.ApplyFont = True
            End If
            If formatting.FillColor.HasValue Then
                Dim color = formatting.FillColor.Value
                Dim pattern As New PatternFill With {.PatternType = If(color.IsEmpty, PatternValues.None, PatternValues.Solid)}
                If Not color.IsEmpty Then pattern.ForegroundColor = New ForegroundColor With {.Rgb = ToSpreadsheetRgb(color)}
                format.FillId = WorksheetStyleTransfer.AppendStyleRecord(styles.Fills, New Fill(pattern))
                format.ApplyFill = True
            End If
            If changesAlignment Then
                Dim original = If(If(format.ApplyAlignment?.Value, True), format.Alignment, Nothing)
                Dim alignment = If(DirectCast(If(original, baseFormat?.Alignment)?.CloneNode(True), Alignment), New Alignment())
                If formatting.HorizontalAlignment.HasValue Then
                    alignment.Horizontal = New HorizontalAlignmentValues(formatting.HorizontalAlignment.Value.ToString().ToLowerInvariant())
                End If
                If formatting.VerticalAlignment.HasValue Then
                    alignment.Vertical = New VerticalAlignmentValues(formatting.VerticalAlignment.Value.ToString().ToLowerInvariant())
                End If
                If formatting.WrapText.HasValue Then alignment.WrapText = formatting.WrapText.Value
                format.Alignment = alignment
                format.ApplyAlignment = True
            End If
            If changesBorders Then ApplyBorderChanges(styles, format, baseFormat, formatting, edges)
        End Sub

        Private Shared Sub ApplyBorderChanges(styles As Stylesheet, format As CellFormat, baseFormat As CellFormat,
                                              formatting As ExcelCellFormatting, edges As RangeEdges)
            Dim borderId = If(If(format.ApplyBorder?.Value, True), format.BorderId?.Value, Nothing)
            Dim original = styles.Borders.Elements(Of Border)().ElementAtOrDefault(CInt(If(borderId, If(baseFormat?.BorderId?.Value, 0UI))))
            Dim border = If(DirectCast(original?.CloneNode(True), Border), New Border())
            UpdateBorderEdge(border, New LeftBorder(), If((edges And RangeEdges.Left) <> 0, formatting.OutsideBorder, formatting.InsideBorder))
            UpdateBorderEdge(border, New RightBorder(), If((edges And RangeEdges.Right) <> 0, formatting.OutsideBorder, formatting.InsideBorder))
            UpdateBorderEdge(border, New TopBorder(), If((edges And RangeEdges.Top) <> 0, formatting.OutsideBorder, formatting.InsideBorder))
            UpdateBorderEdge(border, New BottomBorder(), If((edges And RangeEdges.Bottom) <> 0, formatting.OutsideBorder, formatting.InsideBorder))
            format.BorderId = WorksheetStyleTransfer.AppendStyleRecord(styles.Borders, border)
            format.ApplyBorder = True
        End Sub

        Private Shared Sub UpdateBorderEdge(Of T As BorderPropertiesType)(border As Border, edge As T, change As ExcelCellBorder)
            If change Is Nothing Then Return
            Dim previous = border.GetFirstChild(Of T)()
            If previous IsNot Nothing Then previous.Remove()
            If change.Style <> ExcelBorderLineStyle.None Then
                edge.Style = New BorderStyleValues(change.Style.ToString().ToLowerInvariant())
                Dim color = change.Color.GetValueOrDefault()
                edge.Color = If(color.IsEmpty, New Color With {.Auto = True}, New Color With {.Rgb = ToSpreadsheetRgb(color)})
            End If
            border.AddChild(edge)
        End Sub

        Private Shared Function ToSpreadsheetRgb(color As System.Drawing.Color) As String
            Return "FF" & (color.ToArgb() And &HFFFFFF).ToString("X6", CultureInfo.InvariantCulture)
        End Function

        Private Shared Function FindOrAddNumberFormat(styles As Stylesheet, code As String) As UInteger
            If code.Equals("General", StringComparison.OrdinalIgnoreCase) Then Return 0UI
            If styles.NumberingFormats Is Nothing Then styles.NumberingFormats = New NumberingFormats()
            Dim existing = styles.NumberingFormats.Elements(Of NumberingFormat)().ToArray()
            Dim matching = existing.FirstOrDefault(Function(record) String.Equals(record.FormatCode?.Value, code, StringComparison.Ordinal))
            If matching IsNot Nothing Then Return matching.NumberFormatId.Value
            Dim used = New HashSet(Of UInteger)(existing.Select(Function(record) record.NumberFormatId.Value))
            Dim index = 164UI
            While used.Contains(index)
                index += 1UI
            End While
            WorksheetStyleTransfer.AppendStyleRecord(styles.NumberingFormats, New NumberingFormat With {.NumberFormatId = index, .FormatCode = code})
            Return index
        End Function
    End Class

End Namespace
