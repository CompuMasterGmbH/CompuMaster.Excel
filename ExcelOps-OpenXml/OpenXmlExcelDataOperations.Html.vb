Imports System.Globalization
Imports System.Net
Imports System.Text
Imports System.Xml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NumberFormat = ExcelNumberFormat.NumberFormat

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <inheritdoc/>
        Protected Overrides Sub ExportSheetToHtmlInternal(worksheetName As String, sb As StringBuilder, options As HtmlSheetExportOptions)
            ValidateHtmlExportRange(options)
            Dim headers = options.EffectiveTableHeaderRowIndexes()
            Dim part = GetWorksheetPart(worksheetName)
            Dim data = GetSheetData(part)
            Dim merges = MergedCells(worksheetName)
            Dim cells = EnumerateWorksheetCells(data).ToList()
            If cells.Count = 0 AndAlso merges.Count = 0 Then
                sb.AppendLine(options.EffectiveHtmlForEmptySheet())
                Return
            End If
            Dim firstRow = If(cells.Count = 0, Integer.MaxValue, cells.Min(Function(c) c.RowIndex))
            Dim firstColumn = If(cells.Count = 0, Integer.MaxValue, cells.Min(Function(c) c.ColumnIndex))
            Dim lastRow = If(cells.Count = 0, -1, cells.Max(Function(c) c.RowIndex))
            Dim lastColumn = If(cells.Count = 0, -1, cells.Max(Function(c) c.ColumnIndex))
            For Each merge In merges
                firstRow = Math.Min(firstRow, merge.AddressStart.RowIndex)
                firstColumn = Math.Min(firstColumn, merge.AddressStart.ColumnIndex)
                lastRow = Math.Max(lastRow, merge.AddressEnd.RowIndex)
                lastColumn = Math.Max(lastColumn, merge.AddressEnd.ColumnIndex)
            Next
            firstRow = Math.Max(firstRow, options.FirstRowIndex)
            firstColumn = Math.Max(firstColumn, options.FirstColumnIndex)
            lastRow = Math.Min(lastRow, options.LastRowIndex.GetValueOrDefault(lastRow))
            lastColumn = Math.Min(lastColumn, options.LastColumnIndex.GetValueOrDefault(lastColumn))
            If firstRow > lastRow OrElse firstColumn > lastColumn Then
                sb.AppendLine(options.EffectiveHtmlForEmptySheet())
                Return
            End If
            Dim includedMerges = merges.Where(Function(m) m.AddressStart.RowIndex <= lastRow AndAlso m.AddressEnd.RowIndex >= firstRow AndAlso
                m.AddressStart.ColumnIndex <= lastColumn AndAlso m.AddressEnd.ColumnIndex >= firstColumn).ToList()
            For Each merge In includedMerges
                If merge.AddressStart.RowIndex < firstRow OrElse merge.AddressEnd.RowIndex > lastRow OrElse
                    merge.AddressStart.ColumnIndex < firstColumn OrElse merge.AddressEnd.ColumnIndex > lastColumn Then
                    Throw New ArgumentException("The HTML export range partially intersects merged cells '" & merge.LocalAddress() &
                        "' on worksheet '" & worksheetName & "'. Export the entire merged range or exclude it.", NameOf(options))
                End If
            Next

            Dim rows = cells.ToLookup(Function(c) c.RowIndex)
            Dim rowsByIndex = EnumerateWorksheetRows(data).ToDictionary(Function(item) item.Index, Function(item) item.Value)
            Dim stylesheet = WorkbookPart.WorkbookStylesPart?.Stylesheet
            Dim palette = GetThemePalette()
            Dim styleCache As New Dictionary(Of UInteger, String)
            Dim formatCache As New Dictionary(Of UInteger, NumberFormat)
            sb.Append("<table class=""").Append(WebUtility.HtmlEncode(options.TableCssClassName)).Append(""">")
            For rowIndex = firstRow To lastRow
                Dim rowCells = rows(rowIndex).ToDictionary(Function(c) c.ColumnIndex, Function(c) c.Value)
                Dim row As Row = Nothing
                rowsByIndex.TryGetValue(rowIndex, row)
                Dim currentRow = rowIndex
                Dim rowMerges = includedMerges.Where(Function(m) m.AddressStart.RowIndex <= currentRow AndAlso m.AddressEnd.RowIndex >= currentRow).ToArray()
                Dim tag = If(headers IsNot Nothing AndAlso headers.Contains(rowIndex), "th", "td")
                sb.Append("<tr>")
                Dim columnIndex = firstColumn
                While columnIndex <= lastColumn
                    Dim merge = Array.Find(rowMerges, Function(m) m.AddressStart.ColumnIndex <= columnIndex AndAlso m.AddressEnd.ColumnIndex >= columnIndex)
                    If merge Is Nothing OrElse (merge.AddressStart.RowIndex = rowIndex AndAlso merge.AddressStart.ColumnIndex = columnIndex) Then
                        Dim cell As Cell = Nothing
                        rowCells.TryGetValue(columnIndex, cell)
                        Dim index = WorksheetCellStyleIndex(part.Worksheet, row, cell, columnIndex)
                        Dim format = GetCachedNumberFormat(stylesheet, index, formatCache)
                        Dim content As String
                        If IsMissingFormulaResult(cell) Then
                            Dim currentColumn = columnIndex
                            content = ResolveHtmlCellText(worksheetName, currentRow, currentColumn, options,
                                Function() ReadHtmlCellText(worksheetName, currentRow, currentColumn, cell, format))
                        Else
                            content = ReadHtmlCellText(worksheetName, currentRow, columnIndex, cell, format)
                        End If
                        If Not styleCache.ContainsKey(index) Then styleCache.Add(index, HtmlCellStyle(stylesheet, index, palette))
                        AppendHtmlCell(sb, tag, merge, styleCache(index), content)
                    End If
                    columnIndex = If(merge Is Nothing, columnIndex + 1, merge.AddressEnd.ColumnIndex + 1)
                End While
                sb.AppendLine("</tr>")
            Next
            sb.AppendLine("</table>")
        End Sub

        Private Function ReadHtmlCellText(sheetName As String, rowIndex As Integer, columnIndex As Integer, cell As Cell, format As NumberFormat) As String
            Dim value = ReadCellValue(sheetName, rowIndex, columnIndex, cell)
            ' EPPlus 4 renders Boolean cells as 0/1 in its HTML tables.
            If TypeOf value Is Boolean Then value = If(DirectCast(value, Boolean), 1.0, 0.0)
            Return FormatCellValue(value, format)
        End Function

        Private Shared Sub AppendHtmlCell(sb As StringBuilder, tag As String, merge As ExcelRange, style As String, content As String)
            sb.Append("<").Append(tag)
            If merge IsNot Nothing Then
                Dim rows = merge.AddressEnd.RowIndex - merge.AddressStart.RowIndex + 1
                Dim columns = merge.AddressEnd.ColumnIndex - merge.AddressStart.ColumnIndex + 1
                If rows > 1 Then sb.Append(" rowspan=""").Append(rows).Append(""""c)
                If columns > 1 Then sb.Append(" colspan=""").Append(columns).Append(""""c)
            End If
            sb.Append(" style=""").Append(style).Append(""">")
            sb.Append(WebUtility.HtmlEncode(If(String.IsNullOrEmpty(content), " ", content)))
            sb.Append("</").Append(tag).Append(">"c)
        End Sub

        Private Function GetThemePalette() As String()
            Dim part = WorkbookPart.ThemePart
            If part Is Nothing Then Return ReadThemeColorPalette(Nothing)
            Dim xml As New XmlDocument With {.XmlResolver = Nothing}
            Using input = part.GetStream(IO.FileMode.Open, IO.FileAccess.Read),
                  reader = XmlReader.Create(input, New XmlReaderSettings With {.DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing})
                xml.Load(reader)
            End Using
            Return ReadThemeColorPalette(xml)
        End Function

        Private Shared Function HtmlCellStyle(stylesheet As Stylesheet, styleIndex As UInteger, palette As String()) As String
            Dim format = GetCellFormat(stylesheet, styleIndex)
            Dim alignment = format?.Alignment
            Dim css As New List(Of String)
            Dim horizontal = If(alignment?.Horizontal?.Value, HorizontalAlignmentValues.Left)
            If horizontal = HorizontalAlignmentValues.Center Then
                css.Add("text-align:center")
            ElseIf horizontal = HorizontalAlignmentValues.Right Then
                css.Add("text-align:right")
            ElseIf horizontal = HorizontalAlignmentValues.Justify Then
                css.Add("text-align:justify")
            Else
                css.Add("text-align:left")
            End If
            Dim font = stylesheet?.Fonts?.Elements(Of Font)().ElementAtOrDefault(CInt(If(format?.FontId?.Value, 0UI)))
            If font?.Bold IsNot Nothing AndAlso If(font.Bold.Val?.Value, True) Then css.Add("font-weight:bold")
            If font?.Italic IsNot Nothing AndAlso If(font.Italic.Val?.Value, True) Then css.Add("font-style:italic")
            If font?.Underline IsNot Nothing AndAlso If(font.Underline.Val?.Value, UnderlineValues.Single) <> UnderlineValues.None Then
                css.Add("text-decoration:underline")
            End If
            Dim foreground = ResolveColor(font?.Color, palette)
            If foreground IsNot Nothing Then css.Add("color:" & foreground)
            Dim fill = stylesheet?.Fills?.Elements(Of Fill)().ElementAtOrDefault(CInt(If(format?.FillId?.Value, 0UI)))?.PatternFill
            Dim pattern = If(fill?.PatternType?.Value, PatternValues.None)
            If pattern <> PatternValues.None Then
                Dim fillFront = ResolveColor(fill.ForegroundColor, palette)
                Dim fillBack = ResolveColor(fill.BackgroundColor, palette)
                Dim background = If(pattern = PatternValues.Solid, If(fillFront, fillBack), If(fillBack, fillFront))
                If background IsNot Nothing Then css.Add("background-color:" & background)
            End If
            If If(alignment?.WrapText?.Value, False) Then css.Add("white-space:pre-wrap")
            Return String.Join(";", css)
        End Function

        Private Shared Function ResolveColor(color As ColorType, palette As String()) As String
            If color Is Nothing Then Return Nothing
            Dim css As String = Nothing
            Dim rgb = color.Rgb?.Value
            If rgb IsNot Nothing AndAlso (rgb.Length = 6 OrElse rgb.Length = 8) Then
                Dim colorValue As UInteger
                If UInteger.TryParse(rgb.Substring(rgb.Length - 6), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, colorValue) Then
                    css = "#" & colorValue.ToString("X6", CultureInfo.InvariantCulture)
                End If
            ElseIf color.Theme IsNot Nothing AndAlso color.Theme.Value < CUInt(palette.Length) Then
                css = palette(CInt(color.Theme.Value))
            ElseIf color.Indexed IsNot Nothing AndAlso color.Indexed.Value > 0UI AndAlso color.Indexed.Value < 64UI Then
                css = IndexedColor(CInt(color.Indexed.Value))
            End If
            Return If(css Is Nothing, Nothing, ApplyTint(css, If(color.Tint?.Value, 0.0)))
        End Function

        Private Shared Function GetCellFormat(stylesheet As Stylesheet, styleIndex As UInteger) As CellFormat
            Return stylesheet?.CellFormats?.Elements(Of CellFormat)().ElementAtOrDefault(CInt(styleIndex))
        End Function

        Private Shared Function WorksheetCellStyleIndex(worksheet As Worksheet, row As Row, cell As Cell, columnIndex As Integer) As UInteger
            If cell?.StyleIndex IsNot Nothing Then Return cell.StyleIndex.Value
            If If(row?.CustomFormat?.Value, False) AndAlso row.StyleIndex IsNot Nothing Then Return row.StyleIndex.Value
            Dim number = CUInt(columnIndex + 1)
            Dim columnStyle = worksheet.Elements(Of Columns)().SelectMany(Function(group) group.Elements(Of Column)()).
                FirstOrDefault(Function(column) column.Min IsNot Nothing AndAlso column.Max IsNot Nothing AndAlso column.Min.Value <= number AndAlso column.Max.Value >= number)
            Return If(columnStyle?.Style?.Value, 0UI)
        End Function

        Private Shared ReadOnly BuiltInFormats As New Dictionary(Of UInteger, String) From {
            {1UI, "0"}, {2UI, "0.00"}, {3UI, "#,##0"}, {4UI, "#,##0.00"},
            {9UI, "0%"}, {10UI, "0.00%"}, {11UI, "0.00E+00"}, {12UI, "# ?/?"}, {13UI, "# ??/??"},
            {14UI, "mm-dd-yy"}, {15UI, "d-mmm-yy"}, {16UI, "d-mmm"}, {17UI, "mmm-yy"},
            {18UI, "h:mm AM/PM"}, {19UI, "h:mm:ss AM/PM"}, {20UI, "h:mm"}, {21UI, "h:mm:ss"},
            {22UI, "m/d/yy h:mm"}, {37UI, "#,##0 ;(#,##0)"}, {38UI, "#,##0 ;[Red](#,##0)"},
            {39UI, "#,##0.00;(#,##0.00)"}, {40UI, "#,##0.00;[Red](#,##0.00)"},
            {45UI, "mm:ss"}, {46UI, "[h]:mm:ss"}, {47UI, "mmss.0"}, {48UI, "##0.0"}, {49UI, "@"}
        }

        Private Shared Function NumberFormatCode(stylesheet As Stylesheet, styleIndex As UInteger) As String
            Dim id = If(GetCellFormat(stylesheet, styleIndex)?.NumberFormatId?.Value, 0UI)
            Dim custom = stylesheet?.NumberingFormats?.Elements(Of NumberingFormat)().
                FirstOrDefault(Function(item) item.NumberFormatId IsNot Nothing AndAlso item.NumberFormatId.Value = id)?.FormatCode?.Value
            If custom IsNot Nothing Then Return custom
            Dim code As String = Nothing
            Return If(BuiltInFormats.TryGetValue(id, code), code, "General")
        End Function

        Private Shared Function GetCachedNumberFormat(stylesheet As Stylesheet, styleIndex As UInteger, cache As Dictionary(Of UInteger, NumberFormat)) As NumberFormat
            Dim format As NumberFormat = Nothing
            If Not cache.TryGetValue(styleIndex, format) Then
                format = New NumberFormat(NumberFormatCode(stylesheet, styleIndex))
                cache.Add(styleIndex, format)
            End If
            Return format
        End Function

        Private Function FormatCellValue(value As Object, format As NumberFormat) As String
            Dim culture = CultureInfo.CurrentCulture
            If value Is Nothing Then Return String.Empty
            If Not format.IsValid OrElse TypeOf value Is Boolean Then Return Convert.ToString(value, culture)
            Return format.Format(value, culture, If(WorkbookPart.Workbook.WorkbookProperties?.Date1904?.Value, False))
        End Function

        ''' <inheritdoc/>
        Public Overrides Function LookupCellFormat(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            ValidateCoordinates(rowIndex, columnIndex)
            Dim part = GetWorksheetPart(sheetName)
            Dim row = FindRow(GetSheetData(part), rowIndex, False)
            Dim cell = FindCell(row, rowIndex, columnIndex, False)
            Return NumberFormatCode(WorkbookPart.WorkbookStylesPart?.Stylesheet, WorksheetCellStyleIndex(part.Worksheet, row, cell, columnIndex))
        End Function

        ''' <inheritdoc/>
        Public Overrides Function LookupCellFormattedText(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Dim value = ReadLegacyCellValueAsObject(sheetName, rowIndex, columnIndex)
            Return FormatCellValue(value, New NumberFormat(LookupCellFormat(sheetName, rowIndex, columnIndex)))
        End Function

        Private Shared Iterator Function EnumerateWorksheetRows(data As SheetData) As IEnumerable(Of (Index As Integer, Value As Row))
            Dim rowIndex = -1
            For Each row In data.Elements(Of Row)()
                rowIndex = If(row.RowIndex Is Nothing, rowIndex + 1, CInt(row.RowIndex.Value) - 1)
                Yield (rowIndex, row)
            Next
        End Function

        Private Shared Iterator Function EnumerateWorksheetCells(data As SheetData) As IEnumerable(Of (RowIndex As Integer, ColumnIndex As Integer, Value As Cell))
            For Each row In EnumerateWorksheetRows(data)
                Dim columnIndex As Integer = -1
                For Each cell In row.Value.Elements(Of Cell)()
                    columnIndex = If(cell.CellReference Is Nothing, columnIndex + 1,
                        New ExcelCell(Nothing, cell.CellReference.Value, ExcelCell.ValueTypes.All).ColumnIndex)
                    Yield (row.Index, columnIndex, cell)
                Next
            Next
        End Function

        ''' <inheritdoc/>
        Protected Overrides Function MergedCells(sheetName As String) As List(Of ExcelRange)
            Return GetWorksheetPart(sheetName).Worksheet.Elements(Of MergeCells)().
                SelectMany(Function(element) element.Elements(Of MergeCell)()).
                Select(Function(merge) New ExcelRange(sheetName, merge.Reference.Value)).ToList()
        End Function

        ''' <inheritdoc/>
        Public Overrides Function IsMergedCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            ValidateCoordinates(rowIndex, columnIndex)
            Return MergedCells(sheetName).Any(Function(m) rowIndex >= m.AddressStart.RowIndex AndAlso rowIndex <= m.AddressEnd.RowIndex AndAlso
                columnIndex >= m.AddressStart.ColumnIndex AndAlso columnIndex <= m.AddressEnd.ColumnIndex)
        End Function

    End Class

End Namespace
