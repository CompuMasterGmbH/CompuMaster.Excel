Imports System.Globalization
Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Theme = DocumentFormat.OpenXml.Drawing

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Function WorksheetThemesDiffer(destination As WorkbookPart) As Boolean
            Dim destinationTheme = destination.ThemePart?.Theme
            Return destinationTheme IsNot Nothing AndAlso
                Not String.Equals(WorkbookPart.ThemePart?.Theme?.OuterXml, destinationTheme.OuterXml, StringComparison.Ordinal)
        End Function

        Private Sub TransferWorksheetStyles(destination As WorkbookPart, worksheets As IEnumerable(Of Worksheet))
            Dim sourceStyles = CurrentStyles
            If sourceStyles Is Nothing Then Return
            Dim materializeTheme = WorksheetThemesDiffer(destination)
            Dim targetPart = destination.WorkbookStylesPart
            If targetPart Is Nothing Then
                targetPart = destination.AddNewPart(Of WorkbookStylesPart)()
                targetPart.Stylesheet = DirectCast(sourceStyles.CloneNode(True), Stylesheet)
                targetPart.Stylesheet.Save()
            End If
            If Not materializeTheme AndAlso sourceStyles.OuterXml = targetPart.Stylesheet.OuterXml Then Return

            Dim transfer As New WorksheetStyleTransfer(sourceStyles, targetPart.Stylesheet,
                If(materializeTheme, GetThemePalette(), Nothing), WorkbookPart.ThemePart?.Theme)
            For Each worksheet In worksheets
                transfer.Apply(worksheet)
            Next
            targetPart.Stylesheet.Save()
        End Sub

        Private NotInheritable Class WorksheetStyleTransfer
            Private ReadOnly _source As Stylesheet
            Private ReadOnly _target As Stylesheet
            Private ReadOnly _fonts As New Dictionary(Of UInteger, UInteger)
            Private ReadOnly _fills As New Dictionary(Of UInteger, UInteger)
            Private ReadOnly _borders As New Dictionary(Of UInteger, UInteger)
            Private ReadOnly _numbers As New Dictionary(Of UInteger, UInteger)
            Private ReadOnly _baseFormats As New Dictionary(Of UInteger, UInteger)
            Private ReadOnly _formats As New Dictionary(Of UInteger, UInteger)

            Private ReadOnly _themeColors As String()
            Private ReadOnly _sourceTheme As Theme.Theme
            Private ReadOnly _resolveIndexedColors As Boolean

            Friend Sub New(source As Stylesheet, target As Stylesheet, themeColors As String(), sourceTheme As Theme.Theme)
                _source = source
                _target = target
                _themeColors = themeColors
                _sourceTheme = sourceTheme
                If _source.CellFormats Is Nothing OrElse _target.CellFormats Is Nothing Then
                    Throw New NotSupportedException("Transferring worksheet styles requires cell format records in both workbooks.")
                End If
                Dim originalPalette = _source.Colors?.IndexedColors
                Dim destinationPalette = _target.Colors?.IndexedColors
                _resolveIndexedColors = (originalPalette Is Nothing) <> (destinationPalette Is Nothing) OrElse
                    (originalPalette IsNot Nothing AndAlso Not SameStyleRecord(originalPalette, destinationPalette))
            End Sub

            Friend Sub Apply(worksheet As Worksheet)
                Dim defaultFormat = MapCellFormat(0UI)
                For Each entry In EnumerateWorksheetCells(worksheet.GetFirstChild(Of SheetData)())
                    Dim cell = entry.Value
                    If cell.StyleIndex IsNot Nothing Then
                        cell.StyleIndex = MapCellFormat(cell.StyleIndex.Value)
                    ElseIf defaultFormat <> 0UI AndAlso WorksheetCellStyleIndex(worksheet, TryCast(cell.Parent, Row), cell, entry.ColumnIndex) = 0UI Then
                        cell.StyleIndex = defaultFormat
                    End If
                Next
                For Each row In worksheet.Descendants(Of Row)()
                    If row.StyleIndex IsNot Nothing Then row.StyleIndex = MapCellFormat(row.StyleIndex.Value)
                Next
                For Each phonetics In worksheet.Descendants(Of PhoneticProperties)()
                    If phonetics.FontId IsNot Nothing Then
                        phonetics.FontId = MapStyleRecord(Of Font)(_source.Fonts, _target.Fonts, _fonts, phonetics.FontId.Value)
                    End If
                Next
                For Each column In worksheet.Descendants(Of Column)()
                    If column.Style IsNot Nothing Then column.Style = MapCellFormat(column.Style.Value)
                Next
            End Sub

            Private Function MapCellFormat(index As UInteger) As UInteger
                Dim mapped As UInteger
                If _formats.TryGetValue(index, mapped) Then Return mapped
                Dim format = ReadStyleRecord(Of CellFormat)(_source.CellFormats, index)
                RemapFormatComponents(format)
                If format.FormatId IsNot Nothing Then format.FormatId = MapBaseFormat(format.FormatId.Value)
                mapped = AppendStyleRecord(_target.CellFormats, format)
                _formats.Add(index, mapped)
                Return mapped
            End Function

            Private Function MapBaseFormat(index As UInteger) As UInteger
                Dim mapped As UInteger
                If _baseFormats.TryGetValue(index, mapped) Then Return mapped
                Dim format = ReadStyleRecord(Of CellFormat)(_source.CellStyleFormats, index)
                RemapFormatComponents(format)
                If _target.CellStyleFormats Is Nothing Then _target.CellStyleFormats = New CellStyleFormats()
                mapped = AppendStyleRecord(_target.CellStyleFormats, format)
                _baseFormats.Add(index, mapped)
                Return mapped
            End Function

            Private Sub RemapFormatComponents(format As CellFormat)
                If format.FontId IsNot Nothing Then
                    If _target.Fonts Is Nothing Then _target.Fonts = New Fonts()
                    format.FontId = MapStyleRecord(Of Font)(_source.Fonts, _target.Fonts, _fonts, format.FontId.Value)
                End If
                If format.FillId IsNot Nothing Then
                    If _target.Fills Is Nothing Then _target.Fills = New Fills()
                    format.FillId = MapStyleRecord(Of Fill)(_source.Fills, _target.Fills, _fills, format.FillId.Value)
                End If
                If format.BorderId IsNot Nothing Then
                    If _target.Borders Is Nothing Then _target.Borders = New Borders()
                    format.BorderId = MapStyleRecord(Of Border)(_source.Borders, _target.Borders, _borders, format.BorderId.Value)
                End If
                If format.NumberFormatId IsNot Nothing AndAlso format.NumberFormatId.Value >= 164UI Then
                    format.NumberFormatId = MapNumberFormat(format.NumberFormatId.Value)
                End If
            End Sub

            Private Function MapNumberFormat(index As UInteger) As UInteger
                Dim mapped As UInteger
                If _numbers.TryGetValue(index, mapped) Then Return mapped
                Dim sourceFormat = _source.NumberingFormats?.Elements(Of NumberingFormat)().FirstOrDefault(
                    Function(record) If(record.NumberFormatId?.Value, UInteger.MaxValue) = index)
                If sourceFormat?.FormatCode Is Nothing Then Throw New InvalidDataException("A worksheet style refers to an undefined number format.")
                If _target.NumberingFormats Is Nothing Then _target.NumberingFormats = New NumberingFormats()
                Dim records = _target.NumberingFormats.Elements(Of NumberingFormat)()
                Dim code = sourceFormat.FormatCode.Value
                Dim existing = records.FirstOrDefault(Function(format) String.Equals(format.FormatCode?.Value, code, StringComparison.Ordinal))
                If existing IsNot Nothing Then
                    mapped = existing.NumberFormatId.Value
                Else
                    Dim used As New HashSet(Of UInteger)(records.Select(Function(record) record.NumberFormatId.Value))
                    mapped = 164UI
                    While used.Contains(mapped)
                        mapped += 1UI
                    End While
                    Dim format = DirectCast(sourceFormat.CloneNode(True), NumberingFormat)
                    format.NumberFormatId = mapped
                    AppendStyleRecord(_target.NumberingFormats, format)
                End If
                _numbers.Add(index, mapped)
                Return mapped
            End Function

            Private Function MapStyleRecord(Of T As OpenXmlElement)(source As OpenXmlCompositeElement, target As OpenXmlCompositeElement,
                                                                          map As Dictionary(Of UInteger, UInteger), index As UInteger) As UInteger
                Dim mapped As UInteger
                If map.TryGetValue(index, mapped) Then Return mapped
                Dim candidate = ReadStyleRecord(Of T)(source, index)
                MaterializeStyleReferences(candidate)
                mapped = AppendStyleRecord(target, candidate)
                map.Add(index, mapped)
                Return mapped
            End Function

            Private Sub MaterializeStyleReferences(record As OpenXmlElement)
                If _themeColors Is Nothing AndAlso Not _resolveIndexedColors Then Return
                For Each color In record.Descendants(Of ColorType)()
                    If color.Rgb IsNot Nothing Then Continue For
                    Dim hex As String = Nothing
                    If _themeColors IsNot Nothing AndAlso color.Theme IsNot Nothing Then
                        hex = ResolveColor(color, _themeColors)
                        If hex Is Nothing Then Throw New InvalidDataException("A style refers to an undefined theme color.")
                        color.Theme = Nothing
                    ElseIf _resolveIndexedColors AndAlso color.Indexed IsNot Nothing AndAlso color.Indexed.Value < 64UI Then
                        Dim colorIndex = CInt(color.Indexed.Value)
                        Dim rgb = _source.Colors?.IndexedColors?.Elements(Of RgbColor)().ElementAtOrDefault(colorIndex)?.Rgb?.Value
                        If rgb IsNot Nothing Then
                            hex = "#" & rgb.Substring(System.Math.Max(0, rgb.Length - 6))
                        Else
                            hex = IndexedColor(colorIndex)
                        End If
                        If hex Is Nothing Then Throw New InvalidDataException("A style refers to an undefined indexed color.")
                        hex = ApplyTint(hex, If(color.Tint?.Value, 0.0))
                        color.Indexed = Nothing
                    End If
                    If hex IsNot Nothing Then
                        color.Rgb = New HexBinaryValue("FF" & hex.Substring(1))
                        color.Tint = Nothing
                    End If
                Next

                Dim font = TryCast(record, Font)
                If _themeColors IsNot Nothing AndAlso font?.FontScheme?.Val IsNot Nothing Then
                    Dim scheme = _sourceTheme?.ThemeElements?.FontScheme
                    Dim latin As Theme.LatinFont = Nothing
                    If font.FontScheme.Val.Value = FontSchemeValues.Major Then latin = scheme?.MajorFont?.GetFirstChild(Of Theme.LatinFont)()
                    If font.FontScheme.Val.Value = FontSchemeValues.Minor Then latin = scheme?.MinorFont?.GetFirstChild(Of Theme.LatinFont)()
                    If Not String.IsNullOrEmpty(latin?.Typeface?.Value) Then font.FontName = New FontName() With {.Val = latin.Typeface.Value}
                    font.FontScheme = Nothing
                End If
            End Sub

            Private Shared Function ReadStyleRecord(Of T As OpenXmlElement)(source As OpenXmlCompositeElement, index As UInteger) As T
                Dim record = If(index <= Integer.MaxValue, source?.Elements(Of T)().ElementAtOrDefault(CInt(index)), Nothing)
                If record Is Nothing Then Throw New InvalidDataException("A worksheet style points outside the workbook's style table.")
                Return DirectCast(record.CloneNode(True), T)
            End Function

            Friend Shared Function AppendStyleRecord(Of T As OpenXmlElement)(target As OpenXmlCompositeElement, candidate As T) As UInteger
                Dim index As UInteger = 0
                For Each record In target.Elements(Of T)()
                    If SameStyleRecord(record, candidate) Then Return index
                    index += 1UI
                Next
                target.AppendChild(candidate)
                target.SetAttribute(New OpenXmlAttribute("count", String.Empty, (index + 1UI).ToString(CultureInfo.InvariantCulture)))
                Return index
            End Function

            Private Shared Function SameStyleRecord(first As OpenXmlElement, second As OpenXmlElement) As Boolean
                If first.LocalName <> second.LocalName OrElse first.NamespaceUri <> second.NamespaceUri OrElse
                   first.ChildElements.Count <> second.ChildElements.Count Then Return False
                Dim attributes = first.GetAttributes()
                Dim other = second.GetAttributes()
                If attributes.Count <> other.Count OrElse attributes.Any(Function(attribute) Not other.Any(
                    Function(item) item.LocalName = attribute.LocalName AndAlso item.NamespaceUri = attribute.NamespaceUri AndAlso
                        item.Value = attribute.Value)) Then Return False
                If Not first.HasChildren Then Return first.InnerText = second.InnerText
                For index = 0 To first.ChildElements.Count - 1
                    If Not SameStyleRecord(first.ChildElements(index), second.ChildElements(index)) Then Return False
                Next
                Return True
            End Function
        End Class

    End Class

End Namespace
