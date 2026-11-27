Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly MetadataRed As New Regex("\[(red|rot)\]", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly MetadataNumber As New Regex("^[#0,]+(?:\.(?<decimals>[0#]+))?(?<suffix>%|E[+-]0+)?$", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

            Private Function EvaluateMetadataFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                      formulaCell As ExcelCell) As Object
                If arguments.Count = 0 OrElse arguments.Count > If(name = "INFO", 1, 2) Then Return New FormulaError("#VALUE!")
                MarkVolatile()
                If name = "INFO" Then Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) InfoValue(values(0)))
                Dim kind = Scalar(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf kind Is FormulaError Then Return kind
                Dim reference = If(arguments.Count = 1, Nothing, EvaluateExpression(arguments(1), formulaCell))
                If TypeOf reference Is FormulaError Then Return reference
                Dim area = TryCast(reference, FormulaArea)
                If reference IsNot Nothing AndAlso area Is Nothing Then Return New FormulaError("#VALUE!")

                Dim source = If(area Is Nothing, MetadataActiveWorkbook(), ResolveReference(area))
                Dim cell = If(area Is Nothing, source.MetadataActiveCell(formulaCell),
                    New ExcelCell(area.Sheet, area.FirstRow, area.FirstColumn, ExcelCell.ValueTypes.All))
                If cell Is Nothing Then Return New FormulaError("#REF!")
                Return CellMetadataValue(TextValue(kind).ToLowerInvariant(), source, cell, formulaCell)
            End Function

            Private Function MetadataActiveWorkbook() As FormulaCalculation
                Dim session = _run.Session
                Dim active As ExcelDataOperationsBase = Nothing
                If session?.ActiveWorkbook IsNot Nothing AndAlso session.Workbooks.TryGetValue(session.ActiveWorkbook, active) AndAlso
                    TypeOf active Is OpenXmlExcelDataOperations AndAlso Not active.IsClosed Then Return ForWorkbook(DirectCast(active, OpenXmlExcelDataOperations))
                Return Me
            End Function

            Private Function MetadataActiveCell(formulaCell As ExcelCell) As ExcelCell
                If _run.Session?.ActiveCell IsNot Nothing AndAlso MetadataActiveWorkbook() Is Me Then Return _run.Session.ActiveCell
                Dim sheet = _owner.SelectedSheetName()
                If sheet Is Nothing Then Return formulaCell
                Dim address = GetMetadataSheetView(sheet)?.Elements(Of Selection)().FirstOrDefault()?.ActiveCell?.Value
                Return New ExcelCell(sheet, If(address, "A1"), ExcelCell.ValueTypes.All)
            End Function

            Private Function GetMetadataSheetView(sheet As String) As SheetView
                Dim snapshot As WorksheetSnapshot = Nothing
                If sheet Is Nothing OrElse Not _worksheets.TryGetValue(sheet, snapshot) Then Return Nothing
                Return snapshot.Part.Worksheet.SheetViews?.Elements(Of SheetView)().
                    FirstOrDefault(Function(view) If(view.WorkbookViewId?.Value, 0UI) = 0UI)
            End Function

            Private Function InfoValue(kind As Object) As Object
                Dim german = ConversionCulture.TwoLetterISOLanguageName = "de"
                Dim active = MetadataActiveWorkbook()
                Select Case TextValue(kind).ToLowerInvariant()
                    Case "directory", "verzeichnis"
                        Return PathWithSeparator(Environment.CurrentDirectory)
                    Case "numfile", "dateienzahl"
                        If _run.Session Is Nothing Then Return CDbl(_worksheets.Count)
                        Return CDbl(_run.Session.Workbooks.Values.OfType(Of OpenXmlExcelDataOperations)().
                            Where(Function(book) Not book.IsClosed).Sum(Function(book) book.WorkbookSheets().Count))
                    Case "origin", "ursprung"
                        Dim sheet = active._owner.SelectedSheetName()
                        Dim address = If(active.GetMetadataSheetView(sheet)?.TopLeftCell?.Value, "A1")
                        Dim origin As New ExcelCell(sheet, address, ExcelCell.ValueTypes.All)
                        Dim calculation = active._owner.WorkbookPart.Workbook.CalculationProperties
                        Dim mode = Convert.ToString(calculation?.ReferenceMode?.InnerText, CultureInfo.InvariantCulture)
                        If String.Equals(mode, "R1C1", StringComparison.OrdinalIgnoreCase) Then
                            Return "$A:" & If(german, "Z", "R") & (origin.RowIndex + 1).ToString(CultureInfo.InvariantCulture) &
                                If(german, "S", "C") & (origin.ColumnIndex + 1).ToString(CultureInfo.InvariantCulture)
                        End If
                        Return "$A:" & origin.Address(False, True, True)
                    Case "osversion", "sysversion"
                        Return RuntimeInformation.OSDescription
                    Case "recalc"
                        Dim mode = active._owner.WorkbookPart.Workbook.CalculationProperties?.CalculationMode?.InnerText
                        Dim manual = String.Equals(mode, "manual", StringComparison.OrdinalIgnoreCase)
                        Return If(german, If(manual, "Manuell", "Automatisch"), If(manual, "Manual", "Automatic"))
                    Case "release", "version"
                        Return GetType(OpenXmlExcelDataOperations).Assembly.GetName().Version.ToString()
                    Case "system"
                        If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then Return "pcdos"
                        If RuntimeInformation.IsOSPlatform(OSPlatform.OSX) Then Return "mac"
                        Return "linux"
                    Case "memavail", "memused", "totmem"
                        Return New FormulaError("#N/A")
                    Case Else
                        Return New FormulaError("#VALUE!")
                End Select
            End Function

            Private Function CellMetadataValue(kind As String, source As FormulaCalculation, reference As ExcelCell, formulaCell As ExcelCell) As Object
                Dim snapshot As WorksheetSnapshot = Nothing
                If Not source._worksheets.TryGetValue(reference.SheetName, snapshot) Then Return New FormulaError("#REF!")
                Dim german = ConversionCulture.TwoLetterISOLanguageName = "de"
                Select Case kind
                    Case "address", "adresse"
                        Dim address = reference.Address(False, True, True)
                        If source Is Me AndAlso String.Equals(reference.SheetName, formulaCell.SheetName, StringComparison.OrdinalIgnoreCase) Then Return address
                        Dim file = Path.GetFileName(ReferencePath(source._owner))
                        Dim qualifier = If(String.IsNullOrEmpty(file), "", "[" & file & "]") & reference.SheetName
                        Return "'" & qualifier.Replace("'", "''") & "'!" & address
                    Case "filename", "dateiname"
                        Dim filePath = ReferencePath(source._owner)
                        If String.IsNullOrEmpty(filePath) Then Return String.Empty
                        Return PathWithSeparator(Path.GetDirectoryName(filePath)) & "[" & Path.GetFileName(filePath) & "]" & reference.SheetName
                    Case "col", "spalte" : Return CDbl(reference.ColumnIndex + 1)
                    Case "row", "zeile" : Return CDbl(reference.RowIndex + 1)
                    Case "contents", "inhalt"
                        Return If(ReadMetadataCell(source, reference), 0.0)
                    Case "type", "typ"
                        Dim value = ReadMetadataCell(source, reference)
                        If value Is Nothing Then Return "b"
                        If TypeOf value Is String Then Return "l"
                        Return If(german, "w", "v")
                    Case "width", "breite"
                        Return MetadataColumnWidth(snapshot.Part.Worksheet, reference.ColumnIndex, source._owner.CurrentStyles)
                End Select

                Dim cell As Cell = Nothing
                snapshot.Cells.TryGetValue((reference.RowIndex, reference.ColumnIndex), cell)
                Dim row = If(TryCast(cell?.Parent, Row), FindRow(GetSheetData(snapshot.Part), reference.RowIndex, False))
                Dim styles = source._owner.CurrentStyles
                Dim styleIndex = WorksheetCellStyleIndex(snapshot.Part.Worksheet, row, cell, reference.ColumnIndex)
                Dim style = GetCellFormat(styles, styleIndex)
                Select Case kind
                    Case "protect", "schutz"
                        Dim original = styles?.CellStyleFormats?.Elements(Of CellFormat)().ElementAtOrDefault(CInt(If(style?.FormatId?.Value, 0UI)))
                        Dim locked = If(If(style?.ApplyProtection?.Value, True), style?.Protection?.Locked?.Value, Nothing)
                        Return If(If(locked, If(original?.Protection?.Locked?.Value, True)), 1.0, 0.0)
                    Case "prefix", "präfix"
                        If Not TypeOf ReadMetadataCell(source, reference) Is String Then Return String.Empty
                        Select Case style?.Alignment?.Horizontal?.InnerText
                            Case "left" : Return "'"
                            Case "right" : Return """"
                            Case "center", "centerContinuous" : Return "^"
                            Case "fill" : Return "\"
                            Case Else : Return If(If(style?.QuotePrefix?.Value, False), "'", "")
                        End Select
                    Case "color", "farbe", "parentheses", "klammern", "format"
                        Return NumberFormatMetadata(kind, NumberFormatCode(styles, styleIndex), If(style?.NumberFormatId?.Value, 0UI), german)
                    Case Else
                        Return New FormulaError("#VALUE!")
                End Select
            End Function

            Private Function ReadMetadataCell(source As FormulaCalculation, reference As ExcelCell) As Object
                RecordPrecedent(New FormulaPrecedent(New FormulaLocation(source._owner, reference.SheetName, reference.RowIndex, reference.ColumnIndex), reference.RowIndex, reference.ColumnIndex))
                Return source.ReadCell(reference.SheetName, reference.RowIndex, reference.ColumnIndex, trackRead:=False)
            End Function

            Private Shared Function PathWithSeparator(path As String) As String
                If String.IsNullOrEmpty(path) OrElse path.EndsWith(PathSeparator(path), StringComparison.Ordinal) Then Return path
                Return path & PathSeparator(path)
            End Function

            Private Shared Function PathSeparator(pathValue As String) As String
                Return If(pathValue.Contains("\"c), "\", Path.DirectorySeparatorChar.ToString())
            End Function

            Private Shared Function MetadataColumnWidth(worksheet As Worksheet, index As Integer, styles As Stylesheet) As FormulaMatrix
                Dim number = CUInt(index + 1)
                Dim column = worksheet.Elements(Of Columns)().SelectMany(Function(group) group.Elements(Of Column)()).
                    FirstOrDefault(Function(item) item.Min IsNot Nothing AndAlso item.Max IsNot Nothing AndAlso
                        item.Min.Value <= number AndAlso item.Max.Value >= number)
                Dim normal = styles?.CellStyles?.Elements(Of CellStyle)().FirstOrDefault(Function(item) If(item.BuiltinId?.Value, UInteger.MaxValue) = 0)
                Dim format = styles?.CellStyleFormats?.Elements(Of CellFormat)().ElementAtOrDefault(CInt(If(normal?.FormatId?.Value, 0UI)))
                Dim font = styles?.Fonts?.Elements(Of Font)().ElementAtOrDefault(CInt(If(format?.FontId?.Value, 0UI)))
                Dim digitPixels = If(String.Equals(font?.FontName?.Val?.Value, "Carlito", StringComparison.OrdinalIgnoreCase), 8, 7)
                Dim width = If(column?.Width?.Value, worksheet.SheetFormatProperties?.DefaultColumnWidth?.Value)
                Dim defaultWidth = Not If(column?.CustomWidth?.Value, False)
                Dim pixels = If(width.HasValue, ColumnWidthInPixels(width.Value, digitPixels),
                    CLng(Math.Ceiling((If(worksheet.SheetFormatProperties?.BaseColumnWidth?.Value, 8UI) * digitPixels + 5) / 8.0)) * 8)
                Dim characters = If(If(column?.Hidden?.Value, False), 0.0, Math.Max(0, pixels - 5) / CDbl(digitPixels))
                Return New FormulaMatrix(1, 2, {CObj(Math.Round(characters, MidpointRounding.AwayFromZero)), defaultWidth})
            End Function

            Private Shared Function NumberFormatMetadata(kind As String, code As String, id As UInteger, german As Boolean) As Object
                Dim sections = code.Split(";"c)
                Dim positive = sections(0).Trim().Replace("\(", "(").Replace("\)", ")")
                Dim red = sections.Length > 1 AndAlso MetadataRed.IsMatch(sections(1))
                Dim parentheses = positive.StartsWith("(", StringComparison.Ordinal) AndAlso positive.EndsWith(")", StringComparison.Ordinal)
                If kind = "color" OrElse kind = "farbe" Then Return If(red, 1.0, 0.0)
                If kind = "parentheses" OrElse kind = "klammern" Then Return If(parentheses, 1.0, 0.0)
                If parentheses Then positive = positive.Substring(1, positive.Length - 2)
                Dim result = If(german, "S", "G")
                Select Case id
                    Case 14UI, 15UI : result = "D1"
                    Case 16UI : result = "D2"
                    Case 17UI : result = "D3"
                    Case 22UI : result = "D4"
                    Case 19UI : result = If(german, "U1", "D6")
                    Case 18UI : result = If(german, "U2", "D7")
                    Case 21UI : result = If(german, "U3", "D8")
                    Case 20UI : result = If(german, "U4", "D9")
                    Case Else
                        Select Case positive.ToLowerInvariant()
                            Case "dd.mm.yy", "dd/mm/yy", "m/d/yy", "d-mmm-yy" : result = "D1"
                            Case "h:mm:ss", "hh:mm:ss" : result = If(german, "U3", "D8")
                            Case "h:mm", "hh:mm" : result = If(german, "U4", "D9")
                            Case Else
                                Dim numeric = MetadataNumber.Match(positive)
                                If numeric.Success Then
                                    Dim suffix = numeric.Groups("suffix").Value
                                    Dim prefix = If(suffix = "%", "P", If(suffix.Length > 0, If(german, "E", "S"),
                                        If(positive.Contains(","c), If(german, ".", ","), "F")))
                                    result = prefix & numeric.Groups("decimals").Length.ToString(CultureInfo.InvariantCulture)
                                End If
                        End Select
                End Select
                Return result & If(red, "-", "") & If(parentheses, "()", "")
            End Function

        End Class

    End Class

End Namespace
