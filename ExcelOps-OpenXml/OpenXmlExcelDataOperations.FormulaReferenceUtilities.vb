Imports System.Globalization
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Class FormulaReferenceUnion
            Friend ReadOnly Areas As IReadOnlyList(Of FormulaArea)

            Friend Sub New(areas As IReadOnlyList(Of FormulaArea))
                Me.Areas = areas
            End Sub
        End Class

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly IndirectAddress As New Regex(
                "\A(?:(?<sheet>'(?:[^']|'')+'|[^\s'!]+)!)?(?<address>[^!]+)\z", RegexOptions.CultureInvariant)
            Private Shared ReadOnly A1Coordinate As New Regex("\A\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})\z",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly A1Column As New Regex("\A\$?([A-Z]{1,3})\z", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly A1Row As New Regex("\A\$?([1-9][0-9]{0,6})\z", RegexOptions.CultureInvariant)
            Private Shared ReadOnly EnglishRowColumn As New Regex("\AR([1-9][0-9]*|\[[+-]?[0-9]+\])?C([1-9][0-9]*|\[[+-]?[0-9]+\])?\z",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly GermanRowColumn As New Regex("\AZ([1-9][0-9]*|\([+-]?[0-9]+\))?S([1-9][0-9]*|\([+-]?[0-9]+\))?\z",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly UnquotedSheet As New Regex("\A[\p{L}_][\p{L}\p{M}\p{Nd}_.]*\z", RegexOptions.CultureInvariant)
            Friend Function EvaluateReferenceUnion(expressions As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Dim areas As New List(Of FormulaArea)
                For Each expression In expressions
                    Dim value = EvaluateExpression(expression, formulaCell)
                    If TypeOf value Is FormulaError Then Return value
                    Dim union = TryCast(value, FormulaReferenceUnion)
                    If union IsNot Nothing Then
                        areas.AddRange(union.Areas)
                        Continue For
                    End If
                    Dim area = TryCast(value, FormulaArea)
                    If area Is Nothing Then Return New FormulaError("#VALUE!")
                    areas.Add(area)
                Next
                Return New FormulaReferenceUnion(areas)
            End Function

            Private Iterator Function ReferenceArgumentValues(value As Object) As IEnumerable(Of Object)
                Dim area = TryCast(value, FormulaArea)
                If area IsNot Nothing Then
                    For Each item In RangeValues(area)
                        Yield item
                    Next
                    Return
                End If
                Dim union = TryCast(value, FormulaReferenceUnion)
                If union IsNot Nothing Then
                    For Each member In union.Areas
                        For Each item In RangeValues(member)
                            Yield item
                        Next
                    Next
                    Return
                End If
                Dim matrix = TryCast(value, FormulaMatrix)
                If matrix IsNot Nothing Then
                    For Each item In matrix.Values
                        For Each entry In ReferenceArgumentValues(item)
                            Yield entry
                        Next
                    Next
                    Return
                End If
                Yield value
            End Function

            Private Function EvaluateReferenceUtility(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                      formulaCell As ExcelCell) As Object
                Select Case name
                    Case "ADDRESS"
                        If arguments.Count < 2 OrElse arguments.Count > 5 Then Return New FormulaError("#VALUE!")
                        Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) AddressValue(values))
                    Case "INDIRECT"
                        MarkVolatile()
                        If arguments.Count < 1 OrElse arguments.Count > 2 Then Return New FormulaError("#VALUE!")
                        Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) IndirectValue(values, formulaCell))
                    Case "OFFSET"
                        MarkVolatile()
                        If arguments.Count < 3 OrElse arguments.Count > 5 Then Return New FormulaError("#VALUE!")
                        Dim input = EvaluateExpression(arguments(0), formulaCell)
                        If TypeOf input Is FormulaError Then Return input
                        Dim area = TryCast(input, FormulaArea)
                        If area Is Nothing Then Return New FormulaError("#VALUE!")
                        Return BroadcastReferenceArguments(arguments.Skip(1).ToArray(), formulaCell, Function(values) OffsetValue(area, values))
                    Case "CHOOSE"
                        Return EvaluateChoose(arguments, formulaCell)
                End Select

                If name = "SHEETS" AndAlso arguments.Count = 0 Then Return CDbl(_owner.WorkbookSheets().Count)
                If name = "SHEET" AndAlso arguments.Count = 0 Then Return WorksheetNumber(Me, formulaCell.SheetName)
                If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                Dim value = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf value Is FormulaError Then Return value
                Dim reference = TryCast(value, FormulaArea)
                Dim unionReference = TryCast(value, FormulaReferenceUnion)
                If unionReference IsNot Nothing Then
                    reference = unionReference.Areas(0)
                    If unionReference.Areas.Any(Function(area) area.SourceWorkbook IsNot reference.SourceWorkbook OrElse
                        Not Object.Equals(area.WorkbookIndex, reference.WorkbookIndex) OrElse
                         Not String.Equals(area.Sheet, reference.Sheet, StringComparison.OrdinalIgnoreCase)) Then Return New FormulaError("#VALUE!")
                    If name = "AREAS" Then Return CDbl(unionReference.Areas.Count)
                End If

                If name = "SHEET" AndAlso TypeOf value Is String Then Return WorksheetNumber(Me, DirectCast(value, String))
                If reference Is Nothing Then Return New FormulaError(If(name = "SHEET" OrElse name = "SHEETS", "#N/A", "#VALUE!"))
                Dim source As FormulaCalculation
                Try
                    source = ResolveReference(reference)
                Catch ex As NotSupportedException When name = "FORMULATEXT" OrElse name = "ISFORMULA"
                    Return If(name = "ISFORMULA", CObj(False), New FormulaError("#N/A"))
                End Try
                If name = "SHEET" Then Return WorksheetNumber(source, reference.Sheet)
                If name = "SHEETS" Then Return If(source._worksheets.ContainsKey(reference.Sheet), CObj(1.0), New FormulaError("#REF!"))
                If name = "AREAS" Then Return If(source._worksheets.ContainsKey(reference.Sheet), CObj(1.0), New FormulaError("#REF!"))

                Dim worksheet As WorksheetSnapshot = Nothing
                If Not source._worksheets.TryGetValue(reference.Sheet, worksheet) Then Return New FormulaError("#REF!")
                RecordPrecedent(New FormulaPrecedent(New FormulaLocation(source._owner, reference.Sheet, reference.FirstRow, reference.FirstColumn),
                    reference.LastRow, reference.LastColumn))
                Dim rows = reference.LastRow - reference.FirstRow + 1
                Dim columns = reference.LastColumn - reference.FirstColumn + 1
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim cell As Cell = Nothing
                        worksheet.Cells.TryGetValue((reference.FirstRow + row, reference.FirstColumn + column), cell)
                        If name = "ISFORMULA" Then
                            results(row * columns + column) = cell?.CellFormula IsNot Nothing AndAlso Not IsSpillPlaceholder(cell)
                        Else
                            Dim formula = cell?.CellFormula?.Text
                            results(row * columns + column) = If(String.IsNullOrEmpty(formula), CObj(New FormulaError("#N/A")), DisplayFormula(formula))
                        End If
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Shared Function WorksheetNumber(source As FormulaCalculation, sheet As String) As Object
                Dim index = source._owner.WorkbookSheets().FindIndex(
                    Function(item) String.Equals(item.Name.Value, sheet, StringComparison.OrdinalIgnoreCase))
                If index < 0 Then Return New FormulaError("#N/A")
                Return CDbl(index + 1)
            End Function

            Private Function BroadcastReferenceArguments(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell,
                                                         calculate As Func(Of Object(), Object),
                                                         Optional checkErrors As Boolean = True) As Object
                Dim matrices = arguments.Select(Function(argument) AsMatrix(EvaluateExpression(argument, formulaCell))).ToArray()
                Dim rows = matrices.Max(Function(matrix) matrix.Rows)
                Dim columns = matrices.Max(Function(matrix) matrix.Columns)
                Dim results = AllocateMatrixValues(rows, columns)
                Dim values(arguments.Count - 1) As Object
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        For index = 0 To matrices.Length - 1
                            values(index) = matrices(index).BroadcastAt(row, column)
                        Next
                        Dim failure = If(checkErrors, values.OfType(Of FormulaError)().FirstOrDefault(), Nothing)
                        results(row * columns + column) = If(failure, calculate(values))
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function AddressValue(values As Object()) As Object
                Dim row = ReadFinancialNumber(values(0))
                If TypeOf row Is FormulaError Then Return row
                Dim column = ReadFinancialNumber(values(1))
                If TypeOf column Is FormulaError Then Return column
                Dim kind = ReadFinancialNumber(If(values.ElementAtOrDefault(2), 1.0))
                If TypeOf kind Is FormulaError Then Return kind
                Dim format = LogicalOptionValue(If(values.ElementAtOrDefault(3), True))
                If TypeOf format Is FormulaError Then Return format
                Dim rowNumber = Math.Truncate(DirectCast(row, Double))
                Dim columnNumber = Math.Truncate(DirectCast(column, Double))
                Dim mode = Math.Truncate(DirectCast(kind, Double))
                If mode < 1.0 OrElse mode > 4.0 Then Return New FormulaError("#VALUE!")
                Dim a1 = DirectCast(format, Boolean)
                Dim absoluteRow = mode <= 2.0
                Dim absoluteColumn = mode = 1.0 OrElse mode = 3.0
                If Not AddressCoordinateFits(rowNumber, RowCount, a1 OrElse absoluteRow) OrElse
                    Not AddressCoordinateFits(columnNumber, ColumnCount, a1 OrElse absoluteColumn) Then Return New FormulaError("#VALUE!")

                Dim german = ConversionCulture.TwoLetterISOLanguageName = "de"
                Dim address As String
                If a1 Then
                    address = New ExcelCell(Nothing, CInt(rowNumber) - 1, CInt(columnNumber) - 1, ExcelCell.ValueTypes.All).
                        Address(False, absoluteColumn, absoluteRow)
                Else
                    Dim opening = If(german, "(", "[")
                    Dim closing = If(german, ")", "]")
                    Dim rowText = rowNumber.ToString(CultureInfo.InvariantCulture)
                    Dim columnText = columnNumber.ToString(CultureInfo.InvariantCulture)
                    address = If(german, "Z", "R") & If(absoluteRow, rowText, If(rowNumber = 0.0, "", opening & rowText & closing)) &
                        If(german, "S", "C") & If(absoluteColumn, columnText, If(columnNumber = 0.0, "", opening & columnText & closing))
                End If
                If values.Length < 5 OrElse values(4) Is Nothing Then Return address
                Dim sheet = TextValue(values(4))
                Dim reserved = If(german, {"WAHR", "FALSCH", "R", "C", "Z", "S"}, {"TRUE", "FALSE", "R", "C"})
                If sheet <> "" AndAlso (Not UnquotedSheet.IsMatch(sheet) OrElse A1Coordinate.IsMatch(sheet) OrElse
                    EnglishRowColumn.IsMatch(sheet) OrElse GermanRowColumn.IsMatch(sheet) OrElse reserved.Contains(sheet, StringComparer.OrdinalIgnoreCase)) Then
                    sheet = "'" & sheet.Replace("'", "''") & "'"
                End If
                Return sheet & "!" & address
            End Function

            Private Shared Function AddressCoordinateFits(value As Double, limit As Integer, absolute As Boolean) As Boolean
                Return If(absolute, value >= 1.0 AndAlso value <= limit, Math.Abs(value) < limit)
            End Function

            Private Function IndirectValue(values As Object(), formulaCell As ExcelCell) As Object
                Dim style = LogicalOptionValue(If(values.ElementAtOrDefault(1), True))
                If TypeOf style Is FormulaError Then Return style
                Dim text = TextValue(values(0))
                Dim matched = IndirectAddress.Match(text)
                If Not matched.Success Then Return New FormulaError("#REF!")
                Dim qualifier = matched.Groups("sheet").Value
                Dim sourceSheet = formulaCell.SheetName
                If qualifier <> "" Then
                    sourceSheet = If(qualifier(0) = "'"c, qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'"), qualifier)
                End If
                Dim sourceWorkbook As OpenXmlExcelDataOperations = Nothing
                Dim source = Me
                If sourceSheet.Contains("]"c) Then
                    sourceWorkbook = FindWorkbookForTextReference(sourceSheet, sourceSheet)
                    If sourceWorkbook Is Nothing Then Return New FormulaError("#REF!")
                    source = ForWorkbook(sourceWorkbook)
                End If
                If Not source._worksheets.ContainsKey(sourceSheet) Then Return New FormulaError("#REF!")

                Dim address = matched.Groups("address").Value
                Dim parts = address.Split(":"c)
                Dim first As (Row As Integer, Column As Integer)
                Dim last As (Row As Integer, Column As Integer)
                If DirectCast(style, Boolean) Then
                    If parts.Length <= 2 AndAlso TryA1Coordinates(parts(0), first) AndAlso
                        TryA1Coordinates(parts(parts.Length - 1), last) Then Return MakeArea(sourceSheet, first, last, Nothing, sourceWorkbook)
                    If parts.Length = 2 AndAlso TryWholeReference(parts, first, last) Then Return MakeArea(sourceSheet, first, last, Nothing, sourceWorkbook)
                Else
                    If parts.Length <= 2 AndAlso TryRowColumnCoordinates(parts(0), formulaCell, first) AndAlso
                        TryRowColumnCoordinates(parts(parts.Length - 1), formulaCell, last) Then Return MakeArea(sourceSheet, first, last, Nothing, sourceWorkbook)
                End If
                If Not IsValidDefinedNameToken(address) Then Return New FormulaError("#REF!")
                If sourceWorkbook IsNot Nothing Then Return New FormulaError("#REF!")
                Dim definition = EvaluateDefinedName(address, If(qualifier = "", Nothing, sourceSheet), formulaCell)
                Return If(TryCast(definition, FormulaArea), CObj(New FormulaError("#REF!")))
            End Function

            Private Shared Function TryA1Coordinates(text As String, ByRef result As (Row As Integer, Column As Integer)) As Boolean
                Dim coordinate = A1Coordinate.Match(text)
                If Not coordinate.Success Then Return False
                Dim row = Integer.Parse(coordinate.Groups(2).Value, CultureInfo.InvariantCulture) - 1
                Dim column = ExcelCell.CalculateColumnIndex(coordinate.Groups(1).Value.ToUpperInvariant())
                result = (row, column)
                Return row < RowCount AndAlso column < ColumnCount
            End Function

            Private Shared Function TryWholeReference(parts As String(), ByRef first As (Row As Integer, Column As Integer),
                                                      ByRef last As (Row As Integer, Column As Integer)) As Boolean
                Dim left = A1Column.Match(parts(0))
                Dim right = A1Column.Match(parts(1))
                If left.Success AndAlso right.Success Then
                    first = (0, ExcelCell.CalculateColumnIndex(left.Groups(1).Value.ToUpperInvariant()))
                    last = (RowCount - 1, ExcelCell.CalculateColumnIndex(right.Groups(1).Value.ToUpperInvariant()))
                    Return first.Column < ColumnCount AndAlso last.Column < ColumnCount
                End If
                Dim top = A1Row.Match(parts(0))
                Dim bottom = A1Row.Match(parts(1))
                If Not top.Success OrElse Not bottom.Success Then Return False
                first = (Integer.Parse(top.Groups(1).Value, CultureInfo.InvariantCulture) - 1, 0)
                last = (Integer.Parse(bottom.Groups(1).Value, CultureInfo.InvariantCulture) - 1, ColumnCount - 1)
                Return first.Row < RowCount AndAlso last.Row < RowCount
            End Function

            Private Function TryRowColumnCoordinates(text As String, cell As ExcelCell, ByRef result As (Row As Integer, Column As Integer)) As Boolean
                Dim pattern = If(ConversionCulture.TwoLetterISOLanguageName = "de", GermanRowColumn, EnglishRowColumn)
                Dim matched = pattern.Match(text)
                If Not matched.Success Then Return False
                Dim row = RelativeCoordinate(matched.Groups(1).Value, cell.RowIndex)
                Dim column = RelativeCoordinate(matched.Groups(2).Value, cell.ColumnIndex)
                If row < 0 OrElse row >= RowCount OrElse column < 0 OrElse column >= ColumnCount Then Return False
                result = (CInt(row), CInt(column))
                Return True
            End Function

            Private Shared Function RelativeCoordinate(text As String, origin As Integer) As Long
                If text.Length = 0 Then Return origin
                Dim relative = text(0) = "["c OrElse text(0) = "("c
                Dim number As Integer
                If Not Integer.TryParse(If(relative, text.Substring(1, text.Length - 2), text),
                    NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, number) Then Return -1
                Return If(relative, CLng(origin) + number, CLng(number) - 1L)
            End Function

            Private Function OffsetValue(area As FormulaArea, values As Object()) As Object
                Dim numbers = {0.0, 0.0, CDbl(area.LastRow - area.FirstRow + 1), CDbl(area.LastColumn - area.FirstColumn + 1)}
                For index = 0 To values.Length - 1
                    If index >= 2 AndAlso values(index) Is Nothing Then Continue For
                    Dim value = ReadFinancialNumber(values(index))
                    If TypeOf value Is FormulaError Then Return value
                    numbers(index) = Math.Truncate(DirectCast(value, Double))
                Next
                Dim top = area.FirstRow + numbers(0)
                Dim left = area.FirstColumn + numbers(1)
                Dim height = numbers(2)
                Dim width = numbers(3)
                If height = 0.0 OrElse width = 0.0 Then Return New FormulaError("#REF!")
                Dim bottom = top + height - Math.Sign(height)
                Dim right = left + width - Math.Sign(width)
                If Math.Min(top, bottom) < 0 OrElse Math.Max(top, bottom) >= RowCount OrElse
                    Math.Min(left, right) < 0 OrElse Math.Max(left, right) >= ColumnCount Then Return New FormulaError("#REF!")
                Return MakeArea(area.Sheet, (CInt(top), CInt(left)), (CInt(bottom), CInt(right)), area.WorkbookIndex, area.SourceWorkbook)
            End Function

            Private Shared Function MakeArea(sheet As String, first As (Row As Integer, Column As Integer),
                                             last As (Row As Integer, Column As Integer), workbookIndex As Integer?,
                                             Optional sourceWorkbook As OpenXmlExcelDataOperations = Nothing) As FormulaArea
                Return New FormulaArea(sheet, New ExcelCell(Nothing, first.Row, first.Column, ExcelCell.ValueTypes.All),
                    New ExcelCell(Nothing, last.Row, last.Column, ExcelCell.ValueTypes.All), workbookIndex, sourceWorkbook)
            End Function

            Private Function EvaluateChoose(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 255 Then Return New FormulaError("#VALUE!")
                Dim input = EvaluateExpression(arguments(0), formulaCell)
                If Not TypeOf input Is FormulaMatrix AndAlso Not IsMultiValue(input) Then
                    Dim choice = ChooseIndex(Scalar(input), arguments.Count - 1)
                    If TypeOf choice Is FormulaError Then Return choice
                    Return EvaluateExpression(arguments(DirectCast(choice, Integer)), formulaCell)
                End If
                Dim indexes = AsMatrix(input)
                Dim candidates = arguments.Skip(1).Select(Function(argument) AsMatrix(EvaluateExpression(argument, formulaCell))).ToArray()
                Dim rows = Math.Max(indexes.Rows, candidates.Max(Function(matrix) matrix.Rows))
                Dim columns = Math.Max(indexes.Columns, candidates.Max(Function(matrix) matrix.Columns))
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim choice = ChooseIndex(indexes.BroadcastAt(row, column), candidates.Length)
                        results(row * columns + column) = If(TypeOf choice Is FormulaError, choice,
                            candidates(DirectCast(choice, Integer) - 1).BroadcastAt(row, column))
                    Next
                Next
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function ChooseIndex(value As Object, count As Integer) As Object
                Dim converted = ReadFinancialNumber(value)
                If TypeOf converted Is FormulaError Then Return converted
                Dim index = Math.Truncate(DirectCast(converted, Double))
                If index < 1.0 OrElse index > count Then Return New FormulaError("#VALUE!")
                Return CInt(index)
            End Function

        End Class

    End Class

End Namespace
