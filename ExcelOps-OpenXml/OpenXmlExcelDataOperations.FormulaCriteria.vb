Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation
            Private Const MaxMaterializedReferenceCells As Integer = 100000

            Private Function EvaluateCriteriaFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                      formulaCell As ExcelCell) As Object
                Dim countOnly = name = "COUNTIF" OrElse name = "COUNTIFS"
                Dim multiple = name.EndsWith("IFS", StringComparison.Ordinal)
                Dim firstCriterion = If(multiple AndAlso Not countOnly, 1, 0)
                Dim criterionCount = If(multiple, (arguments.Count - firstCriterion) \ 2, 1)
                If multiple Then
                    If criterionCount < 1 OrElse (arguments.Count - firstCriterion) Mod 2 <> 0 Then Return New FormulaError("#VALUE!")
                ElseIf arguments.Count < 2 OrElse arguments.Count > If(countOnly, 2, 3) Then
                    Return New FormulaError("#VALUE!")
                End If

                Dim ranges As New List(Of (Area As FormulaArea, Matches As Predicate(Of Object)))
                Dim firstArea As FormulaArea = Nothing
                For index = 0 To criterionCount - 1
                    Dim argumentIndex = firstCriterion + index * 2
                    Dim input = EvaluateExpression(arguments(argumentIndex), formulaCell)
                    If TypeOf input Is FormulaError Then Return input
                    Dim area = TryCast(input, FormulaArea)
                    If area Is Nothing Then Return New FormulaError("#VALUE!")
                    If firstArea Is Nothing Then firstArea = area
                    If Not SameSize(firstArea, area) Then Return New FormulaError("#VALUE!")
                    Dim value = Scalar(EvaluateExpression(arguments(argumentIndex + 1), formulaCell))
                    ranges.Add((area, MakeCriterion(value)))
                Next

                Dim amounts As FormulaArea = Nothing
                If Not countOnly Then
                    If Not multiple AndAlso arguments.Count = 2 Then
                        amounts = firstArea
                    Else
                        Dim input = EvaluateExpression(arguments(If(multiple, 0, 2)), formulaCell)
                        If TypeOf input Is FormulaError Then Return input
                        amounts = TryCast(input, FormulaArea)
                        If amounts Is Nothing Then Return New FormulaError("#VALUE!")
                        If Not multiple Then
                            Dim lastRow = amounts.FirstRow + firstArea.LastRow - firstArea.FirstRow
                            Dim lastColumn = amounts.FirstColumn + firstArea.LastColumn - firstArea.FirstColumn
                            ValidateCoordinates(lastRow, lastColumn)
                            amounts = SliceArea(amounts, 0, 0, lastRow - amounts.FirstRow, lastColumn - amounts.FirstColumn)
                        ElseIf Not SameSize(firstArea, amounts) Then
                            Return New FormulaError("#VALUE!")
                        End If
                    End If
                End If

                Dim columns = firstArea.LastColumn - firstArea.FirstColumn + 1
                Dim size = CLng(firstArea.LastRow - firstArea.FirstRow + 1) * columns
                Dim count As Long
                Dim matches As IEnumerable(Of Object)
                If size > MaxMaterializedReferenceCells Then
                    matches = FindSparseCriteriaMatches(ranges, amounts, columns, size, count)
                Else
                    matches = DenseCriteriaMatches(ranges, amounts)
                End If

                Dim minimum = name = "MINIFS"
                Dim maximum = name = "MAXIFS"
                Dim average = name = "AVERAGEIF" OrElse name = "AVERAGEIFS"
                Dim result = If(minimum, Double.PositiveInfinity, If(maximum, Double.NegativeInfinity, 0.0))
                For Each value In matches
                    If countOnly Then
                        count += 1
                        Continue For
                    End If
                    If TypeOf value Is FormulaError Then Return value
                    If Not TypeOf value Is Double Then Continue For
                    count += 1
                    Dim number = DirectCast(value, Double)
                    If minimum Then
                        result = Math.Min(result, number)
                    ElseIf maximum Then
                        result = Math.Max(result, number)
                    Else
                        result += number
                    End If
                Next

                If countOnly Then Return CDbl(count)
                If average Then
                    If count = 0 Then Return New FormulaError("#DIV/0!")
                    Return FiniteNumber(result / count)
                End If
                Return If(count = 0, 0.0, FiniteNumber(result))
            End Function

            Private Iterator Function DenseCriteriaMatches(ranges As List(Of (Area As FormulaArea, Matches As Predicate(Of Object))),
                                                           amounts As FormulaArea) As IEnumerable(Of Object)
                Dim dense = ranges.Select(Function(range) (Values:=MaterializeRange(range.Area), Matches:=range.Matches)).ToArray()
                Dim output As Object() = Nothing
                If amounts IsNot Nothing Then
                    output = If(amounts Is ranges(0).Area, dense(0).Values, MaterializeRange(amounts))
                End If
                For index = 0 To dense(0).Values.Length - 1
                    Dim position = index
                    If dense.All(Function(input) input.Matches(input.Values(position))) Then Yield If(output Is Nothing, Nothing, output(index))
                Next
            End Function

            Private Function FindSparseCriteriaMatches(ranges As List(Of (Area As FormulaArea, Matches As Predicate(Of Object))),
                                                       amounts As FormulaArea, columns As Integer, size As Long,
                                                       ByRef blankMatches As Long) As IEnumerable(Of Object)
                Dim bindings = ranges.Select(Function(range) (Reader:=BindSparseRange(range.Area), Matches:=range.Matches)).ToArray()
                Dim output = If(amounts Is Nothing, Nothing, BindSparseRange(amounts))
                Dim locations As New HashSet(Of Long)
                If output IsNot Nothing Then
                    output.AddStoredPositions(locations, columns)
                Else
                    Dim required = bindings.FirstOrDefault(Function(binding) Not binding.Matches(Nothing)).Reader
                    If required IsNot Nothing Then
                        required.AddStoredPositions(locations, columns)
                    Else
                        For Each binding In bindings
                            binding.Reader.AddStoredPositions(locations, columns)
                        Next
                        blankMatches = size - locations.Count
                    End If
                End If
                Return SparseCriteriaMatches(bindings, output, locations.OrderBy(Function(index) index), columns)
            End Function

            Private Iterator Function SparseCriteriaMatches(ranges As (Reader As SparseRangeReader, Matches As Predicate(Of Object))(),
                                                            amounts As SparseRangeReader, candidates As IEnumerable(Of Long),
                                                            columns As Integer) As IEnumerable(Of Object)
                For Each index In candidates
                    Dim row = CInt(index \ columns)
                    Dim column = CInt(index Mod columns)
                    If ranges.All(Function(input) input.Matches(input.Reader.Read(row, column))) Then
                        Yield If(amounts Is Nothing, Nothing, amounts.Read(row, column))
                    End If
                Next
            End Function

            Private Function BindSparseRange(area As FormulaArea) As SparseRangeReader
                Dim source = ResolveReference(area)
                Dim worksheet As WorksheetSnapshot = Nothing
                If Not source._worksheets.TryGetValue(area.Sheet, worksheet) Then Throw New NotSupportedException("The referenced worksheet is unavailable.")
                RecordPrecedent(New FormulaPrecedent(New FormulaLocation(source._owner, area.Sheet, area.FirstRow, area.FirstColumn), area.LastRow, area.LastColumn))
                RecalculateReferencedSpills(source, worksheet, area)
                Return New SparseRangeReader(source, worksheet, area)
            End Function

            Private NotInheritable Class SparseRangeReader
                Private ReadOnly _source As FormulaCalculation
                Private ReadOnly _worksheet As WorksheetSnapshot
                Private ReadOnly _area As FormulaArea

                Friend Sub New(source As FormulaCalculation, worksheet As WorksheetSnapshot, area As FormulaArea)
                    _source = source
                    _worksheet = worksheet
                    _area = area
                End Sub

                Friend Function Read(row As Integer, column As Integer) As Object
                    Return _source.ReadCell(_area.Sheet, row + _area.FirstRow, column + _area.FirstColumn, trackRead:=False)
                End Function

                Friend Sub AddStoredPositions(positions As HashSet(Of Long), columns As Integer)
                    For Each coordinate In _worksheet.Cells.Keys.Concat(_worksheet.SpillChildren.Keys)
                        If coordinate.RowIndex >= _area.FirstRow AndAlso coordinate.RowIndex <= _area.LastRow AndAlso
                            coordinate.ColumnIndex >= _area.FirstColumn AndAlso coordinate.ColumnIndex <= _area.LastColumn Then
                            positions.Add(CLng(coordinate.RowIndex - _area.FirstRow) * columns + coordinate.ColumnIndex - _area.FirstColumn)
                        End If
                    Next
                End Sub
            End Class

            Private Function MaterializeRange(area As FormulaArea) As Object()
                Dim rows = area.LastRow - area.FirstRow + 1
                Dim columns = area.LastColumn - area.FirstColumn + 1
                If CLng(rows) * columns > MaxMaterializedReferenceCells Then
                    Throw New NotSupportedException("The range contains more cells than this function can materialize.")
                End If
                Dim source = ResolveReference(area)
                If Not source._worksheets.ContainsKey(area.Sheet) Then Throw New NotSupportedException("The referenced worksheet is unavailable.")
                RecordPrecedent(New FormulaPrecedent(New FormulaLocation(source._owner, area.Sheet, area.FirstRow, area.FirstColumn), area.LastRow, area.LastColumn))
                Dim values(rows * columns - 1) As Object
                Dim index = 0
                For row = area.FirstRow To area.LastRow
                    For column = area.FirstColumn To area.LastColumn
                        values(index) = source.ReadCell(area.Sheet, row, column, trackRead:=False)
                        index += 1
                    Next
                Next
                Return values
            End Function

            Private Shared Function SameSize(first As FormulaArea, second As FormulaArea) As Boolean
                Return first.LastRow - first.FirstRow = second.LastRow - second.FirstRow AndAlso
                    first.LastColumn - first.FirstColumn = second.LastColumn - second.FirstColumn
            End Function

            Private Shared Function SliceArea(area As FormulaArea, firstRow As Integer, firstColumn As Integer,
                                              lastRow As Integer, lastColumn As Integer) As FormulaArea
                Return New FormulaArea(area.Sheet,
                    New ExcelCell(Nothing, area.FirstRow + firstRow, area.FirstColumn + firstColumn, ExcelCell.ValueTypes.All),
                    New ExcelCell(Nothing, area.FirstRow + lastRow, area.FirstColumn + lastColumn, ExcelCell.ValueTypes.All), area.WorkbookIndex, area.SourceWorkbook)
            End Function

            Private Function MakeCriterion(value As Object) As Predicate(Of Object)
                If value Is Nothing Then Return NumericCriterion("=", 0.0)
                Dim criterionError = TryCast(value, FormulaError)
                If criterionError IsNot Nothing Then Return ErrorCriterion(criterionError.Code, negate:=False)
                If TypeOf value Is Double Then Return NumericCriterion("=", DirectCast(value, Double))
                If TypeOf value Is Boolean Then
                    Dim expected = DirectCast(value, Boolean)
                    Return Function(actual) TypeOf actual Is Boolean AndAlso DirectCast(actual, Boolean) = expected
                End If

                Dim text = DirectCast(value, String)
                Dim operation = "="
                If text.StartsWith("<=", StringComparison.Ordinal) OrElse text.StartsWith(">=", StringComparison.Ordinal) OrElse text.StartsWith("<>", StringComparison.Ordinal) Then
                    operation = text.Substring(0, 2)
                    text = text.Substring(2)
                ElseIf text.Length > 0 AndAlso "=<>".IndexOf(text(0)) >= 0 Then
                    operation = text.Substring(0, 1)
                    text = text.Substring(1)
                End If
                If operation = "=" OrElse operation = "<>" Then
                    Dim errorCode = CriterionErrorCode(text)
                    If errorCode IsNot Nothing Then Return ErrorCriterion(errorCode, negate:=operation = "<>")
                End If
                Dim number As Double
                If Double.TryParse(text, NumberStyles.Float, ConversionCulture, number) AndAlso Not Double.IsInfinity(number) AndAlso Not Double.IsNaN(number) Then
                    Return NumericCriterion(operation, number)
                End If
                If text.Length = 0 Then
                    If operation = "=" Then Return Function(actual) actual Is Nothing OrElse (TypeOf actual Is String AndAlso DirectCast(actual, String).Length = 0)
                    If operation = "<>" Then Return Function(actual) actual IsNot Nothing
                End If
                Dim pattern = BuildWildcard(text)
                If operation = "=" OrElse operation = "<>" Then
                    Return Function(actual)
                               Dim equal = TypeOf actual Is String AndAlso MatchesWildcard(pattern, DirectCast(actual, String))
                               Return If(operation = "=", equal, Not equal)
                           End Function
                End If
                Return Function(actual) TypeOf actual Is String AndAlso EvaluateComparison(operation, actual, text)
            End Function

            Private Function CriterionErrorCode(text As String) As String
                Dim label = text.ToUpperInvariant()
                Select Case label
                    Case "#NULL!", "#DIV/0!", "#NAME?"
                        Return label
                End Select
                If ConversionCulture.TwoLetterISOLanguageName = "de" Then
                    Select Case label
                        Case "#NV" : Return "#N/A"
                        Case "#WERT!" : Return "#VALUE!"
                        Case "#BEZUG!" : Return "#REF!"
                        Case "#ZAHL!" : Return "#NUM!"
                    End Select
                Else
                    Select Case label
                        Case "#N/A", "#VALUE!", "#REF!", "#NUM!"
                            Return label
                    End Select
                End If
                Return Nothing
            End Function

            Private Shared Function ErrorCriterion(errorCode As String, negate As Boolean) As Predicate(Of Object)
                Return Function(actual) (TypeOf actual Is FormulaError AndAlso DirectCast(actual, FormulaError).Code = errorCode) <> negate
            End Function

            Private Function NumericCriterion(operation As String, number As Double) As Predicate(Of Object)
                Return Function(actual)
                           If TypeOf actual Is Double Then Return EvaluateComparison(operation, actual, number)
                           If operation = "<>" Then Return True
                           If operation <> "=" OrElse Not TypeOf actual Is String Then Return False
                           Dim parsed As Double
                           Return Double.TryParse(DirectCast(actual, String), NumberStyles.Float, ConversionCulture, parsed) AndAlso parsed = number
                       End Function
            End Function

            Private Shared Function BuildWildcard(text As String, Optional searchWithinText As Boolean = False) As Regex
                Dim pattern As New StringBuilder(If(searchWithinText, "", "\A"))
                Dim index = 0
                While index < text.Length
                    Dim character = text(index)
                    If character = "~"c AndAlso index + 1 < text.Length AndAlso
                        (searchWithinText OrElse "*?~".IndexOf(text(index + 1)) >= 0) Then
                        index += 1
                        pattern.Append(Regex.Escape(text(index).ToString()))
                    ElseIf character = "*"c Then
                        pattern.Append(".*")
                        While index + 1 < text.Length AndAlso text(index + 1) = "*"c
                            index += 1
                        End While
                    ElseIf character = "?"c Then
                        pattern.Append("."c)
                    ElseIf Not (searchWithinText AndAlso character = "~"c) Then
                        pattern.Append(Regex.Escape(character.ToString()))
                    End If
                    index += 1
                End While
                If Not searchWithinText Then pattern.Append("\z")
                Dim options = RegexOptions.CultureInvariant Or RegexOptions.Singleline
                If Not searchWithinText Then options = options Or RegexOptions.IgnoreCase
                Return New Regex(pattern.ToString(), options, TimeSpan.FromMilliseconds(200))
            End Function

            Private Shared Function MatchesWildcard(pattern As Regex, text As String) As Boolean
                Try
                    Return pattern.IsMatch(text)
                Catch ex As RegexMatchTimeoutException
                    Throw New NotSupportedException("The wildcard comparison exceeds the supported complexity.", ex)
                End Try
            End Function
        End Class

    End Class

End Namespace
