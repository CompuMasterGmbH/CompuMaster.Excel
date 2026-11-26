Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateTextDelimiterFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                          formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 6 Then Return New FormulaError("#VALUE!")
                Dim supplied(5) As Object
                For index = 0 To arguments.Count - 1
                    supplied(index) = EvaluateExpression(arguments(index), formulaCell)
                Next
                Dim split = name = "TEXTSPLIT"
                Dim delimiters As New List(Of (Text As String, IsRow As Boolean))
                Dim delimiterError = ReadTextDelimiters(supplied(1), False, split, delimiters)
                If delimiterError IsNot Nothing Then Return delimiterError
                If split Then
                    delimiterError = ReadTextDelimiters(supplied(2), True, True, delimiters)
                    If delimiterError IsNot Nothing Then Return delimiterError
                    If delimiters.Count = 0 OrElse IsMultiValue(supplied(5)) Then Return New FormulaError("#VALUE!")
                End If

                Dim subjects = AsMatrix(supplied(0))
                Dim firstOption = AsMatrix(If(split, supplied(3), supplied(2)))
                Dim caseMode = AsMatrix(If(split, supplied(4), supplied(3)))
                Dim endMode = AsMatrix(If(split, Nothing, supplied(4)))
                Dim replacements = AsMatrix(If(split, Nothing, supplied(5)))
                Dim rows = Math.Max(subjects.Rows, Math.Max(firstOption.Rows, Math.Max(caseMode.Rows, Math.Max(endMode.Rows, replacements.Rows))))
                Dim columns = Math.Max(subjects.Columns, Math.Max(firstOption.Columns, Math.Max(caseMode.Columns, Math.Max(endMode.Columns, replacements.Columns))))
                Dim results = AllocateMatrixValues(rows, columns)
                Dim padding = If(split AndAlso arguments.Count = 6, If(Scalar(supplied(5)), 0.0), New FormulaError("#N/A"))

                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim subject = subjects.BroadcastAt(row, column)
                        Dim first = firstOption.BroadcastAt(row, column)
                        Dim caseValue = caseMode.BroadcastAt(row, column)
                        Dim result As Object
                        If TypeOf subject Is FormulaError Then
                            result = subject
                        ElseIf split Then
                            result = SplitText(TextValue(subject), delimiters, first, caseValue, padding)
                        Else
                            Dim instance = If(supplied(2) Is Nothing, 1.0, first)
                            Dim fallback = If(supplied(5) Is Nothing, New FormulaError("#N/A"), replacements.BroadcastAt(row, column))
                            result = SelectDelimitedText(name, TextValue(subject), delimiters, instance, supplied(2) IsNot Nothing,
                                                         caseValue, endMode.BroadcastAt(row, column), fallback)
                        End If
                        If results.Length = 1 Then Return result
                        Dim nestedResult = TryCast(result, FormulaMatrix)
                        results(row * columns + column) = If(nestedResult Is Nothing, result, nestedResult.At(0, 0))
                    Next
                Next
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function ReadTextDelimiters(input As Object, isRow As Boolean, splitting As Boolean,
                                                target As List(Of (Text As String, IsRow As Boolean))) As FormulaError
                If input Is Nothing AndAlso splitting Then Return Nothing
                For Each item In AsMatrix(input).Values
                    If TypeOf item Is FormulaError Then Return DirectCast(item, FormulaError)
                    Dim delimiter = TextValue(item)
                    If splitting AndAlso delimiter.Length = 0 Then Return New FormulaError("#VALUE!")
                    target.Add((delimiter, isRow))
                Next
                Return Nothing
            End Function

            Private Function SelectDelimitedText(name As String, subject As String, delimiters As List(Of (Text As String, IsRow As Boolean)),
                                                instance As Object, explicitInstance As Boolean, caseMode As Object,
                                                endMode As Object, fallback As Object) As Object
                Dim occurrence = NumberValue(instance)
                If TypeOf occurrence Is FormulaError Then Return occurrence
                Dim caseValue = ReadTextMatchMode(caseMode)
                If TypeOf caseValue Is FormulaError Then Return caseValue
                Dim endValue = ReadTextMatchMode(endMode)
                If TypeOf endValue Is FormulaError Then Return endValue

                Dim count = Math.Floor(DirectCast(occurrence, Double))
                If count = 0.0 Then Return New FormulaError("#VALUE!")
                Dim reverse = count < 0.0
                If delimiters.Any(Function(delimiter) delimiter.Text.Length = 0) Then
                    Return DelimitedSubstring(name, subject, If(reverse, subject.Length, 0), 0)
                End If
                If Math.Abs(count) > subject.Length Then
                    If Not explicitInstance AndAlso subject.Length = 0 Then Return DelimiterFallback(fallback)
                    Return New FormulaError("#VALUE!")
                End If

                Dim comparison = If(DirectCast(caseValue, Boolean), StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
                Dim cursor = If(reverse, subject.Length, 0)
                For offset = 1 To CInt(Math.Abs(count))
                    Dim delimiter = FindTextDelimiter(subject, delimiters, cursor, reverse, comparison)
                    If delimiter.Index < 0 Then
                        If offset = CInt(Math.Abs(count)) AndAlso DirectCast(endValue, Boolean) Then
                            Return DelimitedSubstring(name, subject, If(reverse, 0, subject.Length), 0)
                        End If
                        Return DelimiterFallback(fallback)
                    End If
                    If offset = CInt(Math.Abs(count)) Then Return DelimitedSubstring(name, subject, delimiter.Index, delimiter.Length)
                    cursor = If(reverse, delimiter.Index, delimiter.Index + delimiter.Length)
                Next
                Return DelimiterFallback(fallback)
            End Function

            Private Function DelimiterFallback(value As Object) As Object
                If TypeOf value Is FormulaError Then Return value
                Return TextValue(value)
            End Function

            Private Shared Function DelimitedSubstring(name As String, subject As String, index As Integer, length As Integer) As String
                If name = "TEXTBEFORE" Then Return subject.Substring(0, index)
                Return subject.Substring(index + length)
            End Function

            Private Function SplitText(subject As String, delimiters As List(Of (Text As String, IsRow As Boolean)),
                                       ignoreEmpty As Object, caseMode As Object, padding As Object) As Object
                If subject.Length = 0 Then Return New FormulaError("#VALUE!")
                Dim emptyMode = LogicalOptionValue(ignoreEmpty)
                If TypeOf emptyMode Is FormulaError Then Return emptyMode
                Dim caseValue = ReadTextMatchMode(caseMode, truncateTowardZero:=True)
                If TypeOf caseValue Is FormulaError Then Return caseValue
                Dim ignore = DirectCast(emptyMode, Boolean)
                Dim comparison = If(DirectCast(caseValue, Boolean), StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)

                Dim remaining As New List(Of List(Of Object))
                Dim current As New List(Of Object)
                Dim cursor = 0
                Do
                    Dim delimiter = FindTextDelimiter(subject, delimiters, cursor, False, comparison)
                    Dim last = If(delimiter.Index < 0, subject.Length, delimiter.Index)
                    If last > cursor OrElse Not ignore Then current.Add(subject.Substring(cursor, last - cursor))
                    If delimiter.Index < 0 OrElse delimiter.IsRow Then
                        If current.Count > 0 Then remaining.Add(current)
                        current = New List(Of Object)
                    End If
                    If delimiter.Index < 0 Then Exit Do
                    cursor = delimiter.Index + delimiter.Length
                Loop
                If remaining.Count = 0 Then Return New FormulaError(CalcErrorCode)

                Dim columns = remaining.Max(Function(row) row.Count)
                If remaining.Count = 1 AndAlso columns = 1 Then Return remaining(0)(0)
                Dim values = AllocateMatrixValues(remaining.Count, columns)
                For row = 0 To remaining.Count - 1
                    For column = 0 To columns - 1
                        values(row * columns + column) = If(column < remaining(row).Count, remaining(row)(column), padding)
                    Next
                Next
                Return New FormulaMatrix(remaining.Count, columns, values)
            End Function

            Private Function ReadTextMatchMode(input As Object, Optional truncateTowardZero As Boolean = False) As Object
                Dim numeric = NumberValue(input)
                If TypeOf numeric Is FormulaError Then Return numeric
                Dim value = DirectCast(numeric, Double)
                Dim mode = If(truncateTowardZero, Math.Truncate(value), Math.Floor(value))
                If mode <> 0.0 AndAlso mode <> 1.0 Then Return New FormulaError("#VALUE!")
                Return mode = 1.0
            End Function

            Private Shared Function FindTextDelimiter(subject As String, delimiters As List(Of (Text As String, IsRow As Boolean)),
                                                      cursor As Integer, reverse As Boolean, comparison As StringComparison
                                                      ) As (Index As Integer, Length As Integer, IsRow As Boolean)
                Dim best As (Index As Integer, Length As Integer, IsRow As Boolean) = (-1, 0, False)
                If reverse AndAlso cursor = 0 Then Return best
                For Each delimiter In delimiters
                    Dim index = If(reverse, subject.LastIndexOf(delimiter.Text, cursor - 1, cursor, comparison),
                                   subject.IndexOf(delimiter.Text, cursor, comparison))
                    If index < 0 Then Continue For

                    Dim nearer = best.Index < 0 OrElse If(reverse, index > best.Index, index < best.Index)
                    Dim preferredAtSamePosition = index = best.Index AndAlso
                        If(delimiter.IsRow = best.IsRow, delimiter.Text.Length > best.Length, Not delimiter.IsRow)
                    If nearer OrElse preferredAtSamePosition Then best = (index, delimiter.Text.Length, delimiter.IsRow)
                Next
                Return best
            End Function

        End Class

    End Class

End Namespace
