Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateOrderStatistic(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                    formulaCell As ExcelCell) As Object
                Dim values As New List(Of Double)
                If name = "MEDIAN" Then
                    If arguments.Count = 0 Then Return New FormulaError("#VALUE!")
                    For Each argument In arguments
                        Dim failure = AddOrderStatisticValues(EvaluateExpression(argument, formulaCell), values)
                        If failure IsNot Nothing Then Return failure
                    Next
                    If values.Count = 0 Then Return New FormulaError("#NUM!")
                    values.Sort()
                    Return InterpolateOrderedValues(values, (values.Count - 1) * 0.5)
                End If

                Dim isRank = name.StartsWith("RANK", StringComparison.Ordinal)
                If arguments.Count < 2 OrElse arguments.Count > If(isRank, 3, 2) Then Return New FormulaError("#VALUE!")
                Dim data = EvaluateExpression(arguments(If(isRank, 1, 0)), formulaCell)
                If isRank AndAlso Not TypeOf data Is FormulaArea Then Return New FormulaError("#VALUE!")
                Dim dataError = AddOrderStatisticValues(data, values)
                If dataError IsNot Nothing Then Return dataError
                values.Sort()

                Dim parameter = AsMatrix(EvaluateExpression(arguments(If(isRank, 0, 1)), formulaCell))
                Dim order = AsMatrix(If(isRank AndAlso arguments.Count = 3, EvaluateExpression(arguments(2), formulaCell), Nothing))
                Dim rows = Math.Max(parameter.Rows, order.Rows)
                Dim columns = Math.Max(parameter.Columns, order.Columns)
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim number = NumberValue(parameter.BroadcastAt(row, column))
                        Dim result As Object
                        If TypeOf number Is FormulaError Then
                            result = number
                        ElseIf isRank Then
                            result = EvaluateRank(values, DirectCast(number, Double), order.BroadcastAt(row, column), name = "RANK.AVG")
                        Else
                            result = SelectOrderStatistic(name, values, DirectCast(number, Double))
                        End If
                        results(row * columns + column) = result
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function AddOrderStatisticValues(input As Object, target As List(Of Double)) As FormulaError
                Dim area = TryCast(input, FormulaArea)
                Dim matrix = TryCast(input, FormulaMatrix)
                If area Is Nothing AndAlso matrix Is Nothing Then
                    Dim number = NumberValue(input)
                    If TypeOf number Is FormulaError Then Return DirectCast(number, FormulaError)
                    target.Add(DirectCast(number, Double))
                    Return Nothing
                End If

                Dim values = If(area IsNot Nothing, RangeValues(area), CType(matrix.Values, IEnumerable(Of Object)))
                For Each value In values
                    If TypeOf value Is FormulaError Then Return DirectCast(value, FormulaError)
                    If TypeOf value Is Double Then target.Add(DirectCast(value, Double))
                Next
                Return Nothing
            End Function

            Private Function EvaluateRank(values As List(Of Double), number As Double, order As Object, averageTies As Boolean) As Object
                Dim direction = LogicalOptionValue(order)
                If TypeOf direction Is FormulaError Then Return direction
                Dim before = RankBoundary(values, number, upper:=False)
                If before >= values.Count OrElse values(before) <> number Then Return New FormulaError("#N/A")
                If DirectCast(direction, Boolean) AndAlso Not averageTies Then Return CDbl(before + 1)
                Dim after = RankBoundary(values, number, upper:=True)
                Dim position = If(DirectCast(direction, Boolean), before + 1.0, values.Count - after + 1.0)
                If averageTies Then position += (after - before - 1) / 2.0
                Return position
            End Function

            Private Shared Function RankBoundary(values As List(Of Double), number As Double, upper As Boolean) As Integer
                Dim low = 0
                Dim high = values.Count
                While low < high
                    Dim middle = low + (high - low) \ 2
                    If values(middle) < number OrElse (upper AndAlso values(middle) = number) Then
                        low = middle + 1
                    Else
                        high = middle
                    End If
                End While
                Return low
            End Function

            Private Shared Function SelectOrderStatistic(name As String, values As List(Of Double), number As Double) As Object
                If name = "SMALL" OrElse name = "LARGE" Then
                    If number < 1 OrElse number > values.Count Then Return New FormulaError("#NUM!")
                    Dim index = If(name = "SMALL", Math.Floor(number) - 1, Math.Floor(values.Count - number))
                    Return values(CInt(index))
                End If

                If values.Count = 0 OrElse number < 0 Then Return New FormulaError("#NUM!")
                Dim percentile = If(name.StartsWith("QUARTILE", StringComparison.Ordinal), Math.Floor(number) / 4.0, number)
                Dim exclusive = name.EndsWith(".EXC", StringComparison.Ordinal)
                If percentile > 1.0 OrElse (exclusive AndAlso (percentile = 0.0 OrElse percentile = 1.0)) Then
                    Return New FormulaError("#NUM!")
                End If
                Dim position = If(exclusive, (values.Count + 1.0) * percentile - 1.0, (values.Count - 1.0) * percentile)
                If position < 0 OrElse position > values.Count - 1 Then Return New FormulaError("#NUM!")
                Return InterpolateOrderedValues(values, position)
            End Function

            Private Shared Function InterpolateOrderedValues(values As List(Of Double), position As Double) As Double
                Dim lower = CInt(Math.Floor(position))
                Dim fraction = position - lower
                If fraction = 0.0 OrElse values(lower) = values(lower + 1) Then Return values(lower)
                Return (1.0 - fraction) * values(lower) + fraction * values(lower + 1)
            End Function

        End Class

    End Class

End Namespace
