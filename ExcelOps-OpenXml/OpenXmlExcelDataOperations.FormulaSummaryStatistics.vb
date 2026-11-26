Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateSummaryStatistic(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                      formulaCell As ExcelCell) As Object
                If arguments.Count = 0 Then Return New FormulaError("#VALUE!")
                Dim isMode = name.StartsWith("MODE", StringComparison.Ordinal)
                Dim policy As StatisticInput
                Select Case name
                    Case "MODE", "MODE.SNGL", "MODE.MULT"
                        policy = StatisticInput.Mode
                    Case "AVERAGEA"
                        policy = StatisticInput.AverageWithText
                    Case "MINA", "MAXA", "VARA", "VARPA", "STDEVA", "STDEVPA"
                        policy = StatisticInput.IncludeReferenceTextAndBooleans
                    Case Else
                        policy = StatisticInput.Numeric
                End Select
                Dim values As New List(Of Double)
                For Each argument In arguments
                    Dim failure = AddStatisticValues(EvaluateExpression(argument, formulaCell), values, policy)
                    If failure IsNot Nothing Then Return failure
                Next
                If isMode Then Return Modes(values, name = "MODE.MULT")
                Select Case name
                    Case "MINA" : Return If(values.Count = 0, 0.0, values.Min())
                    Case "MAXA" : Return If(values.Count = 0, 0.0, values.Max())
                End Select

                Dim count = values.Count
                Dim isDeviation = name = "DEVSQ" OrElse name = "AVEDEV"
                If count = 0 Then Return New FormulaError(If(isDeviation, "#NUM!", "#DIV/0!"))
                If name = "AVERAGEA" Then Return FiniteNumber(values.Average())

                Dim sample = name = "VAR" OrElse name = "VAR.S" OrElse name = "VARA" OrElse
                    name = "STDEV" OrElse name = "STDEV.S" OrElse name = "STDEVA"
                If sample AndAlso count < 2 Then Return New FormulaError("#DIV/0!")

                ' Center the inputs before calculating the mean to preserve small deviations on large offsets.
                Dim origin = values(0)
                Dim mean = values.Average(Function(value) value - origin)
                Dim total As Double
                For Each value In values
                    Dim deviation = (value - origin) - mean
                    total += If(name = "AVEDEV", Math.Abs(deviation), deviation * deviation)
                Next
                If name = "DEVSQ" Then Return FiniteNumber(total)
                Dim result = total / If(sample, count - 1, count)
                If name.StartsWith("STDEV", StringComparison.Ordinal) Then result = Math.Sqrt(result)
                Return FiniteNumber(result)
            End Function

            Private Shared Function Modes(values As List(Of Double), multiple As Boolean) As Object
                Dim counts As New Dictionary(Of Double, Integer)
                Dim highest = 1
                For Each value In values
                    Dim occurrences As Integer
                    counts.TryGetValue(value, occurrences)
                    occurrences += 1
                    counts(value) = occurrences
                    highest = Math.Max(highest, occurrences)
                Next
                If highest < 2 Then Return New FormulaError("#N/A")
                If Not multiple Then Return values.First(Function(value) counts(value) = highest)

                Dim seen As New HashSet(Of Double)
                Dim results = values.Where(Function(value) counts(value) = highest AndAlso seen.Add(value)).
                    Select(Function(value) CObj(value)).ToArray()
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(results.Length, 1, results)
            End Function

        End Class

    End Class

End Namespace
