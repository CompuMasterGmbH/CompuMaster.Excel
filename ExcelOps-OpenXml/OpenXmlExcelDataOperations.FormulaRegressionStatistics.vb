Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluatePairedStatistic(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                Dim forecasting = name = "FORECAST" OrElse name = "FORECAST.LINEAR"
                If arguments.Count <> If(forecasting, 3, 2) Then Return New FormulaError("#VALUE!")
                Dim queries As FormulaMatrix = Nothing
                If forecasting Then
                    Dim estimates = EvaluateExpression(arguments(0), formulaCell)
                    If TypeOf estimates Is FormulaError Then Return estimates
                    Dim rawQueries = AsMatrix(estimates)
                    queries = New FormulaMatrix(rawQueries.Rows, rawQueries.Columns, rawQueries.Values.Select(
                        Function(value) ReadFinancialNumber(value)).ToArray())
                    If queries.Values.Length = 1 AndAlso TypeOf queries.Values(0) Is FormulaError Then Return queries.Values(0)
                End If

                Dim start = If(forecasting, 1, 0)
                Dim first = AsMatrix(EvaluateExpression(arguments(start), formulaCell)).Values
                Dim second = AsMatrix(EvaluateExpression(arguments(start + 1), formulaCell)).Values
                Dim state As New PairedStatistics()
                Dim failure = ReadPairedStatistics(first, second, state)
                If Not forecasting Then Return If(failure, PairedStatisticValue(name, state))
                If failure Is Nothing AndAlso state.SumXX = 0.0 Then failure = New FormulaError("#DIV/0!")
                Dim results = queries.Values.Select(Function(value) ForecastValue(value, state, failure)).ToArray()
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(queries.Rows, queries.Columns, results)
            End Function

            Private Structure PairedStatistics
                Friend Count As Integer
                Friend MeanX As Double
                Friend MeanY As Double
                Friend SumXX As Double
                Friend SumYY As Double
                Friend SumXY As Double
            End Structure

            Private Shared Function ReadPairedStatistics(first As Object(), second As Object(), ByRef result As PairedStatistics) As FormulaError
                If first.Length <> second.Length Then Return New FormulaError("#N/A")
                Dim x As New List(Of Double)
                Dim y As New List(Of Double)
                For index = 0 To first.Length - 1
                    Dim yValue = first(index)
                    Dim xValue = second(index)
                    If TypeOf yValue Is FormulaError Then Return DirectCast(yValue, FormulaError)
                    If TypeOf xValue Is FormulaError Then Return DirectCast(xValue, FormulaError)
                    If Not TypeOf yValue Is Double OrElse Not TypeOf xValue Is Double Then Continue For
                    x.Add(DirectCast(xValue, Double))
                    y.Add(DirectCast(yValue, Double))
                Next
                If x.Count = 0 Then Return New FormulaError("#DIV/0!")
                result.Count = x.Count
                result.MeanX = MeanFromFirst(x)
                result.MeanY = MeanFromFirst(y)
                For index = 0 To x.Count - 1
                    Dim dx = x(index) - result.MeanX
                    Dim dy = y(index) - result.MeanY
                    result.SumXX += dx * dx
                    result.SumYY += dy * dy
                    result.SumXY += dx * dy
                Next
                Return Nothing
            End Function

            Private Shared Function MeanFromFirst(values As List(Of Double)) As Double
                Dim origin = values(0)
                Return origin + values.Average(Function(value) value - origin)
            End Function

            Private Shared Function PairedStatisticValue(name As String, state As PairedStatistics) As Object
                Select Case name
                    Case "COVARIANCE.P", "COVAR"
                        Return FiniteNumber(state.SumXY / state.Count)
                    Case "COVARIANCE.S"
                        If state.Count < 2 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber(state.SumXY / (state.Count - 1))
                End Select
                If state.SumXX = 0.0 Then Return New FormulaError("#DIV/0!")
                Dim slope = state.SumXY / state.SumXX
                Select Case name
                    Case "SLOPE"
                        Return FiniteNumber(slope)
                    Case "INTERCEPT"
                        Return FiniteNumber(state.MeanY - slope * state.MeanX)
                    Case "STEYX"
                        If state.Count < 3 Then Return New FormulaError("#DIV/0!")
                        Dim residual = state.SumYY - state.SumXY * state.SumXY / state.SumXX
                        If Double.IsInfinity(residual) OrElse Double.IsNaN(residual) Then Return New FormulaError("#NUM!")
                        Return FiniteNumber(Math.Sqrt(Math.Max(0.0, residual) / (state.Count - 2)))
                    Case Else
                        If state.SumYY = 0.0 Then Return New FormulaError("#DIV/0!")
                        Dim correlation = state.SumXY / Math.Sqrt(state.SumXX) / Math.Sqrt(state.SumYY)
                        Return FiniteNumber(If(name = "RSQ", correlation * correlation, correlation))
                End Select
            End Function

            Private Shared Function ForecastValue(value As Object, state As PairedStatistics, failure As FormulaError) As Object
                If TypeOf value Is FormulaError Then Return value
                If failure IsNot Nothing Then Return failure
                Return FiniteNumber(state.MeanY + (DirectCast(value, Double) - state.MeanX) * (state.SumXY / state.SumXX))
            End Function

            Private Function EvaluateShapeStatistic(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                Dim trim = name = "TRIMMEAN"
                If arguments.Count = 0 OrElse (trim AndAlso arguments.Count <> 2) Then Return New FormulaError("#VALUE!")
                Dim values As New List(Of Double)
                For index = 0 To If(trim, 0, arguments.Count - 1)
                    Dim failure = AddStatisticValues(EvaluateExpression(arguments(index), formulaCell), values)
                    If failure IsNot Nothing Then Return failure
                Next
                If trim Then
                    values.Sort()
                    Dim fraction = AsMatrix(EvaluateExpression(arguments(1), formulaCell))
                    Dim results = fraction.Values.Select(Function(value) TrimmedMean(values, ReadFinancialNumber(value))).ToArray()
                    If results.Length = 1 Then Return results(0)
                    Return New FormulaMatrix(fraction.Rows, fraction.Columns, results)
                End If

                If name = "GEOMEAN" OrElse name = "HARMEAN" Then
                    If values.Count = 0 Then Return New FormulaError(If(name = "GEOMEAN", "#NUM!", "#N/A"))
                    If values.Any(Function(value) value <= 0.0) Then Return New FormulaError("#NUM!")
                    If values.Count = 1 Then Return values(0)
                    If name = "GEOMEAN" Then Return FiniteNumber(Math.Exp(values.Average(Function(value) Math.Log(value))))
                    Dim smallest = values.Min()
                    Return FiniteNumber(smallest * (values.Count / values.Sum(Function(value) smallest / value)))
                End If

                Dim count = CDbl(values.Count)
                If count < If(name = "KURT", 4.0, 3.0) Then Return New FormulaError("#DIV/0!")
                Dim mean = MeanFromFirst(values)
                Dim square = values.Sum(Function(value) (value - mean) * (value - mean))
                If square = 0.0 Then Return New FormulaError("#DIV/0!")
                Dim population = name = "SKEW.P"
                Dim deviation = Math.Sqrt(square / If(population, count, count - 1.0))
                Dim fourth = name = "KURT"
                Dim moment = values.Sum(Function(value)
                                           Dim z = (value - mean) / deviation
                                           Dim z2 = z * z
                                           Return If(fourth, z2 * z2, z * z2)
                                       End Function)
                If population Then Return FiniteNumber(moment / count)
                If Not fourth Then Return FiniteNumber(count * moment / ((count - 1.0) * (count - 2.0)))
                Return FiniteNumber(count * (count + 1.0) * moment / ((count - 1.0) * (count - 2.0) * (count - 3.0)) -
                    3.0 * (count - 1.0) * (count - 1.0) / ((count - 2.0) * (count - 3.0)))
            End Function

            Private Shared Function TrimmedMean(values As List(Of Double), fraction As Object) As Object
                If TypeOf fraction Is FormulaError Then Return fraction
                Dim trim = DirectCast(fraction, Double)
                If trim < 0.0 OrElse trim >= 1.0 OrElse values.Count = 0 Then Return New FormulaError("#NUM!")
                Dim cut = CInt(Math.Floor(values.Count * trim * 0.5))
                Dim remaining = values.Count - 2 * cut
                Return FiniteNumber(values.Skip(cut).Take(remaining).Average())
            End Function

            Private Function EvaluateFisherTransform(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                Return EvaluateProbabilityArguments(arguments, formulaCell, name = "SQRTPI", -1,
                    Function(numbers, cumulative) FisherTransformValue(name, numbers(0)))
            End Function

            Private Shared Function FisherTransformValue(name As String, number As Double) As Object
                Select Case name
                    Case "FISHER"
                        If Math.Abs(number) >= 1.0 Then Return New FormulaError("#NUM!")
                        Return 0.5 * Math.Log((1.0 + number) / (1.0 - number))
                    Case "FISHERINV"
                        If Math.Abs(number) >= 20.0 Then Return CDbl(Math.Sign(number))
                        Dim exponential = Math.Exp(2.0 * number)
                        Return (exponential - 1.0) / (exponential + 1.0)
                    Case Else
                        If number < 0.0 Then Return New FormulaError("#NUM!")
                        Return FiniteNumber(Math.Sqrt(number * Math.PI))
                End Select
            End Function

        End Class

    End Class

End Namespace
