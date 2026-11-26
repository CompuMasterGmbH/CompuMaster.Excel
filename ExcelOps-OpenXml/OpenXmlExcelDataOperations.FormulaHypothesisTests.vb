Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateHypothesisTest(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                    formulaCell As ExcelCell) As Object
                Select Case name
                    Case "CHISQ.TEST", "CHITEST"
                        If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                        Dim observed = AsMatrix(EvaluateExpression(arguments(0), formulaCell))
                        Dim expected = AsMatrix(EvaluateExpression(arguments(1), formulaCell))
                        Return ChiSquareTest(observed, expected)
                    Case "F.TEST", "FTEST"
                        If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                        Dim first = AsMatrix(EvaluateExpression(arguments(0), formulaCell))
                        Dim second = AsMatrix(EvaluateExpression(arguments(1), formulaCell))
                        Return VarianceRatioTest(first, second)
                    Case "T.TEST", "TTEST"
                        If arguments.Count <> 4 Then Return New FormulaError("#VALUE!")
                        Return EvaluateStudentTest(arguments, formulaCell)
                    Case Else
                        If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                        Return EvaluateNormalTest(arguments, formulaCell)
                End Select
            End Function

            Private Shared Function ChiSquareTest(observed As FormulaMatrix, expected As FormulaMatrix) As Object
                If observed.Values.Length <> expected.Values.Length Then Return New FormulaError("#N/A")
                Dim total = 0.0
                Dim count = 0
                For index = 0 To observed.Values.Length - 1
                    Dim actual = observed.Values(index)
                    Dim estimate = expected.Values(index)
                    If TypeOf actual Is FormulaError Then Return actual
                    If TypeOf estimate Is FormulaError Then Return estimate
                    If observed.Values.Length = 1 Then Return New FormulaError("#N/A")
                    If Not TypeOf actual Is Double OrElse Not TypeOf estimate Is Double Then Continue For
                    Dim prediction = DirectCast(estimate, Double)
                    If prediction < 0.0 Then Return New FormulaError("#NUM!")
                    If prediction = 0.0 Then Return New FormulaError("#DIV/0!")
                    Dim difference = DirectCast(actual, Double) - prediction
                    total += difference * difference / prediction
                    count += 1
                Next
                If count = 0 Then Return New FormulaError("#DIV/0!")
                Dim degrees = If(observed.Rows = 1 OrElse observed.Columns = 1, observed.Values.Length - 1,
                    (observed.Rows - 1) * (observed.Columns - 1))
                Return FiniteNumber(NormalizeWorksheetUnderflow(GammaProbability(degrees * 0.5, total * 0.5, True)))
            End Function

            Private Structure NumericSample
                Friend Count As Integer
                Friend Mean As Double
                Friend SumSquares As Double
            End Structure

            Private Shared Function SampleStatistics(values As Object(), ByRef sample As NumericSample) As FormulaError
                Dim numbers As New List(Of Double)
                For Each value In values
                    If TypeOf value Is FormulaError Then Return DirectCast(value, FormulaError)
                    If TypeOf value Is Double Then numbers.Add(DirectCast(value, Double))
                Next
                sample = SampleStatistics(numbers)
                Return Nothing
            End Function

            Private Shared Function SampleStatistics(values As List(Of Double)) As NumericSample
                If values.Count = 0 Then Return New NumericSample()
                Dim mean = MeanFromFirst(values)
                Return New NumericSample With {
                    .Count = values.Count,
                    .Mean = mean,
                    .SumSquares = values.Sum(Function(value) (value - mean) * (value - mean))
                }
            End Function

            Private Shared Function VarianceRatioTest(first As FormulaMatrix, second As FormulaMatrix) As Object
                Dim one As NumericSample
                Dim two As NumericSample
                Dim failure = SampleStatistics(first.Values, one)
                If failure IsNot Nothing Then Return failure
                failure = SampleStatistics(second.Values, two)
                If failure IsNot Nothing Then Return failure
                If one.Count < 2 OrElse two.Count < 2 OrElse one.SumSquares = 0.0 OrElse two.SumSquares = 0.0 Then
                    Return New FormulaError("#DIV/0!")
                End If
                Dim ratio = one.SumSquares / two.SumSquares
                Dim complement = 1.0 / (1.0 + ratio)
                Dim fraction = If(Double.IsPositiveInfinity(ratio), 1.0, ratio / (1.0 + ratio))
                Dim alpha = 0.5 * (one.Count - 1)
                Dim beta = 0.5 * (two.Count - 1)
                Dim left = BetaProbability(fraction, complement, alpha, beta)
                Dim right = BetaProbability(complement, fraction, beta, alpha)
                Return FiniteNumber(NormalizeWorksheetUnderflow(2.0 * Math.Min(left, right)))
            End Function

            Private Function EvaluateStudentTest(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Dim first = AsMatrix(EvaluateExpression(arguments(0), formulaCell))
                Dim second = AsMatrix(EvaluateExpression(arguments(1), formulaCell))
                Dim parameters = {arguments(2), arguments(3)}
                Dim probabilities(2) As Object
                Return BroadcastReferenceArguments(parameters, formulaCell,
                    Function(values)
                        Dim tailsValue = ReadFinancialNumber(values(0))
                        If TypeOf tailsValue Is FormulaError Then Return tailsValue
                        Dim typeValue = ReadFinancialNumber(values(1))
                        If TypeOf typeValue Is FormulaError Then Return typeValue
                        Dim tails = Math.Truncate(DirectCast(tailsValue, Double))
                        Dim method = Math.Truncate(DirectCast(typeValue, Double))
                        If tails < 1.0 OrElse tails > 2.0 OrElse method < 1.0 OrElse method > 3.0 Then Return New FormulaError("#NUM!")
                        Dim index = CInt(method) - 1
                        If probabilities(index) Is Nothing Then probabilities(index) = StudentTestProbability(first.Values, second.Values, CInt(method))
                        Dim probability = probabilities(index)
                        If TypeOf probability Is FormulaError OrElse tails = 2.0 Then Return probability
                        Return NormalizeWorksheetUnderflow(DirectCast(probability, Double) * 0.5)
                    End Function)
            End Function

            Private Shared Function StudentTestProbability(first As Object(), second As Object(), method As Integer) As Object
                If method = 1 Then
                    If first.Length <> second.Length Then Return New FormulaError("#N/A")
                    Dim differences As New List(Of Double)
                    For index = 0 To first.Length - 1
                        If TypeOf first(index) Is FormulaError Then Return first(index)
                        If TypeOf second(index) Is FormulaError Then Return second(index)
                        If TypeOf first(index) Is Double AndAlso TypeOf second(index) Is Double Then
                            differences.Add(DirectCast(first(index), Double) - DirectCast(second(index), Double))
                        End If
                    Next
                    Dim paired = SampleStatistics(differences)
                    If paired.Count < 2 OrElse paired.SumSquares = 0.0 Then Return New FormulaError("#DIV/0!")
                    Dim t = Math.Abs(paired.Mean) * Math.Sqrt(paired.Count) / Math.Sqrt(paired.SumSquares / (paired.Count - 1))
                    Return FiniteNumber(NormalizeWorksheetUnderflow(StudentTwoTail(t, paired.Count - 1)))
                End If

                Dim one As NumericSample
                Dim two As NumericSample
                Dim failure = SampleStatistics(first, one)
                If failure IsNot Nothing Then Return failure
                failure = SampleStatistics(second, two)
                If failure IsNot Nothing Then Return failure
                If one.Count < 2 OrElse two.Count < 2 Then Return New FormulaError("#DIV/0!")
                Dim degrees As Double
                Dim standardError As Double
                If method = 2 Then
                    degrees = one.Count + two.Count - 2.0
                    standardError = Math.Sqrt((one.SumSquares + two.SumSquares) / degrees * (1.0 / one.Count + 1.0 / two.Count))
                Else
                    Dim left = one.SumSquares / (CDbl(one.Count) * (one.Count - 1))
                    Dim right = two.SumSquares / (CDbl(two.Count) * (two.Count - 1))
                    Dim total = left + right
                    If total = 0.0 Then Return New FormulaError("#DIV/0!")
                    Dim share = left / total
                    degrees = 1.0 / (share * share / (one.Count - 1) + (1.0 - share) * (1.0 - share) / (two.Count - 1))
                    standardError = Math.Sqrt(total)
                End If
                If standardError = 0.0 Then Return New FormulaError("#DIV/0!")
                Dim statistic = Math.Abs(one.Mean - two.Mean) / standardError
                Return FiniteNumber(NormalizeWorksheetUnderflow(StudentTwoTail(statistic, degrees)))
            End Function

            Private Function EvaluateNormalTest(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Dim input = EvaluateExpression(arguments(0), formulaCell)
                Dim data = If(TypeOf input Is FormulaArea OrElse TypeOf input Is FormulaMatrix,
                    AsMatrix(input), New FormulaMatrix(1, 1, {ReadFinancialNumber(input)}))
                Dim sample As NumericSample
                Dim failure = SampleStatistics(data.Values, sample)
                Dim parameters = arguments.Skip(1).ToArray()
                Return BroadcastReferenceArguments(parameters, formulaCell, Function(values) NormalTestValue(values, sample, failure))
            End Function

            Private Function NormalTestValue(values As Object(), sample As NumericSample, failure As FormulaError) As Object
                Dim populationMean = ReadFinancialNumber(values(0))
                If TypeOf populationMean Is FormulaError Then Return populationMean
                Dim deviation As Object = Nothing
                If values.Length = 2 Then
                    deviation = ReadFinancialNumber(values(1))
                    If TypeOf deviation Is FormulaError Then Return deviation
                End If
                If failure IsNot Nothing Then Return failure
                If sample.Count = 0 Then Return New FormulaError("#N/A")
                If sample.Count = 1 Then Return New FormulaError("#DIV/0!")
                Dim sigma = If(deviation Is Nothing, Math.Sqrt(sample.SumSquares / (sample.Count - 1)), DirectCast(deviation, Double))
                If deviation IsNot Nothing AndAlso sigma <= 0.0 Then Return New FormulaError("#NUM!")
                If sigma = 0.0 Then Return New FormulaError("#DIV/0!")
                Dim statistic = (DirectCast(populationMean, Double) - sample.Mean) * Math.Sqrt(sample.Count) / sigma
                Return FiniteNumber(NormalizeWorksheetUnderflow(StandardNormalProbability(statistic)))
            End Function

        End Class

    End Class

End Namespace
