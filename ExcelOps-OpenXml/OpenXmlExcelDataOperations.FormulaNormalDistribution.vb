Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const InverseSquareRootTwoPi As Double = 0.3989422804014327
            Private Const SquareRootTwo As Double = 1.4142135623730951
            Private Const TwiceInverseSquareRootPi As Double = 1.1283791670955126

            Private Function EvaluateNormalFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                    formulaCell As ExcelCell) As Object
                Dim engineering = name.StartsWith("ERF", StringComparison.Ordinal)
                Dim cumulativeIndex = -1
                Dim minimum As Integer
                Select Case name
                    Case "NORM.S.DIST"
                        minimum = 2
                        cumulativeIndex = 1
                    Case "NORM.DIST", "NORMDIST", "LOGNORM.DIST"
                        minimum = 4
                        cumulativeIndex = 3
                    Case "LOGNORMDIST", "STANDARDIZE", "NORM.INV", "NORMINV", "LOGNORM.INV", "LOGINV", "CONFIDENCE", "CONFIDENCE.NORM"
                        minimum = 3
                    Case Else
                        minimum = 1
                End Select
                If arguments.Count < minimum OrElse arguments.Count > If(name = "ERF", 2, minimum) Then Return New FormulaError("#VALUE!")

                Dim evaluated = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                If engineering Then
                    If evaluated.Any(Function(value) TypeOf value Is FormulaArea AndAlso IsMultiValue(value)) Then Return New FormulaError("#VALUE!")
                    If evaluated.Any(Function(value) value Is Nothing) Then Return New FormulaError("#N/A")
                End If
                Dim inputs = evaluated.Select(Function(value) AsMatrix(value)).ToArray()
                Dim rows = inputs.Max(Function(input) input.Rows)
                Dim columns = inputs.Max(Function(input) input.Columns)
                Dim results = AllocateMatrixValues(rows, columns)
                Dim numbers(2) As Double
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim cumulative = True
                        Dim failure As FormulaError = Nothing
                        For index = 0 To inputs.Length - 1
                            Dim input = inputs(index).BroadcastAt(row, column)
                            Dim number As Object
                            If index = cumulativeIndex Then
                                number = LogicalOptionValue(input)
                            ElseIf engineering AndAlso TypeOf input Is Boolean Then
                                number = New FormulaError("#VALUE!")
                            Else
                                number = ReadFinancialNumber(input)
                            End If
                            If TypeOf number Is FormulaError Then
                                failure = DirectCast(number, FormulaError)
                                Exit For
                            End If
                            If index = cumulativeIndex Then
                                cumulative = DirectCast(number, Boolean)
                            Else
                                numbers(index) = DirectCast(number, Double)
                            End If
                        Next
                        results(row * columns + column) = If(failure, NormalFunctionValue(name, numbers, inputs.Length, cumulative))
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Shared Function NormalFunctionValue(name As String, values As Double(), count As Integer, cumulative As Boolean) As Object
                Dim x = values(0)
                Dim mean = values(1)
                Dim deviation = values(2)
                Select Case name
                    Case "ERF", "ERF.PRECISE"
                        Dim result = ErrorIntegral(x)
                        If count = 2 Then result = ErrorIntegral(mean) - result
                        Return NormalizeWorksheetUnderflow(result)
                    Case "ERFC", "ERFC.PRECISE"
                        Dim complement = PositiveErrorComplement(Math.Abs(x))
                        Return NormalizeWorksheetUnderflow(If(x < 0.0, 2.0 - complement, complement))
                    Case "PHI"
                        Return NormalizeWorksheetUnderflow(StandardNormalDensity(x))
                    Case "GAUSS"
                        Return NormalizeWorksheetUnderflow(ErrorIntegral(x / SquareRootTwo) * 0.5)
                    Case "NORMSDIST", "NORM.S.DIST"
                        Return NormalizeWorksheetUnderflow(If(cumulative, StandardNormalProbability(x), StandardNormalDensity(x)))
                    Case "NORM.S.INV", "NORMSINV"
                        If x <= 0.0 OrElse x >= 1.0 Then Return New FormulaError("#NUM!")
                        Return StandardNormalQuantile(x)
                    Case "CONFIDENCE", "CONFIDENCE.NORM"
                        Dim sampleSize = Math.Truncate(deviation)
                        Dim probability = 1.0 - x / 2.0
                        If x <= 0.0 OrElse probability >= 1.0 OrElse x >= 1.0 OrElse mean <= 0.0 OrElse sampleSize < 1.0 Then
                            Return New FormulaError("#NUM!")
                        End If
                        Return FiniteNumber(StandardNormalQuantile(probability) * (mean / Math.Sqrt(sampleSize)))
                End Select

                If deviation <= 0.0 Then Return New FormulaError("#NUM!")
                Select Case name
                    Case "NORM.INV", "NORMINV", "LOGNORM.INV", "LOGINV"
                        If x <= 0.0 OrElse x >= 1.0 Then Return New FormulaError("#NUM!")
                        Dim quantile = mean + deviation * StandardNormalQuantile(x)
                        If name = "LOGNORM.INV" OrElse name = "LOGINV" Then quantile = Math.Exp(quantile)
                        Return FiniteNumber(NormalizeWorksheetUnderflow(quantile))
                    Case "LOGNORM.DIST", "LOGNORMDIST"
                        If x <= 0.0 Then Return New FormulaError("#NUM!")
                        Dim z = (Math.Log(x) - mean) / deviation
                        Dim logResult = If(cumulative, StandardNormalProbability(z), StandardNormalDensity(z) / deviation / x)
                        Return FiniteNumber(NormalizeWorksheetUnderflow(logResult))
                    Case Else
                        Dim z = (x - mean) / deviation
                        If name = "STANDARDIZE" Then Return FiniteNumber(NormalizeWorksheetUnderflow(z))
                        Dim result = If(cumulative, StandardNormalProbability(z), StandardNormalDensity(z) / deviation)
                        Return FiniteNumber(NormalizeWorksheetUnderflow(result))
                End Select
            End Function

            Private Shared Function StandardNormalDensity(x As Double) As Double
                Return InverseSquareRootTwoPi * Math.Exp(-0.5 * x * x)
            End Function

            Private Shared Function StandardNormalProbability(x As Double) As Double
                Dim tail = 0.5 * PositiveErrorComplement(Math.Abs(x) / SquareRootTwo)
                Return If(x < 0.0, tail, 1.0 - tail)
            End Function

            Private Shared Function StandardNormalQuantile(probability As Double) As Double
                If probability = 0.5 Then Return 0.0
                Dim tail = Math.Min(probability, 1.0 - probability)
                Dim lower = 0.0
                Dim upper = 40.0
                ' Invert the smaller tail directly to avoid rounding it to one.
                For iteration = 1 To 64
                    Dim middle = (lower + upper) * 0.5
                    If PositiveErrorComplement(middle / SquareRootTwo) * 0.5 > tail Then
                        lower = middle
                    Else
                        upper = middle
                    End If
                Next
                Dim result = (lower + upper) * 0.5
                Return If(probability < 0.5, -result, result)
            End Function

            Private Shared Function ErrorIntegral(x As Double) As Double
                Dim magnitude = Math.Abs(x)
                If magnitude < 1.0E-8 Then Return TwiceInverseSquareRootPi * x
                Dim result = If(magnitude < Math.Sqrt(1.5), SmallErrorIntegral(magnitude), 1.0 - PositiveErrorComplement(magnitude))
                Return If(x < 0.0, -result, result)
            End Function

            Private Shared Function SmallErrorIntegral(x As Double) As Double
                Dim square = x * x
                Dim term = 1.0
                Dim sum = term
                For index = 1 To 64
                    term *= square / (index + 0.5)
                    sum += term
                    If term <= sum * 1.0E-16 Then Exit For
                Next
                Return TwiceInverseSquareRootPi * x * Math.Exp(-square) * sum
            End Function

            Private Shared Function PositiveErrorComplement(x As Double) As Double
                If x >= 28.0 Then Return 0.0
                Dim square = x * x
                If square < 1.5 Then Return 1.0 - SmallErrorIntegral(x)

                ' Use the upper incomplete gamma continued fraction so that small tails retain their precision.
                Dim denominator = square + 0.5
                Dim first = 1.0E+300
                Dim second = 1.0 / denominator
                Dim fraction = second
                For index = 1 To 128
                    Dim coefficient = -CDbl(index) * (index - 0.5)
                    denominator += 2.0
                    first = denominator + coefficient / first
                    second = 1.0 / (denominator + coefficient * second)
                    Dim change = first * second
                    fraction *= change
                    If Math.Abs(change - 1.0) <= 2.0E-16 Then Exit For
                Next
                Return TwiceInverseSquareRootPi * 0.5 * x * Math.Exp(-square) * fraction
            End Function

        End Class

    End Class

End Namespace
