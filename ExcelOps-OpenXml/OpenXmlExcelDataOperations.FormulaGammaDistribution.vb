Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const LogSquareRootTwoPi As Double = 0.9189385332046727

            Private Function EvaluateGammaFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                   formulaCell As ExcelCell) As Object
                Dim minimum As Integer
                Dim cumulativeIndex = -1
                Select Case name
                    Case "GAMMA", "GAMMALN", "GAMMALN.PRECISE"
                        minimum = 1
                    Case "CHISQ.DIST.RT", "CHIDIST", "CHISQ.INV", "CHISQ.INV.RT", "CHIINV"
                        minimum = 2
                    Case "GAMMA.INV", "GAMMAINV"
                        minimum = 3
                    Case "CHISQ.DIST", "EXPON.DIST", "EXPONDIST"
                        minimum = 3
                        cumulativeIndex = 2
                    Case Else
                        minimum = 4
                        cumulativeIndex = 3
                End Select
                If arguments.Count <> minimum Then Return New FormulaError("#VALUE!")
                Return EvaluateProbabilityArguments(arguments, formulaCell, False, cumulativeIndex,
                    Function(numbers, cumulative) GammaFunctionValue(name, numbers, cumulative))
            End Function

            Private Shared Function GammaFunctionValue(name As String, numbers As Double(), cumulative As Boolean) As Object
                Dim x = numbers(0)
                Dim shape = numbers(1)
                Dim scale = numbers(2)
                Select Case name
                    Case "GAMMALN", "GAMMALN.PRECISE"
                        If x <= 0.0 Then Return New FormulaError("#NUM!")
                        Return FiniteNumber(PositiveLogGamma(x))
                    Case "GAMMA"
                        Return WorksheetGamma(x)
                    Case "EXPON.DIST", "EXPONDIST"
                        If x < 0.0 OrElse shape <= 0.0 Then Return New FormulaError("#NUM!")
                        Dim power = x * shape
                        Dim result = If(cumulative, ExponentialProbability(power), shape * Math.Exp(-power))
                        Return FiniteNumber(NormalizeWorksheetUnderflow(result))
                    Case "WEIBULL.DIST", "WEIBULL"
                        Return WeibullValue(x, shape, scale, cumulative)
                End Select

                Dim upperTail = name = "CHISQ.DIST.RT" OrElse name = "CHIDIST" OrElse name = "CHISQ.INV.RT" OrElse name = "CHIINV"
                Dim inverse = name = "GAMMA.INV" OrElse name = "GAMMAINV" OrElse name = "CHISQ.INV" OrElse
                    name = "CHISQ.INV.RT" OrElse name = "CHIINV"
                Dim chiSquare = name.StartsWith("CHI", StringComparison.Ordinal)
                If chiSquare Then
                    Dim degrees = Math.Truncate(shape)
                    If degrees < 1.0 OrElse (degrees > 1.0E+10 AndAlso Not (upperTail AndAlso inverse)) Then
                        Return New FormulaError("#NUM!")
                    End If
                    shape = degrees * 0.5
                    scale = 2.0
                End If
                If x < 0.0 OrElse shape <= 0.0 OrElse scale <= 0.0 Then Return New FormulaError("#NUM!")
                If inverse Then
                    If x > 1.0 OrElse (x = 1.0 AndAlso Not upperTail) OrElse (x = 0.0 AndAlso upperTail) Then Return New FormulaError("#NUM!")
                    If x = If(upperTail, 1.0, 0.0) Then Return 0.0
                    Dim result = InverseGammaProbability(x, shape, upperTail) * scale
                    Return FiniteNumber(NormalizeWorksheetUnderflow(result))
                End If
                If cumulative Then
                    Return FiniteNumber(NormalizeWorksheetUnderflow(GammaProbability(shape, x / scale, upperTail)))
                End If
                If x = 0.0 Then
                    If chiSquare AndAlso shape = 1.0 Then Return 0.5
                    If shape <= 1.0 Then Return New FormulaError("#NUM!")
                    Return 0.0
                End If
                Dim normalized = x / scale
                If Double.IsPositiveInfinity(normalized) Then Return 0.0
                Dim density = Math.Exp(GammaLogPrefix(shape, normalized) - Math.Log(x))
                Return FiniteNumber(NormalizeWorksheetUnderflow(density))
            End Function

            Private Shared Function WeibullValue(x As Double, shape As Double, scale As Double, cumulative As Boolean) As Object
                If x < 0.0 OrElse shape <= 0.0 OrElse scale <= 0.0 Then Return New FormulaError("#NUM!")
                Dim power = Math.Pow(x / scale, shape)
                If cumulative Then Return NormalizeWorksheetUnderflow(ExponentialProbability(power))
                Dim denominator = Math.Pow(scale, shape)
                If Double.IsInfinity(denominator) Then Return New FormulaError("#NUM!")
                Dim numerator = If(x = 0.0, 0.0, shape * Math.Pow(x, shape - 1.0))
                If denominator = 0.0 Then Return New FormulaError(If(numerator = 0.0, "#DIV/0!", "#NUM!"))
                If x = 0.0 Then Return 0.0
                Return FiniteNumber(NormalizeWorksheetUnderflow(numerator / denominator * Math.Exp(-power)))
            End Function

            Private Shared Function ExponentialProbability(power As Double) As Double
                If power > 0.01 Then Return 1.0 - Math.Exp(-power)
                ' Summing the first terms of -expm1(-power) avoids cancellation near zero.
                Return power * (1.0 - power * (0.5 - power * (1.0 / 6.0 - power * (1.0 / 24.0 -
                    power * (1.0 / 120.0 - power * (1.0 / 720.0 - power / 5040.0))))))
            End Function

            Private Shared Function WorksheetGamma(x As Double) As Object
                Dim integerPart = Math.Truncate(x)
                If x = integerPart AndAlso (x <= 0.0 OrElse x > 171.0) Then Return New FormulaError("#NUM!")
                If x = integerPart Then
                    Dim factorial = 1.0
                    For factor As Integer = 2 To CInt(integerPart) - 1
                        factorial *= factor
                    Next
                    Return factorial
                End If
                If x > 0.0 Then
                    If x < 1.0E-4 Then Return FiniteNumber(Math.Exp(LogGammaOnePlus(x)) / x)
                    Return FiniteNumber(Math.Exp(PositiveLogGamma(x)))
                End If

                ' Reduce the sine argument before reflecting negative values across zero.
                Dim fraction = x - integerPart
                Dim sine = Math.Sin(Math.PI * fraction)
                If integerPart Mod 2.0 <> 0.0 Then sine = -sine
                Dim reflected = Math.Exp(Math.Log(Math.PI) - Math.Log(Math.Abs(sine)) - PositiveLogGamma(1.0 - x))
                Return FiniteNumber(NormalizeWorksheetUnderflow(If(sine < 0.0, -reflected, reflected)))
            End Function

            Private Shared Function PositiveLogGamma(x As Double) As Double
                If x = 1.0 OrElse x = 2.0 Then Return 0.0
                If x < 1.0E-4 Then Return LogGammaOnePlus(x) - Math.Log(x)
                Dim shift = 0.0
                While x < 12.0
                    shift -= Math.Log(x)
                    x += 1.0
                End While
                Return shift + ((x - 0.5) * Math.Log(x) - x + LogSquareRootTwoPi + StirlingCorrection(x))
            End Function

            Private Shared Function LogGammaOnePlus(x As Double) As Double
                ' The first three Taylor terms resolve log Gamma(1 + x) for x below 0.0001.
                Return x * (-0.5772156649015329 + x * (0.8224670334241132 - 0.4006856343865314 * x))
            End Function

            Private Shared Function StirlingCorrection(x As Double) As Double
                Dim inverse = 1.0 / x
                Dim square = inverse * inverse
                Return inverse * (1.0 / 12.0 + square * (-1.0 / 360.0 + square * (1.0 / 1260.0 +
                    square * (-1.0 / 1680.0 + square * (1.0 / 1188.0 + square * (-691.0 / 360360.0))))))
            End Function

            Private Shared Function GammaLogPrefix(shape As Double, x As Double) As Double
                If x <= 0.0 Then Return Double.NegativeInfinity
                If shape < 8.0 Then Return shape * Math.Log(x) - x - PositiveLogGamma(shape)
                Dim relative = (x - shape) / shape
                Dim deviance = If(Math.Abs(relative) < 0.25, shape * LogOnePlusMinusArgument(relative),
                                  shape * (Math.Log(x) - Math.Log(shape)) + shape - x)
                Return deviance + 0.5 * Math.Log(shape) - LogSquareRootTwoPi - StirlingCorrection(shape)
            End Function

            Private Shared Function LogOnePlusMinusArgument(relative As Double) As Double
                If Math.Abs(relative) >= 0.25 Then Return Math.Log(1.0 + relative) - relative
                Dim power = -relative * relative
                Dim sum = power * 0.5
                For index = 3 To 32
                    power *= -relative
                    Dim term = power / index
                    sum += term
                    If Math.Abs(term) <= Math.Abs(sum) * 1.0E-16 Then Exit For
                Next
                Return sum
            End Function

            Private Shared Function GammaProbability(shape As Double, x As Double, upperTail As Boolean) As Double
                If x <= 0.0 Then Return If(upperTail, 1.0, 0.0)
                If Double.IsPositiveInfinity(x) Then Return If(upperTail, 0.0, 1.0)
                Dim relative = (x - shape) / shape
                If shape >= 1.0E+6 AndAlso Math.Abs(relative) < 0.25 Then
                    ' Use the leading correction of the uniform approximation in https://dlmf.nist.gov/8.12.
                    Dim eta = Math.Sign(relative) * Math.Sqrt(-2.0 * LogOnePlusMinusArgument(relative))
                    Dim normal = eta * Math.Sqrt(shape)
                    Dim coefficient = If(Math.Abs(relative) < 0.001,
                        -1.0 / 3.0 + relative * (1.0 / 12.0 - relative * 23.0 / 540.0),
                        1.0 / relative - 1.0 / eta)
                    Dim correction = StandardNormalDensity(normal) / Math.Sqrt(shape) * coefficient
                    If upperTail Then Return Math.Max(0.0, Math.Min(1.0, StandardNormalProbability(-normal) + correction))
                    Return Math.Max(0.0, Math.Min(1.0, StandardNormalProbability(normal) - correction))
                End If

                Dim prefix = GammaLogPrefix(shape, x)
                If x < shape + 1.0 Then
                    Dim term = 1.0
                    Dim sum = term
                    For index = 1 To 32768
                        term *= x / (shape + index)
                        sum += term
                        If term <= sum * 1.0E-16 Then Exit For
                    Next
                    Dim lower = Math.Min(1.0, Math.Exp(prefix - Math.Log(shape)) * sum)
                    Return If(upperTail, 1.0 - lower, lower)
                End If

                ' Evaluate the ratio defining the upper incomplete gamma without subtracting a small tail from one.
                Const tiny As Double = 1.0E-300
                Dim denominator = x - shape + 1.0
                Dim first = 1.0 / tiny
                Dim second = 1.0 / denominator
                Dim ratio = second
                For index = 1 To 32768
                    Dim coefficient = CDbl(index) * (shape - index)
                    denominator += 2.0
                    first = denominator + coefficient / first
                    If Math.Abs(first) < tiny Then first = tiny
                    second = denominator + coefficient * second
                    If Math.Abs(second) < tiny Then second = tiny
                    second = 1.0 / second
                    Dim change = first * second
                    ratio *= change
                    If Math.Abs(change - 1.0) <= 2.0E-16 Then Exit For
                Next
                Dim upper = Math.Min(1.0, Math.Exp(prefix) * ratio)
                Return If(upperTail, upper, 1.0 - upper)
            End Function

            Private Shared Function InverseGammaProbability(probability As Double, shape As Double, upperTail As Boolean) As Double
                Dim tail = If(probability <= 0.5, probability, 1.0 - probability)
                Dim invertUpper = If(probability <= 0.5, upperTail, Not upperTail)
                Dim lower = 0.0
                Dim upper = Math.Max(1.0, shape)
                If Not invertUpper AndAlso shape <= 8.0 Then
                    ' Integrating the leading power term gives a scale for roots extremely close to zero.
                    Dim initial = Math.Exp((Math.Log(tail) + PositiveLogGamma(shape + 1.0)) / shape)
                    If initial = 0.0 Then Return 0.0
                    upper = 2.0 * initial
                End If
                For iteration = 1 To 1024
                    Dim value = GammaProbability(shape, upper, invertUpper)
                    If If(invertUpper, value <= tail, value >= tail) Then Exit For
                    lower = upper
                    upper *= 2.0
                    If Double.IsInfinity(upper) Then Return upper
                Next
                For iteration = 1 To 80
                    Dim middle = lower + (upper - lower) * 0.5
                    If middle = lower OrElse middle = upper Then Return middle
                    Dim value = GammaProbability(shape, middle, invertUpper)
                    If If(invertUpper, value > tail, value < tail) Then
                        lower = middle
                    Else
                        upper = middle
                    End If
                Next
                Return lower + (upper - lower) * 0.5
            End Function

        End Class

    End Class

End Namespace
