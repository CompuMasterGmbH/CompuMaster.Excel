Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateBetaFamily(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                formulaCell As ExcelCell) As Object
                Dim minimum As Integer
                Dim cumulativeIndex = -1
                Dim defaults As Double() = Nothing
                Select Case name
                    Case "BETA.DIST"
                        minimum = 4
                        cumulativeIndex = 3
                        defaults = {0.0, 0.0, 0.0, 0.0, 0.0, 1.0}
                    Case "BETADIST", "BETA.INV", "BETAINV"
                        minimum = 3
                        defaults = {0.0, 0.0, 0.0, 0.0, 1.0}
                    Case "T.DIST"
                        minimum = 3
                        cumulativeIndex = 2
                    Case "F.DIST"
                        minimum = 4
                        cumulativeIndex = 3
                    Case "T.DIST.RT", "T.DIST.2T", "T.INV", "T.INV.2T", "TINV"
                        minimum = 2
                    Case Else
                        minimum = 3
                End Select
                If arguments.Count < minimum OrElse arguments.Count > If(defaults Is Nothing, minimum, minimum + 2) Then
                    Return New FormulaError("#VALUE!")
                End If
                Return EvaluateProbabilityArguments(arguments, formulaCell, False, cumulativeIndex,
                    Function(numbers, cumulative) BetaFamilyValue(name, numbers, cumulative), defaults)
            End Function

            Private Shared Function BetaFamilyValue(name As String, numbers As Double(), cumulative As Boolean) As Object
                Dim x = numbers(0)
                Dim first = numbers(1)
                Dim second = numbers(2)
                Select Case name
                    Case "BETA.DIST", "BETADIST", "BETA.INV", "BETAINV"
                        Dim offset = If(name = "BETA.DIST", 4, 3)
                        Return WorksheetBetaValue(name, x, first, second, numbers(offset), numbers(offset + 1), cumulative)
                    Case "T.DIST", "T.DIST.RT", "T.DIST.2T", "TDIST", "T.INV", "T.INV.2T", "TINV", "CONFIDENCE.T"
                        Return StudentValue(name, x, first, second, cumulative)
                    Case Else
                        Return FisherValue(name, x, Math.Truncate(first), Math.Truncate(second), cumulative)
                End Select
            End Function

            Private Shared Function WorksheetBetaValue(name As String, x As Double, alpha As Double, beta As Double,
                                                       minimum As Double, maximum As Double, cumulative As Boolean) As Object
                If alpha <= 0.0 OrElse beta <= 0.0 OrElse maximum <= minimum Then Return New FormulaError("#NUM!")
                Dim width = maximum - minimum
                If name = "BETA.INV" OrElse name = "BETAINV" Then
                    If x <= 0.0 OrElse x >= 1.0 Then Return New FormulaError("#NUM!")
                    Dim quantile = InverseBeta(x, alpha, beta, False)
                    Return FiniteNumber(minimum + width * quantile.Value)
                End If
                If x < minimum OrElse x > maximum Then Return New FormulaError("#NUM!")
                Dim fraction = (x - minimum) / width
                Dim complement = (maximum - x) / width
                If cumulative Then Return NormalizeWorksheetUnderflow(BetaProbability(fraction, complement, alpha, beta))
                If fraction = 0.0 OrElse complement = 0.0 Then
                    If If(fraction = 0.0, alpha, beta) <= 1.0 Then Return New FormulaError("#NUM!")
                    Return 0.0
                End If
                Dim logarithm = BetaLogPrefix(fraction, complement, alpha, beta) -
                    LogBetaSide(fraction, complement) - LogBetaSide(complement, fraction) - Math.Log(width)
                Return FiniteNumber(NormalizeWorksheetUnderflow(Math.Exp(logarithm)))
            End Function

            Private Shared Function StudentValue(name As String, x As Double, degrees As Double, extra As Double, cumulative As Boolean) As Object
                If name = "CONFIDENCE.T" Then
                    Dim size = Math.Truncate(extra)
                    If x <= 0.0 OrElse x >= 1.0 OrElse degrees <= 0.0 OrElse size < 1.0 Then Return New FormulaError("#NUM!")
                    If size = 1.0 Then Return New FormulaError("#DIV/0!")
                    Return FiniteNumber(StudentTwoTailQuantile(x, size - 1.0) * (degrees / Math.Sqrt(size)))
                End If
                degrees = Math.Truncate(degrees)
                Dim densityOnly = name = "T.DIST" AndAlso Not cumulative
                If degrees < 1.0 Then Return New FormulaError(If(densityOnly, "#DIV/0!", "#NUM!"))
                If degrees > 1.0E+10 AndAlso Not densityOnly Then Return New FormulaError("#NUM!")
                Select Case name
                    Case "T.INV", "T.INV.2T", "TINV"
                        Dim twoTailed = name <> "T.INV"
                        If x < 0.0 OrElse If(twoTailed, x >= 2.0, x >= 1.0) Then Return New FormulaError("#NUM!")
                        If x = 0.0 Then Return New FormulaError(If(degrees = 1.0, "#DIV/0!", "#NUM!"))
                        Dim tail = If(twoTailed, Math.Min(x, 2.0 - x), 2.0 * Math.Min(x, 1.0 - x))
                        Dim negative = If(twoTailed, x > 1.0, x < 0.5)
                        Dim critical = StudentTwoTailQuantile(tail, degrees)
                        Return FiniteNumber(If(negative, -critical, critical))
                    Case "T.DIST.2T", "TDIST"
                        If x < 0.0 Then Return New FormulaError("#NUM!")
                        If name = "T.DIST.2T" Then Return NormalizeWorksheetUnderflow(StudentTwoTail(x, degrees))
                        Dim tails = Math.Truncate(extra)
                        If tails <> 1.0 AndAlso tails <> 2.0 Then Return New FormulaError("#NUM!")
                        Return NormalizeWorksheetUnderflow(StudentTwoTail(x, degrees) * (tails * 0.5))
                    Case "T.DIST.RT"
                        Dim tail = StudentTwoTail(Math.Abs(x), degrees) * 0.5
                        Return NormalizeWorksheetUnderflow(If(x >= 0.0, tail, 1.0 - tail))
                    Case Else
                        If cumulative Then
                            Dim tail = StudentTwoTail(Math.Abs(x), degrees) * 0.5
                            Return NormalizeWorksheetUnderflow(If(x <= 0.0, tail, 1.0 - tail))
                        End If
                        Dim ratio = x * x / degrees
                        If Double.IsInfinity(ratio) Then Return 0.0
                        Dim logPower = If(ratio < 0.25, LogOnePlusMinusArgument(ratio) + ratio, Math.Log(1.0 + ratio))
                        Dim scale = LogBetaFunction(0.5, degrees * 0.5) + 0.5 * Math.Log(degrees)
                        Return FiniteNumber(NormalizeWorksheetUnderflow(Math.Exp(-0.5 * (degrees + 1.0) * logPower - scale)))
                End Select
            End Function

            Private Shared Function StudentTwoTail(x As Double, degrees As Double) As Double
                If x = 0.0 Then Return 1.0
                If degrees = 1.0 Then Return 2.0 / Math.PI * Math.Atan(1.0 / x)
                If degrees >= 1.0E+7 AndAlso x < 40.0 Then
                    Dim square = x * x
                    Dim density = StandardNormalDensity(x)
                    Dim firstCorrection = (square + 1.0) * x * density / (4.0 * degrees)
                    Dim secondCorrection = (((3.0 * square - 7.0) * square - 5.0) * square - 3.0) *
                        x * density / (96.0 * degrees * degrees)
                    Return 2.0 * (StandardNormalProbability(-x) + firstCorrection + secondCorrection)
                End If
                Dim ratio = x * x / degrees
                Dim complement = If(Double.IsInfinity(ratio), 1.0, ratio / (1.0 + ratio))
                Return BetaProbability(1.0 / (1.0 + ratio), complement, degrees * 0.5, 0.5)
            End Function

            Private Shared Function StudentTwoTailQuantile(tail As Double, degrees As Double) As Double
                If tail = 1.0 Then Return 0.0
                If degrees = 1.0 Then
                    If tail > 0.5 Then Return Math.Tan(Math.PI * 0.5 * (1.0 - tail))
                    Return 1.0 / Math.Tan(Math.PI * 0.5 * tail)
                End If
                Dim lower = 0.0
                Dim upper = 1.0
                While StudentTwoTail(upper, degrees) > tail
                    lower = upper
                    upper *= 2.0
                    If Double.IsInfinity(upper) Then Return upper
                End While
                For iteration = 1 To 80
                    Dim middle = lower + (upper - lower) * 0.5
                    If middle = lower OrElse middle = upper Then Return middle
                    If StudentTwoTail(middle, degrees) > tail Then
                        lower = middle
                    Else
                        upper = middle
                    End If
                Next
                Return lower + (upper - lower) * 0.5
            End Function

            Private Shared Function FisherValue(name As String, x As Double, first As Double, second As Double,
                                                cumulative As Boolean) As Object
                If x < 0.0 OrElse first < 1.0 OrElse second < 1.0 OrElse first > 1.0E+10 OrElse second > 1.0E+10 Then
                    Return New FormulaError("#NUM!")
                End If
                Dim alpha = first * 0.5
                Dim beta = second * 0.5
                Dim rightTail = name = "F.INV.RT" OrElse name = "FINV" OrElse name = "F.DIST.RT" OrElse name = "FDIST"
                If name = "F.INV" OrElse name = "F.INV.RT" OrElse name = "FINV" Then
                    If x > 1.0 OrElse x = If(rightTail, 0.0, 1.0) Then Return New FormulaError("#NUM!")
                    If x = If(rightTail, 1.0, 0.0) Then Return 0.0
                    If first = 1.0 AndAlso second = 1.0 Then
                        Dim tail = Math.Min(x, 1.0 - x)
                        Dim tangent = Math.Tan(Math.PI * 0.5 * tail)
                        Dim squared = tangent * tangent
                        Return FiniteNumber(NormalizeWorksheetUnderflow(If((x > 0.5) <> rightTail, 1.0 / squared, squared)))
                    End If
                    Dim quantile = InverseBeta(x, alpha, beta, rightTail)
                    Return FiniteNumber(NormalizeWorksheetUnderflow((quantile.Value / quantile.Complement) * (second / first)))
                End If
                If x = 0.0 AndAlso Not cumulative Then
                    If first < 2.0 Then Return New FormulaError("#NUM!")
                    Return If(first = 2.0, 1.0, 0.0)
                End If
                Dim ratio = x * (first / second)
                Dim complement = 1.0 / (1.0 + ratio)
                Dim fraction = If(Double.IsInfinity(ratio), 1.0, ratio / (1.0 + ratio))
                If rightTail Then Return NormalizeWorksheetUnderflow(BetaProbability(complement, fraction, beta, alpha))
                If cumulative Then Return NormalizeWorksheetUnderflow(BetaProbability(fraction, complement, alpha, beta))
                Dim result = Math.Exp(BetaLogPrefix(fraction, complement, alpha, beta) - Math.Log(x))
                Return FiniteNumber(NormalizeWorksheetUnderflow(result))
            End Function

            Private Shared Function LogBetaFunction(alpha As Double, beta As Double) As Double
                Dim small = Math.Min(alpha, beta)
                Dim large = Math.Max(alpha, beta)
                Dim total = small + large
                If large < 12.0 Then Return PositiveLogGamma(small) + PositiveLogGamma(large) - PositiveLogGamma(total)
                Dim relative = small / large
                Dim logarithm = LogOnePlusMinusArgument(relative) + relative
                Dim shift = small * Math.Log(large) + (total - 0.5) * logarithm - small
                Return PositiveLogGamma(small) - shift + StirlingCorrection(large) - StirlingCorrection(total)
            End Function

            Private Shared Function LogBetaSide(value As Double, complement As Double) As Double
                If value <= 0.5 Then Return Math.Log(value)
                Return LogOnePlusMinusArgument(-complement) - complement
            End Function

            Private Shared Function BetaLogPrefix(value As Double, complement As Double, alpha As Double, beta As Double) As Double
                If value <= 0.0 OrElse complement <= 0.0 Then Return Double.NegativeInfinity
                If alpha < 12.0 OrElse beta < 12.0 Then
                    Return alpha * LogBetaSide(value, complement) + beta * LogBetaSide(complement, value) - LogBetaFunction(alpha, beta)
                End If
                Dim total = alpha + beta
                Dim first = alpha / total
                Dim second = beta / total
                Dim offset = If(first <= second, value - first, second - complement)
                Dim firstChange = offset / first
                Dim secondChange = -offset / second
                Dim deviance As Double
                If Math.Abs(firstChange) < 0.25 AndAlso Math.Abs(secondChange) < 0.25 Then
                    deviance = alpha * LogOnePlusMinusArgument(firstChange) + beta * LogOnePlusMinusArgument(secondChange)
                Else
                    deviance = alpha * (LogBetaSide(value, complement) - LogBetaSide(first, second)) +
                        beta * (LogBetaSide(complement, value) - LogBetaSide(second, first))
                End If
                Return deviance + 0.5 * (Math.Log(alpha) + Math.Log(beta) - Math.Log(total)) - LogSquareRootTwoPi -
                    StirlingCorrection(alpha) - StirlingCorrection(beta) + StirlingCorrection(total)
            End Function

            Private Shared Function BetaProbability(value As Double, complement As Double, alpha As Double, beta As Double) As Double
                If value <= 0.0 Then Return 0.0
                If complement <= 0.0 Then Return 1.0
                If value = 0.5 AndAlso complement = 0.5 AndAlso alpha = beta Then Return 0.5
                If value < (alpha + 1.0) / (alpha + beta + 2.0) Then
                    Return Math.Max(0.0, Math.Min(1.0, BetaFraction(value, complement, alpha, beta)))
                End If
                Return Math.Max(0.0, Math.Min(1.0, 1.0 - BetaFraction(complement, value, beta, alpha)))
            End Function

            Private Shared Function BetaFraction(value As Double, complement As Double, alpha As Double, beta As Double) As Double
                ' Evaluate the continued fraction defined by https://dlmf.nist.gov/8.17.22.
                Dim numerator = 1.0
                Dim denominator = 0.0
                Dim product = 1.0
                Dim total = alpha + beta
                For index = 0 To 4096
                    Dim doubled = 2.0 * index
                    Dim coefficient = -((alpha + index) / (alpha + doubled)) *
                        ((total + index) / (alpha + doubled + 1.0)) * value
                    Dim change = EvaluateBetaFractionTerm(coefficient, numerator, denominator)
                    product *= change
                    Dim nextIndex = index + 1.0
                    coefficient = (nextIndex / (alpha + doubled + 1.0)) *
                        ((beta - nextIndex) / (alpha + doubled + 2.0)) * value
                    change = EvaluateBetaFractionTerm(coefficient, numerator, denominator)
                    product *= change
                    If Math.Abs(change - 1.0) <= 4.0E-16 Then Exit For
                Next
                Return Math.Exp(BetaLogPrefix(value, complement, alpha, beta) - Math.Log(alpha)) / product
            End Function

            Private Shared Function EvaluateBetaFractionTerm(coefficient As Double, ByRef numerator As Double, ByRef denominator As Double) As Double
                Const tiny As Double = 1.0E-300
                numerator = 1.0 + coefficient / numerator
                If Math.Abs(numerator) < tiny Then numerator = tiny
                denominator = 1.0 + coefficient * denominator
                If Math.Abs(denominator) < tiny Then denominator = tiny
                denominator = 1.0 / denominator
                Return numerator * denominator
            End Function

            Private Shared Function InverseBeta(probability As Double, alpha As Double, beta As Double,
                                                upperTail As Boolean) As (Value As Double, Complement As Double)
                Dim swap = If(upperTail, probability <= 0.5, probability > 0.5)
                Dim target = Math.Min(probability, 1.0 - probability)
                If swap Then
                    Dim quantile = InverseBetaLower(target, beta, alpha)
                    Return (1.0 - quantile, quantile)
                End If
                Dim lowerTail = InverseBetaLower(target, alpha, beta)
                Return (lowerTail, 1.0 - lowerTail)
            End Function

            Private Shared Function InverseBetaLower(probability As Double, alpha As Double, beta As Double) As Double
                If probability = 0.5 AndAlso alpha = beta Then Return 0.5
                Dim lower = 0.0
                Dim upper = 1.0
                If alpha <= 8.0 Then
                    Dim initial = Math.Exp((Math.Log(probability) + Math.Log(alpha) + LogBetaFunction(alpha, beta)) / alpha)
                    If initial = 0.0 Then Return 0.0
                    upper = Math.Min(1.0, 2.0 * initial)
                    While BetaProbability(upper, 1.0 - upper, alpha, beta) < probability
                        lower = upper
                        upper = Math.Min(1.0, upper * 2.0)
                    End While
                End If
                For iteration = 1 To 80
                    Dim middle = lower + (upper - lower) * 0.5
                    If middle = lower OrElse middle = upper Then Return middle
                    If BetaProbability(middle, 1.0 - middle, alpha, beta) < probability Then
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
