Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const CircularArgumentLimit As Double = 134217728.0
            ' These parts represent Excel's 64-bit significand for pi/2 without rounding it to binary64.
            Private Const HalfPiHigh As Double = 1.570796251296997
            Private Const HalfPiMiddle As Double = 7.549789415861596E-8
            Private Const HalfPiLow As Double = 5.3903279409461824E-15
            Private Const SmallestNormalDouble As Double = 2.2250738585072014E-308
            Private Shared ReadOnly ExponentialLimit As Double = Math.Log(Double.MaxValue)

            Private Function EvaluateTranscendental(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                    formulaCell As ExcelCell) As Object
                If name = "PI" Then
                    If arguments.Count <> 0 Then Return New FormulaError("#VALUE!")
                    Return Math.PI
                End If
                Dim minimum = If(name = "ATAN2", 2, 1)
                Dim maximum = If(name = "LOG", 2, minimum)
                If arguments.Count < minimum OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")

                Dim first = AsMatrix(EvaluateExpression(arguments(0), formulaCell))
                Dim second = AsMatrix(If(arguments.Count = 2, EvaluateExpression(arguments(1), formulaCell), 10.0))
                Dim rows = Math.Max(first.Rows, second.Rows)
                Dim columns = Math.Max(first.Columns, second.Columns)
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim x = NumberValue(first.BroadcastAt(row, column))
                        Dim y = NumberValue(second.BroadcastAt(row, column))
                        Dim result As Object
                        If TypeOf x Is FormulaError Then
                            result = x
                        ElseIf TypeOf y Is FormulaError Then
                            result = y
                        Else
                            result = TranscendentalValue(name, DirectCast(x, Double), DirectCast(y, Double))
                        End If
                        values(row * columns + column) = result
                    Next
                Next
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Shared Function TranscendentalValue(name As String, x As Double, y As Double) As Object
                Select Case name
                    Case "EXP" : Return FiniteNumber(Math.Exp(x))
                    Case "LN" : Return FiniteNumber(Math.Log(x))
                    Case "LOG10" : Return FiniteNumber(Math.Log10(x))
                    Case "LOG"
                        If x <= 0.0 OrElse y <= 0.0 Then Return New FormulaError("#NUM!")
                        If y = 1.0 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber(Math.Log(x) / Math.Log(y))
                    Case "DEGREES" : Return FiniteNumber(x * (180.0 / Math.PI))
                    Case "RADIANS" : Return FiniteNumber(x * (Math.PI / 180.0))
                    Case "SIN", "COS", "TAN", "COT", "SEC", "CSC"
                        Return CircularValue(name, x)
                    Case "ASIN" : Return FiniteNumber(Math.Asin(x))
                    Case "ACOS" : Return FiniteNumber(Math.Acos(x))
                    Case "ATAN" : Return Math.Atan(x)
                    Case "ATAN2"
                        If x = 0.0 AndAlso y = 0.0 Then Return New FormulaError("#DIV/0!")
                        Return Math.Atan2(If(y = 0.0, 0.0, y), x)
                    Case "ACOT" : Return Math.Atan2(1.0, x)
                    Case "ASINH"
                        Dim magnitude = Math.Abs(x)
                        If magnitude < 1.0E-8 Then Return x
                        Dim result = Math.Log(magnitude + Math.Sqrt(magnitude * magnitude + 1.0))
                        Return FiniteNumber(If(x < 0.0, -result, result))
                    Case "ACOSH"
                        Return FiniteNumber(Math.Log(x + Math.Sqrt((x - 1.0) * (x + 1.0))))
                    Case "ATANH"
                        Return InverseHyperbolicTangent(x)
                    Case "ACOTH"
                        If Math.Abs(x) <= 1.0 Then Return New FormulaError("#NUM!")
                        Return InverseHyperbolicTangent(1.0 / x)
                    Case Else
                        Return HyperbolicValue(name, x)
                End Select
            End Function

            Private Shared Function TryCircularParts(angle As Double, ByRef sine As Double, ByRef cosine As Double) As Boolean
                If Double.IsNaN(angle) OrElse Math.Abs(angle) >= CircularArgumentLimit Then Return False

                ' Reduce small enough to avoid platform rounding differences near poles and large multiples of pi.
                Dim quadrant = CInt(Math.Round(angle / (Math.PI / 2.0)))
                Dim reduced = ((angle - quadrant * HalfPiHigh) - quadrant * HalfPiMiddle) - quadrant * HalfPiLow
                Dim reducedSine = Math.Sin(reduced)
                Dim reducedCosine = Math.Cos(reduced)
                If Math.Abs(reduced) = Math.PI / 4.0 Then
                    ' Excel uses adjacent binary64 values for sine and cosine at this stored angle.
                    reducedSine = If(reduced < 0.0, -0.70710678118654746, 0.70710678118654746)
                    reducedCosine = 0.70710678118654757
                End If
                Select Case quadrant And 3
                    Case 0 : sine = reducedSine : cosine = reducedCosine
                    Case 1 : sine = reducedCosine : cosine = -reducedSine
                    Case 2 : sine = -reducedSine : cosine = -reducedCosine
                    Case Else : sine = -reducedCosine : cosine = reducedSine
                End Select
                Return True
            End Function

            Private Shared Function CircularValue(name As String, angle As Double) As Object
                Dim sine As Double, cosine As Double
                If Not TryCircularParts(angle, sine, cosine) Then Return New FormulaError("#NUM!")
                Select Case name
                    Case "SIN" : Return sine
                    Case "COS" : Return cosine
                    Case "TAN" : Return TranscendentalQuotient(sine, cosine)
                    Case "COT" : Return TranscendentalQuotient(cosine, sine)
                    Case "SEC" : Return TranscendentalQuotient(1.0, cosine)
                    Case Else : Return TranscendentalQuotient(1.0, sine)
                End Select
            End Function

            Private Shared Function InverseHyperbolicTangent(value As Double) As Object
                If Math.Abs(value) < 1.0E-8 Then Return value
                Return FiniteNumber(0.5 * Math.Log((1.0 + value) / (1.0 - value)))
            End Function

            Private Shared Function HyperbolicValue(name As String, value As Double) As Object
                Select Case name
                    Case "TANH" : Return Math.Tanh(value)
                    Case "COTH" : Return TranscendentalQuotient(1.0, Math.Tanh(value))
                    Case "SINH", "COSH"
                        If Math.Abs(value) > ExponentialLimit Then Return New FormulaError("#NUM!")
                        Return FiniteNumber(If(name = "SINH", Math.Sinh(value), Math.Cosh(value)))
                    Case Else
                        If name = "CSCH" AndAlso value = 0.0 Then Return New FormulaError("#DIV/0!")
                        Dim denominator = If(name = "CSCH", Math.Sinh(value), Math.Cosh(value))
                        Dim result = 1.0 / denominator
                        If Math.Abs(result) < SmallestNormalDouble Then Return 0.0
                        Return FiniteNumber(result)
                End Select
            End Function

            Private Shared Function TranscendentalQuotient(numerator As Double, denominator As Double) As Object
                If denominator = 0.0 Then Return New FormulaError("#DIV/0!")
                Return FiniteNumber(numerator / denominator)
            End Function

        End Class

    End Class

End Namespace
