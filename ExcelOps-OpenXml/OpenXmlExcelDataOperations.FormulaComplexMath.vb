Imports System.Numerics

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function CalculateComplexMath(name As String, operand As ComplexOperand, second As Object) As Object
                Dim value = operand.Value
                Select Case name
                    Case "IMEXP"
                        Dim magnitude = Math.Exp(value.Real)
                        Dim saturated = Double.IsInfinity(magnitude)
                        Return FormatPolarComplex(If(saturated, Double.MaxValue, magnitude), value.Imaginary, operand.Unit, saturated)
                    Case "IMLN", "IMLOG10", "IMLOG2"
                        Dim magnitude = Complex.Abs(value)
                        If magnitude = 0.0 Then Return New FormulaError("#NUM!")
                        Dim scale = If(name = "IMLN", 1.0, Math.Log(If(name = "IMLOG2", 2.0, 10.0)))
                        Return FormatComplexValue(New Complex(Math.Log(magnitude) / scale, value.Phase / scale), operand.Unit)
                    Case "IMSQRT"
                        Return FormatPolarComplex(Math.Sqrt(Complex.Abs(value)), value.Phase / 2.0, operand.Unit)
                    Case "IMPOWER"
                        If TypeOf second Is FormulaError Then Return second
                        Dim rawExponent = ConvertWorksheetText("VALUE", second)
                        If TypeOf rawExponent Is FormulaError Then Return rawExponent
                        Dim exponent = DirectCast(rawExponent, Double)
                        If value = Complex.Zero AndAlso exponent <= 0.0 Then Return New FormulaError("#NUM!")
                        If exponent = 2.0 AndAlso value.Real <> 0.0 AndAlso value.Imaginary <> 0.0 Then
                            Return FormatComplexValue(value * value, operand.Unit)
                        End If
                        Return FormatPolarComplex(Math.Pow(Complex.Abs(value), exponent), value.Phase * exponent, operand.Unit)
                    Case Else
                        Return EvaluateComplexTrigonometric(name, operand)
                End Select
            End Function

            Private Function FormatPolarComplex(radius As Double, angle As Double, unit As Char, Optional saturated As Boolean = False) As Object
                Dim sine As Double, cosine As Double
                If Not TryCircularParts(angle, sine, cosine) Then Return New FormulaError("#NUM!")
                If saturated AndAlso (Math.Abs(sine) = 1.0 OrElse Math.Abs(cosine) = 1.0) Then Return New FormulaError("#NUM!")
                Return FormatComplexValue(New Complex(radius * cosine, radius * sine), unit)
            End Function

            Private Function EvaluateComplexTrigonometric(name As String, operand As ComplexOperand) As Object
                Dim value = operand.Value
                Dim hyperbolic = name = "IMSINH" OrElse name = "IMCOSH" OrElse name = "IMSECH" OrElse name = "IMCSCH"
                Dim circularAngle = If(hyperbolic, value.Imaginary, value.Real)
                Dim hyperbolicAngle = If(hyperbolic, value.Real, value.Imaginary)
                Dim sine As Double, cosine As Double
                If Not TryCircularParts(circularAngle, sine, cosine) Then Return New FormulaError("#NUM!")

                If (name = "IMTAN" OrElse name = "IMCOT") AndAlso hyperbolicAngle <> 0.0 Then
                    Dim tangent = name = "IMTAN"
                    Dim sign = If(tangent, 1.0, -1.0)
                    Dim doubledImaginary = 2.0 * hyperbolicAngle
                    If Math.Abs(doubledImaginary) > ExponentialLimit Then
                        Return FormatComplexValue(New Complex(0.0, sign * Math.Sign(hyperbolicAngle)), operand.Unit)
                    End If
                    Dim sinDouble As Double, cosDouble As Double
                    If Not TryCircularParts(2.0 * circularAngle, sinDouble, cosDouble) Then Return New FormulaError("#NUM!")
                    Dim denominator = Math.Cosh(doubledImaginary) + sign * cosDouble
                    Return FormatComplexValue(New Complex(sinDouble / denominator, sign * Math.Sinh(doubledImaginary) / denominator), operand.Unit)
                End If
                If Math.Abs(hyperbolicAngle) > ExponentialLimit Then
                    If name = "IMSEC" OrElse name = "IMCSC" OrElse name = "IMSECH" OrElse name = "IMCSCH" Then Return "0"
                    Return New FormulaError("#NUM!")
                End If

                Dim sinh = Math.Sinh(hyperbolicAngle)
                Dim cosh = Math.Cosh(hyperbolicAngle)
                Dim first As Complex, second As Complex
                If hyperbolic Then
                    first = New Complex(sinh * cosine, cosh * sine)
                    second = New Complex(cosh * cosine, sinh * sine)
                Else
                    first = New Complex(sine * cosh, cosine * sinh)
                    second = New Complex(cosine * cosh, -sine * sinh)
                End If
                Dim result As Complex
                Select Case name
                    Case "IMSIN", "IMSINH" : result = first
                    Case "IMCOS", "IMCOSH" : result = second
                    Case "IMTAN" : result = first / second
                    Case "IMCOT" : result = second / first
                    Case "IMCSC", "IMCSCH" : result = Complex.One / first
                    Case Else : result = Complex.One / second
                End Select
                Return FormatComplexValue(result, operand.Unit)
            End Function

        End Class

    End Class

End Namespace
