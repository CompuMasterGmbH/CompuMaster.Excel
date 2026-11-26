Imports MathNet.Numerics

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            ' Modified Bessel polynomial approximations (Abramowitz and Stegun, sections 9.8.1–9.8.8).
            Private Shared ReadOnly I0Near As Double() = {1.0, 3.5156229, 3.0899424, 1.2067492, 0.2659732, 0.0360768, 0.0045813}
            Private Shared ReadOnly I0Far As Double() = {0.39894228, 0.01328592, 0.00225319, -0.00157565, 0.00916281, -0.02057706, 0.02635537, -0.01647633, 0.00392377}
            Private Shared ReadOnly I1Near As Double() = {0.5, 0.87890594, 0.51498869, 0.15084934, 0.02658733, 0.00301532, 0.00032411}
            Private Shared ReadOnly I1Far As Double() = {0.39894228, -0.03988024, -0.00362018, 0.00163801, -0.01031555, 0.02282967, -0.02895312, 0.01787654, -0.00420059}
            Private Shared ReadOnly K0Near As Double() = {-0.57721566, 0.4227842, 0.23069756, 0.0348859, 0.00262698, 0.0001075, 0.0000074}
            Private Shared ReadOnly K0Far As Double() = {1.25331414, -0.07832358, 0.02189568, -0.01062446, 0.00587872, -0.0025154, 0.00053208}
            Private Shared ReadOnly K1Near As Double() = {1.0, 0.15443144, -0.67278579, -0.18156897, -0.01919402, -0.00110404, -0.00004686}
            Private Shared ReadOnly K1Far As Double() = {1.25331414, 0.23498619, -0.0365562, 0.01504268, -0.00780353, 0.00325614, -0.00068245}

            Private Shared Function BesselPolynomial(coefficients As Double(), x As Double) As Double
                Dim result = coefficients(coefficients.Length - 1)
                For index = coefficients.Length - 2 To 0 Step -1
                    result = result * x + coefficients(index)
                Next
                Return result
            End Function

            Private Shared Function ModifiedBesselI(rank As Double, x As Double) As Double
                If rank > 1.0 Then
                    If rank > 4096.0 OrElse x > 700.0 OrElse x < 1.0E-150 Then Return SpecialFunctions.BesselI(rank, x)
                    Dim stepSize = 2.0 / x
                    Dim start = 2 * (CInt(rank) + CInt(Math.Floor(Math.Sqrt(40.0 * rank))))
                    Dim following = 0.0
                    Dim current = 1.0
                    Dim selected = 0.0
                    For order = start To 1 Step -1
                        Dim previous = following + order * stepSize * current
                        following = current
                        current = previous
                        If Math.Abs(current) > 1.0E10 Then
                            following *= 1.0E-10
                            current *= 1.0E-10
                            selected *= 1.0E-10
                        End If
                        If order = rank Then selected = following
                    Next
                    Return selected / current * ModifiedBesselI(0.0, x)
                End If
                Dim t = x / 3.75
                If x < 3.75 Then
                    If rank = 0.0 Then Return BesselPolynomial(I0Near, t * t)
                    Return x * BesselPolynomial(I1Near, t * t)
                End If
                Dim scaled = BesselPolynomial(If(rank = 0.0, I0Far, I1Far), 1.0 / t)
                If x < 709.0 Then Return Math.Exp(x) / Math.Sqrt(x) * scaled
                Return Math.Exp(x - 0.5 * Math.Log(x) + Math.Log(scaled))
            End Function

            Private Shared Function ModifiedBesselK(rank As Double, x As Double) As Double
                If rank > 4096.0 OrElse x > 680.0 OrElse x < 1.0E-150 Then Return SpecialFunctions.BesselK(rank, x)
                Dim k0, k1 As Double
                If x <= 2.0 Then
                    Dim t = x * x / 4.0
                    Dim log = Math.Log(x / 2.0)
                    k0 = -log * ModifiedBesselI(0.0, x) + BesselPolynomial(K0Near, t)
                    k1 = log * ModifiedBesselI(1.0, x) + BesselPolynomial(K1Near, t) / x
                Else
                    Dim t = 2.0 / x
                    Dim scale = Math.Exp(-x) / Math.Sqrt(x)
                    k0 = scale * BesselPolynomial(K0Far, t)
                    k1 = scale * BesselPolynomial(K1Far, t)
                End If
                If rank = 0.0 Then Return k0
                For order = 1 To CInt(rank) - 1
                    Dim nextTerm = k0 + order * (2.0 / x) * k1
                    If Double.IsInfinity(nextTerm) Then Return nextTerm
                    k0 = k1
                    k1 = nextTerm
                Next
                Return k1
            End Function

        End Class

    End Class

End Namespace
