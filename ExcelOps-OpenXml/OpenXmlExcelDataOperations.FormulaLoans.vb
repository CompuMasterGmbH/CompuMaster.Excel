Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateLoanFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                 formulaCell As ExcelCell) As Object
                Dim interestOnly = name = "ISPMT"
                If arguments.Count < If(interestOnly, 4, 3) OrElse arguments.Count > If(interestOnly, 4, 5) Then
                    Return New FormulaError("#VALUE!")
                End If

                Dim inputs = arguments.Select(Function(argument) AsMatrix(EvaluateExpression(argument, formulaCell))).ToArray()
                Dim rows = inputs.Max(Function(input) input.Rows)
                Dim columns = inputs.Max(Function(input) input.Columns)
                Dim values = AllocateMatrixValues(rows, columns)
                Dim numbers(4) As Double
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim failure As FormulaError = Nothing
                        For index = 0 To inputs.Length - 1
                            Dim value = inputs(index).BroadcastAt(row, column)
                            Dim number = If(TypeOf value Is String, ConvertWorksheetText("VALUE", value), NumberValue(value))
                            If TypeOf number Is FormulaError Then
                                failure = DirectCast(number, FormulaError)
                                Exit For
                            End If
                            numbers(index) = DirectCast(number, Double)
                        Next
                        values(row * columns + column) = If(failure Is Nothing, LoanValue(name, numbers), failure)
                    Next
                Next
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Shared Function LoanValue(name As String, values As Double()) As Object
                Dim rate = values(0)
                If name = "ISPMT" Then
                    If values(2) = 0.0 Then Return New FormulaError("#DIV/0!")
                    Return FiniteNumber(-values(3) * rate * (1.0 - values(1) / values(2)))
                End If

                Dim periods = values(1)
                Dim payment = values(2)
                Dim balance = values(3)
                Dim timing = If(values(4) = 0.0, 1.0, 1.0 + rate)
                Select Case name
                    Case "NPER"
                        If rate = 0.0 Then
                            If periods = 0.0 Then Return New FormulaError("#DIV/0!")
                            Return FiniteNumber(-(payment + balance) / periods)
                        End If
                        Dim adjustedPayment = periods * timing
                        Dim factor = (adjustedPayment - balance * rate) / (adjustedPayment + payment * rate)
                        If rate <= -1.0 OrElse factor <= 0.0 OrElse Double.IsNaN(factor) OrElse Double.IsInfinity(factor) Then
                            Return New FormulaError("#NUM!")
                        End If
                        Dim growth = Math.Log(1.0 + rate)
                        If growth = 0.0 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber(Math.Log(factor) / growth)
                    Case "PMT"
                        If periods = 0.0 OrElse rate <= -1.0 Then Return New FormulaError("#NUM!")
                        If rate = 0.0 Then Return FiniteNumber(-(payment + balance) / periods)
                        Return PeriodicLoanPayment(rate, periods, payment, balance, timing)
                End Select

                If rate = 0.0 Then Return FiniteNumber(-(balance + payment * periods))
                Dim compound = CompoundFactor(1.0 + rate, periods)
                Dim payments = payment * timing * (compound - 1.0) / rate
                If name = "FV" Then Return FiniteNumber(-(balance * compound + payments))
                If compound = 0.0 Then Return New FormulaError("#DIV/0!")
                Return FiniteNumber(-(balance + payments) / compound)
            End Function

            Private Shared Function CompoundFactor(base As Double, periods As Double) As Double
                If Math.Abs(periods) > Integer.MaxValue OrElse Math.Truncate(periods) <> periods Then Return Math.Pow(base, periods)
                ' Excel compounds whole periods with repeated squaring, including for small rates.
                Dim exponent = CInt(Math.Abs(periods))
                Dim factor = 1.0
                While exponent > 0
                    If (exponent And 1) <> 0 Then factor *= base
                    exponent >>= 1
                    If exponent <> 0 Then base *= base
                End While
                Return If(periods < 0.0, 1.0 / factor, factor)
            End Function

            Private Shared Function PeriodicLoanPayment(rate As Double, periods As Double, principal As Double,
                                                       target As Double, timing As Double) As Object
                ' PMT retains the small difference from one even when the compounded factor rounds to one.
                Dim growth = periods * LogarithmOfOnePlus(rate)
                If growth >= 0.0 Then
                    Dim discount = Math.Exp(-growth)
                    Dim discountAnnuity = -ExponentialMinusOne(-growth) / rate
                    Return FiniteNumber(-(principal + target * discount) / (discountAnnuity * timing))
                End If
                Dim compound = Math.Exp(growth)
                Dim annuity = ExponentialMinusOne(growth) / rate
                Return FiniteNumber(-(target + principal * compound) / (annuity * timing))
            End Function

            Private Shared Function LogarithmOfOnePlus(value As Double) As Double
                If Math.Abs(value) >= 0.0001 Then Return Math.Log(1.0 + value)
                Dim fraction = value / (2.0 + value)
                Dim square = fraction * fraction
                Return 2.0 * fraction * (1.0 + square * (1.0 / 3.0 + square / 5.0))
            End Function

            Private Shared Function ExponentialMinusOne(value As Double) As Double
                If Math.Abs(value) >= 0.025 Then Return Math.Exp(value) - 1.0
                Dim term = value
                Dim total = term
                For power = 2 To 9
                    term *= value / power
                    total += term
                Next
                Return total
            End Function

        End Class

    End Class

End Namespace
