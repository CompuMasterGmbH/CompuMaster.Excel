Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateLoanFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                 formulaCell As ExcelCell) As Object
                Dim interestOnly = name = "ISPMT"
                Dim cumulative = name = "CUMIPMT" OrElse name = "CUMPRINC"
                Dim installment = name = "IPMT" OrElse name = "PPMT"
                Dim minimum = If(cumulative, 6, If(interestOnly OrElse installment, 4, 3))
                Dim maximum = If(cumulative OrElse installment, 6, If(interestOnly, 4, 5))
                If arguments.Count < minimum OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")

                Dim evaluated = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                If cumulative Then
                    If evaluated.Any(Function(value) TypeOf value Is FormulaArea AndAlso IsMultiValue(value)) Then Return New FormulaError("#VALUE!")
                    If evaluated.Any(Function(value) value Is Nothing) Then Return New FormulaError("#N/A")
                End If
                Dim inputs = evaluated.Select(Function(value) AsMatrix(value)).ToArray()
                Dim rows = inputs.Max(Function(input) input.Rows)
                Dim columns = inputs.Max(Function(input) input.Columns)
                Dim values = AllocateMatrixValues(rows, columns)
                Dim numbers(5) As Double
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim failure As FormulaError = Nothing
                        For index = 0 To inputs.Length - 1
                            Dim value = inputs(index).BroadcastAt(row, column)
                            Dim number As Object
                            If TypeOf value Is String Then
                                number = ConvertWorksheetText("VALUE", value)
                            ElseIf cumulative AndAlso TypeOf value Is Boolean Then
                                number = New FormulaError("#VALUE!")
                            Else
                                number = NumberValue(value)
                            End If
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
                If name = "IPMT" OrElse name = "PPMT" Then Return InstallmentComponent(name = "IPMT", values)
                If name = "CUMIPMT" OrElse name = "CUMPRINC" Then Return AccumulatedInstallments(name = "CUMIPMT", values)
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

            Private Shared Function InstallmentComponent(interestOnly As Boolean, values As Double()) As Object
                Dim rate = values(0)
                Dim period = values(1)
                Dim periods = values(2)
                Dim principal = values(3)
                Dim target = values(4)
                Dim due = values(5) <> 0.0
                If period < 1.0 OrElse period >= periods + 1.0 Then Return New FormulaError("#NUM!")
                If interestOnly AndAlso due AndAlso period = 1.0 Then Return 0.0
                If rate <= -1.0 Then Return New FormulaError("#NUM!")
                If rate = 0.0 Then Return If(interestOnly, 0.0, FiniteNumber(-(principal + target) / periods))

                Dim timing = If(due, 1.0 + rate, 1.0)
                Dim payment = PeriodicLoanPayment(rate, periods, principal, target, timing)
                If TypeOf payment Is FormulaError Then Return payment
                Dim amount = DirectCast(payment, Double)
                If due AndAlso period = 1.0 Then Return amount
                Dim periodicGrowth = LogarithmOfOnePlus(rate)
                Dim loanGrowth = periods * periodicGrowth
                Dim growth = (period - 1.0) * periodicGrowth
                If interestOnly AndAlso Math.Abs(growth) < 0.001 Then
                    Return FiniteNumber(-principal * rate * Math.Exp(growth) / timing - amount * ExponentialMinusOne(growth))
                End If

                ' Calculate principal with the same discount scale as the payment.
                Dim principalComponent As Double
                If loanGrowth >= 0.0 Then
                    principalComponent = (principal + target) * rate / timing *
                        Math.Exp((period - periods - 1.0) * periodicGrowth) / ExponentialMinusOne(-loanGrowth)
                Else
                    principalComponent = -(principal + target) * rate / timing *
                        Math.Exp(growth) / ExponentialMinusOne(loanGrowth)
                End If
                Return FiniteNumber(If(interestOnly, amount - principalComponent, principalComponent))
            End Function

            Private Shared Function AccumulatedInstallments(interestOnly As Boolean, values As Double()) As Object
                Dim rate = values(0)
                Dim periods = values(1)
                Dim principal = values(2)
                Dim start = values(3)
                Dim finish = values(4)
                Dim type = values(5)
                If rate <= 0.0 OrElse periods <= 0.0 OrElse principal <= 0.0 OrElse start < 1.0 OrElse start > finish OrElse
                    finish > periods OrElse (type <> 0.0 AndAlso type <> 1.0) Then Return New FormulaError("#NUM!")

                Dim first = Math.Ceiling(start)
                Dim last = Math.Floor(finish)
                If first > last Then Return 0.0
                Dim count = last - first + 1.0
                Dim timing = If(type = 0.0, 1.0, 1.0 + rate)
                Dim payment = PeriodicLoanPayment(rate, periods, principal, 0.0, timing)
                If TypeOf payment Is FormulaError Then Return payment
                Dim amount = DirectCast(payment, Double)
                If first = 1.0 AndAlso last = periods Then Return FiniteNumber(If(interestOnly, count * amount + principal, -principal))

                Dim immediate = If(type = 1.0 AndAlso first = 1.0, amount, 0.0)
                If immediate <> 0.0 Then first += 1.0
                Dim growth = LogarithmOfOnePlus(rate)
                Dim remaining = last - first + 1.0
                ' Sum the geometric progression of principal payments using discount factors.
                Dim principalSum = immediate - principal / timing * Math.Exp((last - periods) * growth) *
                    ExponentialMinusOne(-remaining * growth) / ExponentialMinusOne(-periods * growth)
                Return FiniteNumber(If(interestOnly, count * amount - principalSum, principalSum))
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
