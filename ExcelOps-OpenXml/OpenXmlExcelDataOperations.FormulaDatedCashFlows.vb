Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const Last1900DateSerial As Double = 2958465.0
            Private Const LowestDatedReturnRate As Double = -0.9999999999999999

            Private Function EvaluateDatedCashFlow(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                   formulaCell As ExcelCell) As Object
                Dim returnRate = name = "XIRR"
                If arguments.Count < If(returnRate, 2, 3) OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim inputs = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                Dim ratesInput = If(returnRate, If(inputs.Length = 3, inputs(2), 0.1), inputs(0))
                If Not returnRate AndAlso TypeOf ratesInput Is FormulaArea AndAlso IsMultiValue(ratesInput) Then
                    Return New FormulaError("#VALUE!")
                End If
                Dim flows = inputs(If(returnRate, 0, 1))
                Dim dates = inputs(If(returnRate, 1, 2))
                Dim rawError = If(TryCast(flows, FormulaError), TryCast(dates, FormulaError))
                Dim payments As Double() = Nothing
                Dim years As Double() = Nothing
                Dim scheduleError = If(rawError, ReadDatedSchedule(flows, dates, returnRate, payments, years))

                Dim rates = AsMatrix(ratesInput)
                Dim results = AllocateMatrixValues(rates.Rows, rates.Columns)
                For index = 0 To results.Length - 1
                    Dim input = rates.Values(index)
                    Dim rate = If(TypeOf input Is Boolean, New FormulaError("#VALUE!"), ReadFinancialNumber(input))
                    Dim result As Object
                    If TypeOf rate Is FormulaError Then
                        result = rate
                    ElseIf rawError IsNot Nothing Then
                        result = rawError
                    ElseIf DirectCast(rate, Double) <= If(returnRate, -1.0, 0.0) Then
                        result = New FormulaError("#NUM!")
                    ElseIf scheduleError IsNot Nothing Then
                        result = scheduleError
                    ElseIf returnRate Then
                        result = DatedReturnRate(payments, years, DirectCast(rate, Double))
                    Else
                        result = FiniteNumber(DatedValue(payments, years, DirectCast(rate, Double)))
                    End If
                    results(index) = result
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rates.Rows, rates.Columns, results)
            End Function

            Private Function ReadDatedSchedule(flows As Object, dateInput As Object, returnRate As Boolean,
                                                ByRef payments As Double(), ByRef years As Double()) As FormulaError
                Dim values = AsMatrix(flows).Values
                Dim dates = AsMatrix(dateInput).Values
                If values.Length <> dates.Length Then Return New FormulaError("#NUM!")
                If returnRate AndAlso values.Length < 2 Then Return New FormulaError("#N/A")
                payments = New Double(values.Length - 1) {}
                years = New Double(dates.Length - 1) {}
                Dim start As Double
                Dim earliest = If(returnRate, 0.0, 1.0)
                Dim latest = If(returnRate, ToExcelDateSerial(DateTime.MaxValue.Date), Last1900DateSerial)
                For index = 0 To values.Length - 1
                    Dim payment = values(index)
                    Dim day = dates(index)
                    If TypeOf payment Is FormulaError Then Return DirectCast(payment, FormulaError)
                    If TypeOf day Is FormulaError Then Return DirectCast(day, FormulaError)
                    If returnRate Then
                        If TypeOf payment Is Boolean OrElse TypeOf day Is Boolean Then Return New FormulaError("#VALUE!")
                        payment = ReadFinancialNumber(payment)
                        day = ReadFinancialNumber(day)
                        If TypeOf payment Is FormulaError Then Return DirectCast(payment, FormulaError)
                        If TypeOf day Is FormulaError Then Return DirectCast(day, FormulaError)
                    Else
                        If Not TypeOf payment Is Double Then Return New FormulaError("#NUM!")
                        If Not TypeOf day Is Double Then Return New FormulaError("#VALUE!")
                    End If
                    Dim serial = Math.Truncate(DirectCast(day, Double))
                    If serial < earliest OrElse serial > latest Then Return New FormulaError("#NUM!")
                    If index = 0 Then start = serial
                    If Not returnRate AndAlso serial < start Then Return New FormulaError("#NUM!")
                    payments(index) = DirectCast(payment, Double)
                    years(index) = (serial - start) / 365.0
                Next
                Return Nothing
            End Function

            Private Shared Function DatedValue(payments As Double(), years As Double(), rate As Double) As Double
                Dim factor = 1.0 + rate
                Dim total = 0.0
                For index = 0 To payments.Length - 1
                    If payments(index) <> 0.0 Then total += payments(index) / Math.Pow(factor, years(index))
                Next
                Return total
            End Function

            Private Shared Function DatedReturnRate(payments As Double(), years As Double(), guess As Double) As Object
                If Not payments.Any(Function(amount) amount > 0.0) OrElse Not payments.Any(Function(amount) amount < 0.0) OrElse
                    years.All(Function(year) year = 0.0) Then Return New FormulaError("#NUM!")
                Dim undiscounted = payments.Sum()
                If undiscounted = 0.0 Then Return 0.0

                Dim opening = payments.First(Function(payment) payment <> 0.0)
                Dim low = 0.0
                Dim high = 0.0
                Dim lowValue = undiscounted
                If guess < 0.0 OrElse Math.Sign(undiscounted) = Math.Sign(opening) Then
                    If guess >= 1.0 OrElse years.Any(Function(year) year < 0.0) Then Return New FormulaError("#NUM!")
                    low = LowestDatedReturnRate
                    lowValue = DatedValue(payments, years, low)
                    If Double.IsNaN(lowValue) OrElse Math.Sign(lowValue) = Math.Sign(undiscounted) Then Return New FormulaError("#NUM!")
                Else
                    high = If(guess = 0.0, 0.00001, guess)
                    Dim highValue = DatedValue(payments, years, high)
                    For attempt = 1 To 75
                        If Double.IsNaN(highValue) OrElse Math.Sign(highValue) <> Math.Sign(lowValue) Then Exit For
                        high *= 2.0
                        highValue = DatedValue(payments, years, high)
                    Next
                    If Double.IsNaN(highValue) OrElse Math.Sign(highValue) = Math.Sign(lowValue) Then Return New FormulaError("#NUM!")
                End If

                For iteration = 1 To 100
                    Dim middle = low + (high - low) / 2.0
                    Dim value = DatedValue(payments, years, middle)
                    If Double.IsNaN(value) Then Return New FormulaError("#NUM!")
                    If value = 0.0 OrElse high - low <= 1.0E-13 * Math.Max(1.0, Math.Abs(middle)) Then Return middle
                    If Math.Sign(value) = Math.Sign(lowValue) Then
                        low = middle
                    Else
                        high = middle
                    End If
                Next
                Return New FormulaError("#NUM!")
            End Function

        End Class

    End Class

End Namespace
