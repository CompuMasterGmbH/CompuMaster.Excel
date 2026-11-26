Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateAccruedPeriodicInterest(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 6 OrElse arguments.Count > 8 Then Return New FormulaError("#VALUE!")
                Dim inputs = arguments.ToArray()
                If inputs(4) Is OmittedFormulaArgument Then inputs(4) = Function(calculation, cell) 1000.0
                If inputs.Length = 8 AndAlso inputs(7) Is OmittedFormulaArgument Then inputs(7) = Function(calculation, cell) True
                Return BroadcastReferenceArguments(inputs, formulaCell, AddressOf AccruedPeriodicInterestValue)
            End Function

            Private Function AccruedPeriodicInterestValue(values As Object()) As Object
                If values.Take(7).Any(Function(value) TypeOf value Is Boolean) Then Return New FormulaError("#VALUE!")
                Dim dates(2) As Integer
                For i = 0 To dates.Length - 1
                    Dim parsed = ReadDateDifferenceInput(values(i))
                    If TypeOf parsed Is FormulaError Then Return parsed
                    Dim serial = DirectCast(parsed, Double)
                    If serial <= 0.0 OrElse Not IsCalendarSerialInRange(serial) Then Return New FormulaError("#NUM!")
                    dates(i) = CInt(Math.Truncate(serial))
                Next
                If dates(0) >= dates(2) Then Return New FormulaError("#NUM!")
                Dim numbers(3) As Double
                For i = 0 To numbers.Length - 1
                    Dim parsed = ReadFinancialNumber(values.ElementAtOrDefault(i + 3))
                    If TypeOf parsed Is FormulaError Then Return parsed
                    numbers(i) = DirectCast(parsed, Double)
                Next
                Dim frequency = Math.Truncate(numbers(2))
                Dim basis = Math.Truncate(numbers(3))
                If numbers(0) <= 0 OrElse numbers(1) <= 0 OrElse numbers(3) < 0 OrElse basis > 4 OrElse
                    (frequency <> 1 AndAlso frequency <> 2 AndAlso frequency <> 4) Then Return New FormulaError("#NUM!")
                Dim method = If(values.Length = 8, values(7), True)
                If TypeOf method Is String Then Return New FormulaError("#VALUE!")
                Dim fromIssue = If(TypeOf method Is Boolean, DirectCast(method, Boolean),
                    TypeOf method Is Double AndAlso DirectCast(method, Double) <> 0.0)
                Dim interest = AccruedPeriodicFraction(dates(0), dates(1), dates(2), CInt(frequency), CInt(basis), fromIssue)
                If Not interest.HasValue Then Return New FormulaError("#NUM!")
                Return FiniteNumber(numbers(0) * numbers(1) / frequency * interest.Value)
            End Function

            Private Function AccruedPeriodicFraction(issue As Integer, anchor As Integer, settlement As Integer, frequency As Integer,
                                                     basis As Integer, fromIssue As Boolean) As Double?
                Dim months = 12 \ frequency
                Dim first = DateParts(anchor)
                Dim anchorMonth As New DateTime(first.Year, first.Month, 1)
                Dim endOfMonth = first.Day = DaysInExcelMonth(first.Year, first.Month)
                Dim periodStart = CouponDate(anchorMonth.AddMonths(-months), first.Day, endOfMonth)
                Dim denominator = If(basis = 1, anchor - periodStart, If(basis = 3, 365.0, 360.0) / frequency)
                If denominator <= 0.0 Then Return Nothing

                Dim periodMonth = anchorMonth.AddMonths(-months)
                If fromIssue Then
                    ' Excel charges each complete quasi-coupon as one period, using the first coupon's denominator for the active period.
                    Dim following = anchor
                    While settlement > following
                        periodMonth = periodMonth.AddMonths(months)
                        periodStart = following
                        If periodMonth.Year = 9999 AndAlso periodMonth.Month + months > 12 Then Exit While
                        following = CouponDate(periodMonth.AddMonths(months), first.Day, endOfMonth)
                    End While
                End If
                Dim fraction = CouponDaySpan(Math.Max(periodStart, issue), settlement, basis) / denominator
                While periodStart > issue
                    periodMonth = periodMonth.AddMonths(-months)
                    Dim preceding = CouponDate(periodMonth, first.Day, endOfMonth)
                    If preceding >= issue Then
                        If fromIssue Then fraction += 1.0
                    Else
                        ' The oldest partial coupon uses its own day-count denominator.
                        Dim partialDays = If(basis = 1 OrElse basis = 4, CouponDaySpan(preceding, periodStart, basis), denominator)
                        If partialDays <= 0.0 Then Return Nothing
                        fraction += CouponDaySpan(issue, periodStart, basis) / partialDays
                    End If
                    periodStart = preceding
                End While
                Return fraction
            End Function

        End Class

    End Class

End Namespace
