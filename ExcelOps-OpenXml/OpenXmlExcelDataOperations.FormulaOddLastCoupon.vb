Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateOddLastCoupon(yield As Boolean, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 7 OrElse arguments.Count > 8 Then Return New FormulaError("#VALUE!")
                Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) OddLastCouponValue(yield, values))
            End Function

            Private Function OddLastCouponValue(yield As Boolean, values As Object()) As Object
                If values.Any(Function(value) TypeOf value Is Boolean) Then Return New FormulaError("#VALUE!")
                Dim dates(2) As Integer
                For i = 0 To dates.Length - 1
                    Dim input = ReadCouponDate(values(i))
                    If TypeOf input Is FormulaError Then Return input
                    dates(i) = DirectCast(input, Integer)
                Next
                Dim settlement = dates(0)
                Dim maturity = dates(1)
                Dim lastInterest = dates(2)
                If lastInterest >= settlement OrElse settlement >= maturity Then Return New FormulaError("#NUM!")
                Dim numbers(4) As Double
                For i = 0 To numbers.Length - 1
                    Dim input = ReadFinancialNumber(values.ElementAtOrDefault(i + 3))
                    If TypeOf input Is FormulaError Then Return input
                    numbers(i) = DirectCast(input, Double)
                Next
                Dim rate = numbers(0)
                Dim quote = numbers(1)
                Dim redemption = numbers(2)
                Dim frequency = Math.Truncate(numbers(3))
                Dim basis = Math.Truncate(numbers(4))
                If rate < 0 OrElse quote < 0 OrElse (yield AndAlso quote = 0.0) OrElse redemption <= 0 OrElse
                    (frequency <> 1 AndAlso frequency <> 2 AndAlso frequency <> 4) OrElse numbers(4) < 0.0 OrElse basis > 4 Then
                    Return New FormulaError("#NUM!")
                End If

                Dim fractions = OddLastCouponFractions(lastInterest, settlement, maturity, CInt(frequency), CInt(basis))
                If Not fractions.HasValue Then Return New FormulaError("#NUM!")
                Dim coupon = (rate * 100.0) / frequency
                Dim dirty = redemption + coupon * fractions.Value.AtMaturity
                Dim accrued = coupon * fractions.Value.AtSettlement
                Dim term = fractions.Value.ForDiscount
                If Not yield Then Return FiniteNumber(dirty / (1.0 + (quote / frequency) * term) - accrued)
                If term = 0.0 Then Return 0.0
                Return FiniteNumber((dirty / (quote + accrued) - 1.0) * frequency / term)
            End Function

            Private Function OddLastCouponFractions(anchor As Integer, settlement As Integer, maturity As Integer, frequency As Integer,
                                                     basis As Integer) As (AtMaturity As Double, AtSettlement As Double, ForDiscount As Double)?
                Dim months = 12 \ frequency
                Dim origin = DateParts(anchor)
                Dim anchorMonth As New DateTime(origin.Year, origin.Month, 1)
                Dim originAtMonthEnd = origin.Day = DaysInExcelMonth(origin.Year, origin.Month)
                Dim previous = anchor
                Dim prospectiveMonth = anchorMonth
                Dim period As Integer
                Dim atMaturity As Double
                Dim atSettlement As Double
                Dim discount As Double
                While previous < maturity
                    If prospectiveMonth.Year = 9999 AndAlso prospectiveMonth.Month + months > 12 Then Return Nothing
                    prospectiveMonth = prospectiveMonth.AddMonths(months)
                    Dim first = DateParts(previous)
                    Dim following = CouponDate(prospectiveMonth, first.Day, endOfMonth:=False)
                    Dim finalEnd = Math.Min(maturity, following)
                    period += 1
                    ' An original coupon date can be later than its rolled date after a short month.
                    ' The remaining coupon days use that period's denominator while discounting ends on the rolled date.
                    If finalEnd < maturity AndAlso maturity <= CouponDate(anchorMonth.AddMonths(period * months), origin.Day, endOfMonth:=originAtMonthEnd) Then
                        finalEnd = maturity
                    End If
                    Dim denominator = OddCouponPaymentDays(previous, following, basis)
                    If denominator <= 0.0 Then Return Nothing
                    atMaturity += OddCouponPaymentDays(previous, finalEnd, basis) / denominator
                    If settlement >= finalEnd Then
                        atSettlement += OddCouponPaymentDays(previous, finalEnd, basis) / denominator
                    ElseIf settlement > previous Then
                        atSettlement += CouponDaySpan(previous, settlement, basis) / denominator
                    End If
                    Dim discountStart = Math.Max(previous, settlement)
                    Dim discountEnd = Math.Min(following, maturity)
                    If discountStart < discountEnd Then discount += CouponDaySpan(discountStart, discountEnd, basis) / denominator
                    previous = finalEnd
                End While
                Return (atMaturity, atSettlement, discount)
            End Function

            Private Function OddCouponPaymentDays(first As Integer, last As Integer, basis As Integer) As Double
                If basis <> 0 Then Return CouponDaySpan(first, last, basis)
                Dim left = DateParts(first)
                Dim right = DateParts(last)
                Dim leftDay = If(left.Month = 2 AndAlso left.Day = DaysInExcelMonth(left.Year, 2), 30, Math.Min(left.Day, 30))
                Dim rightDay = If(right.Month = 2 AndAlso right.Day = DaysInExcelMonth(right.Year, 2), 30, Math.Min(right.Day, 30))
                Return (right.Year - left.Year) * 360.0 + (right.Month - left.Month) * 30.0 + rightDay - leftDay
            End Function

        End Class

    End Class

End Namespace
