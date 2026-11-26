Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateOddFirstCoupon(yield As Boolean, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 8 OrElse arguments.Count > 9 Then Return New FormulaError("#VALUE!")
                Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) OddFirstCouponValue(yield, values))
            End Function

            Private Function OddFirstCouponValue(yield As Boolean, values As Object()) As Object
                If values.Any(Function(value) TypeOf value Is Boolean) Then Return New FormulaError("#VALUE!")
                Dim dates(3) As Integer
                For i = 0 To dates.Length - 1
                    Dim input = ReadCouponDate(values(i))
                    If TypeOf input Is FormulaError Then Return input
                    dates(i) = DirectCast(input, Integer)
                Next
                Dim settlement = dates(0)
                Dim maturity = dates(1)
                Dim issue = dates(2)
                Dim firstCoupon = dates(3)
                If issue >= settlement OrElse settlement >= firstCoupon OrElse firstCoupon >= maturity Then Return New FormulaError("#NUM!")
                Dim numbers(4) As Double
                For i = 0 To numbers.Length - 1
                    Dim input = ReadFinancialNumber(values.ElementAtOrDefault(i + 4))
                    If TypeOf input Is FormulaError Then Return input
                    numbers(i) = DirectCast(input, Double)
                Next
                Dim rate = numbers(0)
                Dim quote = numbers(1)
                Dim redemption = numbers(2)
                Dim frequency = Math.Truncate(numbers(3))
                Dim basis = Math.Truncate(numbers(4))
                If rate < 0 OrElse quote < 0 OrElse (yield AndAlso quote = 0.0) OrElse redemption <= 0 OrElse
                    (frequency <> 1 AndAlso frequency <> 2 AndAlso frequency <> 4) OrElse numbers(4) < 0 OrElse basis > 4 Then
                    Return New FormulaError("#NUM!")
                End If

                Dim first = DateParts(firstCoupon)
                Dim last = DateParts(maturity)
                Dim months = 12 \ CInt(frequency)
                Dim laterCoupons = ((last.Year - first.Year) * 12 + last.Month - first.Month) \ months
                Dim endOfMonth = last.Day = DaysInExcelMonth(last.Year, last.Month)
                Dim maturityMonth As New DateTime(last.Year, last.Month, 1)
                If laterCoupons = 0 OrElse CouponDate(maturityMonth.AddMonths(-laterCoupons * months), last.Day, endOfMonth) <> firstCoupon Then
                    Return New FormulaError("#NUM!")
                End If
                Dim fractions = OddFirstCouponFractions(issue, settlement, firstCoupon, months, CInt(basis), endOfMonth)
                If Not fractions.HasValue Then Return New FormulaError("#NUM!")
                Dim coupon = (rate * 100.0) / frequency
                Dim firstPayment = coupon * fractions.Value.Payment
                Dim accrued = coupon * fractions.Value.Accrued
                Dim term = fractions.Value.ForDiscount
                Dim presentValue = Function(perPeriod As Double) OddFirstDirtyPrice(laterCoupons, term, firstPayment, coupon, redemption, perPeriod)
                If Not yield Then Return FiniteNumber(presentValue(quote / frequency) - accrued)
                Return InvertCouponPrice(quote + accrued, frequency, presentValue)
            End Function

            Private Function OddFirstCouponFractions(issue As Integer, settlement As Integer, firstCoupon As Integer, months As Integer,
                                                      basis As Integer, endOfMonth As Boolean) As (Payment As Double, Accrued As Double, ForDiscount As Double)?
                Dim first = DateParts(firstCoupon)
                Dim month As New DateTime(first.Year, first.Month, 1)
                Dim prior = CouponDate(month.AddMonths(-months), first.Day, endOfMonth)
                Dim normal = OddFirstPeriodDays(prior, firstCoupon, months, basis)
                If normal <= 0 Then Return Nothing
                Dim issueDays = CouponDaySpan(issue, firstCoupon, basis)
                If issue = prior AndAlso issueDays = normal Then Return Nothing

                If issue >= prior Then
                    Dim elapsed = CouponDaySpan(issue, settlement, basis) / normal
                    Dim discount = CouponDaySpan(settlement, firstCoupon, basis) / normal
                    If (basis = 0 OrElse basis = 4) AndAlso issue = prior AndAlso issueDays > normal Then discount = 1.0 - elapsed
                    Return (issueDays / normal, elapsed, discount)
                End If

                ' The coupon and accrual roll the day across short months. Discounting keeps the anchored coupon schedule.
                Dim cashPrior = CouponDate(month.AddMonths(-months), first.Day, endOfMonth:=False)
                Dim firstPayment = 1.0 + OddFirstFractions(issue, cashPrior, firstCoupon, months, basis, endOfMonth:=False, rollDays:=True, couponPayment:=True)
                Dim accrued = OddFirstFractions(issue, settlement, firstCoupon, months, basis, endOfMonth:=False, rollDays:=True)
                Dim remaining As Double
                If basis = 0 OrElse basis = 4 Then
                    remaining = OddFirstFractions(issue, firstCoupon, firstCoupon, months, basis, endOfMonth, rollDays:=False) -
                        OddFirstFractions(issue, settlement, firstCoupon, months, basis, endOfMonth, rollDays:=False)
                Else
                    remaining = OddFirstFractions(settlement, firstCoupon, firstCoupon, months, basis, endOfMonth, rollDays:=False)
                End If
                ' Excel counts another quasi-coupon on an end-of-month schedule if the
                ' settlement within the last interval is before that month's last day.
                If endOfMonth AndAlso settlement > prior Then
                    Dim settled = DateParts(settlement)
                    If settled.Day < DaysInExcelMonth(settled.Year, settled.Month) Then remaining += 1.0
                End If
                Return (firstPayment, accrued, remaining)
            End Function

            Private Function OddFirstFractions(startDate As Integer, endDate As Integer, firstCoupon As Integer, months As Integer, basis As Integer,
                                               endOfMonth As Boolean, rollDays As Boolean, Optional couponPayment As Boolean = False) As Double
                Dim anchor = DateParts(firstCoupon)
                Dim month As New DateTime(anchor.Year, anchor.Month, 1)
                Dim following = firstCoupon
                Dim day = anchor.Day
                Dim sum As Double
                While following > startDate
                    month = month.AddMonths(-months)
                    Dim prior = CouponDate(month, day, endOfMonth)
                    Dim first = Math.Max(prior, startDate)
                    Dim last = Math.Min(following, endDate)
                    If first < last Then
                        ' When settlement falls on the last quasi-coupon date, Excel discounts the first
                        ' payment using the actual length of that coupon for Actual/360 and Actual/365.
                        Dim actualDaysAtBoundary = startDate = prior AndAlso following = firstCoupon AndAlso (basis = 2 OrElse basis = 3)
                        Dim normalizeWholePeriod = If(couponPayment, prior > startDate, Not rollDays AndAlso Not actualDaysAtBoundary)
                        If normalizeWholePeriod AndAlso first = prior AndAlso last = following Then
                            sum += 1.0
                        Else
                            sum += CouponDaySpan(first, last, basis) / OddFirstPeriodDays(prior, following, months, basis)
                        End If
                    End If
                    If rollDays Then day = DateParts(prior).Day
                    following = prior
                End While
                Return sum
            End Function

            Private Shared Function OddFirstPeriodDays(previous As Integer, following As Integer, months As Integer, basis As Integer) As Double
                Return If(basis = 1, following - previous, If(basis = 3, 365.0, 360.0) * months / 12.0)
            End Function

            Private Shared Function OddFirstDirtyPrice(coupons As Integer, duration As Double, firstCoupon As Double, coupon As Double,
                                                       redemption As Double, yieldPerPeriod As Double) As Double
                Dim factor = 1.0 + yieldPerPeriod
                If factor = 1.0 Then Return firstCoupon + coupons * coupon + redemption
                Dim finalDiscount = Math.Pow(factor, -coupons)
                Dim cash = firstCoupon + redemption * finalDiscount
                If coupon <> 0.0 Then
                    If Math.Abs(yieldPerPeriod) * coupons < 0.0001 Then
                        Dim discount = 1.0
                        For i = 1 To coupons
                            discount /= factor
                            cash += coupon * discount
                        Next
                    Else
                        cash += coupon * (1.0 - finalDiscount) / yieldPerPeriod
                    End If
                End If
                Return cash * Math.Pow(factor, -duration)
            End Function

        End Class

    End Class

End Namespace
