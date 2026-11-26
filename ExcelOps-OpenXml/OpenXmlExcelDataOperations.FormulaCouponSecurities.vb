Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Structure CouponSchedule
                Friend Number As Integer
                Friend BeforeSettlement As Integer
                Friend AfterSettlement As Integer
                Friend Elapsed As Double
                Friend Remaining As Double
                Friend PeriodDays As Double
            End Structure

            Private Function EvaluateCouponSecurity(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Dim couponMetric = name.StartsWith("COUP", StringComparison.Ordinal)
                Dim duration = name.EndsWith("DURATION", StringComparison.Ordinal)
                Dim required = If(couponMetric, 3, If(duration, 5, 6))
                If arguments.Count < required OrElse arguments.Count > required + 1 Then Return New FormulaError("#VALUE!")
                Return BroadcastReferenceArguments(arguments, formulaCell,
                    Function(values) CouponSecurityValue(name, couponMetric, duration, required, values))
            End Function

            Private Function CouponSecurityValue(name As String, couponMetric As Boolean, duration As Boolean, required As Integer,
                                                 values As Object()) As Object
                If values.Any(Function(input) TypeOf input Is Boolean) Then Return New FormulaError("#VALUE!")
                Dim settlement = ReadCouponDate(values(0))
                If TypeOf settlement Is FormulaError Then Return settlement
                Dim maturity = ReadCouponDate(values(1))
                If TypeOf maturity Is FormulaError Then Return maturity
                Dim first = DirectCast(settlement, Integer)
                Dim last = DirectCast(maturity, Integer)
                If first >= last Then Return New FormulaError("#NUM!")
                Dim frequencyValue = ReadFinancialNumber(values(required - 1))
                If TypeOf frequencyValue Is FormulaError Then Return frequencyValue
                Dim basisValue = ReadFinancialNumber(If(values.Length > required, values(required), Nothing))
                If TypeOf basisValue Is FormulaError Then Return basisValue
                Dim frequency = Math.Truncate(DirectCast(frequencyValue, Double))
                Dim basis = Math.Truncate(DirectCast(basisValue, Double))
                If (frequency <> 1 AndAlso frequency <> 2 AndAlso frequency <> 4) OrElse DirectCast(basisValue, Double) < 0.0 OrElse basis > 4 Then
                    Return New FormulaError("#NUM!")
                End If
                Dim schedule = ReadCouponSchedule(first, last, CInt(frequency), CInt(basis))
                If schedule.PeriodDays <= 0 OrElse schedule.Number <= 0 Then Return New FormulaError("#NUM!")
                If couponMetric Then
                    Select Case name
                        Case "COUPDAYBS" : Return schedule.Elapsed
                        Case "COUPDAYS"
                            If basis = 1 Then Return ReportedActualCouponDays(schedule, last, useStoredDay:=Uses1904Dates AndAlso Not TypeOf values(1) Is String)
                            Return schedule.PeriodDays
                        Case "COUPDAYSNC" : Return schedule.Remaining
                        Case "COUPNCD" : Return CDbl(schedule.AfterSettlement - If(Uses1904Dates, DateSystemsOffset, 0))
                        Case "COUPNUM" : Return CDbl(schedule.Number)
                        Case Else : Return CDbl(schedule.BeforeSettlement - If(Uses1904Dates, DateSystemsOffset, 0))
                    End Select
                End If

                Dim interest = ReadFinancialNumber(values(2))
                If TypeOf interest Is FormulaError Then Return interest
                Dim value = ReadFinancialNumber(values(3))
                If TypeOf value Is FormulaError Then Return value
                Dim redemption = ReadFinancialNumber(If(duration, CObj(100.0), values(4)))
                If TypeOf redemption Is FormulaError Then Return redemption
                Dim rate = DirectCast(interest, Double)
                Dim level = DirectCast(value, Double)
                Dim face = DirectCast(redemption, Double)
                If rate < 0.0 OrElse face <= 0.0 OrElse If(name = "YIELD", level <= 0.0, level < 0.0) Then
                    Return New FormulaError("#NUM!")
                End If
                Dim coupon = 100.0 * rate / frequency
                Dim fraction = (schedule.PeriodDays - schedule.Elapsed) / schedule.PeriodDays
                Dim accrued = coupon * schedule.Elapsed / schedule.PeriodDays
                If duration Then
                    Return FiniteNumber(CouponDuration(schedule.Number, fraction, coupon, level / frequency) /
                                        If(name = "MDURATION", frequency + level, frequency))
                ElseIf name = "PRICE" Then
                    Return FiniteNumber(DirtyCouponPrice(schedule.Number, fraction, coupon, face, level / frequency) - accrued)
                Else
                    If schedule.Number = 1 Then
                        Dim yieldBasis = If(basis = 0.0 OrElse basis = 4.0, CInt(basis), 1)
                        Dim elapsed = CouponDaySpan(schedule.BeforeSettlement, first, yieldBasis)
                        Dim remaining = CouponDaySpan(first, schedule.AfterSettlement, yieldBasis)
                        Dim period = If(yieldBasis = 1, CDbl(schedule.AfterSettlement - schedule.BeforeSettlement), schedule.PeriodDays)
                        fraction = remaining / period
                        accrued = coupon * elapsed / period
                    End If
                    Return CalculateCouponYield(schedule.Number, fraction, coupon, face, level + accrued, frequency)
                End If
            End Function

            Private Function ReadCouponDate(value As Object) As Object
                If TypeOf value Is Boolean Then Return New FormulaError("#VALUE!")
                Dim raw = ReadDateDifferenceInput(value)
                If TypeOf raw Is FormulaError Then Return raw
                Dim serial = DirectCast(raw, Double)
                If Not IsCalendarSerialInRange(serial) Then Return New FormulaError("#NUM!")
                ' Excel's financial functions apply the extra 1904 offset to numeric dates.
                Dim adjusted = CInt(Math.Truncate(serial)) + If(Uses1904Dates AndAlso Not TypeOf value Is String, DateSystemsOffset, 0)
                If Not IsCalendarSerialInRange(adjusted) Then Return New FormulaError("#NUM!")
                Return adjusted
            End Function

            Private Function ReadCouponSchedule(settlement As Integer, maturity As Integer, frequency As Integer, basis As Integer) As CouponSchedule
                Dim first = DateParts(settlement)
                Dim last = DateParts(maturity)
                Dim months = 12 \ frequency
                Dim intervals = ((last.Year - first.Year) * 12 + last.Month - first.Month) \ months
                Dim lastMonth As New DateTime(last.Year, last.Month, 1)
                Dim endOfMonth = last.Day = DaysInExcelMonth(last.Year, last.Month)
                Dim previous = CouponDate(lastMonth.AddMonths(-intervals * months), last.Day, endOfMonth)
                If previous > settlement Then
                    intervals += 1
                    previous = CouponDate(lastMonth.AddMonths(-intervals * months), last.Day, endOfMonth)
                End If
                Dim following = CouponDate(lastMonth.AddMonths(-(intervals - 1) * months), last.Day, endOfMonth)
                Dim periodDays As Double
                Select Case basis
                    Case 1 : periodDays = following - previous
                    Case 3 : periodDays = 365.0 / frequency
                    Case Else : periodDays = 360.0 / frequency
                End Select
                Dim elapsed = CouponDaySpan(previous, settlement, basis)
                Dim remaining = If(basis = 0, periodDays - elapsed, CouponDaySpan(settlement, following, basis))
                Return New CouponSchedule With {
                    .Number = intervals, .BeforeSettlement = previous, .AfterSettlement = following,
                    .Elapsed = elapsed, .Remaining = remaining, .PeriodDays = periodDays
                }
            End Function

            Private Function CouponDate(month As DateTime, targetDay As Integer, endOfMonth As Boolean) As Integer
                Dim last = DaysInExcelMonth(month.Year, month.Month)
                Dim day = If(endOfMonth, last, Math.Min(targetDay, last))
                Return CInt(ToExcelDateSerial(month)) + day - 1
            End Function

            Private Function ReportedActualCouponDays(schedule As CouponSchedule, maturity As Integer, useStoredDay As Boolean) As Double
                Dim anchor = DateParts(maturity)
                Dim day = anchor.Day
                If useStoredDay Then
                    Dim stored = DateParts(maturity - DateSystemsOffset)
                    If stored.Month = anchor.Month Then day = stored.Day
                End If
                If anchor.Month = 2 AndAlso day = 29 Then Return schedule.PeriodDays
                Dim first = DateParts(schedule.BeforeSettlement)
                Dim last = DateParts(schedule.AfterSettlement)
                Dim previousDay = Math.Min(day, DaysInExcelMonth(first.Year, first.Month))
                Dim nextDay = Math.Min(day, DaysInExcelMonth(last.Year, last.Month))
                If anchor.Month <> 2 Then
                    If first.Month = 2 Then previousDay = Math.Min(day, DaysInExcelMonth(anchor.Year, 2))
                    If last.Month = 2 Then nextDay = Math.Min(day, DaysInExcelMonth(anchor.Year, 2))
                End If
                Return schedule.PeriodDays + nextDay - last.Day - previousDay + first.Day
            End Function

            Private Function CouponDaySpan(first As Integer, last As Integer, basis As Integer) As Double
                Select Case basis
                    Case 0 : Return YearFractionUsDays(DateParts(first), DateParts(last))
                    Case 4 : Return Days360(first, last, european:=True)
                    Case Else : Return last - first
                End Select
            End Function

            Private Shared Function DirtyCouponPrice(coupons As Integer, fraction As Double, amount As Double, redemption As Double, yieldPerPeriod As Double) As Double
                If coupons = 1 Then Return (redemption + amount) / (1.0 + fraction * yieldPerPeriod)
                Dim factor = 1.0 + yieldPerPeriod
                If factor = 1.0 Then Return coupons * amount + redemption
                Dim firstDiscount = Math.Pow(factor, -fraction)
                Dim lastDiscount = Math.Pow(factor, -(coupons - 1.0))
                ' Repeated discounting avoids subtracting nearly equal powers when the yield is near zero.
                If Math.Abs(yieldPerPeriod) * coupons < 0.0001 Then
                    Dim discount = firstDiscount
                    Dim sum = 0.0
                    For payment = 1 To coupons
                        sum += discount
                        discount /= factor
                    Next
                    Return amount * sum + redemption * firstDiscount * lastDiscount
                End If
                Dim couponSeries = (factor - lastDiscount) / yieldPerPeriod
                Return firstDiscount * (amount * couponSeries + redemption * lastDiscount)
            End Function

            Private Shared Function CouponDuration(coupons As Integer, fraction As Double, amount As Double, yieldPerPeriod As Double) As Double
                Dim factor = 1.0 + yieldPerPeriod
                Dim present = 0.0
                Dim weighted = 0.0
                Dim discount = 1.0
                ' The first discount factor cancels from the weighted price fraction.
                For payment = 0 To coupons - 1
                    Dim cash = (amount + If(payment = coupons - 1, 100.0, 0.0)) * discount
                    present += cash
                    weighted += cash * (fraction + payment)
                    discount /= factor
                Next
                If present = 0.0 AndAlso amount = 0.0 Then Return fraction + coupons - 1.0
                Return weighted / present
            End Function

            Private Shared Function CalculateCouponYield(coupons As Integer, fraction As Double, amount As Double, redemption As Double,
                                                         price As Double, frequency As Double) As Object
                If fraction <= 0.0 AndAlso coupons = 1 Then Return New FormulaError("#DIV/0!")
                If coupons = 1 Then Return FiniteNumber(((redemption + amount) / price - 1.0) / fraction * frequency)
                Return InvertCouponPrice(price, frequency, Function(yield) DirtyCouponPrice(coupons, fraction, amount, redemption, yield))
            End Function

            Private Shared Function InvertCouponPrice(price As Double, frequency As Double, presentValue As Func(Of Double, Double)) As Object
                Dim lower = -0.999999999999999
                Dim upper = 1.0
                While presentValue(upper) > price
                    upper *= 2
                    If Double.IsInfinity(upper) Then Return New FormulaError("#NUM!")
                End While
                For index = 0 To 79
                    Dim middle = lower + (upper - lower) / 2.0
                    If middle = lower OrElse middle = upper Then Exit For
                    Dim candidate = presentValue(middle)
                    If candidate > price OrElse Double.IsNaN(candidate) Then
                        lower = middle
                    Else
                        upper = middle
                    End If
                Next
                Return FiniteNumber((lower + (upper - lower) / 2.0) * frequency)
            End Function

        End Class

    End Class

End Namespace
