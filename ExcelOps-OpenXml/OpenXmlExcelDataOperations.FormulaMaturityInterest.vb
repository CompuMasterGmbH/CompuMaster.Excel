Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateMaturityInterest(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                       formulaCell As ExcelCell) As Object
                Dim accruedOnly = name = "ACCRINTM"
                If arguments.Count < If(accruedOnly, 3, 5) OrElse arguments.Count > If(accruedOnly, 5, 6) Then
                    Return New FormulaError("#VALUE!")
                End If
                Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) MaturityInterestValue(name, accruedOnly, values))
            End Function

            Private Function MaturityInterestValue(name As String, accruedOnly As Boolean, values As Object()) As Object
                If values.Any(Function(value) TypeOf value Is Boolean) Then Return New FormulaError("#VALUE!")
                Dim issue = ReadCouponDate(values(If(accruedOnly, 0, 2)))
                If TypeOf issue Is FormulaError Then Return issue
                Dim settlement = ReadCouponDate(values(If(accruedOnly, 1, 0)))
                If TypeOf settlement Is FormulaError Then Return settlement
                Dim first = DirectCast(issue, Integer)
                Dim settle = DirectCast(settlement, Integer)
                If first > settle OrElse (first = settle AndAlso Not accruedOnly) Then Return New FormulaError("#NUM!")
                Dim basisInput = If(values.Length = If(accruedOnly, 5, 6), values(values.Length - 1), Nothing)
                Dim basisValue = ReadFinancialNumber(basisInput)
                If TypeOf basisValue Is FormulaError Then Return basisValue
                Dim basis = Math.Truncate(DirectCast(basisValue, Double))
                If DirectCast(basisValue, Double) < 0.0 OrElse basis > 4.0 Then Return New FormulaError("#NUM!")
                Dim interest = ReadFinancialNumber(values(If(accruedOnly, 2, 3)))
                If TypeOf interest Is FormulaError Then Return interest
                Dim rate = DirectCast(interest, Double)
                If rate < 0.0 Then Return New FormulaError("#NUM!")
                Dim yearLength = If(basis = 1, ActualYearLength(DateParts(first), DateParts(settle)), If(basis = 3, 365.0, 360.0))
                Dim accruedDays = CouponDaySpan(first, settle, CInt(basis))
                If accruedOnly Then
                    Dim par = ReadFinancialNumber(If(values.Length >= 4 AndAlso values(3) IsNot Nothing, values(3), 1000.0))
                    If TypeOf par Is FormulaError Then Return par
                    Dim face = DirectCast(par, Double)
                    If face <= 0.0 OrElse rate <= 0.0 Then Return New FormulaError("#NUM!")
                    Return FiniteNumber(face * rate * accruedDays / yearLength)
                End If

                Dim maturity = ReadCouponDate(values(1))
                If TypeOf maturity Is FormulaError Then Return maturity
                Dim last = DirectCast(maturity, Integer)
                If last <= settle Then Return New FormulaError("#NUM!")
                Dim priceOrYield = ReadFinancialNumber(values(4))
                If TypeOf priceOrYield Is FormulaError Then Return priceOrYield
                Dim level = DirectCast(priceOrYield, Double)
                If If(name = "PRICEMAT", level < 0.0, level <= 0.0) Then Return New FormulaError("#NUM!")
                Dim totalDays = CouponDaySpan(first, last, CInt(basis))
                Dim remaining = (totalDays - accruedDays) / yearLength
                Dim accrued = 100.0 * rate * accruedDays / yearLength
                Dim payoff = 100.0 * (1.0 + rate * totalDays / yearLength)
                If name = "PRICEMAT" Then Return FiniteNumber(payoff / (1.0 + level * remaining) - accrued)
                If remaining = 0.0 Then Return New FormulaError("#DIV/0!")
                Return FiniteNumber((payoff / (level + accrued) - 1.0) / remaining)
            End Function

        End Class

    End Class

End Namespace
