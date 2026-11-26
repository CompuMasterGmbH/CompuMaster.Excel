Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateDiscountedSecurity(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                        formulaCell As ExcelCell) As Object
                Dim treasury = name.StartsWith("TBILL", StringComparison.Ordinal)
                Dim minimum = If(treasury, 3, 4)
                If arguments.Count < minimum OrElse arguments.Count > If(treasury, minimum, minimum + 1) Then Return New FormulaError("#VALUE!")
                Return BroadcastReferenceArguments(arguments, formulaCell,
                    Function(values) DiscountedSecurityValue(name, treasury, values), checkErrors:=False)
            End Function

            Private Function DiscountedSecurityValue(name As String, treasury As Boolean, values As Object()) As Object
                Dim basis = 0.0
                If Not treasury AndAlso values.Length = 5 Then
                    If TypeOf values(4) Is Boolean Then Return New FormulaError("#VALUE!")
                    Dim mode = ReadFinancialNumber(values(4))
                    If TypeOf mode Is FormulaError Then Return mode
                    basis = Math.Truncate(DirectCast(mode, Double))
                    If basis < 0.0 OrElse basis > 4.0 Then Return New FormulaError("#NUM!")
                End If
                Dim settlement = ReadYearFractionDate(values(0))
                If TypeOf settlement Is FormulaError Then Return settlement
                Dim maturity = ReadYearFractionDate(values(1))
                If TypeOf maturity Is FormulaError Then Return maturity
                Dim amounts(1) As Double
                For index = 2 To If(treasury, 2, 3)
                    If TypeOf values(index) Is Boolean Then Return New FormulaError("#VALUE!")
                    Dim amount = ReadFinancialNumber(values(index))
                    If TypeOf amount Is FormulaError Then Return amount
                    amounts(index - 2) = DirectCast(amount, Double)
                Next

                Dim first = DirectCast(settlement, Integer)
                Dim last = DirectCast(maturity, Integer)
                If first >= last OrElse amounts(0) <= 0.0 Then Return New FormulaError("#NUM!")
                If treasury Then Return TreasuryBillValue(name, first, last, amounts(0))
                If amounts(1) <= 0.0 Then Return New FormulaError("#NUM!")
                If basis <> 2.0 AndAlso basis <> 3.0 AndAlso Not IsCalendarSerialInRange(last) Then Return New FormulaError("#NUM!")
                Dim fraction = SecurityYearFraction(first, last, CInt(basis))
                If fraction <= 0.0 Then Return New FormulaError("#NUM!")
                Select Case name
                    Case "RECEIVED"
                        Dim discount = 1.0 - amounts(1) * fraction
                        If discount <= 0.0 Then Return New FormulaError("#NUM!")
                        Return FiniteNumber(amounts(0) / discount)
                    Case "PRICEDISC"
                        Return FiniteNumber(amounts(1) * (1.0 - amounts(0) * fraction))
                    Case Else
                        Dim baseAmount = If(name = "DISC", amounts(1), amounts(0))
                        Return FiniteNumber((amounts(1) - amounts(0)) / baseAmount / fraction)
                End Select
            End Function

            Private Function SecurityYearFraction(first As Integer, last As Integer, basis As Integer) As Double
                Select Case basis
                    Case 0
                        Return YearFractionUsDays(DateParts(first), DateParts(last)) / 360.0
                    Case 1
                        Return (last - first) / ActualYearLength(DateParts(first), DateParts(last))
                    Case 2
                        Return (last - first) / 360.0
                    Case 3
                        Return (last - first) / 365.0
                    Case Else
                        Return Days360(first, last, european:=True) / 360.0
                End Select
            End Function

            Private Function TreasuryBillValue(name As String, settlement As Integer, maturity As Integer, input As Double) As Object
                Dim days = maturity - settlement
                If days > 366 OrElse (days = 366 AndAlso (Not IsCalendarSerialInRange(maturity) OrElse
                    Not BillTermIncludesLeapDay(settlement, maturity))) Then Return New FormulaError("#NUM!")
                If name = "TBILLYIELD" Then Return FiniteNumber((100.0 - input) / input * (360.0 / days))

                Dim unroundedDiscount = input * days / 360.0
                Dim price = 1.0 - unroundedDiscount
                If price < 0.0 OrElse (price = 0.0 AndAlso name = "TBILLEQ") Then Return New FormulaError("#NUM!")
                If name = "TBILLPRICE" Then Return FiniteNumber(100.0 * price)
                If days <= 182 Then Return FiniteNumber(365.0 * input / (360.0 - input * days))

                ' Solve the bond-equivalent yield with one semiannual compounding period.
                Dim yearFraction = days / If(days = 366, 366.0, 365.0)
                Dim gain = unroundedDiscount / price
                Return FiniteNumber(2.0 * gain / (yearFraction + Math.Sqrt(yearFraction * yearFraction + 2.0 * (yearFraction - 0.5) * gain)))
            End Function

            Private Function BillTermIncludesLeapDay(settlement As Integer, maturity As Integer) As Boolean
                For candidateYear As Integer = DateParts(settlement).Year To DateParts(maturity).Year
                    If Not DateTime.IsLeapYear(candidateYear) Then Continue For
                    Dim serial = ToExcelDateSerial(New DateTime(candidateYear, 2, 29))
                    If serial >= settlement AndAlso serial <= maturity Then Return True
                Next
                Return False
            End Function

        End Class

    End Class

End Namespace
