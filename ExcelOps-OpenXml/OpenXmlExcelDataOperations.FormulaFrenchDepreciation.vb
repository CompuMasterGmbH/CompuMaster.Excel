Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateFrenchDepreciation(declining As Boolean, arguments As IReadOnlyList(Of FormulaExpression),
                                                        formulaCell As ExcelCell) As Object
                If arguments.Count < 6 OrElse arguments.Count > 7 Then Return New FormulaError("#VALUE!")
                Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) FrenchDepreciationValue(declining, values))
            End Function

            Private Function FrenchDepreciationValue(declining As Boolean, values As Object()) As Object
                If values.Any(Function(value) TypeOf value Is Boolean) Then Return New FormulaError("#VALUE!")
                Dim purchased = ReadCouponDate(values(1))
                If TypeOf purchased Is FormulaError Then Return purchased
                Dim firstPeriod = ReadCouponDate(values(2))
                If TypeOf firstPeriod Is FormulaError Then Return firstPeriod
                Dim startDate = DirectCast(purchased, Integer)
                Dim endDate = DirectCast(firstPeriod, Integer)
                Dim numbers(4) As Double
                Dim argumentIndexes As Integer() = {0, 3, 4, 5, 6}
                For i = 0 To argumentIndexes.Length - 1
                    Dim input = ReadFinancialNumber(values.ElementAtOrDefault(argumentIndexes(i)))
                    If TypeOf input Is FormulaError Then Return input
                    numbers(i) = DirectCast(input, Double)
                Next
                Dim cost = numbers(0)
                Dim salvage = numbers(1)
                Dim rawPeriod = numbers(2)
                Dim rate = numbers(3)
                Dim basis = Math.Truncate(numbers(4))
                If cost <= 0 OrElse salvage < 0 OrElse salvage > cost OrElse rawPeriod < 0 OrElse rate <= 0 OrElse
                   numbers(4) < 0 OrElse basis = 2 OrElse basis > 4 OrElse startDate > endDate Then Return New FormulaError("#NUM!")
                Dim fraction = FrenchFirstPeriodFraction(startDate, endDate, CInt(basis))
                Dim period = Math.Truncate(rawPeriod)
                Dim maximum = cost - salvage
                Dim life = 1.0 / rate
                If Not declining Then
                    Dim annual = cost * rate
                    Dim initial = Math.Min(maximum, annual * fraction)
                    If rawPeriod = 0.0 Then Return FiniteNumber(initial)
                    If period = 0.0 Then Return FiniteNumber(annual)
                    If rawPeriod > Math.Ceiling(life) Then Return 0.0
                    Return FiniteNumber(Math.Max(0.0, Math.Min(annual, maximum - initial - (period - 1.0) * annual)))
                End If

                If life <= 2.0 Then Return New FormulaError("#NUM!")
                If rawPeriod > 0.0 AndAlso period = 0.0 Then
                    Throw New NotSupportedException("AMORDEGRC returns inconsistent results in Excel for periods strictly between zero and one.")
                End If
                Dim effectiveRate = rate * If(life <= 4.0, 1.5, If(life <= 6.0, 2.0, 2.5))
                Dim first = Math.Round(Math.Min(maximum, (cost * effectiveRate) * fraction), MidpointRounding.AwayFromZero)
                If period = 0.0 Then Return FiniteNumber(first)
                Dim transition = Math.Ceiling(life) - 2.0
                If period > transition + 1.0 Then Return 0.0
                Dim opening = (cost - first) * CompoundFactor(1.0 - effectiveRate, Math.Min(period, transition) - 1.0)
                If period > transition Then opening /= 2.0
                If opening <= salvage Then Return 0.0
                Dim multiplier = If(period < transition, effectiveRate, If(period = transition, 0.5, 1.0))
                Return FiniteNumber(Math.Round(opening * multiplier, MidpointRounding.AwayFromZero))
            End Function

            Private Function FrenchFirstPeriodFraction(startDate As Integer, endDate As Integer, basis As Integer) As Double
                If startDate = endDate Then Return 1.0
                If basis = 4 Then Return Days360(startDate, endDate, european:=True) / 360.0
                Dim start = DateParts(startDate)
                Dim finish = DateParts(endDate)
                If basis = 0 Then Return YearFractionUsDays(start, finish) / 360.0
                If start.Year <> 1900 AndAlso start.Month = 2 AndAlso start.Day = 29 Then startDate -= 1
                If finish.Year <> 1900 AndAlso finish.Month = 2 AndAlso finish.Day = 29 Then endDate -= 1
                If startDate = endDate Then Return 1.0
                Dim daysInYear = If(basis = 1 AndAlso DateTime.IsLeapYear(start.Year), 366.0, 365.0)
                Return (endDate - startDate) / daysInYear
            End Function

        End Class

    End Class

End Namespace
