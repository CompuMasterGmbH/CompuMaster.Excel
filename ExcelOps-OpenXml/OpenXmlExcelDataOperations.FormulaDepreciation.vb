Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateDepreciationOrRate(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                         formulaCell As ExcelCell) As Object
                Dim declining = name = "DB" OrElse name = "DDB"
                Dim addIn = name = "EFFECT" OrElse name = "NOMINAL" OrElse name = "DOLLARDE" OrElse name = "DOLLARFR"
                Dim minimum = If(addIn, 2, If(name = "SLN", 3, 4))
                If arguments.Count < minimum OrElse arguments.Count > If(declining, 5, minimum) Then Return New FormulaError("#VALUE!")

                Dim evaluated = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                If addIn Then
                    If evaluated.Any(Function(value) TypeOf value Is FormulaArea AndAlso IsMultiValue(value)) Then Return New FormulaError("#VALUE!")
                    If evaluated.Any(Function(value) value Is Nothing) Then Return New FormulaError("#N/A")
                End If
                Dim inputs = evaluated.Select(Function(value) AsMatrix(value)).ToArray()
                Dim rows = inputs.Max(Function(input) input.Rows)
                Dim columns = inputs.Max(Function(input) input.Columns)
                Dim results = AllocateMatrixValues(rows, columns)
                Dim numbers(4) As Double
                If declining AndAlso arguments.Count = 4 Then numbers(4) = If(name = "DB", 12.0, 2.0)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim failure As FormulaError = Nothing
                        For index = 0 To inputs.Length - 1
                            Dim input = inputs(index).BroadcastAt(row, column)
                            Dim number = If(addIn AndAlso TypeOf input Is Boolean, New FormulaError("#VALUE!"), ReadFinancialNumber(input))
                            If TypeOf number Is FormulaError Then
                                failure = DirectCast(number, FormulaError)
                                Exit For
                            End If
                            numbers(index) = DirectCast(number, Double)
                        Next
                        results(row * columns + column) = If(failure, DepreciationOrRateValue(name, numbers))
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Shared Function DepreciationOrRateValue(name As String, values As Double()) As Object
                Dim cost = values(0)
                Dim salvage = values(1)
                Dim life = values(2)
                Dim period = values(3)
                Select Case name
                    Case "SLN"
                        If life = 0.0 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber((cost - salvage) / life)
                    Case "SYD"
                        If salvage < 0.0 OrElse life <= 0.0 OrElse period <= 0.0 OrElse period > life Then Return New FormulaError("#NUM!")
                        Return FiniteNumber((cost - salvage) * 2.0 * (life - period + 1.0) / (life * (life + 1.0)))
                    Case "DB", "DDB"
                        If cost < 0.0 OrElse salvage < 0.0 OrElse life <= 0.0 OrElse period <= 0.0 Then Return New FormulaError("#NUM!")
                        If name = "DB" Then Return FixedDecliningDepreciation(cost, salvage, life, period, values(4))
                        Return DoubleDecliningDepreciation(cost, salvage, life, period, values(4))
                    Case "EFFECT", "NOMINAL"
                        Dim compounds = Math.Truncate(salvage)
                        If cost <= 0.0 OrElse compounds < 1.0 Then Return New FormulaError("#NUM!")
                        If name = "EFFECT" Then Return FiniteNumber(CompoundFactor(1.0 + cost / compounds, compounds) - 1.0)
                        Return FiniteNumber(compounds * (Math.Pow(1.0 + cost, 1.0 / compounds) - 1.0))
                    Case Else
                        Return ConvertDollarFraction(cost, salvage, name = "DOLLARFR")
                End Select
            End Function

            Private Shared Function FixedDecliningDepreciation(cost As Double, salvage As Double, life As Double,
                                                                period As Double, rawMonths As Double) As Object
                Dim months = Math.Truncate(rawMonths)
                Dim year = Math.Truncate(period)
                Dim fullLife = Math.Truncate(life)
                Dim finalPartial = year > fullLife
                If months < 1.0 OrElse months > 12.0 OrElse year > fullLife + If(months < 12.0, 1.0, 0.0) Then
                    Return New FormulaError("#NUM!")
                End If
                If cost = 0.0 Then Return 0.0
                Dim rate = Math.Round(1.0 - Math.Pow(salvage / cost, 1.0 / life), 3, MidpointRounding.AwayFromZero)
                Dim first = cost * rate * months / 12.0
                If year <= 1.0 Then Return FiniteNumber(first)
                Dim opening = (cost - first) * CompoundFactor(1.0 - rate, year - 2.0)
                Dim fraction = If(finalPartial, (12.0 - months) / 12.0, 1.0)
                Return FiniteNumber(opening * rate * fraction)
            End Function

            Private Shared Function DoubleDecliningDepreciation(cost As Double, salvage As Double, life As Double,
                                                                 period As Double, factor As Double) As Object
                If factor <= 0.0 OrElse period > life Then Return New FormulaError("#NUM!")
                Dim rate = Math.Min(1.0, factor / life)
                Dim opening = cost * Math.Pow(1.0 - rate, Math.Max(0.0, period - 1.0))
                Return FiniteNumber(Math.Max(0.0, Math.Min(opening * rate, opening - salvage)))
            End Function

            Private Function EvaluateVariableDecliningDepreciation(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 5 OrElse arguments.Count > 7 Then Return New FormulaError("#VALUE!")
                Return BroadcastReferenceArguments(arguments, formulaCell, AddressOf VariableDecliningDepreciationValue)
            End Function

            Private Function VariableDecliningDepreciationValue(values As Object()) As Object
                Dim inputs(5) As Double
                inputs(5) = 2.0
                For index = 0 To Math.Min(5, values.Length - 1)
                    Dim value = ReadFinancialNumber(values(index))
                    If TypeOf value Is FormulaError Then Return value
                    inputs(index) = DirectCast(value, Double)
                Next

                Dim noSwitch = False
                If values.Length = 7 Then
                    If TypeOf values(6) Is String Then Return New FormulaError("#VALUE!")
                    noSwitch = DirectCast(NumberValue(values(6)), Double) <> 0.0
                End If
                Return VariableDecliningDepreciation(inputs(0), inputs(1), inputs(2), inputs(3), inputs(4), inputs(5), noSwitch)
            End Function

            Private Shared Function VariableDecliningDepreciation(cost As Double, salvage As Double, life As Double, startPeriod As Double,
                                                                  endPeriod As Double, factor As Double, noSwitch As Boolean) As Object
                If cost < 0.0 OrElse salvage < 0.0 OrElse life < 0.0 OrElse factor < 0.0 OrElse
                   startPeriod < 0.0 OrElse endPeriod < startPeriod OrElse endPeriod > life Then Return New FormulaError("#NUM!")
                If life = 0.0 Then Return New FormulaError("#DIV/0!")
                If startPeriod = endPeriod Then
                    If Not noSwitch AndAlso startPeriod = life AndAlso
                       ((salvage = 0.0 AndAlso (cost = 0.0 OrElse factor >= life)) OrElse (cost <= salvage AndAlso factor = 0.0)) Then
                        Return New FormulaError("#DIV/0!")
                    End If
                    Return 0.0
                End If

                Dim duration = endPeriod - startPeriod
                Dim depreciable = cost - salvage
                If depreciable < 0.0 Then Return If(noSwitch OrElse startPeriod >= 1.0, 0.0, depreciable * Math.Min(1.0, duration))
                If Not noSwitch AndAlso life < 1.0 Then Return FiniteNumber(depreciable * duration / life)

                Dim rate = Math.Min(1.0, factor / life)
                If noSwitch Then
                    Return FiniteNumber(DecliningDepreciationTotal(cost, salvage, endPeriod, rate) -
                                        DecliningDepreciationTotal(cost, salvage, startPeriod, rate))
                End If

                Dim line = depreciable / life
                Dim firstDecline = Math.Min(cost * rate, depreciable)
                If line > firstDecline OrElse (life = 1.0 AndAlso factor = 1.0 AndAlso salvage = 0.0) Then Return FiniteNumber(line * duration)

                ' Excel can apply the fractional part of the starting period before the whole periods.
                Dim fraction = startPeriod - Math.Floor(startPeriod)
                Dim offset = startPeriod
                If fraction > 0.0 AndAlso (factor > 1.0 OrElse Math.Min(depreciable, line) <= firstDecline) Then offset = fraction
                Dim opening = cost - DepreciationWithSwitch(cost, salvage, life, offset, rate, openingPeriod:=True)
                Return FiniteNumber(DepreciationWithSwitch(opening, salvage, life - offset, endPeriod - offset, rate, openingPeriod:=False) -
                                    DepreciationWithSwitch(opening, salvage, life - offset, startPeriod - offset, rate, openingPeriod:=False))
            End Function

            Private Shared Function DecliningDepreciationTotal(cost As Double, salvage As Double, period As Double, rate As Double) As Double
                Dim wholePeriods = Math.Floor(period)
                Dim book = Math.Max(salvage, cost * Math.Pow(1.0 - rate, wholePeriods))
                Dim fractionalDepreciation = Math.Min(book * rate, book - salvage) * (period - wholePeriods)
                Return cost - book + fractionalDepreciation
            End Function

            Private Shared Function DepreciationWithSwitch(cost As Double, salvage As Double, life As Double, period As Double,
                                                           rate As Double, openingPeriod As Boolean) As Double
                If period >= 1048576.0 Then Return DepreciationWithSwitchOverLongInterval(cost, salvage, life, period, rate, openingPeriod)

                Dim book = cost
                Dim elapsed = 0.0
                ' Before this point the remaining life is too long for straight-line depreciation to exceed the declining rate.
                Dim guaranteedDecliningPeriods = Math.Floor(Math.Min(period, life - 1.0 / rate))
                If guaranteedDecliningPeriods > 1024.0 Then
                    elapsed = guaranteedDecliningPeriods
                    book = Math.Max(salvage, cost * Math.Pow(1.0 - rate, elapsed))
                End If
                While elapsed < period AndAlso book > salvage
                    Dim remaining = book - salvage
                    Dim declining = Math.Min(book * rate, remaining)
                    Dim span = Math.Min(1.0, period - elapsed)
                    If Not openingPeriod OrElse span = 1.0 Then
                        Dim straight = remaining / (life - elapsed)
                        If Math.Min(remaining, straight) > declining Then
                            Return cost - book + straight * (period - elapsed)
                        End If
                        book -= declining * span
                    ElseIf elapsed = 0.0 Then
                        book -= declining * span
                    Else
                        book -= Math.Min(book * rate * span, remaining)
                    End If
                    elapsed += 1.0
                End While
                Return cost - book
            End Function

            Private Shared Function DepreciationWithSwitchOverLongInterval(cost As Double, salvage As Double, life As Double,
                                                                           period As Double, rate As Double, openingPeriod As Boolean) As Double
                Dim wholePeriods = Math.Floor(period)
                Dim last = If(openingPeriod, wholePeriods, Math.Ceiling(period)) - 1.0
                Dim retention = 1.0 - rate
                If rate > 0.0 AndAlso retention > 0.0 AndAlso cost > salvage Then
                    ' The straight-line advantage increases while the next declining period stays above salvage.
                    If salvage > 0.0 AndAlso retention < 1.0 Then
                        last = Math.Min(last, Math.Floor((Math.Log(salvage) - Math.Log(cost)) / Math.Log(retention) - 1.0))
                    End If
                    If last >= 0.0 AndAlso StraightLineDepreciationWins(cost, salvage, life, rate, last) Then
                        Dim first = 0.0
                        While first < last
                            Dim middle = Math.Floor(first + (last - first) / 2.0)
                            If middle <= first OrElse middle >= last Then Exit While
                            If StraightLineDepreciationWins(cost, salvage, life, rate, middle) Then
                                last = middle
                            Else
                                first = middle + 1.0
                            End If
                        End While
                        Dim switchPeriod = If(StraightLineDepreciationWins(cost, salvage, life, rate, first), first, last)
                        Dim book = cost * Math.Pow(retention, switchPeriod)
                        Return cost - book + (book - salvage) / (life - switchPeriod) * (period - switchPeriod)
                    End If
                End If

                If Not openingPeriod Then Return DecliningDepreciationTotal(cost, salvage, period, rate)
                Dim opening = Math.Max(salvage, cost * Math.Pow(retention, wholePeriods))
                Return cost - opening + Math.Min(opening * rate * (period - wholePeriods), opening - salvage)
            End Function

            Private Shared Function StraightLineDepreciationWins(cost As Double, salvage As Double, life As Double,
                                                                 rate As Double, elapsed As Double) As Boolean
                Dim book = cost * Math.Pow(1.0 - rate, elapsed)
                Dim remaining = book - salvage
                Return Math.Min(remaining, remaining / (life - elapsed)) > book * rate
            End Function

            Private Shared Function ConvertDollarFraction(amount As Double, rawDenominator As Double, fractional As Boolean) As Object
                If rawDenominator < 0.0 Then Return New FormulaError("#NUM!")
                Dim denominator = Math.Truncate(rawDenominator)
                If denominator = 0.0 Then Return New FormulaError("#DIV/0!")
                Dim scale = Math.Pow(10.0, Math.Ceiling(Math.Log10(denominator)))
                Dim whole = Math.Truncate(amount)
                Dim multiplier = If(fractional, denominator / scale, scale / denominator)
                Return FiniteNumber(whole + (amount - whole) * multiplier)
            End Function

        End Class

    End Class

End Namespace
