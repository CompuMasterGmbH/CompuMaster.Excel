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
