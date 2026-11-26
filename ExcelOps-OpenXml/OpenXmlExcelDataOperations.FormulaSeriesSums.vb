Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateSeriesSum(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count <> 4 Then Return New FormulaError("#VALUE!")
                Dim inputs = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                If inputs.Take(3).Any(Function(input) TypeOf input Is FormulaArea AndAlso IsMultiValue(input)) Then
                    Return New FormulaError("#VALUE!")
                End If
                If inputs.Any(Function(input) input Is Nothing) Then Return New FormulaError("#N/A")

                Dim parameters = inputs.Take(3).Select(Function(input) AsMatrix(input)).ToArray()
                Dim coefficients = AsMatrix(inputs(3))
                Dim terms(coefficients.Values.Length - 1) As Double
                Dim coefficientError As FormulaError = Nothing
                For index = 0 To terms.Length - 1
                    Dim raw = coefficients.Values(index)
                    Dim number As Object
                    If terms.Length = 1 Then
                        number = SeriesNumber(raw)
                    Else
                        number = If(raw, 0.0)
                        If Not TypeOf number Is Double AndAlso Not TypeOf number Is FormulaError Then number = New FormulaError("#VALUE!")
                    End If
                    If TypeOf number Is FormulaError Then
                        coefficientError = DirectCast(number, FormulaError)
                        Exit For
                    End If
                    terms(index) = DirectCast(number, Double)
                Next

                Dim rows = parameters.Max(Function(matrix) matrix.Rows)
                Dim columns = parameters.Max(Function(matrix) matrix.Columns)
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim baseValue = SeriesNumber(parameters(0).BroadcastAt(row, column))
                        Dim first = SeriesNumber(parameters(1).BroadcastAt(row, column))
                        Dim increment = SeriesNumber(parameters(2).BroadcastAt(row, column))
                        Dim result As Object
                        If TypeOf baseValue Is FormulaError Then
                            result = baseValue
                        ElseIf TypeOf first Is FormulaError Then
                            result = first
                        ElseIf TypeOf increment Is FormulaError Then
                            result = increment
                        ElseIf coefficientError IsNot Nothing Then
                            result = coefficientError
                        Else
                            result = SumPowerSeries(DirectCast(baseValue, Double), DirectCast(first, Double), DirectCast(increment, Double), terms)
                        End If
                        results(row * columns + column) = result
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function SeriesNumber(value As Object) As Object
                If TypeOf value Is Boolean Then Return New FormulaError("#VALUE!")
                Return If(TypeOf value Is String, ConvertWorksheetText("VALUE", value), NumberValue(value))
            End Function

            Private Shared Function SumPowerSeries(baseValue As Double, first As Double, increment As Double, terms As Double()) As Object
                Dim total = 0.0
                For index = 0 To terms.Length - 1
                    Dim exponent = first + index * increment
                    If baseValue = 0.0 AndAlso exponent <= 0.0 Then Return New FormulaError("#NUM!")
                    ' Excel contributes zero for a negative base with a fractional exponent.
                    If baseValue < 0.0 AndAlso Math.Truncate(exponent) <> exponent Then Continue For
                    total += terms(index) * Math.Pow(baseValue, exponent)
                Next
                Return FiniteNumber(total)
            End Function

            Private Function EvaluatePairwiseSquares(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                Dim first = EvaluateExpression(arguments(0), formulaCell)
                Dim second = EvaluateExpression(arguments(1), formulaCell)
                If first Is Nothing OrElse second Is Nothing Then Return New FormulaError("#VALUE!")
                If TypeOf first Is FormulaError Then Return first
                If TypeOf second Is FormulaError Then Return second
                Dim left = AsMatrix(first).Values
                Dim right = AsMatrix(second).Values
                If (left.Length = 1 AndAlso left(0) Is Nothing) OrElse (right.Length = 1 AndAlso right(0) Is Nothing) Then
                    Return New FormulaError("#VALUE!")
                End If
                If left.Length <> right.Length Then Return New FormulaError("#N/A")
                Dim total = 0.0
                Dim matched = False
                For index = 0 To left.Length - 1
                    Dim xValue = left(index)
                    Dim yValue = right(index)
                    If TypeOf xValue Is FormulaError Then Return xValue
                    If TypeOf yValue Is FormulaError Then Return yValue
                    If Not TypeOf xValue Is Double OrElse Not TypeOf yValue Is Double Then Continue For
                    Dim x = DirectCast(xValue, Double)
                    Dim y = DirectCast(yValue, Double)
                    matched = True
                    Select Case name
                        Case "SUMX2MY2" : total += x * x - y * y
                        Case "SUMX2PY2" : total += x * x + y * y
                        Case Else : total += (x - y) * (x - y)
                    End Select
                Next
                If Not matched Then Return New FormulaError("#DIV/0!")
                Return FiniteNumber(total)
            End Function

        End Class

    End Class

End Namespace
