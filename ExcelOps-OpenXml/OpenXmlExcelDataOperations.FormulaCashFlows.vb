Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateCashFlowFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                Dim netPresentValue = name = "NPV"
                Dim modifiedReturn = name = "MIRR"
                Dim minimum = If(netPresentValue, 2, If(modifiedReturn, 3, 1))
                Dim maximum = If(netPresentValue, 255, If(modifiedReturn, 3, 2))
                If arguments.Count < minimum OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")

                Dim inputs = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                If Not netPresentValue AndAlso TypeOf inputs(0) Is FormulaError Then Return inputs(0)
                Dim rateArgument = If(netPresentValue, 0, 1)
                Dim firstRate = AsMatrix(If(inputs.Length > rateArgument, inputs(rateArgument), 0.1))
                Dim secondRate = AsMatrix(If(modifiedReturn, inputs(2), 0.0))
                Dim cashFlows As New List(Of Double)
                Dim cashFlowError As FormulaError = Nothing
                Dim firstFlow = If(netPresentValue, 1, 0)
                Dim lastFlow = If(netPresentValue, inputs.Length - 1, 0)
                For index = firstFlow To lastFlow
                    cashFlowError = AppendCashFlows(inputs(index), netPresentValue, name = "IRR", cashFlows)
                    If cashFlowError IsNot Nothing Then Exit For
                Next

                Dim rows = Math.Max(firstRate.Rows, secondRate.Rows)
                Dim columns = Math.Max(firstRate.Columns, secondRate.Columns)
                Dim output = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim rate = ReadFinancialNumber(firstRate.BroadcastAt(row, column))
                        Dim reinvestment = ReadFinancialNumber(secondRate.BroadcastAt(row, column))
                        Dim result As Object
                        If TypeOf rate Is FormulaError Then
                            result = rate
                        ElseIf TypeOf reinvestment Is FormulaError Then
                            result = reinvestment
                        ElseIf cashFlowError IsNot Nothing Then
                            result = cashFlowError
                        Else
                            result = CalculateCashFlow(name, cashFlows, DirectCast(rate, Double), DirectCast(reinvestment, Double))
                        End If
                        output(row * columns + column) = result
                    Next
                Next
                If output.Length = 1 Then Return output(0)
                Return New FormulaMatrix(rows, columns, output)
            End Function

            Private Function AppendCashFlows(input As Object, coerceScalar As Boolean, replaceCollectionErrors As Boolean,
                                             payments As List(Of Double)) As FormulaError
                If TypeOf input Is FormulaError Then Return DirectCast(input, FormulaError)
                If TypeOf input Is FormulaArea OrElse TypeOf input Is FormulaMatrix Then
                    For Each value In AsMatrix(input).Values
                        If TypeOf value Is FormulaError Then
                            Return If(replaceCollectionErrors, New FormulaError("#VALUE!"), DirectCast(value, FormulaError))
                        End If
                        If TypeOf value Is Double Then payments.Add(DirectCast(value, Double))
                    Next
                ElseIf coerceScalar Then
                    Dim number = ReadFinancialNumber(input)
                    If TypeOf number Is FormulaError Then Return DirectCast(number, FormulaError)
                    payments.Add(DirectCast(number, Double))
                ElseIf TypeOf input Is Double Then
                    payments.Add(DirectCast(input, Double))
                End If
                Return Nothing
            End Function

            Private Function ReadFinancialNumber(value As Object) As Object
                Return If(TypeOf value Is String, ConvertWorksheetText("VALUE", value), NumberValue(value))
            End Function

            Private Shared Function CalculateCashFlow(name As String, payments As List(Of Double), rate As Double,
                                                      reinvestment As Double) As Object
                If name = "NPV" Then
                    If rate = -1.0 Then Return New FormulaError("#DIV/0!")
                    Dim discount = 1.0 / (1.0 + rate)
                    Dim factor = discount
                    Dim value = 0.0
                    For Each payment In payments
                        value += payment * factor
                        factor *= discount
                    Next
                    Return FiniteNumber(value)
                End If
                If name = "IRR" Then
                    If rate <= -1.0 Then Return New FormulaError("#VALUE!")
                    If Not payments.Any(Function(payment) payment < 0.0) OrElse Not payments.Any(Function(payment) payment > 0.0) Then
                        Return New FormulaError("#NUM!")
                    End If
                    Return InternalRateOfReturn(payments, rate)
                End If
                Return ModifiedRateOfReturn(payments, rate, reinvestment)
            End Function

            Private Shared Function InternalRateOfReturn(payments As List(Of Double), guess As Double) As Object
                Dim scale = payments.Max(Function(payment) Math.Abs(payment))
                Dim coefficients = payments.Select(Function(payment) payment / scale).ToArray()
                Dim discount = 1.0 / (1.0 + guess)
                ' Solve the cash-flow polynomial in the discount factor to keep the rate above -100%.
                For iteration = 1 To 20
                    Dim value = coefficients(coefficients.Length - 1)
                    Dim derivative = 0.0
                    For index = coefficients.Length - 2 To 0 Step -1
                        derivative = derivative * discount + value
                        value = value * discount + coefficients(index)
                    Next
                    If value = 0.0 Then Return FiniteNumber(1.0 / discount - 1.0)
                    Dim updated = discount - value / derivative
                    If Double.IsNaN(updated) OrElse Double.IsInfinity(updated) OrElse updated <= 0.0 Then
                        discount /= 2.0
                        Continue For
                    End If
                    If Math.Abs(updated - discount) <= 1.0E-14 * Math.Max(1.0, discount) Then
                        Return FiniteNumber(1.0 / updated - 1.0)
                    End If
                    discount = updated
                Next
                Return New FormulaError("#NUM!")
            End Function

            Private Shared Function ModifiedRateOfReturn(payments As List(Of Double), finance As Double, reinvestment As Double) As Object
                If payments.Count < 2 OrElse reinvestment = -1.0 Then Return New FormulaError("#DIV/0!")
                Dim future = 0.0
                Dim present = 0.0
                Dim periods = payments.Count - 1
                For index = 0 To periods
                    Dim payment = payments(index)
                    If payment > 0.0 Then
                        future += payment * CompoundFactor(1.0 + reinvestment, periods - index)
                    ElseIf payment < 0.0 Then
                        Dim discount = CompoundFactor(1.0 + finance, index)
                        If discount = 0.0 Then Return New FormulaError("#DIV/0!")
                        present -= payment / discount
                    End If
                Next
                If present = 0.0 Then Return New FormulaError("#DIV/0!")
                Dim ratio = future / present
                If ratio < 0.0 AndAlso periods Mod 2 = 0 Then Return New FormulaError("#NUM!")
                Dim root = Math.Pow(Math.Abs(ratio), 1.0 / periods)
                Return FiniteNumber(If(ratio < 0.0, -root, root) - 1.0)
            End Function

        End Class

    End Class

End Namespace
