Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateGrowthSchedule(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                Dim principalInput = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf principalInput Is FormulaArea AndAlso IsMultiValue(principalInput) Then Return New FormulaError("#VALUE!")
                Dim principals = AsMatrix(principalInput)
                Dim schedule = EvaluateExpression(arguments(1), formulaCell)
                Dim factor = ReadGrowthSchedule(schedule)
                Dim results = AllocateMatrixValues(principals.Rows, principals.Columns)
                For index = 0 To results.Length - 1
                    Dim principal = principals.Values(index)
                    Dim number = If(TypeOf principal Is Boolean, New FormulaError("#VALUE!"), ReadFinancialNumber(principal))
                    If TypeOf number Is FormulaError Then
                        results(index) = number
                    ElseIf TypeOf factor Is FormulaError Then
                        results(index) = factor
                    Else
                        results(index) = FiniteNumber(DirectCast(number, Double) * DirectCast(factor, Double))
                    End If
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(principals.Rows, principals.Columns, results)
            End Function

            Private Function ReadGrowthSchedule(input As Object) As Object
                If input Is Nothing Then Return New FormulaError("#N/A")
                If Not TypeOf input Is FormulaArea AndAlso Not TypeOf input Is FormulaMatrix Then
                    If TypeOf input Is Boolean Then Return New FormulaError("#VALUE!")
                    Dim rate = ReadFinancialNumber(input)
                    If TypeOf rate Is FormulaError Then Return rate
                    Return FiniteNumber(1.0 + DirectCast(rate, Double))
                End If
                Dim compound = 1.0
                For Each item In AsMatrix(input).Values
                    If TypeOf item Is FormulaError Then Return item
                    If item Is Nothing Then Continue For
                    If Not TypeOf item Is Double Then Return New FormulaError("#VALUE!")
                    compound *= 1.0 + DirectCast(item, Double)
                Next
                Return FiniteNumber(compound)
            End Function

            Private Shared Function ConstantGrowthRate(name As String, values As Double()) As Object
                Dim parameter = values(0)
                Dim present = values(1)
                Dim future = values(2)
                If parameter <= 0.0 Then Return New FormulaError("#NUM!")
                If name = "RRI" Then
                    If present = 0.0 Then Return New FormulaError("#NUM!")
                    Dim ratio = future / present
                    If ratio < 0.0 Then Return New FormulaError("#NUM!")
                    If ratio = 1.0 Then Return 0.0
                    Return FiniteNumber(Math.Exp(Math.Log(ratio) * (1.0 / parameter)) - 1.0)
                End If
                If present <= 0.0 OrElse future <= 0.0 Then Return New FormulaError("#NUM!")
                Dim logGrowth = Math.Log(1.0 + parameter)
                If logGrowth = 0.0 Then Return New FormulaError("#DIV/0!")
                Return FiniteNumber((Math.Log(future) - Math.Log(present)) / logGrowth)
            End Function

            Private Shared Function EstimateLoanRate(values As Double()) As Object
                Dim periods = values(0)
                Dim payment = values(1)
                Dim present = values(2)
                Dim future = values(3)
                Dim type = If(values(4) = 0.0, 0.0, 1.0)
                Dim rate = values(5)
                If rate <= -1.0 Then Return New FormulaError("#VALUE!")
                If periods <= 0.0 OrElse Math.Min(payment, Math.Min(present, future)) >= 0.0 OrElse
                    Math.Max(payment, Math.Max(present, future)) <= 0.0 Then Return New FormulaError("#NUM!")
                If payment = 0.0 Then Return FiniteNumber(Math.Pow(-future / present, 1.0 / periods) - 1.0)

                Dim triangular = periods * (periods - 1.0) / 2.0
                Dim quadratic = triangular * (periods - 2.0) / 3.0
                Dim cubic = quadratic * (periods - 3.0) / 4.0
                For iteration = 1 To 20
                    Dim exponent = periods * LogarithmOfOnePlus(rate)
                    Dim compound = Math.Exp(exponent)
                    Dim compoundDerivative = periods * compound / (1.0 + rate)
                    Dim annuity As Double, annuityDerivative As Double
                    If Math.Abs(exponent) < 0.0001 Then
                        annuity = periods + rate * (triangular + rate * (quadratic + rate * cubic))
                        annuityDerivative = triangular + rate * (2.0 * quadratic + 3.0 * rate * cubic)
                    Else
                        Dim difference = ExponentialMinusOne(exponent)
                        annuity = difference / rate
                        annuityDerivative = (compoundDerivative * rate - difference) / (rate * rate)
                    End If
                    Dim timing = 1.0 + type * rate
                    Dim balance = present * compound + payment * timing * annuity + future
                    Dim derivative = present * compoundDerivative + payment * (timing * annuityDerivative + type * annuity)
                    If derivative = 0.0 Then Return New FormulaError("#NUM!")
                    Dim estimate = rate - balance / derivative
                    If estimate <= -1.0 OrElse Double.IsNaN(estimate) OrElse Double.IsInfinity(estimate) Then Return New FormulaError("#NUM!")
                    If Math.Abs(estimate - rate) < 1.0E-7 Then Return estimate
                    rate = estimate
                Next
                Return New FormulaError("#NUM!")
            End Function

        End Class

    End Class

End Namespace
