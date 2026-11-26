Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const MaximumExactInteger As Long = 9007199254740992L

            Private Function EvaluateCombinatorialFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                           formulaCell As ExcelCell) As Object
                If arguments.Count = 0 Then Return New FormulaError("#VALUE!")
                If name = "SUMSQ" Then
                    Dim inputs As New List(Of Double)
                    For Each argument In arguments
                        Dim failure = AddStatisticValues(EvaluateExpression(argument, formulaCell), inputs)
                        If failure IsNot Nothing Then Return failure
                    Next
                    Return FiniteNumber(inputs.Sum(Function(value) value * value))
                End If
                If name = "GCD" OrElse name = "LCM" OrElse name = "MULTINOMIAL" Then
                    Return EvaluateIntegerAggregate(name, arguments, formulaCell)
                End If

                Dim factorial = name = "FACT" OrElse name = "FACTDOUBLE"
                If arguments.Count <> If(factorial, 1, 2) Then Return New FormulaError("#VALUE!")
                Dim firstInput = EvaluateExpression(arguments(0), formulaCell)
                If name = "FACTDOUBLE" AndAlso TypeOf firstInput Is FormulaArea AndAlso IsMultiValue(firstInput) Then Return New FormulaError("#VALUE!")
                Dim first = AsMatrix(firstInput)
                Dim second = AsMatrix(If(factorial, 0.0, EvaluateExpression(arguments(1), formulaCell)))
                Dim rows = Math.Max(first.Rows, second.Rows)
                Dim columns = Math.Max(first.Columns, second.Columns)
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim raw = first.BroadcastAt(row, column)
                        Dim x = If(name = "FACTDOUBLE" AndAlso TypeOf raw Is Boolean, New FormulaError("#VALUE!"), NumberValue(raw))
                        Dim y = NumberValue(second.BroadcastAt(row, column))
                        Dim result As Object
                        If TypeOf x Is FormulaError Then
                            result = x
                        ElseIf TypeOf y Is FormulaError Then
                            result = y
                        ElseIf factorial Then
                            result = FactorialValue(DirectCast(x, Double), name = "FACTDOUBLE")
                        Else
                            result = CombinatorialValue(name, DirectCast(x, Double), DirectCast(y, Double))
                        End If
                        values(row * columns + column) = result
                    Next
                Next
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Shared Function FactorialValue(number As Double, doubleFactorial As Boolean) As Object
                ' FACT accepts a small tolerance below the next integer. FACTDOUBLE truncates the input.
                Dim last = If(doubleFactorial, Math.Truncate(number), Math.Floor(number + 2.5E-7))
                If number < If(doubleFactorial, -1.0, 0.0) OrElse last > If(doubleFactorial, 300.0, 170.0) Then Return New FormulaError("#NUM!")
                Dim result = 1.0
                For factor As Integer = CInt(last) To 2 Step -If(doubleFactorial, 2, 1)
                    result *= factor
                Next
                Return result
            End Function

            Private Shared Function CombinatorialValue(name As String, number As Double, chosen As Double) As Object
                Dim available = Math.Truncate(number)
                Dim selected = Math.Truncate(chosen)
                If name = "COMBINA" AndAlso available = 0.0 AndAlso selected = 0.0 Then Return 1.0
                If If(name = "COMBINA", available, number) < 0.0 OrElse
                    If(name = "PERMUTATIONA", selected, chosen) < 0.0 Then Return New FormulaError("#NUM!")
                If name = "COMBINA" AndAlso available <= Integer.MaxValue AndAlso selected = 0.0 Then Return 1.0
                If available >= Integer.MaxValue OrElse selected >= Integer.MaxValue Then Return New FormulaError("#NUM!")
                If name = "PERMUTATIONA" Then
                    If selected = 0.0 Then Return 1.0
                    Return FiniteNumber(Math.Pow(available, selected))
                End If
                If name = "COMBINA" Then
                    If available = 0.0 OrElse available + selected - 1.0 >= Integer.MaxValue Then Return New FormulaError("#NUM!")
                    Return FiniteNumber(BinomialCoefficient(CInt(available + selected - 1), CInt(selected)))
                End If
                If selected > available Then Return New FormulaError("#NUM!")
                If name = "COMBIN" Then Return FiniteNumber(BinomialCoefficient(CInt(available), CInt(selected)))

                If selected > 170.0 Then Return New FormulaError("#NUM!")
                Dim product = 1.0
                For index = 0 To CInt(selected) - 1
                    product *= available - index
                Next
                Return FiniteNumber(product)
            End Function

            Private Shared Function BinomialCoefficient(available As Integer, selected As Integer) As Double
                Dim limit = Math.Min(selected, available - selected)
                If limit > 1024 Then Return Double.PositiveInfinity
                Dim coefficient = 1.0
                For index = 1 To limit
                    coefficient *= (available - limit + index) / CDbl(index)
                    If Double.IsInfinity(coefficient) Then Return coefficient
                Next
                Return coefficient
            End Function

            Private Function EvaluateIntegerAggregate(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                Dim inputs As New List(Of Long)
                For index = 0 To arguments.Count - 1
                    Dim input = EvaluateExpression(arguments(index), formulaCell)
                    If input Is Nothing AndAlso index = 0 Then Return New FormulaError("#N/A")
                    Dim area = TryCast(input, FormulaArea)
                    Dim matrix = TryCast(input, FormulaMatrix)
                    Dim values As IEnumerable(Of Object)
                    If area IsNot Nothing AndAlso IsMultiValue(area) Then
                        Dim expectedCount = CLng(area.LastRow - area.FirstRow + 1) * (area.LastColumn - area.FirstColumn + 1)
                        Dim present = RangeValues(area).ToArray()
                        values = present
                        If present.LongLength < expectedCount Then inputs.Add(0)
                    ElseIf matrix IsNot Nothing Then
                        values = matrix.Values
                    Else
                        Dim scalarInput = Scalar(input)
                        If scalarInput Is Nothing AndAlso arguments.Count = 1 Then Return New FormulaError("#VALUE!")
                        If scalarInput Is Nothing Then Continue For
                        values = {scalarInput}
                    End If

                    For Each value In values
                        If TypeOf value Is Boolean Then Return New FormulaError("#VALUE!")
                        Dim number = NumberValue(value)
                        If TypeOf number Is FormulaError Then Return number
                        Dim raw = DirectCast(number, Double)
                        If raw < 0.0 OrElse raw > MaximumExactInteger Then Return New FormulaError("#NUM!")
                        inputs.Add(CLng(Math.Truncate(raw)))
                    Next
                Next

                If name = "MULTINOMIAL" Then
                    Dim total As Integer
                    Dim result = 1.0
                    For Each count In inputs
                        If count > 170 - total Then Return New FormulaError("#NUM!")
                        total += CInt(count)
                        result *= BinomialCoefficient(total, CInt(count))
                    Next
                    Return result
                End If

                If name = "GCD" Then
                    Dim common As Long
                    For Each value In inputs
                        common = GreatestCommonDivisor(common, value)
                    Next
                    Return CDbl(common)
                End If
                If inputs.Count = 0 OrElse inputs.Contains(0) Then Return 0.0
                Dim multiple As Long = 1
                For Each value In inputs
                    Dim reduced = multiple \ GreatestCommonDivisor(multiple, value)
                    If reduced > MaximumExactInteger \ value Then Return New FormulaError("#NUM!")
                    multiple = reduced * value
                Next
                Return CDbl(multiple)
            End Function

            Private Shared Function GreatestCommonDivisor(first As Long, second As Long) As Long
                While second <> 0
                    Dim remainder = first Mod second
                    first = second
                    second = remainder
                End While
                Return first
            End Function

        End Class

    End Class

End Namespace
