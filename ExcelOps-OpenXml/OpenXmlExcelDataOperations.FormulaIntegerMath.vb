Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const LargestBitOperand As Long = 281474976710655L

            Private Function EvaluateIntegerMath(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                formulaCell As ExcelCell) As Object
                Dim unary = name = "EVEN" OrElse name = "ODD"
                Dim hasMode = name.EndsWith(".MATH", StringComparison.Ordinal)
                Dim optionalSignificance = hasMode OrElse name.EndsWith(".PRECISE", StringComparison.Ordinal) OrElse name = "ISO.CEILING"
                Dim minimum = If(unary OrElse optionalSignificance, 1, 2)
                Dim maximum = If(unary, 1, If(hasMode, 3, 2))
                If arguments.Count < minimum OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")

                Dim firstValue = EvaluateExpression(arguments(0), formulaCell)
                Dim secondValue = If(arguments.Count > 1, EvaluateExpression(arguments(1), formulaCell), Nothing)
                Dim addInRules = name = "MROUND" OrElse name = "QUOTIENT"
                If addInRules AndAlso ((TypeOf firstValue Is FormulaArea AndAlso IsMultiValue(firstValue)) OrElse
                    (TypeOf secondValue Is FormulaArea AndAlso IsMultiValue(secondValue))) Then Return New FormulaError("#VALUE!")
                Dim first = AsMatrix(firstValue)
                Dim second = AsMatrix(If(secondValue, If(optionalSignificance, 1.0, 0.0)))
                Dim third = AsMatrix(If(arguments.Count > 2, EvaluateExpression(arguments(2), formulaCell), 0.0))
                Dim rows = Math.Max(first.Rows, Math.Max(second.Rows, third.Rows))
                Dim columns = Math.Max(first.Columns, Math.Max(second.Columns, third.Columns))
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim firstInput = first.BroadcastAt(row, column)
                        Dim secondInput = second.BroadcastAt(row, column)
                        Dim number = If(addInRules AndAlso TypeOf firstInput Is Boolean, New FormulaError("#VALUE!"), NumberValue(firstInput))
                        Dim significance = If(addInRules AndAlso TypeOf secondInput Is Boolean, New FormulaError("#VALUE!"), NumberValue(secondInput))
                        Dim mode = NumberValue(third.BroadcastAt(row, column))
                        Dim value As Object
                        If TypeOf number Is FormulaError Then
                            value = number
                        ElseIf TypeOf significance Is FormulaError Then
                            value = significance
                        ElseIf TypeOf mode Is FormulaError Then
                            value = mode
                        Else
                            value = IntegerMathValue(name, DirectCast(number, Double), DirectCast(significance, Double), DirectCast(mode, Double))
                        End If
                        results(row * columns + column) = value
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Shared Function IntegerMathValue(name As String, number As Double, operand As Double, mode As Double) As Object
                If name.StartsWith("BIT", StringComparison.Ordinal) Then Return BitMathValue(name, number, operand)
                Select Case name
                    Case "EVEN", "ODD"
                        Dim magnitude = Math.Ceiling(Math.Abs(number))
                        If (magnitude Mod 2.0 = 0.0) = (name = "ODD") Then magnitude += 1.0
                        Return If(number < 0.0, -magnitude, magnitude)
                    Case "QUOTIENT"
                        If operand = 0.0 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber(Math.Truncate(number / operand))
                    Case "MROUND"
                        Return RoundToNearestMultiple(number, operand)
                    Case Else
                        Return RoundToMultiple(name, number, operand, mode)
                End Select
            End Function

            Private Shared Function BitMathValue(name As String, x As Double, y As Double) As Object
                If Not IsBitOperand(x) Then Return New FormulaError("#NUM!")
                Dim first = CLng(x)
                If name = "BITLSHIFT" OrElse name = "BITRSHIFT" Then
                    If first = 0 Then Return 0.0
                    Dim amount = Math.Truncate(y)
                    If Math.Abs(amount) > 53.0 Then Return New FormulaError("#NUM!")
                    If Math.Abs(amount) > 48.0 Then Return 0.0

                    Dim shift = CInt(If(name = "BITLSHIFT", amount, -amount))
                    If shift <= 0 Then Return CDbl(first >> -shift)
                    If first > (LargestBitOperand >> shift) Then Return New FormulaError("#NUM!")
                    Return CDbl(first << shift)
                End If
                If Not IsBitOperand(y) Then Return New FormulaError("#NUM!")
                Dim second = CLng(y)
                Select Case name
                    Case "BITAND" : Return CDbl(first And second)
                    Case "BITOR" : Return CDbl(first Or second)
                    Case Else : Return CDbl(first Xor second)
                End Select
            End Function

            Private Shared Function IsBitOperand(value As Double) As Boolean
                Return value >= 0.0 AndAlso value <= LargestBitOperand AndAlso Math.Truncate(value) = value
            End Function

            Private Shared Function RoundToNearestMultiple(number As Double, stepSize As Double) As Object
                If stepSize = 0.0 OrElse number = 0.0 Then Return 0.0
                If (number < 0.0) <> (stepSize < 0.0) Then Return New FormulaError("#NUM!")
                Dim nearest = Math.Abs(number / stepSize)
                If Double.IsInfinity(nearest) Then Return New FormulaError("#NUM!")
                Dim whole = Math.Floor(nearest)
                Dim remainder = nearest - whole
                ' Excel preserves some decimal midpoint ties close to the unrounded quotient.
                If remainder >= 0.5 OrElse (remainder >= 0.5 - 4.0E-15 AndAlso
                    ToWorksheetPrecision(number) = number AndAlso ToWorksheetPrecision(stepSize) = stepSize) Then whole += 1.0
                Return FiniteNumber(whole * stepSize)
            End Function

            Private Shared Function RoundToMultiple(name As String, number As Double, significance As Double, mode As Double) As Object
                If number = 0.0 Then Return 0.0
                If significance = 0.0 Then
                    If name = "FLOOR" Then Return New FormulaError("#DIV/0!")
                    Return 0.0
                End If
                Dim legacy = name = "FLOOR" OrElse name = "CEILING"
                If legacy AndAlso number > 0.0 AndAlso significance < 0.0 Then Return New FormulaError("#NUM!")

                Dim stepSize = Math.Abs(significance)
                Dim quotient = number / stepSize
                If Double.IsInfinity(quotient) Then Return New FormulaError("#NUM!")
                Dim upward = Not name.StartsWith("FLOOR", StringComparison.Ordinal)
                If (legacy AndAlso significance < 0.0) OrElse (Not legacy AndAlso number < 0.0 AndAlso mode <> 0.0) Then
                    upward = Not upward
                End If
                Dim adjusted = ToWorksheetPrecision(quotient)
                Return FiniteNumber(If(upward, Math.Ceiling(adjusted), Math.Floor(adjusted)) * stepSize)
            End Function

            Private Shared Function ToWorksheetPrecision(value As Double) As Double
                Dim rounded As Double
                If Double.TryParse(value.ToString("G15", CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, rounded) AndAlso
                    Not Double.IsInfinity(rounded) Then Return rounded
                Return value
            End Function

        End Class

    End Class

End Namespace
