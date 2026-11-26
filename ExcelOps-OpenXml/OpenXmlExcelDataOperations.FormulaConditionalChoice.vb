Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateIf(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim condition = EvaluateExpression(arguments(0), formulaCell)
                If Not TypeOf condition Is FormulaMatrix AndAlso Not IsMultiValue(condition) Then
                    Dim truth = LogicalOptionValue(Scalar(condition))
                    If TypeOf truth Is FormulaError Then Return truth
                    If DirectCast(truth, Boolean) Then Return If(EvaluateExpression(arguments(1), formulaCell), 0.0)
                    If arguments.Count = 3 Then Return If(EvaluateExpression(arguments(2), formulaCell), 0.0)
                    Return False
                End If

                ' For an array condition, both branches participate in the output dimensions.
                Dim conditions = AsMatrix(condition)
                Dim whenTrue = AsMatrix(If(EvaluateExpression(arguments(1), formulaCell), 0.0))
                Dim whenFalse = AsMatrix(If(arguments.Count = 3, If(EvaluateExpression(arguments(2), formulaCell), 0.0), False))
                Dim rows = Math.Max(conditions.Rows, Math.Max(whenTrue.Rows, whenFalse.Rows))
                Dim columns = Math.Max(conditions.Columns, Math.Max(whenTrue.Columns, whenFalse.Columns))
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim truth = LogicalOptionValue(conditions.BroadcastAt(row, column))
                        Dim result = truth
                        If Not TypeOf truth Is FormulaError Then
                            Dim selected = If(DirectCast(truth, Boolean), whenTrue, whenFalse)
                            result = If(selected.BroadcastAt(row, column), 0.0)
                        End If
                        values(row * columns + column) = result
                    Next
                Next
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Function EvaluateConditionalChoice(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                       formulaCell As ExcelCell) As Object
                Dim isSwitch = name = "SWITCH"
                Dim firstPair = If(isSwitch, 1, 0)
                Dim hasDefault = isSwitch AndAlso arguments.Count Mod 2 = 0
                Dim endOfPairs = arguments.Count - If(hasDefault, 1, 0)
                If endOfPairs < firstPair + 2 OrElse (endOfPairs - firstPair) Mod 2 <> 0 Then Return New FormulaError("#VALUE!")

                Dim operands = arguments.Select(Function(expression) New Lazy(Of FormulaMatrix)(
                    Function() AsMatrix(EvaluateExpression(expression, formulaCell)))).ToArray()
                Dim rows = 1
                Dim columns = 1
                Dim read As Func(Of Integer, Integer, Integer, Object) =
                    Function(index, row, column)
                        Dim matrix = operands(index).Value
                        rows = Math.Max(rows, matrix.Rows)
                        columns = Math.Max(columns, matrix.Columns)
                        Return matrix.BroadcastAt(row, column)
                    End Function

                ' Only arguments that contribute to at least one branch grow the result.
                Dim evaluatedRows As Integer
                Dim evaluatedColumns As Integer
                Dim values As Object()
                Do
                    evaluatedRows = rows
                    evaluatedColumns = columns
                    values = AllocateMatrixValues(evaluatedRows, evaluatedColumns)
                    For row = 0 To evaluatedRows - 1
                        For column = 0 To evaluatedColumns - 1
                            values(row * evaluatedColumns + column) =
                                EvaluateConditionalChoiceCell(isSwitch, firstPair, endOfPairs, hasDefault, read, row, column)
                        Next
                    Next
                Loop While rows <> evaluatedRows OrElse columns <> evaluatedColumns
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Function EvaluateConditionalChoiceCell(isSwitch As Boolean, firstPair As Integer, endOfPairs As Integer,
                                                           hasDefault As Boolean, read As Func(Of Integer, Integer, Integer, Object),
                                                           row As Integer, column As Integer) As Object
                Dim target As Object = Nothing
                If isSwitch Then
                    target = If(read(0, row, column), 0.0)
                    If TypeOf target Is FormulaError Then Return target
                End If
                For index = firstPair To endOfPairs - 1 Step 2
                    Dim condition = read(index, row, column)
                    If TypeOf condition Is FormulaError Then Return condition
                    If isSwitch Then
                        If SwitchEquals(target, If(condition, 0.0)) Then Return If(read(index + 1, row, column), 0.0)
                    Else
                        Dim truth = LogicalOptionValue(condition)
                        If TypeOf truth Is FormulaError Then Return truth
                        If DirectCast(truth, Boolean) Then Return If(read(index + 1, row, column), 0.0)
                    End If
                Next
                If hasDefault Then Return If(read(endOfPairs, row, column), 0.0)
                Return New FormulaError("#N/A")
            End Function

            Private Shared Function SwitchEquals(first As Object, second As Object) As Boolean
                If TypeOf first Is Double AndAlso TypeOf second Is Double Then
                    Dim left = DirectCast(first, Double)
                    Dim right = DirectCast(second, Double)
                    Return left = right OrElse left.ToString("G15", CultureInfo.InvariantCulture) = right.ToString("G15", CultureInfo.InvariantCulture)
                End If
                Return EvaluateComparison("=", first, second)
            End Function

        End Class

    End Class

End Namespace
