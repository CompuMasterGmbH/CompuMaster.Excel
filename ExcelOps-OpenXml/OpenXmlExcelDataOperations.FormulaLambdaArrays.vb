Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateLambdaArray(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                formulaCell As ExcelCell) As Object
                Dim expected = If(name = "BYROW" OrElse name = "BYCOL", 2, 3)
                If If(name = "MAP", arguments.Count < 2, arguments.Count <> expected) Then Return New FormulaError("#VALUE!")

                If name = "MAKEARRAY" Then Return MakeLambdaArray(arguments, formulaCell)
                Dim sourceIndex = If(name = "SCAN" OrElse name = "REDUCE", 1, 0)
                Dim source = EvaluateExpression(arguments(sourceIndex), formulaCell)
                If source Is Nothing Then Return New FormulaError("#VALUE!")
                Dim lambda = TryCast(EvaluateExpression(arguments(arguments.Count - 1), formulaCell), FormulaLambda)
                Dim parameters = If(name = "MAP", arguments.Count - 1, If(sourceIndex = 1, 2, 1))
                If lambda Is Nothing OrElse Not lambda.Accepts(parameters) Then Return New FormulaError("#VALUE!")

                Dim input = AsMatrix(source)
                Select Case name
                    Case "BYROW", "BYCOL"
                        Return ApplyLambdaToAxes(input, name = "BYROW", lambda, formulaCell)
                    Case "MAP"
                        Dim matrices As New List(Of FormulaMatrix) From {input}
                        For index = 1 To arguments.Count - 2
                            Dim additional = EvaluateExpression(arguments(index), formulaCell)
                            If Not TypeOf additional Is FormulaArea AndAlso Not TypeOf additional Is FormulaMatrix Then Return New FormulaError("#VALUE!")
                            matrices.Add(AsMatrix(additional))
                        Next
                        Return MapLambdaArrays(matrices, lambda, formulaCell)
                    Case Else
                        Dim initial = EvaluateExpression(arguments(0), formulaCell)
                        Return FoldLambdaArray(input, initial, name = "SCAN", lambda, formulaCell)
                End Select
            End Function

            Private Function MakeLambdaArray(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Dim rowsValue = NumberValue(Scalar(EvaluateExpression(arguments(0), formulaCell)))
                If TypeOf rowsValue Is FormulaError Then Return rowsValue
                Dim columnsValue = NumberValue(Scalar(EvaluateExpression(arguments(1), formulaCell)))
                If TypeOf columnsValue Is FormulaError Then Return columnsValue
                Dim rows = DirectCast(rowsValue, Double)
                Dim columns = DirectCast(columnsValue, Double)
                If rows < 1.0 OrElse columns < 1.0 OrElse rows > RowCount OrElse columns > ColumnCount OrElse
                    rows <> Math.Truncate(rows) OrElse columns <> Math.Truncate(columns) Then Return New FormulaError("#VALUE!")

                Dim lambda = TryCast(EvaluateExpression(arguments(2), formulaCell), FormulaLambda)
                If lambda Is Nothing OrElse Not lambda.Accepts(2) Then Return New FormulaError("#VALUE!")
                Dim values = AllocateMatrixValues(CInt(rows), CInt(columns))
                For row = 0 To CInt(rows) - 1
                    For column = 0 To CInt(columns) - 1
                        Dim value = CallFormulaLambda(lambda, {row + 1.0, column + 1.0}, formulaCell)
                        If IsNestedLambdaResult(value) Then Return New FormulaError(CalcErrorCode)
                        values(row * CInt(columns) + column) = If(Scalar(value), 0.0)
                    Next
                Next
                Return LambdaArrayResult(CInt(rows), CInt(columns), values)
            End Function

            Private Function ApplyLambdaToAxes(input As FormulaMatrix, byRow As Boolean, lambda As FormulaLambda,
                                                formulaCell As ExcelCell) As Object
                Dim rows = If(byRow, input.Rows, 1)
                Dim columns = If(byRow, 1, input.Columns)
                Dim values = AllocateMatrixValues(rows, columns)
                For index = 0 To values.Length - 1
                    Dim position = index
                    Dim slice = If(byRow,
                        Enumerable.Range(0, input.Columns).Select(Function(column) input.At(position, column)).ToArray(),
                        Enumerable.Range(0, input.Rows).Select(Function(row) input.At(row, position)).ToArray())
                    Dim vector As New FormulaMatrix(If(byRow, 1, input.Rows), If(byRow, input.Columns, 1), slice)
                    Dim value = CallFormulaLambda(lambda, {vector}, formulaCell)
                    If IsNestedLambdaResult(value) Then Return New FormulaError(CalcErrorCode)
                    values(index) = If(Scalar(value), 0.0)
                Next
                Return LambdaArrayResult(rows, columns, values)
            End Function

            Private Function MapLambdaArrays(inputs As IReadOnlyList(Of FormulaMatrix), lambda As FormulaLambda,
                                             formulaCell As ExcelCell) As Object
                Dim rows = inputs.Max(Function(matrix) matrix.Rows)
                Dim columns = inputs.Max(Function(matrix) matrix.Columns)
                Dim values = AllocateMatrixValues(rows, columns)
                Dim operands(inputs.Count - 1) As Object
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim withinInputs = True
                        For index = 0 To operands.Length - 1
                            If row >= inputs(index).Rows OrElse column >= inputs(index).Columns Then
                                withinInputs = False
                                Exit For
                            End If
                            operands(index) = inputs(index).At(row, column)
                        Next
                        Dim value = If(withinInputs, CallFormulaLambda(lambda, operands, formulaCell), New FormulaError("#N/A"))
                        If IsNestedLambdaResult(value) Then Return New FormulaError(CalcErrorCode)
                        values(row * columns + column) = If(Scalar(value), 0.0)
                    Next
                Next
                Return LambdaArrayResult(rows, columns, values)
            End Function

            Private Function FoldLambdaArray(input As FormulaMatrix, initial As Object, scan As Boolean, lambda As FormulaLambda,
                                             formulaCell As ExcelCell) As Object
                Dim accumulator = If(initial, input.Values(0))
                Dim start = If(initial Is Nothing, 1, 0)
                Dim values = If(scan, AllocateMatrixValues(input.Rows, input.Columns), Nothing)
                If scan AndAlso start = 1 Then values(0) = If(Scalar(accumulator), 0.0)
                For index = start To input.Values.Length - 1
                    accumulator = CallFormulaLambda(lambda, {accumulator, input.Values(index)}, formulaCell)
                    If Not scan Then Continue For
                    If IsNestedLambdaResult(accumulator) Then Return New FormulaError(CalcErrorCode)
                    values(index) = If(Scalar(accumulator), 0.0)
                Next
                If scan Then Return LambdaArrayResult(input.Rows, input.Columns, values)
                Return accumulator
            End Function

            Private Shared Function IsNestedLambdaResult(input As Object) As Boolean
                Return TypeOf input Is FormulaLambda OrElse IsMultiValue(input)
            End Function

            Private Shared Function LambdaArrayResult(rows As Integer, columns As Integer, values As Object()) As Object
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(rows, columns, values)
            End Function

        End Class

    End Class

End Namespace
