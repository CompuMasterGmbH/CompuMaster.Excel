Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateArrayWindow(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell, drop As Boolean) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim input = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim matrix = AsMatrix(input)
                Dim rowCount = ReadArrayWindowSize(arguments(1), formulaCell, If(drop, 0, matrix.Rows))
                Dim columnCount = If(arguments.Count = 3, ReadArrayWindowSize(arguments(2), formulaCell, If(drop, 0, matrix.Columns)),
                    CType(CDbl(If(drop, 0, matrix.Columns)), Object))
                If TypeOf rowCount Is FormulaError Then Return rowCount
                If TypeOf columnCount Is FormulaError Then Return columnCount

                Dim rows = SelectWindow(matrix.Rows, DirectCast(rowCount, Double), drop)
                Dim columns = SelectWindow(matrix.Columns, DirectCast(columnCount, Double), drop)
                If rows.Count = 0 OrElse columns.Count = 0 Then Return New FormulaError(CalcErrorCode)
                Return SliceMatrix(matrix, rows.Start, columns.Start, rows.Start + rows.Count - 1, columns.Start + columns.Count - 1, preserveShape:=True)
            End Function

            Private Function ReadArrayWindowSize(argument As FormulaExpression, formulaCell As ExcelCell, defaultCount As Integer) As Object
                Dim input = EvaluateExpression(argument, formulaCell)
                If input Is Nothing Then Return CDbl(defaultCount)
                Dim value = NumberValue(Scalar(input))
                If TypeOf value Is FormulaError Then Return value
                Return Math.Truncate(DirectCast(value, Double))
            End Function

            Private Shared Function SelectWindow(size As Integer, requested As Double, drop As Boolean) As (Start As Integer, Count As Integer)
                Dim matched = CInt(Math.Min(CDbl(size), Math.Abs(requested)))
                If drop Then Return (If(requested > 0, matched, 0), size - matched)
                Return (If(requested < 0, size - matched, 0), matched)
            End Function

            Private Function EvaluateArraySelection(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell, byRow As Boolean) As Object
                If arguments.Count < 2 Then Return New FormulaError("#VALUE!")
                Dim input = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim matrix = AsMatrix(input)
                Dim size = If(byRow, matrix.Rows, matrix.Columns)
                Dim selected As New List(Of Integer)
                For index = 1 To arguments.Count - 1
                    Dim positions = AsMatrix(EvaluateExpression(arguments(index), formulaCell))
                    If positions.Rows > 1 AndAlso positions.Columns > 1 Then Return New FormulaError("#VALUE!")
                    For Each position In positions.Values
                        Dim value = NumberValue(position)
                        If TypeOf value Is FormulaError Then Return value
                        Dim number = Math.Truncate(DirectCast(value, Double))
                        If number = 0 OrElse Math.Abs(number) > size Then Return New FormulaError("#VALUE!")
                        selected.Add(CInt(If(number > 0, number - 1, size + number)))
                    Next
                Next

                Return SelectMatrixSlices(matrix, selected, byRow)
            End Function

            Private Shared Function SelectMatrixSlices(matrix As FormulaMatrix, selected As IReadOnlyList(Of Integer),
                                                       byRow As Boolean) As FormulaMatrix
                Dim rows = If(byRow, selected.Count, matrix.Rows)
                Dim columns = If(byRow, matrix.Columns, selected.Count)
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        values(row * columns + column) = matrix.At(If(byRow, selected(row), row), If(byRow, column, selected(column)))
                    Next
                Next
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Function EvaluateArrayStack(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell, vertical As Boolean) As Object
                If arguments.Count = 0 Then Return New FormulaError("#VALUE!")
                Dim arrays = arguments.Select(Function(argument) AsMatrix(EvaluateExpression(argument, formulaCell))).ToArray()
                Dim rows = If(vertical, arrays.Sum(Function(matrix) matrix.Rows), arrays.Max(Function(matrix) matrix.Rows))
                Dim columns = If(vertical, arrays.Max(Function(matrix) matrix.Columns), arrays.Sum(Function(matrix) matrix.Columns))
                Dim values = AllocateMatrixValues(rows, columns)
                Dim padding As New FormulaError("#N/A")
                For index = 0 To values.Length - 1
                    values(index) = padding
                Next

                Dim firstRow = 0
                Dim firstColumn = 0
                For Each matrix In arrays
                    For row = 0 To matrix.Rows - 1
                        Array.Copy(matrix.Values, row * matrix.Columns, values, (firstRow + row) * columns + firstColumn, matrix.Columns)
                    Next
                    If vertical Then
                        firstRow += matrix.Rows
                    Else
                        firstColumn += matrix.Columns
                    End If
                Next
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Shared Function SliceMatrix(matrix As FormulaMatrix, firstRow As Integer, firstColumn As Integer,
                                                lastRow As Integer, lastColumn As Integer, Optional preserveShape As Boolean = False) As Object
                If Not preserveShape AndAlso firstRow = lastRow AndAlso firstColumn = lastColumn Then Return matrix.At(firstRow, firstColumn)
                Dim rows = lastRow - firstRow + 1
                Dim columns = lastColumn - firstColumn + 1
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    Array.Copy(matrix.Values, (firstRow + row) * matrix.Columns + firstColumn, values, row * columns, columns)
                Next
                Return New FormulaMatrix(rows, columns, values)
            End Function

        End Class

    End Class

End Namespace
