Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Class FormulaMatrix
            Friend ReadOnly Rows As Integer
            Friend ReadOnly Columns As Integer
            Friend ReadOnly Values As Object()

            Friend Sub New(rows As Integer, columns As Integer, values As Object())
                Me.Rows = rows
                Me.Columns = columns
                Me.Values = values
            End Sub

            Friend Function At(row As Integer, column As Integer) As Object
                Return Values(row * Columns + column)
            End Function

            Friend Function BroadcastAt(row As Integer, column As Integer) As Object
                If Rows = 1 Then row = 0
                If Columns = 1 Then column = 0
                If row >= Rows OrElse column >= Columns Then Return New FormulaError("#N/A")
                Return At(row, column)
            End Function
        End Class

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared Function IsMultiValue(value As Object) As Boolean
                Dim area = TryCast(value, FormulaArea)
                If area IsNot Nothing Then Return area.FirstRow <> area.LastRow OrElse area.FirstColumn <> area.LastColumn
                Dim matrix = TryCast(value, FormulaMatrix)
                Return matrix IsNot Nothing AndAlso matrix.Values.Length > 1
            End Function

            Private Function AsMatrix(value As Object) As FormulaMatrix
                Dim area = TryCast(value, FormulaArea)
                If area IsNot Nothing Then
                    Return New FormulaMatrix(area.LastRow - area.FirstRow + 1, area.LastColumn - area.FirstColumn + 1, MaterializeRange(area))
                End If
                Return If(TryCast(value, FormulaMatrix), New FormulaMatrix(1, 1, {value}))
            End Function

            Private Shared Function AllocateMatrixValues(rows As Integer, columns As Integer) As Object()
                If CLng(rows) * columns > MaxMaterializedReferenceCells Then
                    Throw New NotSupportedException("The array contains more cells than this function can materialize.")
                End If
                Return New Object(rows * columns - 1) {}
            End Function

            Private Function EvaluateArrayBinary(operation As String, first As FormulaMatrix, second As FormulaMatrix) As FormulaMatrix
                Dim rows = Math.Max(first.Rows, second.Rows)
                Dim columns = Math.Max(first.Columns, second.Columns)
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        values(row * columns + column) = EvaluateBinary(operation, first.BroadcastAt(row, column), second.BroadcastAt(row, column))
                    Next
                Next
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Function EvaluateArrayFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Select Case name
                    Case "SEQUENCE", "_XLFN.SEQUENCE"
                        Return EvaluateSequence(arguments, formulaCell)
                    Case "TRANSPOSE"
                        If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                        Dim input = EvaluateExpression(arguments(0), formulaCell)
                        If TypeOf input Is FormulaError Then Return input
                        Dim matrix = AsMatrix(input)
                        Dim values(matrix.Values.Length - 1) As Object
                        For row = 0 To matrix.Rows - 1
                            For column = 0 To matrix.Columns - 1
                                values(column * matrix.Rows + row) = If(matrix.At(row, column), 0.0)
                            Next
                        Next
                        Return New FormulaMatrix(matrix.Columns, matrix.Rows, values)
                    Case Else
                        Return EvaluateFilter(arguments, formulaCell)
                End Select
            End Function

            Private Function EvaluateSequence(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count = 0 OrElse arguments.Count > 4 Then Return New FormulaError("#VALUE!")
                Dim parameters = {1.0, 1.0, 1.0, 1.0}
                For index = 0 To arguments.Count - 1
                    Dim argument = EvaluateExpression(arguments(index), formulaCell)
                    If argument Is Nothing Then Continue For
                    Dim number = NumberValue(Scalar(argument))
                    If TypeOf number Is FormulaError Then Return number
                    parameters(index) = DirectCast(number, Double)
                Next

                Dim rows = Math.Truncate(parameters(0))
                Dim columns = Math.Truncate(parameters(1))
                If rows < 0 OrElse columns < 0 Then Return New FormulaError("#VALUE!")
                If rows = 0 OrElse columns = 0 Then Return New FormulaError(CalcErrorCode)
                If rows > 1048576 OrElse columns > 16384 Then Return New FormulaError("#NUM!")
                Dim rowCount = CInt(rows)
                Dim columnCount = CInt(columns)
                Dim values = AllocateMatrixValues(rowCount, columnCount)
                For index = 0 To values.Length - 1
                    values(index) = FiniteNumber(parameters(2) + index * parameters(3))
                Next
                Return New FormulaMatrix(rowCount, columnCount, values)
            End Function

            Private Function EvaluateFilter(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim input = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim matrix = AsMatrix(input)
                Dim include = EvaluateExpression(arguments(1), formulaCell)
                If TypeOf include Is FormulaError Then Return include
                Dim mask = AsMatrix(include)
                Dim selectRows = mask.Columns = 1 AndAlso mask.Rows = matrix.Rows
                Dim selectColumns = mask.Rows = 1 AndAlso mask.Columns = matrix.Columns
                If Not selectRows AndAlso Not selectColumns Then Return New FormulaError("#VALUE!")

                Dim indexes As New List(Of Integer)
                For index = 0 To mask.Values.Length - 1
                    If TypeOf mask.Values(index) Is String Then Return New FormulaError("#VALUE!")
                    Dim truth = NumberValue(mask.Values(index))
                    If TypeOf truth Is FormulaError Then Return truth
                    If DirectCast(truth, Double) <> 0.0 Then indexes.Add(index)
                Next
                If indexes.Count = 0 Then
                    If arguments.Count = 3 Then Return If(EvaluateExpression(arguments(2), formulaCell), 0.0)
                    Return New FormulaError(CalcErrorCode)
                End If

                Dim rows = If(selectRows, indexes.Count, matrix.Rows)
                Dim columns = If(selectRows, matrix.Columns, indexes.Count)
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim sourceRow = If(selectRows, indexes(row), row)
                        Dim sourceColumn = If(selectRows, column, indexes(column))
                        values(row * columns + column) = If(matrix.At(sourceRow, sourceColumn), 0.0)
                    Next
                Next
                Return New FormulaMatrix(rows, columns, values)
            End Function

        End Class

    End Class

End Namespace
