Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const SmallestNormalMatrixNumber As Double = 2.2250738585072014E-308

            Private Function EvaluateMatrixAlgebra(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                   formulaCell As ExcelCell) As Object
                Dim multiply = name = "MMULT"
                If arguments.Count <> If(multiply, 2, 1) Then Return New FormulaError("#VALUE!")
                Dim input = EvaluateExpression(arguments(0), formulaCell)
                Dim first = AsMatrix(input)
                If name = "MUNIT" Then Return EvaluateIdentityMatrix(first, TypeOf input Is FormulaMatrix OrElse IsMultiValue(input))
                Dim numbers As Double() = Nothing
                Dim failure = ReadNumericMatrix(first, numbers)
                If failure IsNot Nothing Then Return failure

                If multiply Then
                    Dim second = AsMatrix(EvaluateExpression(arguments(1), formulaCell))
                    Dim factors As Double() = Nothing
                    failure = ReadNumericMatrix(second, factors)
                    If failure IsNot Nothing Then Return failure
                    If first.Columns <> second.Rows Then Return New FormulaError("#VALUE!")
                    Return MultiplyMatrices(first.Rows, first.Columns, second.Columns, numbers, factors)
                End If

                If first.Rows <> first.Columns Then Return New FormulaError("#VALUE!")
                Dim size = first.Rows
                Dim permutation = Enumerable.Range(0, size).ToArray()
                Dim sign = FactorMatrix(numbers, size, permutation)
                If name = "MINVERSE" Then
                    If sign = 0 Then Return New FormulaError("#NUM!")
                    Return InvertFactoredMatrix(numbers, size, permutation)
                End If
                If sign = 0 Then Return 0.0
                Dim determinant = CDbl(sign)
                For index = 0 To size - 1
                    determinant *= numbers(index * size + index)
                Next
                Return FiniteNumber(MatrixNumber(determinant))
            End Function

            Private Function EvaluateIdentityMatrix(input As FormulaMatrix, arrayInput As Boolean) As Object
                If arrayInput Then
                    ' A range of dimensions yields the first element of each identity matrix.
                    Dim results = AllocateMatrixValues(input.Rows, input.Columns)
                    For index = 0 To results.Length - 1
                        Dim dimension = IdentityDimension(input.Values(index))
                        results(index) = If(TypeOf dimension Is FormulaError, dimension, 1.0)
                    Next
                    Return New FormulaMatrix(input.Rows, input.Columns, results)
                End If
                Dim size = IdentityDimension(input.Values(0))
                If TypeOf size Is FormulaError Then Return size
                Dim width = DirectCast(size, Integer)
                Dim values = AllocateMatrixValues(width, width)
                For row = 0 To width - 1
                    For column = 0 To width - 1
                        values(row * width + column) = If(row = column, 1.0, 0.0)
                    Next
                Next
                Return New FormulaMatrix(width, width, values)
            End Function

            Private Function IdentityDimension(value As Object) As Object
                Dim number = NumberValue(value)
                If TypeOf number Is FormulaError Then Return number
                Dim dimension = Math.Truncate(DirectCast(number, Double))
                If dimension < 1.0 Then Return New FormulaError("#VALUE!")
                If dimension > 16384.0 Then Return New FormulaError("#NUM!")
                Return CInt(dimension)
            End Function

            Private Shared Function ReadNumericMatrix(matrix As FormulaMatrix, ByRef numbers As Double()) As FormulaError
                numbers = New Double(matrix.Values.Length - 1) {}
                For index = 0 To numbers.Length - 1
                    Dim value = matrix.Values(index)
                    If TypeOf value Is FormulaError Then Return DirectCast(value, FormulaError)
                    If Not TypeOf value Is Double Then Return New FormulaError("#VALUE!")
                    numbers(index) = MatrixNumber(DirectCast(value, Double))
                Next
                Return Nothing
            End Function

            Private Shared Function MultiplyMatrices(rows As Integer, sharedSize As Integer, columns As Integer,
                                                     left As Double(), right As Double()) As Object
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim sum = 0.0
                        For index = 0 To sharedSize - 1
                            sum += left(row * sharedSize + index) * right(index * columns + column)
                        Next
                        Dim result = FiniteNumber(MatrixNumber(sum))
                        If TypeOf result Is FormulaError Then Return result
                        values(row * columns + column) = result
                    Next
                Next
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Shared Function FactorMatrix(values As Double(), size As Integer, permutation As Integer()) As Integer
                ' Store both triangular factors in place and track partial pivoting separately.
                Dim sign = 1
                For column = 0 To size - 1
                    Dim pivot = column
                    For row = column + 1 To size - 1
                        If Math.Abs(values(row * size + column)) > Math.Abs(values(pivot * size + column)) Then pivot = row
                    Next
                    If values(pivot * size + column) = 0.0 Then Return 0
                    If pivot <> column Then
                        For index = 0 To size - 1
                            Dim previous = values(column * size + index)
                            values(column * size + index) = values(pivot * size + index)
                            values(pivot * size + index) = previous
                        Next
                        Dim previousRow = permutation(column)
                        permutation(column) = permutation(pivot)
                        permutation(pivot) = previousRow
                        sign = -sign
                    End If
                    Dim diagonal = values(column * size + column)
                    For row = column + 1 To size - 1
                        Dim factor = values(row * size + column) / diagonal
                        values(row * size + column) = factor
                        For index = column + 1 To size - 1
                            values(row * size + index) -= factor * values(column * size + index)
                        Next
                    Next
                Next
                Return sign
            End Function

            Private Shared Function InvertFactoredMatrix(factors As Double(), size As Integer, permutation As Integer()) As Object
                Dim inverse = AllocateMatrixValues(size, size)
                Dim offsets(size - 1) As Double
                For column = 0 To size - 1
                    ' Solve the two triangular systems for the current identity column.
                    For row = 0 To size - 1
                        Dim sum = If(permutation(row) = column, 1.0, 0.0)
                        For index = 0 To row - 1
                            sum -= factors(row * size + index) * offsets(index)
                        Next
                        offsets(row) = sum
                    Next
                    For row = size - 1 To 0 Step -1
                        Dim sum = offsets(row)
                        For index = row + 1 To size - 1
                            sum -= factors(row * size + index) * offsets(index)
                        Next
                        offsets(row) = sum / factors(row * size + row)
                        Dim result = FiniteNumber(MatrixNumber(offsets(row)))
                        If TypeOf result Is FormulaError Then Return result
                        inverse(row * size + column) = result
                    Next
                Next
                Return New FormulaMatrix(size, size, inverse)
            End Function

            Private Shared Function MatrixNumber(value As Double) As Double
                Return If(Math.Abs(value) < SmallestNormalMatrixNumber, 0.0, value)
            End Function

        End Class

    End Class

End Namespace
