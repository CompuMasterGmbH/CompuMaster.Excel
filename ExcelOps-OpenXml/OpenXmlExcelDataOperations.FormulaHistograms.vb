Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateFrequency(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                Dim data = EvaluateExpression(arguments(0), formulaCell)
                Dim bins = EvaluateExpression(arguments(1), formulaCell)
                If data Is Nothing OrElse bins Is Nothing Then Return New FormulaError("#VALUE!")

                Dim values = AsMatrix(data).Values
                Dim dataError = values.OfType(Of FormulaError)().FirstOrDefault()
                If dataError IsNot Nothing Then Return dataError
                Dim limits = AsMatrix(bins).Values.OfType(Of Double)().ToArray()
                If limits.Length = 0 Then limits = {0.0}

                ' Process equal cutoffs in their original order and put each count back in its input position.
                Dim sorted = values.OfType(Of Double)().OrderBy(Function(value) value).ToArray()
                Dim output = AllocateMatrixValues(limits.Length + 1, 1)
                Dim position = 0
                For Each index In Enumerable.Range(0, limits.Length).OrderBy(Function(item) limits(item))
                    Dim start = position
                    While position < sorted.Length AndAlso sorted(position) <= limits(index)
                        position += 1
                    End While
                    output(index) = CDbl(position - start)
                Next
                output(limits.Length) = CDbl(sorted.Length - position)
                Return New FormulaMatrix(output.Length, 1, output)
            End Function

            Private Function EvaluateDiscreteProbability(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 3 OrElse arguments.Count > 4 Then Return New FormulaError("#VALUE!")
                Dim data = ProbabilityInput(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf data Is FormulaError Then Return data
                Dim weights = ProbabilityInput(EvaluateExpression(arguments(1), formulaCell))
                If TypeOf weights Is FormulaError Then Return weights
                Dim values = DirectCast(data, FormulaMatrix).Values
                Dim probabilities = DirectCast(weights, FormulaMatrix).Values

                Dim lower = AsMatrix(EvaluateExpression(arguments(2), formulaCell))
                Dim upperArgument = If(arguments.Count = 4, EvaluateExpression(arguments(3), formulaCell), Nothing)
                Dim upper = If(upperArgument Is Nothing, lower, AsMatrix(upperArgument))
                Dim rows = Math.Max(lower.Rows, upper.Rows)
                Dim columns = Math.Max(lower.Columns, upper.Columns)
                Dim validation = ValidateProbabilities(values, probabilities)
                Dim output = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim minimum = NumberValue(lower.BroadcastAt(row, column))
                        Dim maximum = NumberValue(upper.BroadcastAt(row, column))
                        Dim result As Object
                        If TypeOf minimum Is FormulaError Then
                            result = minimum
                        ElseIf TypeOf maximum Is FormulaError Then
                            result = maximum
                        ElseIf validation IsNot Nothing Then
                            result = validation
                        Else
                            result = SumProbabilities(values, probabilities, DirectCast(minimum, Double), DirectCast(maximum, Double))
                        End If
                        output(row * columns + column) = result
                    Next
                Next
                If output.Length = 1 Then Return output(0)
                Return New FormulaMatrix(rows, columns, output)
            End Function

            Private Function ProbabilityInput(value As Object) As Object
                If value Is Nothing Then Return New FormulaError("#VALUE!")
                Dim input = AsMatrix(value)
                If Not TypeOf value Is FormulaMatrix AndAlso input.Values.Length = 1 Then
                    If input.Values(0) Is Nothing Then Return New FormulaError("#VALUE!")
                    If TypeOf input.Values(0) Is FormulaError Then Return input.Values(0)
                End If
                Return input
            End Function

            Private Shared Function ValidateProbabilities(values As Object(), probabilities As Object()) As FormulaError
                If values.Length <> probabilities.Length Then Return New FormulaError("#N/A")
                Dim referenceError = If(values.OfType(Of FormulaError)().FirstOrDefault(), probabilities.OfType(Of FormulaError)().FirstOrDefault())
                If referenceError IsNot Nothing Then Return referenceError
                Dim total = 0.0
                Dim numericPairs = 0
                For index = 0 To values.Length - 1
                    Dim value = values(index)
                    Dim probability = probabilities(index)
                    If Not TypeOf value Is Double OrElse Not TypeOf probability Is Double Then Continue For
                    total += DirectCast(probability, Double)
                    numericPairs += 1
                Next
                If numericPairs = 0 Then Return New FormulaError("#DIV/0!")
                ' Excel accepts a small positive rounding difference when comparing the total with one.
                If total < 1.0 OrElse total > 1.0000000000000011 Then Return New FormulaError("#NUM!")
                Return Nothing
            End Function

            Private Shared Function SumProbabilities(values As Object(), probabilities As Object(), minimum As Double, maximum As Double) As Double
                Dim result = 0.0
                For index = 0 To values.Length - 1
                    If Not TypeOf values(index) Is Double OrElse Not TypeOf probabilities(index) Is Double Then Continue For
                    Dim point = DirectCast(values(index), Double)
                    If point >= minimum AndAlso point <= maximum Then result += DirectCast(probabilities(index), Double)
                Next
                Return result
            End Function

        End Class

    End Class

End Namespace
