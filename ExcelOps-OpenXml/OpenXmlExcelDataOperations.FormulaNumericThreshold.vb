Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateNumericThreshold(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                      formulaCell As ExcelCell) As Object
                If arguments.Count = 0 OrElse arguments.Count > 2 Then Return New FormulaError("#VALUE!")
                Dim firstValue = EvaluateExpression(arguments(0), formulaCell)
                Dim secondValue = If(arguments.Count = 2, EvaluateExpression(arguments(1), formulaCell), Nothing)
                Dim firstIsRange = TypeOf firstValue Is FormulaArea AndAlso IsMultiValue(firstValue)
                Dim secondIsRange = TypeOf secondValue Is FormulaArea AndAlso IsMultiValue(secondValue)
                Dim first = ThresholdInput(firstValue)
                Dim second = ThresholdInput(secondValue)
                Dim rows = Math.Max(first.Rows, second.Rows)
                Dim columns = Math.Max(first.Columns, second.Columns)
                Dim results = AllocateMatrixValues(rows, columns)

                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim anchor = row = 0 AndAlso column = 0
                        Dim comparison = ThresholdNumber(second.BroadcastAt(row, column), anchor AndAlso secondIsRange)
                        If TypeOf comparison Is FormulaError Then
                            results(row * columns + column) = comparison
                            Continue For
                        End If
                        Dim number = If(firstValue Is Nothing, New FormulaError("#N/A"),
                                        ThresholdNumber(first.BroadcastAt(row, column), anchor AndAlso firstIsRange))
                        If TypeOf number Is FormulaError Then
                            results(row * columns + column) = number
                        Else
                            Dim x = DirectCast(number, Double)
                            Dim y = DirectCast(comparison, Double)
                            results(row * columns + column) = If(If(name = "DELTA", x = y, x >= y), 1.0, 0.0)
                        End If
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function ThresholdNumber(value As Object, rangeAtAnchor As Boolean) As Object
                ' Excel's legacy add-in rejects ranges at the anchor even when another argument spills.
                If rangeAtAnchor OrElse TypeOf value Is Boolean Then Return New FormulaError("#VALUE!")
                Return NumberValue(value)
            End Function

            Private Function ThresholdInput(value As Object) As FormulaMatrix
                Dim area = TryCast(value, FormulaArea)
                If area Is Nothing Then Return AsMatrix(value)
                Dim cellValue = ResolveReference(area).ReadCell(area.Sheet, area.FirstRow, area.FirstColumn)
                Return New FormulaMatrix(1, 1, {cellValue})
            End Function

        End Class

    End Class

End Namespace
