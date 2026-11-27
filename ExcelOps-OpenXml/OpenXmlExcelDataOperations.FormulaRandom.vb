Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly WorksheetRandom As New Random()

            Private Function EvaluateRandomFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                MarkVolatile()
                Select Case name
                    Case "RAND"
                        If arguments.Count <> 0 Then Return New FormulaError("#VALUE!")
                        SyncLock WorksheetRandom
                            Return WorksheetRandom.NextDouble()
                        End SyncLock
                    Case "RANDBETWEEN"
                        If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                        Return BroadcastReferenceArguments(arguments, formulaCell, AddressOf RandomBetween)
                    Case Else
                        Return EvaluateRandomArray(arguments, formulaCell)
                End Select
            End Function

            Private Function RandomBetween(arguments As Object()) As Object
                Dim bottom = NumberValue(arguments(0))
                If TypeOf bottom Is FormulaError Then Return bottom
                Dim top = NumberValue(arguments(1))
                If TypeOf top Is FormulaError Then Return top
                Dim lower = DirectCast(bottom, Double)
                Dim upper = DirectCast(top, Double)
                If lower > upper Then Return New FormulaError("#NUM!")
                lower = Math.Ceiling(lower)
                upper = Math.Max(lower, Math.Floor(upper))
                SyncLock WorksheetRandom
                    Return DrawWorksheetInteger(lower, upper)
                End SyncLock
            End Function

            Private Function EvaluateRandomArray(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count > 5 Then Return New FormulaError("#VALUE!")
                Dim parameters = {1.0, 1.0, 0.0, 1.0}
                Dim whole As Boolean
                For index = 0 To arguments.Count - 1
                    Dim input = EvaluateExpression(arguments(index), formulaCell)
                    If input Is Nothing AndAlso index <> 3 Then Continue For
                    Dim value = Scalar(input)
                    If index = 4 Then
                        Dim optionValue = LogicalOptionValue(value)
                        If TypeOf optionValue Is FormulaError Then Return optionValue
                        whole = DirectCast(optionValue, Boolean)
                    Else
                        Dim number = NumberValue(value)
                        If TypeOf number Is FormulaError Then Return number
                        parameters(index) = DirectCast(number, Double)
                    End If
                Next
                Dim rows = Math.Truncate(parameters(0))
                Dim columns = Math.Truncate(parameters(1))
                Dim lower = If(whole, Math.Ceiling(parameters(2)), parameters(2))
                Dim upper = If(whole, Math.Floor(parameters(3)), parameters(3))
                If rows < 0 OrElse columns < 0 OrElse lower > upper Then Return New FormulaError("#VALUE!")
                If rows = 0 OrElse columns = 0 Then Return New FormulaError(CalcErrorCode)
                If rows > 1048576 OrElse columns > 16384 Then Return New FormulaError("#NUM!")
                Dim rowCount = CInt(rows)
                Dim columnCount = CInt(columns)
                Dim values = AllocateMatrixValues(rowCount, columnCount)
                SyncLock WorksheetRandom
                    For index = 0 To values.Length - 1
                        If whole Then
                            values(index) = DrawWorksheetInteger(lower, upper)
                        ElseIf lower = upper Then
                            values(index) = lower
                        Else
                            Dim fraction = WorksheetRandom.NextDouble()
                            values(index) = FiniteNumber((1.0 - fraction) * lower + fraction * upper)
                        End If
                    Next
                End SyncLock
                Return New FormulaMatrix(rowCount, columnCount, values)
            End Function

            Private Shared Function DrawWorksheetInteger(lower As Double, upper As Double) As Object
                If lower = upper Then Return lower
                Dim fraction = WorksheetRandom.NextDouble()
                Dim span = upper - lower + 1.0
                If Double.IsInfinity(span) Then Return FiniteNumber(Math.Floor((1.0 - fraction) * lower + fraction * upper))
                Return FiniteNumber(Math.Min(upper, lower + Math.Floor(fraction * span)))
            End Function

        End Class

    End Class

End Namespace
