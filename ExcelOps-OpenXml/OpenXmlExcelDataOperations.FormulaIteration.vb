Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function ReadCircularValue(cell As Cell, address As ExcelCell, formula As String) As Object
                Dim settings = _owner.WorkbookPart.Workbook.CalculationProperties
                If Not If(settings?.Iterate?.Value, False) Then Throw New CircularReferenceException(address, formula)
                If _run.Active.Peek() IsNot cell Then
                    Throw UnsupportedFormula(address.SheetName, address.RowIndex, address.ColumnIndex, formula,
                                             "Iteration across multiple circular formula cells is outside the supported calculation scope.")
                End If
                If _run.IterationWorkbook IsNot Nothing AndAlso _run.IterationWorkbook IsNot _owner Then
                    Throw UnsupportedFormula(address.SheetName, address.RowIndex, address.ColumnIndex, formula,
                                             "Iteration across workbooks is outside the supported calculation scope.")
                End If
                _run.IterationWorkbook = _owner
                MarkIterationDependents()

                Dim seed As Object = Nothing
                If Not _run.IterationSeeds.TryGetValue(cell, seed) Then seed = ReadIterationSeed(cell, address)
                If Not TypeOf seed Is Double AndAlso Not TypeOf seed Is FormulaError Then
                    Throw UnsupportedFormula(address.SheetName, address.RowIndex, address.ColumnIndex, formula,
                                             "An iterative self-reference requires a numeric starting value.")
                End If
                Return seed
            End Function

            Private Sub MarkIterationDependents()
                _run.IterationCells.UnionWith(_run.Active)
            End Sub

            Private Function ReadIterationSeed(cell As Cell, address As ExcelCell) As Object
                If IsMissingFormulaResult(cell) Then Return 0.0
                If cell.DataType?.Value = CellValues.Error Then Return New FormulaError(_owner.ReadStoredFormulaError(cell))
                Return If(_owner.ReadCellValue(address.SheetName, address.RowIndex, address.ColumnIndex, cell), 0.0)
            End Function

            Private Sub CompleteIterations()
                If _run.IterationsApplied OrElse _run.IterationWorkbook Is Nothing Then Return
                Dim iterationOwner = _run.IterationWorkbook
                Dim iterationCalculator = _run.Workbooks(iterationOwner)
                Dim settings = iterationOwner.WorkbookPart.Workbook.CalculationProperties
                Dim limit = If(settings.IterateCount?.Value, 100UI)
                Dim tolerance = If(settings.IterateDelta?.Value, 0.001)
                If limit < 1UI OrElse limit > 32767UI OrElse tolerance < 0 OrElse
                    Double.IsNaN(tolerance) OrElse Double.IsInfinity(tolerance) Then
                    Dim circular = iterationCalculator._results.First(Function(item) _run.IterationCells.Contains(item.Key))
                    Dim location = circular.Value.Location
                    Throw UnsupportedFormula(location.Sheet, location.Row, location.Column, circular.Key.CellFormula.Text,
                                             "The workbook contains invalid Excel iteration settings.")
                End If

                ' Excel applies the maximum change to the circular cell and the calculated cells that depend on it.
                For pass = 1 To CInt(limit)
                    EnsureIterationIsLocal(iterationOwner)
                    Dim participants = iterationCalculator._results.Where(Function(item) _run.IterationCells.Contains(item.Key)).ToArray()
                    If pass = CInt(limit) OrElse participants.All(
                        Function(item) iterationCalculator.IterationValueIsStable(item.Key, item.Value.Location, item.Value.Result, tolerance)) Then Exit For

                    For Each item In participants
                        _run.IterationSeeds(item.Key) = item.Value.Result
                        iterationCalculator._results.Remove(item.Key)
                    Next
                    For Each item In participants
                        Dim location = item.Value.Location
                        iterationCalculator.ReadCell(location.Sheet, location.Row, location.Column)
                    Next
                Next
                _run.IterationsApplied = True
            End Sub

            Private Sub EnsureIterationIsLocal(iterationOwner As OpenXmlExcelDataOperations)
                For Each item In _run.Workbooks
                    If item.Key Is iterationOwner Then Continue For
                    For Each result In item.Value._results
                        If Not _run.IterationCells.Contains(result.Key) Then Continue For
                        Dim location = result.Value.Location
                        Throw UnsupportedFormula(location.Sheet, location.Row, location.Column, result.Key.CellFormula.Text,
                                                 "Iteration across workbooks is outside the supported calculation scope.")
                    Next
                Next
            End Sub

            Private Function IterationValueIsStable(cell As Cell, location As FormulaLocation, value As Object, tolerance As Double) As Boolean
                Dim previous As Object = Nothing
                If Not _run.IterationSeeds.TryGetValue(cell, previous) Then
                    previous = ReadIterationSeed(cell, New ExcelCell(location.Sheet, location.Row, location.Column, ExcelCell.ValueTypes.All))
                End If
                If TypeOf value Is Double AndAlso TypeOf previous Is Double Then
                    Dim change = Math.Abs(DirectCast(value, Double) - DirectCast(previous, Double))
                    Return change = 0 OrElse change < tolerance
                End If
                Dim valueError = TryCast(value, FormulaError)
                Dim previousError = TryCast(previous, FormulaError)
                If valueError IsNot Nothing OrElse previousError IsNot Nothing Then Return valueError?.Code = previousError?.Code
                Return Object.Equals(value, previous)
            End Function

        End Class

    End Class

End Namespace
