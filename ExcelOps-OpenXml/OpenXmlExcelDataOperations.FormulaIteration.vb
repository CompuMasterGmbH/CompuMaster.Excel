Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function ReadCircularValue(cell As Cell, address As ExcelCell, formula As String) As Object
                Dim settings = _owner.WorkbookPart.Workbook.CalculationProperties
                If Not If(settings?.Iterate?.Value, False) Then Throw New CircularReferenceException(address, formula)
                If _run.InOrderedIterationPass Then
                    Throw UnsupportedFormula(address.SheetName, address.RowIndex, address.ColumnIndex, formula,
                                             "The circular dependency changed during the iteration.")
                End If
                If _run.IterationWorkbook Is Nothing Then _run.IterationWorkbook = _owner
                MarkIterationDependents()
                _run.CircularCells.UnionWith(_run.Active.TakeWhile(Function(member) member IsNot cell))
                _run.CircularCells.Add(cell)
                _run.HasMultipleCellCycle = _run.HasMultipleCellCycle OrElse _run.Active.Peek() IsNot cell
                Return IterationSeed(cell, address, formula)
            End Function

            Private Function IterationSeed(cell As Cell, address As ExcelCell, formula As String) As Object
                Dim seed = PreviousIterationValue(cell, address)
                If Not TypeOf seed Is Double AndAlso Not TypeOf seed Is FormulaError Then
                    Throw UnsupportedFormula(address.SheetName, address.RowIndex, address.ColumnIndex, formula,
                                             "An iterative circular reference requires a numeric starting value.")
                End If
                Return seed
            End Function

            Private Sub MarkIterationDependents()
                _run.IterationCells.UnionWith(_run.Active)
            End Sub

            Private Function PreviousIterationValue(cell As Cell, address As ExcelCell) As Object
                Dim seed As Object = Nothing
                If _run.IterationSeeds.TryGetValue(cell, seed) Then Return seed
                If IsMissingFormulaResult(cell) Then Return 0.0
                If cell.DataType?.Value = CellValues.Error Then Return New FormulaError(_owner.ReadStoredFormulaError(cell))
                Return If(_owner.ReadCellValue(address.SheetName, address.RowIndex, address.ColumnIndex, cell), 0.0)
            End Function

            Private Sub CompleteIterations()
                If _run.IterationsApplied OrElse _run.IterationWorkbook Is Nothing Then Return
                Dim primary = IterationSettings(_run.IterationWorkbook)
                Dim participants = IterationParticipants()
                For Each cycle In participants.Where(Function(item) _run.CircularCells.Contains(item.Key)).
                        GroupBy(Function(item) item.Value.Location.Workbook)
                    Dim settings = IterationSettings(cycle.Key)
                    Dim example = cycle.First()
                    Dim location = example.Value.Location
                    If Not settings.Enabled Then
                        Throw New CircularReferenceException(New ExcelCell(location.Sheet, location.Row, location.Column, ExcelCell.ValueTypes.All),
                                                             example.Key.CellFormula.Text)
                    End If
                    If settings.Limit < 1UI OrElse settings.Limit > 32767UI OrElse settings.Tolerance < 0 OrElse
                        Double.IsNaN(settings.Tolerance) OrElse Double.IsInfinity(settings.Tolerance) Then
                        Throw UnsupportedFormula(location.Sheet, location.Row, location.Column, example.Key.CellFormula.Text,
                                                 "The workbook contains invalid Excel iteration settings.")
                    End If
                    If settings.Limit <> primary.Limit OrElse settings.Tolerance <> primary.Tolerance Then
                        Throw UnsupportedFormula(location.Sheet, location.Row, location.Column, example.Key.CellFormula.Text,
                                                 "Circular calculation across workbooks requires matching iteration settings.")
                    End If
                Next

                Dim ordered = _run.HasMultipleCellCycle OrElse participants.Select(Function(item) item.Value.Location.Workbook).Distinct().Skip(1).Any()
                If ordered Then ReplayIterationPass(participants, ordered)

                ' Excel applies the maximum change to the circular cells and the calculated cells that depend on them.
                For pass = 1 To CInt(primary.Limit)
                    participants = IterationParticipants()
                    If pass = CInt(primary.Limit) OrElse participants.All(
                        Function(item) _run.Workbooks(item.Value.Location.Workbook).IterationValueIsStable(
                            item.Key, item.Value.Location, item.Value.Result, primary.Tolerance)) Then Exit For

                    For Each item In participants
                        _run.IterationSeeds(item.Key) = item.Value.Result
                    Next
                    ReplayIterationPass(participants, ordered)
                Next
                _run.IterationsApplied = True
            End Sub

            Private Shared Function IterationSettings(workbook As OpenXmlExcelDataOperations) As (Enabled As Boolean, Limit As UInteger, Tolerance As Double)
                Dim settings = workbook.WorkbookPart.Workbook.CalculationProperties
                Return (If(settings?.Iterate?.Value, False), If(settings?.IterateCount?.Value, 100UI),
                        If(settings?.IterateDelta?.Value, 0.001))
            End Function

            Private Function IterationParticipants() As KeyValuePair(Of Cell, (Part As WorksheetPart, Location As FormulaLocation, Result As Object))()
                Return _run.Workbooks.Values.SelectMany(Function(calculator) calculator._results).
                    Where(Function(item) _run.IterationCells.Contains(item.Key)).ToArray()
            End Function

            Private Sub ReplayIterationPass(participants As KeyValuePair(Of Cell, (Part As WorksheetPart, Location As FormulaLocation, Result As Object))(),
                                            ordered As Boolean)
                For Each item In participants
                    _run.Workbooks(item.Value.Location.Workbook)._results.Remove(item.Key)
                Next
                Dim workbookOrder = participants.GroupBy(Function(item) item.Value.Location.Workbook).
                    OrderBy(Function(group) IterationWorkbookName(group.Key), StringComparer.OrdinalIgnoreCase).ToArray()
                _run.InOrderedIterationPass = ordered
                If ordered Then _run.PendingIterationWorkbooks.UnionWith(workbookOrder.Select(Function(group) group.Key))
                Try
                    For Each group In workbookOrder
                        _run.PendingIterationWorkbooks.Remove(group.Key)
                        Dim calculator = _run.Workbooks(group.Key)
                        If ordered Then
                            ' Excel refreshes a workbook's circular cells, then its dependents, before visiting the next workbook.
                            Dim sheetOrder = StringComparer.Create(calculator.ConversionCulture, ignoreCase:=True)
                            For Each item In group.Where(Function(entry) _run.CircularCells.Contains(entry.Key)).
                                    OrderBy(Function(entry) entry.Value.Location.Sheet, sheetOrder).
                                    ThenBy(Function(entry) entry.Value.Location.Row).ThenBy(Function(entry) entry.Value.Location.Column)
                                Dim location = item.Value.Location
                                calculator.ReadCell(location.Sheet, location.Row, location.Column, calculateIterationMember:=True)
                            Next
                        End If
                        For Each item In group
                            Dim location = item.Value.Location
                            calculator.ReadCell(location.Sheet, location.Row, location.Column)
                        Next
                    Next
                Finally
                    _run.PendingIterationWorkbooks.Clear()
                    _run.InOrderedIterationPass = False
                End Try
            End Sub

            Private Function IterationWorkbookName(workbook As OpenXmlExcelDataOperations) As String
                ' Excel orders saved workbooks by file name. A save-copy run retains the original path.
                Dim path = ReferencePath(workbook)
                If Not String.IsNullOrEmpty(path) Then Return IO.Path.GetFileName(path)
                If workbook._formulaWorkbookName IsNot Nothing OrElse _run.Session Is Nothing Then Return workbook._formulaWorkbookName
                Return _run.Session.Workbooks.FirstOrDefault(Function(entry) entry.Value Is workbook).Key
            End Function

            Private Function IterationValueIsStable(cell As Cell, location As FormulaLocation, value As Object, tolerance As Double) As Boolean
                Dim previous = PreviousIterationValue(cell, New ExcelCell(location.Sheet, location.Row, location.Column, ExcelCell.ValueTypes.All))
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
