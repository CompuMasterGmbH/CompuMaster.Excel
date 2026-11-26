Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Const SpillErrorCode As String = "#SPILL!"

        Private Shared Function IsSpillPlaceholder(cell As Cell) As Boolean
            Dim formula = cell?.CellFormula
            Return formula IsNot Nothing AndAlso String.IsNullOrEmpty(formula.Text) AndAlso If(formula.CalculateCell?.Value, False) AndAlso
                If(formula.FormulaType?.Value, CellFormulaValues.Normal) = CellFormulaValues.Normal
        End Function

        Private Shared Iterator Function ReadDynamicSpills(metadata As Metadata, cells As IEnumerable(Of Cell)) As IEnumerable(Of SpillArea)
            If metadata Is Nothing Then Return
            For Each cell In cells
                If IsDynamicFormulaAnchor(metadata, cell) Then Yield New SpillArea(cell, Nothing)
            Next
        End Function

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateSpillReference(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                Dim argument = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf argument Is FormulaError Then Return argument
                Dim reference = TryCast(argument, FormulaArea)
                If reference Is Nothing OrElse reference.FirstRow <> reference.LastRow OrElse reference.FirstColumn <> reference.LastColumn Then
                    Return New FormulaError("#REF!")
                End If
                If reference.WorkbookIndex.HasValue OrElse (reference.SourceWorkbook IsNot Nothing AndAlso reference.SourceWorkbook IsNot _owner) Then
                    Throw New NotSupportedException("Spill references across workbooks are outside the supported calculation scope.")
                End If

                Dim worksheet As WorksheetSnapshot = Nothing
                If Not _worksheets.TryGetValue(reference.Sheet, worksheet) Then Return New FormulaError("#REF!")
                Dim coordinates = (reference.FirstRow, reference.FirstColumn)
                Dim spill As SpillArea = Nothing
                If Not worksheet.SpillAnchors.TryGetValue(coordinates, spill) Then Return New FormulaError("#REF!")
                Dim anchorValue = ReadCell(worksheet.Name, coordinates.FirstRow, coordinates.FirstColumn)
                If TypeOf anchorValue Is FormulaError Then Return anchorValue
                Dim matrix As FormulaMatrix = Nothing
                If Not worksheet.SpillResults.TryGetValue(spill, matrix) OrElse matrix Is Nothing Then Return New FormulaError("#REF!")
                Return SliceArea(reference, 0, 0, matrix.Rows - 1, matrix.Columns - 1)
            End Function

            Private Sub RecalculateSpillAnchors(worksheet As WorksheetSnapshot)
                For Each anchor In worksheet.SpillAnchors.Keys
                    ReadCell(worksheet.Name, anchor.RowIndex, anchor.ColumnIndex)
                Next
            End Sub

            Private Function ReadSpillCell(worksheet As WorksheetSnapshot, row As Integer, column As Integer, ByRef value As Object) As Boolean
                Dim coordinates = (row, column)
                Dim owner As SpillArea = Nothing
                If worksheet.SpillChildren.TryGetValue(coordinates, owner) Then
                    value = If(worksheet.SpillResults(owner).At(row - owner.First.Row, column - owner.First.Column), 0.0)
                Else
                    owner = worksheet.SpillAnchors.Values.FirstOrDefault(Function(spill) spill.Contains(row, column))
                    If owner Is Nothing Then Return False
                    ReadCell(worksheet.Name, owner.First.Row, owner.First.Column)
                    If worksheet.SpillChildren.TryGetValue(coordinates, owner) Then
                        value = If(worksheet.SpillResults(owner).At(row - owner.First.Row, column - owner.First.Column), 0.0)
                    Else
                        Return True
                    End If
                End If
                Return True
            End Function

            Private Function StageSpill(worksheet As WorksheetSnapshot, spill As SpillArea, result As Object) As Object
                ' A saved dynamic anchor can return a scalar or an error after its inputs change.
                Dim matrix = TryCast(result, FormulaMatrix)
                If matrix Is Nothing AndAlso TypeOf result Is FormulaArea AndAlso IsMultiValue(result) Then matrix = AsMatrix(result)
                If matrix Is Nothing Then
                    worksheet.SpillResults(spill) = Nothing
                    Return If(Scalar(result), 0.0)
                End If

                If _run.Active.Any(Function(cell) _run.IterationCells.Contains(cell)) Then
                    Throw New NotSupportedException("Spilled array results in an iterative cycle are outside the supported calculation scope.")
                End If
                Dim bottom = spill.First.Row + matrix.Rows - 1
                Dim right = spill.First.Column + matrix.Columns - 1
                If bottom >= RowCount OrElse right >= ColumnCount OrElse
                    SpillIsBlocked(worksheet, spill, bottom, right) Then
                    worksheet.SpillResults(spill) = Nothing
                    Return New FormulaError(SpillErrorCode)
                End If

                If matrix.Values.Any(Function(item) TypeOf item Is FormulaArea) Then
                    matrix = New FormulaMatrix(matrix.Rows, matrix.Columns, matrix.Values.Select(
                        Function(item) If(TypeOf item Is FormulaArea, CObj(New FormulaError("#VALUE!")), item)).ToArray())
                End If
                worksheet.SpillResults(spill) = matrix
                For row = spill.First.Row To bottom
                    For column = spill.First.Column To right
                        If row = spill.First.Row AndAlso column = spill.First.Column Then Continue For
                        worksheet.SpillChildren.Add((row, column), spill)
                    Next
                Next
                Return If(matrix.At(0, 0), 0.0)
            End Function

            Private Function SpillIsBlocked(worksheet As WorksheetSnapshot, spill As SpillArea, bottom As Integer, right As Integer) As Boolean
                For Each range In worksheet.Part.Worksheet.Elements(Of MergeCells)().SelectMany(Function(group) group.Elements(Of MergeCell)()).
                        Select(Function(merge) merge.Reference.Value).
                        Concat(worksheet.Part.TableDefinitionParts.Select(Function(part) part.Table.Reference.Value))
                    Dim bounds = New ExcelRange(worksheet.Name, range)
                    If bounds.AddressStart.RowIndex <= bottom AndAlso bounds.AddressEnd.RowIndex >= spill.First.Row AndAlso
                       bounds.AddressStart.ColumnIndex <= right AndAlso bounds.AddressEnd.ColumnIndex >= spill.First.Column Then Return True
                Next

                For row = spill.First.Row To bottom
                    For column = spill.First.Column To right
                        If row = spill.First.Row AndAlso column = spill.First.Column Then Continue For
                        Dim coordinates = (row, column)
                        Dim cell As Cell = Nothing
                        worksheet.Cells.TryGetValue(coordinates, cell)
                        Dim inPreviousSpill = spill.Contains(row, column)
                        If cell?.CellMetaIndex IsNot Nothing OrElse worksheet.SpillChildren.ContainsKey(coordinates) Then Return True
                        If cell?.CellFormula IsNot Nothing AndAlso Not (inPreviousSpill AndAlso IsSpillPlaceholder(cell)) Then Return True
                        If inPreviousSpill Then Continue For
                        If cell?.CellValue IsNot Nothing OrElse cell?.InlineString IsNot Nothing OrElse
                            _owner.IsGroupedFormulaCell(worksheet.Part, worksheet.Name, row, column, Nothing) Then Return True
                    Next
                Next
                Return False
            End Function

            Private Sub CommitSpillResults()
                For Each worksheet In _worksheets.Values
                    If worksheet.SpillResults.Count = 0 Then Continue For
                    For Each entry In worksheet.SpillResults
                        Dim oldRange = entry.Key
                        Dim matrix = entry.Value
                        Dim bottom = oldRange.First.Row + If(matrix?.Rows, 1) - 1
                        Dim right = oldRange.First.Column + If(matrix?.Columns, 1) - 1
                        For Each item In worksheet.Cells
                            Dim row = item.Key.RowIndex
                            Dim column = item.Key.ColumnIndex
                            If item.Value.CellMetaIndex Is Nothing AndAlso oldRange.Contains(row, column) AndAlso (row > bottom OrElse column > right) AndAlso
                                (item.Value.CellFormula Is Nothing OrElse IsSpillPlaceholder(item.Value)) Then
                                item.Value.CellFormula = Nothing
                                ClearStoredCellResult(item.Value)
                            End If
                        Next
                        If bottom <> oldRange.Last.Row OrElse right <> oldRange.Last.Column Then
                            Dim anchor = worksheet.Cells((oldRange.First.Row, oldRange.First.Column))
                            Dim address = ExcelCell.LocalCellAddress(oldRange.First.Row, oldRange.First.Column)
                            anchor.CellFormula.Reference = If(bottom = oldRange.First.Row AndAlso right = oldRange.First.Column,
                                address, address & ":" & ExcelCell.LocalCellAddress(bottom, right))
                            _owner._formulaRanges.Remove(worksheet.Part)
                            _owner._calculationChainInvalid = True
                        End If
                    Next

                    Dim data = GetSheetData(worksheet.Part)
                    For Each entry In worksheet.SpillChildren.OrderBy(Function(item) item.Key.RowIndex).
                            ThenBy(Function(item) item.Key.ColumnIndex)
                        Dim row = entry.Key.RowIndex
                        Dim column = entry.Key.ColumnIndex
                        Dim spill = entry.Value
                        Dim matrix = worksheet.SpillResults(spill)
                        Dim cell As Cell = Nothing
                        If Not worksheet.Cells.TryGetValue(entry.Key, cell) Then
                            cell = FindCell(FindRow(data, row, True), row, column, True)
                        End If
                        WriteCachedResult(cell, If(matrix.At(row - spill.First.Row, column - spill.First.Column), 0.0))
                    Next
                    For Each entry In worksheet.SpillResults
                        If entry.Value Is Nothing Then Continue For
                        UpdateDimension(worksheet.Part.Worksheet, worksheet.Name,
                            entry.Key.First.Row + entry.Value.Rows - 1, entry.Key.First.Column + entry.Value.Columns - 1)
                    Next
                Next
            End Sub

        End Class

    End Class

End Namespace
