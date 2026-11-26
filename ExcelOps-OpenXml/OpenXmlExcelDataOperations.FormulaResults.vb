Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <summary>Removes stored formula results from every worksheet in the open workbook.</summary>
        ''' <remarks>Formulas are preserved. The workbook requests recalculation when opened in Excel.</remarks>
        Public Sub ResetCellValueFromFormulaCellInWholeWorkbook()
            SyncLock WorkbookSaveLock
                    For Each part In WorkbookPart.WorksheetParts
                        ClearWorksheetFormulaResults(part)
                    Next
            End SyncLock
        End Sub

        ''' <summary>Removes stored formula results from one worksheet in the open workbook.</summary>
        ''' <param name="sheetName">The worksheet name.</param>
        Public Sub ResetCellValueFromFormulaCell(sheetName As String)
            SyncLock WorkbookSaveLock
                    ClearWorksheetFormulaResults(GetWorksheetPart(sheetName))
            End SyncLock
        End Sub

        ''' <summary>Removes the stored result of one formula cell in the open workbook.</summary>
        ''' <param name="cell">The formula cell.</param>
        ''' <exception cref="ArgumentException">The cell does not contain a formula.</exception>
        Public Sub ResetCellValueFromFormulaCell(cell As ExcelCell)
            ValidateAddress(cell)
            ResetCellValueFromFormulaCell(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Sub

        ''' <summary>Removes the stored result of one formula cell in the open workbook.</summary>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="rowIndex">The zero-based row index.</param>
        ''' <param name="columnIndex">The zero-based column index.</param>
        ''' <exception cref="ArgumentException">The cell does not contain a formula.</exception>
        Public Sub ResetCellValueFromFormulaCell(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            SyncLock WorkbookSaveLock
                    ValidateCoordinates(rowIndex, columnIndex)
                    Dim part = GetWorksheetPart(sheetName)
                    Dim cell = FindCell(FindRow(GetSheetData(part), rowIndex, False), rowIndex, columnIndex, False)
                    If cell?.CellFormula Is Nothing Then
                        Throw New ArgumentException("Cell " & New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All).Address(True) & " doesn't contain a formula.")
                    End If
                    Dim changed = ClearFormulaResult(cell)
                    If IsDynamicFormulaAnchor(FormulaCellMetadata, cell) Then
                        changed = ClearSpillChildResults(GetSheetData(part).Descendants(Of Cell)(), {New SpillArea(cell, Nothing)}) OrElse changed
                    End If
                    If changed Then MarkFormulaResultsCleared(part)
            End SyncLock
        End Sub

        Private Sub ClearWorksheetFormulaResults(part As WorksheetPart)
            If ClearFormulaResults(GetSheetData(part), FormulaCellMetadata) Then MarkFormulaResultsCleared(part)
        End Sub

        Private Sub MarkFormulaResultsCleared(part As WorksheetPart)
            _changedSheets.Add(part)
            InvalidateWorkbookFormulaResults()
        End Sub

        Private Sub InvalidateWorkbookFormulaResults()
            _calculationChainInvalid = True
            ResetCalculationDependencies()
            RecalculationRequired = True
            Dim workbook = WorkbookPart.Workbook
            If workbook.CalculationProperties Is Nothing Then workbook.CalculationProperties = New CalculationProperties()
            workbook.CalculationProperties.FullCalculationOnLoad = True
            _workbookChanged = True
        End Sub

        Private Shared Function ClearFormulaResults(data As SheetData, Optional metadata As Metadata = Nothing) As Boolean
            Dim cells = data.Descendants(Of Cell)()
            Dim changed = metadata IsNot Nothing AndAlso ClearSpillChildResults(cells, ReadDynamicSpills(metadata, cells).ToArray())
            For Each cell In cells
                If ClearFormulaResult(cell) Then changed = True
            Next
            Return changed
        End Function

        Private Shared Function ClearSpillChildResults(cells As IEnumerable(Of Cell), spills As IReadOnlyList(Of SpillArea)) As Boolean
            If spills.Count = 0 Then Return False
            Dim changed As Boolean
            For Each cell In cells
                If cell.CellReference Is Nothing OrElse cell.CellMetaIndex IsNot Nothing OrElse
                    (cell.CellFormula IsNot Nothing AndAlso Not IsSpillPlaceholder(cell)) Then Continue For
                Dim address = ReadSharedCellAddress(cell)
                If spills.Any(Function(spill) spill.Contains(address.Row, address.Column)) Then
                    changed = ClearStoredCellResult(cell) OrElse changed
                End If
            Next
            Return changed
        End Function

        Private Shared Function ClearFormulaResult(cell As Cell) As Boolean
            Return cell.CellFormula IsNot Nothing AndAlso ClearStoredCellResult(cell)
        End Function

        Private Shared Function ClearStoredCellResult(cell As Cell) As Boolean
            If cell.CellValue Is Nothing AndAlso cell.InlineString Is Nothing Then Return False
            cell.CellValue = Nothing
            cell.InlineString = Nothing
            cell.DataType = Nothing
            cell.ValueMetaIndex = Nothing
            Return True
        End Function

    End Class

End Namespace
