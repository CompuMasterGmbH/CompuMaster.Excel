Imports System.Globalization
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <inheritdoc/>
        Protected Overrides Sub RecalculateAllInternal()
            EnsureCalculationEnabled()
            RecalculateSheets(WorkSheetNames())
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub RecalculateSheet(sheetName As String)
            SyncLock WorkbookSaveLock
                EnsureCalculationEnabled()
                GetWorksheetPart(sheetName)
                RecalculateSheets({sheetName})
            End SyncLock
        End Sub

        Private Sub RecalculateSheets(sheetNames As IEnumerable(Of String))
            Dim previousLastCells = _recalculationLastCells
            _recalculationLastCells = New Dictionary(Of WorksheetPart, (RowIndex As Integer, ColumnIndex As Integer))
            Try
                Dim calculation As New FormulaCalculation(Me)
                For Each sheetName In sheetNames
                    calculation.RecalculateSheet(sheetName)
                Next
                calculation.Commit()
            Finally
                _recalculationLastCells = previousLastCells
            End Try
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer, throwExceptionOnCalculationError As Boolean)
            SyncLock WorkbookSaveLock
                    EnsureCalculationEnabled()
                    ValidateCoordinates(rowIndex, columnIndex)
                    Dim formula = GetCell(sheetName, rowIndex, columnIndex)?.CellFormula
                    If formula Is Nothing AndAlso Not IsGroupedFormulaCell(GetWorksheetPart(sheetName), sheetName, rowIndex, columnIndex, formula) Then Return
                    Dim calculation As New FormulaCalculation(Me)
                    Dim result = calculation.ReadCell(sheetName, rowIndex, columnIndex)
                    calculation.Commit()
                    If throwExceptionOnCalculationError Then ThrowIfFormulaError(sheetName, rowIndex, columnIndex, result)
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub CalculateMissingFormulaResult(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            RecalculateCell(sheetName, rowIndex, columnIndex, throwExceptionOnCalculationError:=False)
        End Sub

        Private Shared Sub ThrowIfFormulaError(sheetName As String, rowIndex As Integer, columnIndex As Integer, result As Object)
            If Not TypeOf result Is FormulaError Then Return
            Dim address As New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
            Throw New InvalidOperationException("Formula " & address.Address(True) & " evaluated to " & DirectCast(result, FormulaError).Code & ".")
        End Sub

        Private Sub EnsureCalculationEnabled()
            If CalculationModuleDisabled Then
                Throw New FeatureDisabledException("Calculation engine", "Enable the Open XML formula evaluator through ExcelDataOperationsOptions.DisableCalculationEngine.")
            End If
        End Sub

        Private NotInheritable Class FormulaError
            Friend ReadOnly Code As String

            Friend Sub New(code As String)
                Me.Code = code
            End Sub
        End Class

        Private NotInheritable Class FormulaArea
            Friend ReadOnly Sheet As String
            Friend ReadOnly FirstRow As Integer
            Friend ReadOnly FirstColumn As Integer
            Friend ReadOnly LastRow As Integer
            Friend ReadOnly LastColumn As Integer

            Friend Sub New(sheet As String, first As ExcelCell, last As ExcelCell)
                Me.Sheet = sheet
                FirstRow = Math.Min(first.RowIndex, last.RowIndex)
                FirstColumn = Math.Min(first.ColumnIndex, last.ColumnIndex)
                LastRow = Math.Max(first.RowIndex, last.RowIndex)
                LastColumn = Math.Max(first.ColumnIndex, last.ColumnIndex)
            End Sub
        End Class

        Private NotInheritable Partial Class FormulaCalculation
            Private ReadOnly _owner As OpenXmlExcelDataOperations
            Private ReadOnly _worksheets As New Dictionary(Of String, WorksheetSnapshot)(StringComparer.OrdinalIgnoreCase)
            Private ReadOnly _results As New Dictionary(Of Cell, (Part As WorksheetPart, Result As Object))
            Private ReadOnly _active As New HashSet(Of Cell)
            Private _expressionDepth As Integer

            Friend Sub New(owner As OpenXmlExcelDataOperations)
                _owner = owner
                For Each sheet In owner.WorkbookSheets()
                    Dim part = TryCast(owner.WorkbookPart.GetPartById(sheet.Id.Value), WorksheetPart)
                    If part IsNot Nothing Then _worksheets.Add(sheet.Name.Value, New WorksheetSnapshot(sheet.Name.Value, part))
                Next
            End Sub

            Friend Sub RecalculateSheet(sheetName As String)
                Dim worksheet As WorksheetSnapshot = Nothing
                If Not _worksheets.TryGetValue(sheetName, worksheet) Then Throw New ArgumentException("The selected sheet is not a worksheet.", NameOf(sheetName))
                For Each item In worksheet.Cells
                    If item.Value.CellFormula IsNot Nothing Then ReadCell(worksheet.Name, item.Key.RowIndex, item.Key.ColumnIndex)
                Next
            End Sub

            Friend Function EvaluateReplacement(sheetName As String, rowIndex As Integer, columnIndex As Integer, replacement As Cell) As Object
                _worksheets(sheetName).Cells((rowIndex, columnIndex)) = replacement
                Dim result = ReadCell(sheetName, rowIndex, columnIndex)
                _results.Remove(replacement)
                WriteCachedResult(replacement, result)
                Return result
            End Function

            Friend Function ReadCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Object
                Dim worksheet As WorksheetSnapshot = Nothing
                If Not _worksheets.TryGetValue(sheetName, worksheet) Then Return New FormulaError("#REF!")
                Dim cell As Cell = Nothing
                worksheet.Cells.TryGetValue((rowIndex, columnIndex), cell)
                If _owner.IsGroupedFormulaCell(worksheet.Part, worksheet.Name, rowIndex, columnIndex, cell?.CellFormula) Then
                    Throw UnsupportedFormula(worksheet.Name, rowIndex, columnIndex, cell?.CellFormula?.Text, "Grouped, array and data-table formulas are outside the supported calculation scope.")
                End If
                If cell?.CellFormula Is Nothing Then
                    If cell?.DataType?.Value = CellValues.Error Then Return New FormulaError(If(cell.CellValue?.Text, "#VALUE!"))
                    Dim value = _owner.ReadCellValue(worksheet.Name, rowIndex, columnIndex, cell)
                    If TypeOf value Is Double Then Return FiniteNumber(DirectCast(value, Double))
                    If TypeOf value Is DateTime Then Throw New NotSupportedException("ISO date cells are outside the supported calculation scope.")
                    Return value
                End If

                Dim cached As (Part As WorksheetPart, Result As Object) = Nothing
                If _results.TryGetValue(cell, cached) Then Return cached.Result
                Dim formula = cell.CellFormula.Text
                Dim address As New ExcelCell(worksheet.Name, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
                If _active.Contains(cell) Then Throw New CircularReferenceException(address, formula)
                If _active.Count >= 256 Then
                    Throw UnsupportedFormula(worksheet.Name, rowIndex, columnIndex, formula, "The dependency chain exceeds the supported depth.")
                End If
                _active.Add(cell)
                Try
                    Dim expression = FormulaParser.Parse(formula)
                    Dim result = If(Scalar(EvaluateExpression(expression, address)), 0.0)
                    _results.Add(cell, (worksheet.Part, result))
                    Return result
                Catch ex As NotSupportedException When Not TypeOf ex Is FormulaComplexityNotSupportedException
                    Throw UnsupportedFormula(worksheet.Name, rowIndex, columnIndex, formula, ex.Message, ex)
                Catch ex As ArgumentOutOfRangeException
                    Throw UnsupportedFormula(worksheet.Name, rowIndex, columnIndex, formula, "The formula refers to a cell outside the XLSX grid.", ex)
                Finally
                    _active.Remove(cell)
                End Try
            End Function

            Friend Function EvaluateExpression(expression As FormulaExpression, formulaCell As ExcelCell) As Object
                If _expressionDepth >= 512 Then Throw New NotSupportedException("The nested formula expression exceeds the supported evaluation depth.")
                _expressionDepth += 1
                Try
                    Return expression(Me, formulaCell)
                Finally
                    _expressionDepth -= 1
                End Try
            End Function

            Friend Sub Commit()
                For Each item In _results
                    WriteCachedResult(item.Key, item.Value.Result)
                    _owner._changedSheets.Add(item.Value.Part)
                Next
            End Sub

            Private Shared Function UnsupportedFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer, formula As String,
                                                       reason As String, Optional inner As Exception = Nothing) As FormulaComplexityNotSupportedException
                Dim address As New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
                Return New FormulaComplexityNotSupportedException(address, formula, address.Address(True) & ": " & reason, inner)
            End Function

            Private Shared Sub WriteCachedResult(cell As Cell, value As Object)
                cell.InlineString = Nothing
                If TypeOf value Is FormulaError Then
                    cell.DataType = CellValues.Error
                    cell.CellValue = New CellValue(DirectCast(value, FormulaError).Code)
                ElseIf TypeOf value Is Boolean Then
                    cell.DataType = CellValues.Boolean
                    cell.CellValue = New CellValue(If(DirectCast(value, Boolean), "1", "0"))
                ElseIf TypeOf value Is String Then
                    cell.DataType = CellValues.String
                    Dim encoded = EncodeText(DirectCast(value, String))
                    cell.CellValue = New CellValue(encoded)
                Else
                    cell.DataType = CellValues.Number
                    cell.CellValue = New CellValue(DirectCast(value, Double).ToString("R", CultureInfo.InvariantCulture))
                End If
            End Sub

            Private Function Scalar(value As Object) As Object
                Dim area = TryCast(value, FormulaArea)
                If area Is Nothing Then Return value
                If area.FirstRow <> area.LastRow OrElse area.FirstColumn <> area.LastColumn Then
                    Throw New NotSupportedException("A multi-cell reference requires a function that supports ranges.")
                End If
                Return ReadCell(area.Sheet, area.FirstRow, area.FirstColumn)
            End Function

            Private Iterator Function RangeValues(area As FormulaArea) As IEnumerable(Of Object)
                Dim worksheet As WorksheetSnapshot = Nothing
                If Not _worksheets.TryGetValue(area.Sheet, worksheet) Then
                    Yield New FormulaError("#REF!")
                    Return
                End If
                For Each address In worksheet.Cells.Keys
                    If address.RowIndex >= area.FirstRow AndAlso address.RowIndex <= area.LastRow AndAlso
                        address.ColumnIndex >= area.FirstColumn AndAlso address.ColumnIndex <= area.LastColumn Then
                        Yield ReadCell(area.Sheet, address.RowIndex, address.ColumnIndex)
                    End If
                Next
            End Function

            Private NotInheritable Class WorksheetSnapshot
                Private ReadOnly _cells As Lazy(Of Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), Cell))
                Friend ReadOnly Name As String
                Friend ReadOnly Part As WorksheetPart

                Friend Sub New(name As String, part As WorksheetPart)
                    Me.Name = name
                    Me.Part = part
                    _cells = New Lazy(Of Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), Cell))(
                        Function() EnumerateWorksheetCells(GetSheetData(part)).ToDictionary(
                            Function(item) (item.RowIndex, item.ColumnIndex), Function(item) item.Value))
                End Sub

                Friend ReadOnly Property Cells As Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), Cell)
                    Get
                        Return _cells.Value
                    End Get
                End Property
            End Class
        End Class

    End Class

End Namespace
