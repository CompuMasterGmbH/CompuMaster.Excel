Imports System.Globalization
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports WorkbookTextVersion = DocumentFormat.OpenXml.Office.SpreadSheetML.Y2024.WorkbookCompatibilityVersion.Version

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <inheritdoc/>
        Protected Overrides Sub RecalculateAllInternal()
            EnsureCalculationEnabled()
            RecalculateSheets(Nothing)
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
                If sheetNames Is Nothing Then
                    calculation.RecalculateWorkbook()
                Else
                    For Each sheetName In sheetNames
                        calculation.RecalculateSheet(sheetName)
                    Next
                End If
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
                    calculation.ReadCell(sheetName, rowIndex, columnIndex)
                    calculation.Commit()
                    If throwExceptionOnCalculationError Then
                        ThrowIfFormulaError(sheetName, rowIndex, columnIndex, calculation.ReadCell(sheetName, rowIndex, columnIndex))
                    End If
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
            Friend ReadOnly WorkbookIndex As Integer?
            Friend ReadOnly SourceWorkbook As OpenXmlExcelDataOperations
            Friend ReadOnly FirstRow As Integer
            Friend ReadOnly FirstColumn As Integer
            Friend ReadOnly LastRow As Integer
            Friend ReadOnly LastColumn As Integer

            Friend Sub New(sheet As String, first As ExcelCell, last As ExcelCell, workbookIndex As Integer?,
                           Optional sourceWorkbook As OpenXmlExcelDataOperations = Nothing)
                Me.Sheet = sheet
                Me.WorkbookIndex = workbookIndex
                Me.SourceWorkbook = sourceWorkbook
                FirstRow = Math.Min(first.RowIndex, last.RowIndex)
                FirstColumn = Math.Min(first.ColumnIndex, last.ColumnIndex)
                LastRow = Math.Max(first.RowIndex, last.RowIndex)
                LastColumn = Math.Max(first.ColumnIndex, last.ColumnIndex)
            End Sub
        End Class

        Private NotInheritable Partial Class FormulaCalculation
            Private ReadOnly _owner As OpenXmlExcelDataOperations
            Private ReadOnly _unicodeTextPositions As Boolean
            Private ReadOnly _worksheets As New Dictionary(Of String, WorksheetSnapshot)(StringComparer.OrdinalIgnoreCase)
            Private ReadOnly _results As New Dictionary(Of Cell, (Part As WorksheetPart, Location As FormulaLocation, Result As Object))
            Private ReadOnly _run As CalculationRun

            Friend Sub New(owner As OpenXmlExcelDataOperations)
                Me.New(owner, New CalculationRun(owner, owner._formulaSession?.Snapshot()))
            End Sub

            Friend Sub New(owner As OpenXmlExcelDataOperations, session As ExcelFormulaSessionSnapshot, sourcePath As String)
                Me.New(owner, New CalculationRun(owner, session, owner, sourcePath))
            End Sub

            Private Sub New(owner As OpenXmlExcelDataOperations, run As CalculationRun)
                _owner = owner
                _run = run
                Dim compatibility = owner.WorkbookPart.Workbook.Descendants(Of WorkbookTextVersion)().FirstOrDefault()
                _unicodeTextPositions = If(compatibility?.SetVersion?.Value, 1UI) >= 2UI
                run.Workbooks.Add(owner, Me)
                For Each sheet In owner.WorkbookSheets()
                    Dim part = TryCast(owner.WorkbookPart.GetPartById(sheet.Id.Value), WorksheetPart)
                    If part IsNot Nothing Then _worksheets.Add(sheet.Name.Value, New WorksheetSnapshot(owner.FormulaCellMetadata, sheet.Name.Value, part))
                Next
            End Sub

            Friend Sub RecalculateWorkbook()
                For Each worksheet In _worksheets.Values
                    RecalculateSpillAnchors(worksheet)
                Next
                For Each worksheet In _worksheets.Keys
                    RecalculateSheet(worksheet)
                Next
                If _run.TrackDependencies Then _run.FullyCalculated.Add(_owner)
            End Sub

            Friend Sub RecalculateSheet(sheetName As String)
                Dim worksheet As WorksheetSnapshot = Nothing
                If Not _worksheets.TryGetValue(sheetName, worksheet) Then Throw New ArgumentException("The selected sheet is not a worksheet.", NameOf(sheetName))
                RecalculateSpillAnchors(worksheet)
                For Each item In worksheet.Cells
                    If item.Value.CellFormula IsNot Nothing Then ReadCell(worksheet.Name, item.Key.RowIndex, item.Key.ColumnIndex)
                Next
            End Sub

            Friend Function EvaluateReplacement(sheetName As String, rowIndex As Integer, columnIndex As Integer, replacement As Cell) As Object
                Dim worksheet = _worksheets(sheetName)
                Dim coordinates = (rowIndex, columnIndex)
                If replacement.CellMetaIndex IsNot Nothing AndAlso Not worksheet.SpillAnchors.ContainsKey(coordinates) Then
                    worksheet.SpillAnchors.Add(coordinates, New SpillArea(replacement, Nothing))
                End If
                worksheet.Cells(coordinates) = replacement
                ReadCell(sheetName, rowIndex, columnIndex)
                CompleteIterations()
                Dim result = _results(replacement).Result
                _results.Remove(replacement)
                WriteCachedResult(replacement, result)
                Return result
            End Function

            Friend Sub AttachDynamicReplacement(sheetName As String, rowIndex As Integer, columnIndex As Integer, cell As Cell)
                _worksheets(sheetName).Cells((rowIndex, columnIndex)) = cell
            End Sub

            Friend Function ReadCell(sheetName As String, rowIndex As Integer, columnIndex As Integer,
                                     Optional trackRead As Boolean = True, Optional calculateIterationMember As Boolean = False) As Object
                Dim worksheet As WorksheetSnapshot = Nothing
                If Not _worksheets.TryGetValue(sheetName, worksheet) Then Return New FormulaError("#REF!")
                Dim location As New FormulaLocation(_owner, worksheet.Name, rowIndex, columnIndex)
                If trackRead Then RecordPrecedent(New FormulaPrecedent(location, rowIndex, columnIndex))
                Dim cell As Cell = Nothing
                worksheet.Cells.TryGetValue((rowIndex, columnIndex), cell)
                If cell?.CellFormula Is Nothing OrElse IsSpillPlaceholder(cell) Then
                    Dim spillValue As Object = Nothing
                    If ReadSpillCell(worksheet, rowIndex, columnIndex, spillValue) Then Return spillValue
                End If
                Dim spill As SpillArea = Nothing
                If Not worksheet.SpillAnchors.TryGetValue((rowIndex, columnIndex), spill) AndAlso
                   _owner.IsGroupedFormulaCell(worksheet.Part, worksheet.Name, rowIndex, columnIndex, cell?.CellFormula) Then
                    Throw UnsupportedFormula(worksheet.Name, rowIndex, columnIndex, cell?.CellFormula?.Text, "Grouped, array and data-table formulas are outside the supported calculation scope.")
                End If
                If cell?.CellFormula Is Nothing Then
                    If cell?.DataType?.Value = CellValues.Error Then Return New FormulaError(_owner.ReadStoredFormulaError(cell))
                    Dim value = _owner.ReadCellValue(worksheet.Name, rowIndex, columnIndex, cell)
                    If TypeOf value Is Double Then Return FiniteNumber(DirectCast(value, Double))
                    If TypeOf value Is DateTime Then Return ToExcelDateSerial(DirectCast(value, DateTime))
                    Return value
                End If

                _owner.EnsureCalculationEnabled()
                Dim cached As (Part As WorksheetPart, Location As FormulaLocation, Result As Object) = Nothing
                If _results.TryGetValue(cell, cached) Then
                    If _run.IterationCells.Contains(cell) Then MarkIterationDependents()
                    Return cached.Result
                End If
                Dim formula = cell.CellFormula.Text
                Dim address As New ExcelCell(worksheet.Name, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
                If _run.InOrderedIterationPass AndAlso Not calculateIterationMember Then
                    If _run.CircularCells.Contains(cell) Then
                        MarkIterationDependents()
                        Return IterationSeed(cell, address, formula)
                    End If
                    If _run.PendingIterationWorkbooks.Contains(_owner) AndAlso _run.IterationCells.Contains(cell) Then
                        MarkIterationDependents()
                        Return PreviousIterationValue(cell, address)
                    End If
                End If
                If _run.Active.Contains(cell) Then Return ReadCircularValue(cell, address, formula)
                If _run.Active.Count >= 256 Then
                    Throw UnsupportedFormula(worksheet.Name, rowIndex, columnIndex, formula, "The dependency chain exceeds the supported depth.")
                End If
                _run.Active.Push(cell)
                If _run.TrackDependencies Then _run.Frames.Push(New FormulaDependencies(location))
                Try
                    Dim expression = FormulaParser.Parse(formula)
                    Dim value = EvaluateExpression(expression, address)
                    If TypeOf value Is FormulaLambda Then value = New FormulaError(CalcErrorCode)
                    If value Is MissingLambdaArgument Then value = Nothing
                    Dim result = If(spill Is Nothing, If(Scalar(value), 0.0), StageSpill(worksheet, spill, value))
                    _results.Add(cell, (worksheet.Part, location, result))
                    If _run.TrackDependencies Then _run.Dependencies(location) = _run.Frames.Peek()
                    Return result
                Catch ex As NotSupportedException When Not TypeOf ex Is FormulaComplexityNotSupportedException
                    Throw UnsupportedFormula(worksheet.Name, rowIndex, columnIndex, formula, ex.Message, ex)
                Catch ex As ArgumentOutOfRangeException
                    Throw UnsupportedFormula(worksheet.Name, rowIndex, columnIndex, formula, "The formula refers to a cell outside the XLSX grid.", ex)
                Finally
                    If _run.TrackDependencies Then _run.Frames.Pop()
                    _run.Active.Pop()
                End Try
            End Function

            Friend Function EvaluateExpression(expression As FormulaExpression, formulaCell As ExcelCell) As Object
                If _run.ExpressionDepth >= 512 Then Throw New NotSupportedException("The nested formula expression exceeds the supported evaluation depth.")
                _run.ExpressionDepth += 1
                Try
                    Return expression(Me, formulaCell)
                Finally
                    _run.ExpressionDepth -= 1
                End Try
            End Function

            Friend Sub Commit()
                CompleteIterations()
                For Each evaluatedWorkbook In _run.Workbooks.Values
                    If _run.CopyWorkbook IsNot Nothing AndAlso evaluatedWorkbook._owner IsNot _run.CopyWorkbook Then Continue For
                    evaluatedWorkbook.CommitSpillResults()
                    For Each item In evaluatedWorkbook._results
                        WriteCachedResult(item.Key, item.Value.Result)
                        evaluatedWorkbook._owner._changedSheets.Add(item.Value.Part)
                    Next
                Next
                PublishDependencies()
            End Sub

            Private Shared Function UnsupportedFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer, formula As String,
                                                       reason As String, Optional inner As Exception = Nothing) As FormulaComplexityNotSupportedException
                Dim address As New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
                Return New FormulaComplexityNotSupportedException(address, formula, address.Address(True) & ": " & reason, inner)
            End Function

            Private Shared Sub WriteCachedResult(cell As Cell, value As Object)
                cell.InlineString = Nothing
                cell.ValueMetaIndex = Nothing
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
                If value Is MissingLambdaArgument Then Return Nothing
                If TypeOf value Is FormulaReferenceUnion Then Return New FormulaError("#VALUE!")
                Dim matrix = TryCast(value, FormulaMatrix)
                If matrix IsNot Nothing Then
                    If matrix.Values.Length <> 1 Then Throw New NotSupportedException("A multi-cell array result requires spill calculation.")
                    Return If(TypeOf matrix.Values(0) Is FormulaArea, CObj(New FormulaError("#VALUE!")), matrix.Values(0))
                End If
                Dim area = TryCast(value, FormulaArea)
                If area Is Nothing Then Return value
                If area.FirstRow <> area.LastRow OrElse area.FirstColumn <> area.LastColumn Then
                    Throw New NotSupportedException("A multi-cell reference requires a function that supports ranges.")
                End If
                Return ResolveReference(area).ReadCell(area.Sheet, area.FirstRow, area.FirstColumn)
            End Function

            Private Iterator Function RangeValues(area As FormulaArea) As IEnumerable(Of Object)
                Dim source = ResolveReference(area)
                RecordPrecedent(New FormulaPrecedent(New FormulaLocation(source._owner, area.Sheet, area.FirstRow, area.FirstColumn), area.LastRow, area.LastColumn))
                Dim worksheet As WorksheetSnapshot = Nothing
                If Not source._worksheets.TryGetValue(area.Sheet, worksheet) Then
                    Yield New FormulaError("#REF!")
                    Return
                End If
                RecalculateReferencedSpills(source, worksheet, area)
                Dim addresses = worksheet.Cells.Keys.AsEnumerable()
                If worksheet.SpillChildren.Count > 0 Then addresses = addresses.Concat(worksheet.SpillChildren.Keys).Distinct().ToArray()
                For Each address In addresses
                    If address.RowIndex >= area.FirstRow AndAlso address.RowIndex <= area.LastRow AndAlso
                        address.ColumnIndex >= area.FirstColumn AndAlso address.ColumnIndex <= area.LastColumn Then
                        Yield source.ReadCell(area.Sheet, address.RowIndex, address.ColumnIndex, trackRead:=False)
                    End If
                Next
            End Function

            Private Shared Sub RecalculateReferencedSpills(source As FormulaCalculation, worksheet As WorksheetSnapshot, area As FormulaArea)
                For Each spill In worksheet.SpillAnchors.Values
                    If spill.First.Row <= area.LastRow AndAlso spill.Last.Row >= area.FirstRow AndAlso
                       spill.First.Column <= area.LastColumn AndAlso spill.Last.Column >= area.FirstColumn AndAlso
                       Not source._run.Active.Contains(worksheet.Cells((spill.First.Row, spill.First.Column))) Then
                        source.ReadCell(worksheet.Name, spill.First.Row, spill.First.Column, trackRead:=False)
                    End If
                Next
            End Sub

            Private Function ForWorkbook(workbook As OpenXmlExcelDataOperations) As FormulaCalculation
                Dim calculation As FormulaCalculation = Nothing
                If _run.Workbooks.TryGetValue(workbook, calculation) Then Return calculation
                Return New FormulaCalculation(workbook, _run)
            End Function

            Private NotInheritable Class CalculationRun
                Friend ReadOnly Workbooks As New Dictionary(Of OpenXmlExcelDataOperations, FormulaCalculation)
                Friend ReadOnly Active As New Stack(Of Cell)
                Friend ReadOnly IterationCells As New HashSet(Of Cell)
                Friend ReadOnly CircularCells As New HashSet(Of Cell)
                Friend HasMultipleCellCycle As Boolean
                Friend InOrderedIterationPass As Boolean
                Friend ReadOnly PendingIterationWorkbooks As New HashSet(Of OpenXmlExcelDataOperations)
                Friend ReadOnly IterationSeeds As New Dictionary(Of Cell, Object)
                Friend IterationWorkbook As OpenXmlExcelDataOperations
                Friend IterationsApplied As Boolean
                Friend ReadOnly Session As ExcelFormulaSessionSnapshot
                Friend ReadOnly CopyWorkbook As OpenXmlExcelDataOperations
                Friend ReadOnly CopySourcePath As String
                Friend ExpressionDepth As Integer
                Friend ReadOnly LocalTime As Lazy(Of DateTime)

                Friend ReadOnly TrackDependencies As Boolean
                Friend ReadOnly Frames As New Stack(Of FormulaDependencies)
                Friend ReadOnly Dependencies As New Dictionary(Of FormulaLocation, FormulaDependencies)
                Friend ReadOnly FullyCalculated As New HashSet(Of OpenXmlExcelDataOperations)

                Friend Sub New(owner As OpenXmlExcelDataOperations, session As ExcelFormulaSessionSnapshot,
                               Optional copyWorkbook As OpenXmlExcelDataOperations = Nothing, Optional copySourcePath As String = Nothing)
                    Me.Session = session
                    Me.CopyWorkbook = copyWorkbook
                    Me.CopySourcePath = copySourcePath
                    LocalTime = New Lazy(Of DateTime)(If(owner.LoadOptions.FormulaLocalClock, Function() DateTime.Now))
                    TrackDependencies = copyWorkbook Is Nothing AndAlso (owner.LoadOptions.CalculateOnCellEdit OrElse
                        (session IsNot Nothing AndAlso session.Workbooks.Values.OfType(Of OpenXmlExcelDataOperations)().
                            Any(Function(workbook) workbook.LoadOptions.CalculateOnCellEdit)))
                End Sub
            End Class

            Private NotInheritable Class WorksheetSnapshot
                Private ReadOnly _cells As Lazy(Of Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), Cell))
                Friend ReadOnly Name As String
                Friend ReadOnly Part As WorksheetPart
                Private ReadOnly _spillAnchors As Lazy(Of Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), SpillArea))
                Friend ReadOnly SpillResults As New Dictionary(Of SpillArea, FormulaMatrix)
                Friend ReadOnly SpillChildren As New Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), SpillArea)

                Friend Sub New(metadata As Metadata, name As String, part As WorksheetPart)
                    Me.Name = name
                    Me.Part = part
                    _cells = New Lazy(Of Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), Cell))(
                        Function() EnumerateWorksheetCells(GetSheetData(part)).ToDictionary(
                            Function(item) (item.RowIndex, item.ColumnIndex), Function(item) item.Value))
                    _spillAnchors = New Lazy(Of Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), SpillArea))(
                        Function() ReadDynamicSpills(metadata, Cells.Values).ToDictionary(Function(item) (item.First.Row, item.First.Column)))
                End Sub

                Friend ReadOnly Property SpillAnchors As Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), SpillArea)
                    Get
                        Return _spillAnchors.Value
                    End Get
                End Property

                Friend ReadOnly Property Cells As Dictionary(Of (RowIndex As Integer, ColumnIndex As Integer), Cell)
                    Get
                        Return _cells.Value
                    End Get
                End Property
            End Class
        End Class

    End Class

End Namespace
