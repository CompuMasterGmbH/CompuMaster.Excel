Imports System.Runtime.CompilerServices

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations
        Private ReadOnly _formulaDependencies As New Dictionary(Of FormulaLocation, FormulaDependencies)
        Private ReadOnly _pendingCellEdits As New HashSet(Of FormulaLocation)
        Private _completeDependencyRevision As Long?

        Private Sub RecordCellEdit(sheetName As String, rowIndex As Integer, columnIndex As Integer, calculateOnEdit As Boolean)
            Dim session = _formulaSession?.Snapshot()
            If Not LoadOptions.CalculateOnCellEdit AndAlso (session Is Nothing OrElse
                Not session.Workbooks.Values.OfType(Of OpenXmlExcelDataOperations)().Any(Function(workbook) workbook.LoadOptions.CalculateOnCellEdit)) Then Return

            Dim location As New FormulaLocation(Me, sheetName, rowIndex, columnIndex)
            _formulaDependencies.Remove(location)
            _pendingCellEdits.Add(location)
            If calculateOnEdit Then
                Dim calculation As New FormulaCalculation(Me)
                calculation.RecalculateAfterEdit()
            End If
        End Sub

        Private Function UsesAutomaticValueCalculation() As Boolean
            Return LoadOptions.CalculateOnCellEdit AndAlso Not CalculationModuleDisabled AndAlso AutoCalculationEnabledWorkbookSetting
        End Function

        Private Structure FormulaLocation
            Implements IEquatable(Of FormulaLocation)

            Friend ReadOnly Workbook As OpenXmlExcelDataOperations
            Friend ReadOnly Sheet As String
            Friend ReadOnly Row As Integer
            Friend ReadOnly Column As Integer

            Friend Sub New(workbook As OpenXmlExcelDataOperations, sheet As String, row As Integer, column As Integer)
                Me.Workbook = workbook
                Me.Sheet = sheet
                Me.Row = row
                Me.Column = column
            End Sub

            ''' <inheritdoc/>
            Public Overloads Function Equals(other As FormulaLocation) As Boolean Implements IEquatable(Of FormulaLocation).Equals
                Return Workbook Is other.Workbook AndAlso Row = other.Row AndAlso Column = other.Column AndAlso
                    String.Equals(Sheet, other.Sheet, StringComparison.OrdinalIgnoreCase)
            End Function

            ''' <inheritdoc/>
            Public Overrides Function Equals(obj As Object) As Boolean
                Return TypeOf obj Is FormulaLocation AndAlso Equals(DirectCast(obj, FormulaLocation))
            End Function

            ''' <inheritdoc/>
            Public Overrides Function GetHashCode() As Integer
                Return RuntimeHelpers.GetHashCode(Workbook) Xor StringComparer.OrdinalIgnoreCase.GetHashCode(Sheet) Xor (Row << 16) Xor Column
            End Function
        End Structure

        Private Structure FormulaPrecedent
            Private ReadOnly _start As FormulaLocation
            Private ReadOnly _lastRow As Integer
            Private ReadOnly _lastColumn As Integer

            Friend Sub New(start As FormulaLocation, lastRow As Integer, lastColumn As Integer)
                _start = start
                _lastRow = lastRow
                _lastColumn = lastColumn
            End Sub

            Friend Function Includes(cell As FormulaLocation) As Boolean
                Return _start.Workbook Is cell.Workbook AndAlso String.Equals(_start.Sheet, cell.Sheet, StringComparison.OrdinalIgnoreCase) AndAlso
                    _start.Row <= cell.Row AndAlso cell.Row <= _lastRow AndAlso _start.Column <= cell.Column AndAlso cell.Column <= _lastColumn
            End Function
        End Structure

        Private NotInheritable Class FormulaDependencies
            Friend ReadOnly Cell As FormulaLocation
            Friend ReadOnly Precedents As New List(Of FormulaPrecedent)
            Friend IsVolatile As Boolean

            Friend Sub New(cell As FormulaLocation)
                Me.Cell = cell
            End Sub
        End Class

        Private NotInheritable Partial Class FormulaCalculation
            Private Sub RecordPrecedent(precedent As FormulaPrecedent)
                If _run.Frames.Count > 0 Then _run.Frames.Peek().Precedents.Add(precedent)
            End Sub

            Private Sub MarkVolatile()
                If _run.Frames.Count > 0 Then _run.Frames.Peek().IsVolatile = True
            End Sub

            Private Sub PublishDependencies()
                If Not _run.TrackDependencies Then Return
                For Each workbook In _run.FullyCalculated
                    workbook._formulaDependencies.Clear()
                    workbook._completeDependencyRevision = If(_run.Session?.Revision, 0L)
                Next
                For Each item In _run.Dependencies
                    item.Key.Workbook._formulaDependencies(item.Key) = item.Value
                Next
            End Sub

            Friend Sub RecalculateAfterEdit()
                Dim available = If(_run.Session Is Nothing, {_owner},
                    _run.Session.Workbooks.Values.OfType(Of OpenXmlExcelDataOperations)().
                        Where(Function(workbook) Not workbook.IsClosed).ToArray())
                Dim automatic = New HashSet(Of OpenXmlExcelDataOperations)(available.Where(Function(workbook) workbook.UsesAutomaticValueCalculation()))
                If automatic.Count = 0 Then Return

                Dim revision = If(_run.Session?.Revision, 0L)
                For Each workbook In automatic
                    If Not workbook._completeDependencyRevision.HasValue OrElse workbook._completeDependencyRevision.Value <> revision Then
                        ForWorkbook(workbook).RecalculateWorkbook()
                    End If
                Next

                Dim affected As New HashSet(Of FormulaLocation)
                Dim queue As New Queue(Of FormulaLocation)
                Dim include As Action(Of FormulaLocation) =
                    Sub(location)
                        If affected.Add(location) Then queue.Enqueue(location)
                    End Sub
                Dim dependencies = available.SelectMany(Function(workbook) workbook._formulaDependencies.Values).ToArray()
                For Each workbook In available
                    For Each location In workbook._pendingCellEdits
                        include(location)
                    Next
                Next
                For Each formula In dependencies
                    If formula.IsVolatile Then include(formula.Cell)
                Next

                While queue.Count > 0
                    Dim changed = queue.Dequeue()
                    For Each formula In dependencies
                        If formula.Precedents.Any(Function(precedent) precedent.Includes(changed)) Then include(formula.Cell)
                    Next
                End While

                For Each location In affected
                    If automatic.Contains(location.Workbook) AndAlso Not _run.FullyCalculated.Contains(location.Workbook) AndAlso
                       location.Workbook.GetCell(location.Sheet, location.Row, location.Column)?.CellFormula IsNot Nothing Then
                        ForWorkbook(location.Workbook).ReadCell(location.Sheet, location.Row, location.Column)
                    End If
                Next
                Commit()
                For Each workbook In available
                    workbook._pendingCellEdits.Clear()
                    If automatic.Contains(workbook) Then
                        workbook.RecalculationRequired = False
                    Else
                        workbook._completeDependencyRevision = Nothing
                    End If
                Next
            End Sub
        End Class
    End Class

End Namespace
