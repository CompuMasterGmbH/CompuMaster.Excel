Namespace ExcelOps

    ''' <summary>Stores application-registered custom functions and the workbooks in a calculation session.</summary>
    Public NotInheritable Class ExcelFormulaSession
        Private ReadOnly _gate As New Object()
        Private ReadOnly _functions As New Dictionary(Of String, Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue))(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _workbooks As New Dictionary(Of String, ExcelDataOperationsBase)(StringComparer.OrdinalIgnoreCase)
        Private _activeWorkbook As String
        Private _activeCell As ExcelCell

        ''' <summary>Registers a custom scalar formula function for this session.</summary>
        ''' <param name="name">The case-insensitive, invariant name used in the formula.</param>
        ''' <param name="implementation">The application code invoked during calculation.</param>
        ''' <remarks>Engine built-ins take precedence. A second registration with the same name replaces the first.</remarks>
        Public Sub RegisterFunction(name As String, implementation As Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue))
            If String.IsNullOrWhiteSpace(name) Then Throw New ArgumentException("A function name is required.", NameOf(name))
            If implementation Is Nothing Then Throw New ArgumentNullException(NameOf(implementation))
            If Not Text.RegularExpressions.Regex.IsMatch(name, "^[A-Za-z_][A-Za-z0-9_.]*$") Then
                Throw New ArgumentException("The function name must use Excel formula identifier syntax.", NameOf(name))
            End If
            SyncLock _gate
                _functions(name) = implementation
            End SyncLock
        End Sub

        ''' <summary>Removes a custom formula function from the session.</summary>
        ''' <param name="name">The registered function name.</param>
        ''' <returns>True if a registration was removed.</returns>
        Public Function RemoveFunction(name As String) As Boolean
            SyncLock _gate
                Return _functions.Remove(name)
            End SyncLock
        End Function

        ''' <summary>Sets or clears the selection supplied to custom functions.</summary>
        ''' <param name="workbookName">The registered name of the workbook containing the selection.</param>
        ''' <param name="cell">The selected cell, or Nothing to clear the selection.</param>
        Public Sub SetActiveCell(workbookName As String, cell As ExcelCell)
            SyncLock _gate
                If cell Is Nothing Then
                    _activeWorkbook = Nothing
                    _activeCell = Nothing
                    Return
                End If
                If String.IsNullOrEmpty(workbookName) OrElse Not _workbooks.ContainsKey(workbookName) Then
                    Throw New ArgumentException("The workbook is not open in the session.", NameOf(workbookName))
                End If
                If Not cell.ValidateFullCellAddressInclSheetName() Then Throw New InvalidCellAddressException(cell)
                _activeWorkbook = workbookName
                _activeCell = CopyCell(cell)
            End SyncLock
        End Sub

        Friend Sub AttachWorkbook(name As String, workbook As ExcelDataOperationsBase)
            SyncLock _gate
                If _workbooks.ContainsKey(name) Then Throw New ArgumentException("A workbook named " & name & " is already open in this session.", NameOf(name))
                _workbooks.Add(name, workbook)
            End SyncLock
        End Sub

        Friend Sub DetachWorkbook(name As String, workbook As ExcelDataOperationsBase)
            SyncLock _gate
                Dim attached As ExcelDataOperationsBase = Nothing
                If Not _workbooks.TryGetValue(name, attached) OrElse attached IsNot workbook Then Return
                _workbooks.Remove(name)
                If String.Equals(_activeWorkbook, name, StringComparison.OrdinalIgnoreCase) Then
                    _activeWorkbook = Nothing
                    _activeCell = Nothing
                End If
            End SyncLock
        End Sub

        Friend Function Snapshot() As ExcelFormulaSessionSnapshot
            SyncLock _gate
                Return New ExcelFormulaSessionSnapshot(
                    New Dictionary(Of String, ExcelDataOperationsBase)(_workbooks, StringComparer.OrdinalIgnoreCase),
                    New Dictionary(Of String, Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue))(_functions, StringComparer.OrdinalIgnoreCase),
                    _activeWorkbook, CopyCell(_activeCell))
            End SyncLock
        End Function

        Friend Shared Function CopyCell(cell As ExcelCell) As ExcelCell
            If cell Is Nothing Then Return Nothing
            Return New ExcelCell(cell.SheetName, cell.Address, cell.DataType)
        End Function
    End Class

    Friend NotInheritable Class ExcelFormulaSessionSnapshot
        Friend ReadOnly Workbooks As IReadOnlyDictionary(Of String, ExcelDataOperationsBase)
        Friend ReadOnly Functions As IReadOnlyDictionary(Of String, Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue))
        Friend ReadOnly ActiveWorkbook As String
        Friend ReadOnly ActiveCell As ExcelCell

        Friend Sub New(workbooks As IReadOnlyDictionary(Of String, ExcelDataOperationsBase),
                       functions As IReadOnlyDictionary(Of String, Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue)),
                       activeWorkbook As String, activeCell As ExcelCell)
            Me.Workbooks = workbooks
            Me.Functions = functions
            Me.ActiveWorkbook = activeWorkbook
            Me.ActiveCell = activeCell
        End Sub
    End Class

End Namespace
