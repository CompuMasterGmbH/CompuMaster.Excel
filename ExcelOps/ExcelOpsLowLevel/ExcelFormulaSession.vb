Imports System.Globalization

Namespace ExcelOps

    ''' <summary>Stores application-registered custom functions and the workbooks in a calculation session.</summary>
    Public NotInheritable Class ExcelFormulaSession
        Friend ReadOnly WorkbookSyncRoot As New Object()
        Private ReadOnly _functions As New Dictionary(Of String, ExcelFormulaRegistration)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _rtdProviders As New Dictionary(Of String, ExcelFormulaRegistration)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _workbooks As New Dictionary(Of String, ExcelDataOperationsBase)(StringComparer.OrdinalIgnoreCase)
        Private _revision As Long
        Private _activeWorkbook As String
        Private _activeCell As ExcelCell

        ''' <summary>Creates a calculation session without a default formula culture.</summary>
        Public Sub New()
            Me.New(Nothing)
        End Sub

        ''' <summary>Creates a calculation session with an application-supplied default formula culture.</summary>
        ''' <param name="formulaCulture">Default used when an attached Open XML workbook does not specify a culture, or Nothing to leave the default unset.</param>
        Public Sub New(formulaCulture As CultureInfo)
            DefaultFormulaCulture = If(formulaCulture Is Nothing, Nothing, CultureInfo.ReadOnly(formulaCulture))
        End Sub

        ''' <summary>Gets the session's default formula conversion culture, if set.</summary>
        Public ReadOnly Property DefaultFormulaCulture As CultureInfo

        ''' <summary>Registers a custom scalar formula function for this session.</summary>
        ''' <param name="name">The case-insensitive, invariant name used in the formula.</param>
        ''' <param name="implementation">The application code invoked during calculation.</param>
        ''' <remarks>Engine built-ins take precedence. A second registration with the same name replaces the first.</remarks>
        Public Sub RegisterFunction(name As String, implementation As Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue))
            RegisterFunction(name, implementation, isVolatile:=False)
        End Sub

        ''' <inheritdoc cref="RegisterFunction(String, Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue))"/>
        ''' <param name="isVolatile">True to recalculate the function on every automatic calculation pass.</param>
        Public Sub RegisterFunction(name As String, implementation As Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue),
                                    isVolatile As Boolean)
            If String.IsNullOrWhiteSpace(name) Then Throw New ArgumentException("A function name is required.", NameOf(name))
            If implementation Is Nothing Then Throw New ArgumentNullException(NameOf(implementation))
            If Not Text.RegularExpressions.Regex.IsMatch(name, "^[A-Za-z_][A-Za-z0-9_.]*$") Then
                Throw New ArgumentException("The function name must use Excel formula identifier syntax.", NameOf(name))
            End If
            SyncLock WorkbookSyncRoot
                _functions(name) = New ExcelFormulaRegistration(implementation, isVolatile)
                _revision += 1
            End SyncLock
        End Sub

        ''' <summary>Removes a custom formula function from the session.</summary>
        ''' <param name="name">The registered function name.</param>
        ''' <returns>True if a registration was removed.</returns>
        Public Function RemoveFunction(name As String) As Boolean
            SyncLock WorkbookSyncRoot
                If Not _functions.Remove(name) Then Return False
                _revision += 1
                Return True
            End SyncLock
        End Function

        ''' <summary>Registers a data provider for an RTD server identifier in this session.</summary>
        ''' <param name="progId">The case-insensitive identifier supplied as the first RTD argument.</param>
        ''' <param name="implementation">The provider, which receives the formula context and all RTD arguments in order.</param>
        ''' <remarks>Open XML evaluates the provider on automatic calculation passes. A second registration with the same identifier replaces the first.</remarks>
        Public Sub RegisterRtdProvider(progId As String, implementation As Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue))
            If String.IsNullOrWhiteSpace(progId) Then Throw New ArgumentException("An RTD provider identifier is required.", NameOf(progId))
            If implementation Is Nothing Then Throw New ArgumentNullException(NameOf(implementation))
            SyncLock WorkbookSyncRoot
                _rtdProviders(progId) = New ExcelFormulaRegistration(implementation, isVolatile:=True)
                _revision += 1
            End SyncLock
        End Sub

        ''' <summary>Removes the RTD provider for a server identifier.</summary>
        ''' <param name="progId">The registered server identifier.</param>
        ''' <returns>True if a registration was removed.</returns>
        Public Function RemoveRtdProvider(progId As String) As Boolean
            SyncLock WorkbookSyncRoot
                If Not _rtdProviders.Remove(progId) Then Return False
                _revision += 1
                Return True
            End SyncLock
        End Function

        ''' <summary>Sets or clears the selection supplied to custom functions.</summary>
        ''' <param name="workbookName">The registered name of the workbook containing the selection.</param>
        ''' <param name="cell">The selected cell, or Nothing to clear the selection.</param>
        Public Sub SetActiveCell(workbookName As String, cell As ExcelCell)
            SyncLock WorkbookSyncRoot
                If cell Is Nothing Then
                    _activeWorkbook = Nothing
                    _activeCell = Nothing
                    _revision += 1
                    Return
                End If
                If String.IsNullOrEmpty(workbookName) OrElse Not _workbooks.ContainsKey(workbookName) Then
                    Throw New ArgumentException("The workbook is not open in the session.", NameOf(workbookName))
                End If
                If Not cell.ValidateFullCellAddressInclSheetName() Then Throw New InvalidCellAddressException(cell)
                _activeWorkbook = workbookName
                _activeCell = CopyCell(cell)
                _revision += 1
            End SyncLock
        End Sub

        Friend Sub AttachWorkbook(name As String, workbook As ExcelDataOperationsBase)
            SyncLock WorkbookSyncRoot
                If _workbooks.ContainsKey(name) Then Throw New ArgumentException("A workbook named " & name & " is already open in this session.", NameOf(name))
                _workbooks.Add(name, workbook)
                _revision += 1
            End SyncLock
        End Sub

        Friend Sub DetachWorkbook(name As String, workbook As ExcelDataOperationsBase)
            SyncLock WorkbookSyncRoot
                Dim attached As ExcelDataOperationsBase = Nothing
                If Not _workbooks.TryGetValue(name, attached) OrElse attached IsNot workbook Then Return
                _workbooks.Remove(name)
                _revision += 1
                If String.Equals(_activeWorkbook, name, StringComparison.OrdinalIgnoreCase) Then
                    _activeWorkbook = Nothing
                    _activeCell = Nothing
                End If
            End SyncLock
        End Sub

        Friend Sub InvalidateWorkbookDependencies()
            SyncLock WorkbookSyncRoot
                _revision += 1
            End SyncLock
        End Sub

        Friend Function Snapshot() As ExcelFormulaSessionSnapshot
            SyncLock WorkbookSyncRoot
                Return New ExcelFormulaSessionSnapshot(
                    New Dictionary(Of String, ExcelDataOperationsBase)(_workbooks, StringComparer.OrdinalIgnoreCase),
                    New Dictionary(Of String, ExcelFormulaRegistration)(_functions, StringComparer.OrdinalIgnoreCase),
                    New Dictionary(Of String, ExcelFormulaRegistration)(_rtdProviders, StringComparer.OrdinalIgnoreCase),
                    _activeWorkbook, CopyCell(_activeCell), _revision)
            End SyncLock
        End Function

        Friend Shared Function CopyCell(cell As ExcelCell) As ExcelCell
            If cell Is Nothing Then Return Nothing
            Return New ExcelCell(cell.SheetName, cell.Address, cell.DataType)
        End Function
    End Class

    Friend NotInheritable Class ExcelFormulaRegistration
        Friend ReadOnly Implementation As Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue)
        Friend ReadOnly IsVolatile As Boolean

        Friend Sub New(implementation As Func(Of ExcelFormulaContext, IReadOnlyList(Of ExcelFormulaValue), ExcelFormulaValue), isVolatile As Boolean)
            Me.Implementation = implementation
            Me.IsVolatile = isVolatile
        End Sub
    End Class

    Friend NotInheritable Class ExcelFormulaSessionSnapshot
        Friend ReadOnly Workbooks As IReadOnlyDictionary(Of String, ExcelDataOperationsBase)
        Friend ReadOnly Functions As IReadOnlyDictionary(Of String, ExcelFormulaRegistration)
        Friend ReadOnly RtdProviders As IReadOnlyDictionary(Of String, ExcelFormulaRegistration)
        Friend ReadOnly ActiveWorkbook As String
        Friend ReadOnly ActiveCell As ExcelCell
        Friend ReadOnly Revision As Long

        Friend Sub New(workbooks As IReadOnlyDictionary(Of String, ExcelDataOperationsBase),
                       functions As IReadOnlyDictionary(Of String, ExcelFormulaRegistration),
                       rtdProviders As IReadOnlyDictionary(Of String, ExcelFormulaRegistration),
                       activeWorkbook As String, activeCell As ExcelCell, revision As Long)
            Me.Workbooks = workbooks
            Me.Functions = functions
            Me.RtdProviders = rtdProviders
            Me.ActiveWorkbook = activeWorkbook
            Me.ActiveCell = activeCell
            Me.Revision = revision
        End Sub
    End Class

End Namespace
