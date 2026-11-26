Namespace ExcelOps

    ''' <summary>Provides the context of a single custom formula invocation.</summary>
    ''' <remarks>The workbook views are available for the duration of the synchronous invocation.</remarks>
    Public NotInheritable Class ExcelFormulaContext
        Private ReadOnly _formulaCell As ExcelCell
        Private ReadOnly _activeCell As ExcelCell

        Friend Sub New(workbook As ExcelFormulaWorkbookReference, workbookSession As ExcelFormulaWorkbookSession,
                       formulaCell As ExcelCell, activeWorkbookName As String, activeCell As ExcelCell,
                       nativeExcelApplication As Object)
            Me.Workbook = workbook
            Me.WorkbookSession = workbookSession
            _formulaCell = ExcelFormulaSession.CopyCell(formulaCell)
            Me.ActiveWorkbookName = activeWorkbookName
            _activeCell = ExcelFormulaSession.CopyCell(activeCell)
            Me.NativeExcelApplication = nativeExcelApplication
        End Sub

        ''' <summary>Gets a read-only view of the workbook containing the formula.</summary>
        ''' <value>The source workbook of the current invocation.</value>
        Public ReadOnly Property Workbook As ExcelFormulaWorkbookReference

        ''' <summary>Gets the name of the worksheet containing the formula.</summary>
        ''' <value>The worksheet name.</value>
        Public ReadOnly Property Worksheet As String
            Get
                Return _formulaCell.SheetName
            End Get
        End Property

        ''' <summary>Gets the address of the cell containing the formula.</summary>
        ''' <value>A copy of the formula cell address.</value>
        Public ReadOnly Property FormulaCell As ExcelCell
            Get
                Return ExcelFormulaSession.CopyCell(_formulaCell)
            End Get
        End Property

        ''' <summary>Gets the registered workbook name of the active selection.</summary>
        ''' <value>The name, or Nothing when the caller has supplied no selection.</value>
        Public ReadOnly Property ActiveWorkbookName As String

        ''' <summary>Gets the active selection provided by the session.</summary>
        ''' <value>A copy of the selected address, or Nothing when no selection is available.</value>
        Public ReadOnly Property ActiveCell As ExcelCell
            Get
                Return ExcelFormulaSession.CopyCell(_activeCell)
            End Get
        End Property

        ''' <summary>Gets a read-only view of the workbooks in the calculation session.</summary>
        ''' <value>The session workbook view.</value>
        Public ReadOnly Property WorkbookSession As ExcelFormulaWorkbookSession

        ''' <summary>Gets the native Excel application when the engine provides one.</summary>
        ''' <value>An engine-specific application object, or Nothing for a headless engine.</value>
        Public ReadOnly Property NativeExcelApplication As Object
    End Class

    ''' <summary>Provides read access to the workbooks visible to a custom formula.</summary>
    Public NotInheritable Class ExcelFormulaWorkbookSession
        Private ReadOnly _workbooks As IReadOnlyDictionary(Of String, ExcelFormulaWorkbookReference)

        Friend Sub New(workbooks As IReadOnlyDictionary(Of String, ExcelFormulaWorkbookReference))
            _workbooks = workbooks
            WorkbookNames = Array.AsReadOnly(workbooks.Keys.ToArray())
        End Sub

        ''' <summary>Gets the registered names of the visible workbooks.</summary>
        ''' <value>The names in this calculation pass.</value>
        Public ReadOnly Property WorkbookNames As IReadOnlyList(Of String)

        ''' <summary>Finds a workbook by its registered name.</summary>
        ''' <param name="name">The session workbook name.</param>
        ''' <returns>The workbook view, or Nothing if it is unavailable.</returns>
        Public Function FindWorkbook(name As String) As ExcelFormulaWorkbookReference
            Dim workbook As ExcelFormulaWorkbookReference = Nothing
            _workbooks.TryGetValue(name, workbook)
            Return workbook
        End Function
    End Class

    ''' <summary>Provides synchronous read access to a workbook during a formula calculation.</summary>
    Public NotInheritable Class ExcelFormulaWorkbookReference
        Private ReadOnly _readCell As Func(Of ExcelCell, ExcelFormulaValue)

        Friend Sub New(name As String, readCell As Func(Of ExcelCell, ExcelFormulaValue))
            Me.Name = name
            _readCell = readCell
        End Sub

        ''' <summary>Gets the workbook name registered for the calculation session.</summary>
        ''' <value>The session name of the workbook.</value>
        Public ReadOnly Property Name As String

        ''' <summary>Reads a scalar value and resolves any formula dependency before returning.</summary>
        ''' <param name="cell">The address within this workbook, including the worksheet name.</param>
        ''' <returns>The value of the referenced cell.</returns>
        Public Function ReadCell(cell As ExcelCell) As ExcelFormulaValue
            If cell Is Nothing Then Throw New ArgumentNullException(NameOf(cell))
            If Not cell.ValidateFullCellAddressInclSheetName() Then Throw New InvalidCellAddressException(cell)
            Return _readCell(cell)
        End Function
    End Class

End Namespace
