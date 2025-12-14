Option Explicit On
Option Strict On
Imports System.IO
Imports System.Text

Namespace ExcelOps

    ''' <summary>
    ''' Base implementation for common API for the several Excel engines
    ''' </summary>
    Public MustInherit Class ExcelDataOperationsBase

        Public Enum OpenMode As Byte
            OpenExistingFile = 0
            CreateFile = 1
        End Enum

        ''' <summary>
        ''' Create or open a workbook
        ''' </summary>
        ''' <param name="file"></param>
        ''' <param name="mode"></param>
        ''' <param name="autoCalculationOnLoad"></param>
        ''' <param name="calculationModuleDisabled"></param>
        ''' <param name="[readOnly]"></param>
        ''' <param name="passwordForOpening"></param>
        Protected Sub New(file As String, mode As OpenMode, autoCalculationOnLoad As Boolean, calculationModuleDisabled As Boolean, [readOnly] As Boolean, passwordForOpening As String)
            If autoCalculationOnLoad AndAlso calculationModuleDisabled Then Throw New ArgumentException("Calculation engine is disabled, but AutoCalculation requested", NameOf(autoCalculationOnLoad))
            Me.AutoCalculationOnLoad = autoCalculationOnLoad
            Me.CalculationModuleDisabled = calculationModuleDisabled
            Me.ReadOnly = [readOnly]
            Select Case mode
                Case OpenMode.OpenExistingFile
                    Me.PasswordForOpening = passwordForOpening
                    Me.LoadAndInitializeWorkbookFile(file)
                Case OpenMode.CreateFile
                    Me.CreateAndInitializeWorkbookFile(file)
                    Me.ReadOnly = [readOnly] OrElse (file = Nothing)
                    Me.PasswordForOpening = passwordForOpening
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(mode))
            End Select
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        ''' <param name="data"></param>
        ''' <param name="autoCalculationOnLoad"></param>
        ''' <param name="calculationModuleDisabled"></param>
        ''' <param name="passwordForOpening"></param>
        Protected Sub New(data As Byte(), autoCalculationOnLoad As Boolean, calculationModuleDisabled As Boolean, passwordForOpening As String)
            If autoCalculationOnLoad AndAlso calculationModuleDisabled Then Throw New ArgumentException("Calculation engine is disabled, but AutoCalculation requested", NameOf(autoCalculationOnLoad))
            Me.AutoCalculationOnLoad = autoCalculationOnLoad
            Me.CalculationModuleDisabled = calculationModuleDisabled
            Me.ReadOnly = True
            'OpenMode.OpenExistingFile
            Me.PasswordForOpening = passwordForOpening
            Me.LoadAndInitializeWorkbookFile(data)
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        ''' <param name="data"></param>
        ''' <param name="autoCalculationOnLoad"></param>
        ''' <param name="calculationModuleDisabled"></param>
        ''' <param name="passwordForOpening"></param>
        Protected Sub New(data As System.IO.Stream, autoCalculationOnLoad As Boolean, calculationModuleDisabled As Boolean, passwordForOpening As String)
            If autoCalculationOnLoad AndAlso calculationModuleDisabled Then Throw New ArgumentException("Calculation engine is disabled, but AutoCalculation requested", NameOf(autoCalculationOnLoad))
            Me.AutoCalculationOnLoad = autoCalculationOnLoad
            Me.CalculationModuleDisabled = calculationModuleDisabled
            Me.ReadOnly = True
            'OpenMode.OpenExistingFile
            Me.PasswordForOpening = passwordForOpening
            Me.LoadAndInitializeWorkbookFile(data)
        End Sub

        ''' <summary>
        ''' Create a new instance for accessing Excel workbooks (still requires creating or loading of a workbook)
        ''' </summary>
        ''' <param name="autoCalculationOnLoad">Automatically do a full recalculation after workbook has been loaded</param>
        ''' <param name="calculationModuleDisabled">Disables the Excel calculation engine</param>
        Protected Sub New(autoCalculationOnLoad As Boolean, calculationModuleDisabled As Boolean, [readOnly] As Boolean, passwordForOpening As String)
            If autoCalculationOnLoad AndAlso calculationModuleDisabled Then Throw New ArgumentException("Calculation engine is disabled, but AutoCalculation requested", NameOf(autoCalculationOnLoad))
            Me.AutoCalculationOnLoad = autoCalculationOnLoad
            Me.CalculationModuleDisabled = calculationModuleDisabled
            Me.ReadOnly = [readOnly]
            Me.PasswordForOpening = passwordForOpening
        End Sub

        ''' <summary>
        ''' Reload a file from disk
        ''' </summary>
        Public Sub ReloadFromFile()
            Me.LoadAndInitializeWorkbookFile(Me.FilePath)
        End Sub

        ''' <summary>
        ''' A password for opening an excel file
        ''' </summary>
        ''' <returns></returns>
        Public Property PasswordForOpening As String

        ''' <summary>
        ''' Write protection for this filename prevents Save, but still allows SaveAs
        ''' </summary>
        ''' <returns></returns>
        Public Property [ReadOnly] As Boolean

        ''' <summary>
        ''' The calculation module of involved Excel engine might be disabled due to insufficiency/incompleteness of 3rd party Excel (calculation) engines (except for single cell calculations)
        ''' </summary>
        ''' <returns></returns>
        Public Property CalculationModuleDisabled As Boolean

        ''' <summary>
        ''' If enabled, the calculation engine will do a full recalculation after loading a workbook
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property AutoCalculationOnLoad As Boolean

        ''' <summary>
        ''' If enabled, the calculation engine will do a full recalculation after every modification
        ''' </summary>
        ''' <returns></returns>
        Public MustOverride Property AutoCalculationEnabled As Boolean

        <CodeAnalysis.SuppressMessage("Design", "CA1051:Sichtbare Instanzfelder nicht deklarieren")>
        Protected _FilePath As String
        ''' <summary>
        ''' The file path as initialized in constructor (applies for saved files as well as for created, not-saved files with their intended file location on 1st save)
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property FilePath As String
            Get
                Return _FilePath
            End Get
        End Property

        ''' <summary>
        ''' Close the current worksbook (without saving)
        ''' </summary>
        Public MustOverride Sub Close()

        ''' <summary>
        ''' Has the workbook already been closed
        ''' </summary>
        ''' <returns></returns>
        Public MustOverride ReadOnly Property IsClosed As Boolean

        ''' <summary>
        ''' Close the external Excel engine application (if applicable)
        ''' </summary>
        Public MustOverride Sub CloseExcelAppInstance()

        ''' <summary>
        ''' The current workbook file name (as it is known by the Excel engine)
        ''' </summary>
        ''' <returns>Null/Nothing if the file has been created in memory, but hasn't been saved OR the file name from last open/save action</returns>
        Protected Friend MustOverride ReadOnly Property WorkbookFilePath As String

        ''' <summary>
        ''' Save modifications made to the workbook
        ''' </summary>
        Public Sub Save()
            Me.Save(SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
        End Sub

        ''' <summary>
        ''' Save modifications made to the workbook
        ''' </summary>
        Public Sub Save(cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            If Me.ReadOnly = True Then
                Throw New FileReadOnlyException("File is read-only and can't be saved at same location")
            End If
            If FilePath.ToLowerInvariant.EndsWith(".xlsx", False, System.Globalization.CultureInfo.InvariantCulture) AndAlso Me.HasVbaProject Then
                Throw New NotSupportedException("VBA projects are not supported for .xlsx files, run RemoveVbaProject() method, first")
            End If
            If FilePath.ToLowerInvariant.EndsWith(".xlsx", False, System.Globalization.CultureInfo.InvariantCulture) Then 'remove any last bit of a VBA project (HasVbaModule is not 100% sure)
                Me.RemoveVbaProject()
            End If
            If Me.RecalculationRequired Then Me.RecalculateAll()
            If Me.FilePath <> Nothing AndAlso Me.WorkbookFilePath = Nothing Then
                'Created workbook, initial save must provide a file path, so use SaveAs method instead
                Me.SaveAs(Me.FilePath, cachedCalculationsOption)
            Else
                Me.SaveInternal_ApplyCachedCalculationOption(cachedCalculationsOption)
                Dim AutoCalcBuffer As Boolean = Me.AutoCalculationEnabled
                Try
                    Me.AutoCalculationEnabled = True
                    Me.SaveInternal()
                Finally
                    Me.AutoCalculationEnabled = AutoCalcBuffer
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Apply CachedCalculation setting
        ''' </summary>
        ''' <param name="cachedCalculationsOption"></param>
        <CodeAnalysis.SuppressMessage("Naming", "CA1707:Bezeichner dürfen keine Unterstriche enthalten")>
        Protected Overridable Sub SaveInternal_ApplyCachedCalculationOption(cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            If cachedCalculationsOption = SaveOptionsForDisabledCalculationEngines.DefaultBehaviour Then
                cachedCalculationsOption = SaveOptionsForDisabledCalculationEngines.NoReset
            End If
            Select Case cachedCalculationsOption
                Case SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation
                    'do nothing
                Case SaveOptionsForDisabledCalculationEngines.ResetCalculatedValuesForForcedCellRecalculationIfRecalculationRequired
                   'do nothing
                Case SaveOptionsForDisabledCalculationEngines.NoReset
                    'do nothing
                Case Else
                    Throw New NotImplementedException("Invalid option: " & cachedCalculationsOption)
            End Select
        End Sub

        ''' <summary>
        ''' Save modifications made to the workbook
        ''' </summary>
        Protected MustOverride Sub SaveInternal()

        ''' <summary>
        ''' Save workbook as another file
        ''' </summary>
        ''' <param name="filePath"></param>
        <Obsolete("Use overloaded method", True)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub SaveAs(filePath As String)
            Me.SaveAs(filePath, SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
        End Sub

        ''' <summary>
        ''' Save workbook as another file
        ''' </summary>
        ''' <param name="filePath"></param>
        Public Sub SaveAs(filePath As String, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            If Me.ReadOnly = True AndAlso Me._FilePath = filePath AndAlso Me.WorkbookFilePath <> Nothing Then
                Throw New FileReadOnlyException("File """ & filePath & """ is read-only and can't be saved at same location")
            End If
            If filePath.ToLowerInvariant.EndsWith(".xlsx", False, System.Globalization.CultureInfo.InvariantCulture) AndAlso Me.HasVbaProject Then
                Throw New NotSupportedException("VBA projects are not supported for .xlsx files, run RemoveVbaProject() method, first")
            End If
            If filePath.ToLowerInvariant.EndsWith(".xlsx", False, System.Globalization.CultureInfo.InvariantCulture) Then 'remove any last bit of a VBA project (HasVbaModule is not 100% sure)
                Me.RemoveVbaProject()
            End If
            If Me.RecalculationRequired AndAlso Me.CalculationModuleDisabled = False Then Me.RecalculateAll()

            Me.SaveInternal_ApplyCachedCalculationOption(cachedCalculationsOption)
            Dim AutoCalcBuffer As Boolean = Me.AutoCalculationEnabled
            Try
                Me.AutoCalculationEnabled = True
                Me.SaveAsInternal(filePath, cachedCalculationsOption)
            Finally
                Me.AutoCalculationEnabled = AutoCalcBuffer
            End Try

            Me._FilePath = filePath
            Me.ReadOnly = False
        End Sub

        Public Enum SaveOptionsForDisabledCalculationEngines As Byte
            <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)> DefaultBehaviour = 0
            ''' <summary>
            ''' No reset of cached calculation values of formula cells
            ''' </summary>
            NoReset = 1
            ''' <summary>
            ''' Reset cached calculation values of formula cells
            ''' </summary>
            AlwaysResetCalculatedValuesForForcedCellRecalculation = 2
            ''' <summary>
            ''' Reset cached calculation values of formula cells if a recalculation is requested
            ''' </summary>
            ResetCalculatedValuesForForcedCellRecalculationIfRecalculationRequired = 3
        End Enum

        Protected MustOverride Sub SaveAsInternal(fileName As String, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)

        ''' <summary>
        ''' All available sheet names (work sheets + chart sheets)
        ''' </summary>
        ''' <returns></returns>
        Public MustOverride Function SheetNames() As List(Of String)

        ''' <summary>
        ''' Lookup the (zero-based) index number of a sheet
        ''' </summary>
        ''' <param name="sheetName">A sheet name (work sheet or chart sheet)</param>
        ''' <returns>-1 if the sheet name doesn't exist, otherwise its index value</returns>
        Public Overridable Function SheetIndex(ByVal sheetName As String) As Integer
            Return Me.SheetNames.IndexOf(sheetName)
        End Function

        ''' <summary>
        ''' All available work sheet names
        ''' </summary>
        ''' <returns></returns>
        Public MustOverride Function WorkSheetNames() As List(Of String)

        ''' <summary>
        ''' Lookup the (zero-based) index number of a work sheet
        ''' </summary>
        ''' <param name="workSheetName">A work sheet name</param>
        ''' <returns>-1 if the sheet name doesn't exist, otherwise its index value</returns>
        Public Function WorkSheetIndex(ByVal workSheetName As String) As Integer
            Return Me.WorkSheetNames.IndexOf(workSheetName)
        End Function

        ''' <summary>
        ''' All available chart sheet names
        ''' </summary>
        ''' <returns></returns>
        Public MustOverride Function ChartSheetNames() As List(Of String)

        ''' <summary>
        ''' Lookup the (zero-based) index number of a chart sheet
        ''' </summary>
        ''' <param name="chartName">A work sheet name</param>
        ''' <returns>-1 if the sheet name doesn't exist, otherwise its index value</returns>
        Public Function ChartSheetIndex(ByVal chartName As String) As Integer
            Return Me.ChartSheetNames.IndexOf(chartName)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public MustOverride Function LookupCellValue(Of T)(cell As ExcelCell) As T

        ''' <summary>
        ''' Read several cell values
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="range"></param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public Function LookupCellValue(Of T)(range As ExcelRange) As List(Of T)
            Dim Result As New List(Of T)
            For Each Cell As ExcelCell In range
                Result.Add(LookupCellValue(Of T)(Cell))
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public MustOverride Function LookupCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public MustOverride Function TryLookupCellValue(Of T As Structure)(cell As ExcelCell) As T?

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public MustOverride Function TryLookupCellValue(Of T As Structure)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T?

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public MustOverride Function LookupCellValueAsObject(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Object

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public Function LookupCellValueAsObject(cell As ExcelCell) As Object
            If cell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(cell)
            Return Me.LookupCellValueAsObject(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Public MustOverride Function LookupCellFormula(cell As ExcelCell) As String

        ''' <summary>
        ''' Read a cell formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        Public MustOverride Function LookupCellFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String

        Public MustOverride Function LookupCellIsLocked(cell As ExcelCell) As Boolean

        Public MustOverride Function LookupCellIsLocked(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean

        ''' <summary>
        ''' Write a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <param name="value"></param>
        Public MustOverride Sub WriteCellValue(Of T)(cell As ExcelCell, value As T)

        ''' <summary>
        ''' Write a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <param name="value"></param>
        Public MustOverride Sub WriteCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer, value As T)

        ''' <summary>
        ''' Write a cell formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <param name="formula">Formula without leading '=' char</param>
        Public MustOverride Sub WriteCellFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer, formula As String, immediatelyCalculateCellValue As Boolean)

        Private _RecalculationRequired As TriState = TriState.UseDefault
        ''' <summary>
        ''' Modifications require a full recalculation
        ''' </summary>
        ''' <returns></returns>
        Public Property RecalculationRequired As Boolean
            Get
                If _RecalculationRequired = TriState.UseDefault Then
                    Me.RecalculationRequired = False 'Defaults to false
                End If
                If _RecalculationRequired = TriState.True Then
                    Return True
                Else
                    Return False
                End If
            End Get
            Set(value As Boolean)
                If value = True Then 'then AndAlso Me.AutoCalculationEnabled = False Then
                    _RecalculationRequired = TriState.True
                Else
                    'value=False OR 
                    'sub module requests calculation, but is already done by Excel engine automatically
                    _RecalculationRequired = TriState.False
                End If
            End Set
        End Property

        Protected MustOverride Sub LoadWorkbook(file As System.IO.FileInfo)

        Protected Sub LoadAndInitializeWorkbookFile(inputPath As String)
            If inputPath = Nothing Then Throw New ArgumentNullException(NameOf(inputPath))
            '1st, close an exsting workbook instance
            If Me.IsClosed = False Then Me.Close()
            'Load the changed worksheet
            Me._FilePath = inputPath
            Dim file As New System.IO.FileInfo(inputPath)
            If file.Exists = False Then
                Throw New System.IO.FileNotFoundException("Missing file: " & file.ToString, file.ToString)
            End If
            Me.LoadWorkbook(file)
            Me.AutoCalculationEnabled = False
            If Me.AutoCalculationOnLoad Then
                Me.RecalculateAll()
            End If
        End Sub

        Protected MustOverride Sub LoadWorkbook(data As Byte())

        Protected Sub LoadAndInitializeWorkbookFile(data As Byte())
            '1st, close an exsting workbook instance
            If Me.IsClosed = False Then Me.Close()
            'Load the changed worksheet
            Me._FilePath = Nothing
            Me.LoadWorkbook(data)
            Me.AutoCalculationEnabled = False
            If Me.AutoCalculationOnLoad Then
                Me.RecalculateAll()
            End If
        End Sub

        Protected MustOverride Sub LoadWorkbook(data As System.IO.Stream)

        Protected Sub LoadAndInitializeWorkbookFile(data As System.IO.Stream)
            '1st, close an exsting workbook instance
            If Me.IsClosed = False Then Me.Close()
            'Load the changed worksheet
            Me._FilePath = Nothing
            Me.LoadWorkbook(data)
            Me.AutoCalculationEnabled = False
            If Me.AutoCalculationOnLoad Then
                Me.RecalculateAll()
            End If
        End Sub

        Protected MustOverride Sub CreateWorkbook()

        ''' <summary>
        ''' Create a new workbook
        ''' </summary>
        ''' <param name="intendedFilePath">If the file path is already known, the file will be checked to not exist already and the file path will be used for later saving</param>
        Protected Sub CreateAndInitializeWorkbookFile(intendedFilePath As String)
            'Load the changed worksheet
            If intendedFilePath <> Nothing Then
                Me._FilePath = intendedFilePath
                Dim file As New System.IO.FileInfo(intendedFilePath)
                If file.Exists = True Then
                    Throw New FileAlreadyExistsException(Me.FilePath)
                End If
            Else
                Me._FilePath = Nothing
            End If
            Me.CreateWorkbook()
            Me.AutoCalculationEnabled = False
            If Me.AutoCalculationOnLoad Then
                Me.RecalculateAll()
            End If
        End Sub

        Public MustOverride Sub CleanupRangeNames()

        ''' <summary>
        ''' Lookup the last content column index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Function LookupLastContentColumnIndex(sheetName As String) As Integer
            Return Me.LookupLastContentColumnIndex(sheetName, Me.FindLastMergedCellNonEmpty(sheetName))
        End Function

        ''' <summary>
        ''' Lookup the last content column index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Protected Overridable Function LookupLastContentColumnIndex(sheetName As String, lastMergedCell As ExcelCell) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            Dim autoSuggestionLastRowIndex As Integer = LastCell.RowIndex
            Dim autoSuggestedResult As Integer = LastCell.ColumnIndex
            Dim lastMergeCellRowIndex As Integer
            Dim lastMergeCellColumnIndex As Integer
            If lastMergedCell IsNot Nothing Then
                lastMergeCellRowIndex = lastMergedCell.RowIndex
                lastMergeCellColumnIndex = lastMergedCell.ColumnIndex
            Else
                lastMergeCellRowIndex = 0
                lastMergeCellColumnIndex = 0
            End If
            'Find last content cell
            For colCounter As Integer = autoSuggestedResult To lastMergeCellColumnIndex Step -1
                For rowCounter As Integer = lastMergeCellRowIndex To autoSuggestionLastRowIndex
                    If IsEmptyCell(sheetName, rowCounter, colCounter) = False Then
                        Return System.Math.Max(lastMergeCellColumnIndex, colCounter)
                    End If
                Next
            Next
            Return System.Math.Max(lastMergeCellColumnIndex, 0)
        End Function

        ''' <summary>
        ''' Lookup the last content row index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overridable Function LookupLastContentRowIndex(sheetName As String) As Integer
            Return Me.LookupLastContentRowIndex(sheetName, Me.FindLastMergedCellNonEmpty(sheetName))
        End Function

        ''' <summary>
        ''' Lookup the last content row index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Protected Overridable Function LookupLastContentRowIndex(sheetName As String, lastMergedCell As ExcelCell) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            Dim autoSuggestionLastColumnIndex As Integer = LastCell.ColumnIndex
            Dim autoSuggestedResult As Integer = LastCell.RowIndex
            Dim lastMergeCellRowIndex As Integer
            Dim lastMergeCellColumnIndex As Integer
            If lastMergedCell IsNot Nothing Then
                lastMergeCellRowIndex = lastMergedCell.RowIndex
                lastMergeCellColumnIndex = lastMergedCell.ColumnIndex
            Else
                lastMergeCellRowIndex = 0
                lastMergeCellColumnIndex = 0
            End If
            'Find last content cell
            For rowCounter As Integer = autoSuggestedResult To lastMergeCellRowIndex Step -1
                For colCounter As Integer = lastMergeCellColumnIndex To autoSuggestionLastColumnIndex
                    If IsEmptyCell(sheetName, rowCounter, colCounter) = False Then
                        Return System.Math.Max(lastMergeCellRowIndex, rowCounter)
                    End If
                Next
            Next
            Return System.Math.Max(lastMergeCellRowIndex, 0)
        End Function

        ''' <summary>
        ''' Lookup the last content cell (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <remarks>Please note: there might be a performance impact (especially with MS Excel interop) in comparison to <see cref="LookupLastCell(String)"/> to check all relevant cells due to required COM overhead</remarks>
        Public Function LookupLastContentCell(sheetName As String) As ExcelCell
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim LastMergedCellsNonEmpty As ExcelCell = Me.FindLastMergedCellNonEmpty(sheetName)
            Dim CellRowIndex As Integer = Me.LookupLastContentRowIndex(sheetName, LastMergedCellsNonEmpty)
            Dim CellColIndex As Integer = Me.LookupLastContentColumnIndex(sheetName, LastMergedCellsNonEmpty)
            Return New ExcelOps.ExcelCell(sheetName, CellRowIndex, CellColIndex, Nothing)
        End Function

        ''' <summary>
        ''' Lookup the last column index (zero based index) (the last content cell equals to Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Function LookupLastColumnIndex(sheetName As String) As Integer
            Return Me.LookupLastCell(sheetName).ColumnIndex
        End Function

        ''' <summary>
        ''' Lookup the last row index (zero based index) (the last content cell equals to Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Function LookupLastRowIndex(sheetName As String) As Integer
            Return Me.LookupLastCell(sheetName).RowIndex
        End Function

        ''' <summary>
        ''' Lookup the last cell (the last content cell equals to Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Function LookupLastCell(sheetName As String) As ExcelCell

        ''' <summary>
        ''' Lookup the first unlocked cell (search row by row, then column by column)
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <returns></returns>
        Public Function LookupFirstUnlockedCell(sheetName As String) As ExcelCell
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            Dim FoundFirstUnlockedCell As ExcelCell = Nothing
            For MyRowCounterIndex As Integer = 0 To LastCell.RowIndex
                For MyColCounterIndex As Integer = 0 To LastCell.ColumnIndex
                    If Me.LookupCellIsLocked(sheetName, MyRowCounterIndex, MyColCounterIndex) = False Then
                        FoundFirstUnlockedCell = New ExcelCell(sheetName, MyRowCounterIndex, MyColCounterIndex, ExcelCell.ValueTypes.All)
                    End If
                    If FoundFirstUnlockedCell IsNot Nothing Then
                        Exit For
                    End If
                Next
                If FoundFirstUnlockedCell IsNot Nothing Then
                    Exit For
                End If
            Next
            Return FoundFirstUnlockedCell
        End Function

        ''' <summary>
        ''' Lookup the first unlocked cell (search row by row, then column by column) or alternatively the first cell (A1)
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <returns></returns>
        Public Function LookupFirstUnlockedCellOrFirstCell(sheetName As String) As ExcelCell
            Dim FoundFirstUnlockedCell As ExcelCell = Me.LookupFirstUnlockedCell(sheetName)
            If FoundFirstUnlockedCell IsNot Nothing Then
                Return FoundFirstUnlockedCell
            Else
                Return New ExcelCell(sheetName, "A1", ExcelCell.ValueTypes.All)
            End If
        End Function

        ''' <summary>
        ''' Lookup the row index (zero based index)
        ''' </summary>
        ''' <param name="cell"></param>
        Public MustOverride Function LookupRowIndex(cell As ExcelOps.ExcelCell) As Integer

        ''' <summary>
        ''' Lookup the column index (zero based index)
        ''' </summary>
        ''' <param name="cell"></param>
        Public MustOverride Function LookupColumnIndex(cell As ExcelOps.ExcelCell) As Integer

        ''' <summary>
        ''' Remove specified rows
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="startrowIndex">0-based row number</param>
        ''' <param name="rows">Number of rows to remove</param>
        Public MustOverride Sub RemoveRows(sheetName As String, startRowIndex As Integer, rows As Integer)

        ''' <summary>
        ''' Remove a sheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Sub RemoveSheet(sheetName As String)

        Public Sub AddSheet(sheetName As String)
            Me.AddSheet(sheetName, Nothing)
        End Sub

        Public MustOverride Sub AddSheet(sheetName As String, beforeSheetName As String)

        ''' <summary>
        ''' Determine if a cell contains empty content
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">Zero-based index</param>
        ''' <param name="columnIndex">Zero-based index</param>
        Public MustOverride Function IsEmptyCell(sheetName As String, ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Boolean

        ''' <summary>
        ''' Determine if a cell contains empty content
        ''' </summary>
        Public Function IsEmptyCell(cell As ExcelCell) As Boolean
            If cell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(cell)
            Return Me.IsEmptyCell(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <summary>
        ''' Recalculate everything
        ''' </summary>
        Public Sub RecalculateAll()
            Me.RecalculateAllInternal()
            Me.RecalculationRequired = False
        End Sub

        ''' <summary>
        ''' Recalculate everything
        ''' </summary>
        Protected MustOverride Sub RecalculateAllInternal()

        ''' <summary>
        ''' Recalculate all cells of a sheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Sub RecalculateSheet(sheetName As String)

        ''' <summary>
        ''' Recalculate a cell based on its formula
        ''' </summary>
        ''' <param name="cell"></param>
        Public Sub RecalculateCell(cell As ExcelCell)
            If cell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(cell)
            Me.RecalculateCell(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Sub

        ''' <summary>
        ''' Recalculate a cell based on its formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        Public Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.RecalculateCell(sheetName, rowIndex, columnIndex, True)
        End Sub

        ''' <summary>
        ''' Recalculate a cell based on its formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        Public MustOverride Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer, throwExceptionOnCalculationError As Boolean)

        ''' <summary>
        ''' Try to lookup the cell's value to a string anyhow
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public MustOverride Function LookupCellFormattedText(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String

        ''' <summary>
        ''' Try to lookup the cell's value to a string anyhow
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LookupCellFormattedText(cell As ExcelCell) As String
            If cell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(cell)
            Return Me.LookupCellFormattedText(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <summary>
        ''' Read the cell format string
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Public MustOverride Function LookupCellFormat(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String

        ''' <summary>
        ''' Read the cell format string
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LookupCellFormat(cell As ExcelCell) As String
            If cell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(cell)
            Return Me.LookupCellFormat(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <summary>
        ''' Unprotect all sheets
        ''' </summary>
        Public Sub UnprotectSheets()
            For Each sheetName As String In Me.SheetNames
                Me.UnprotectSheet(sheetName)
            Next
        End Sub

        ''' <summary>
        ''' Unprotect all sheets
        ''' </summary>
        Public Sub UnhideSheets()
            For Each sheetName As String In Me.SheetNames
                Me.UnhideSheet(sheetName)
            Next
        End Sub

        ''' <summary>
        ''' Unprotect selected sheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Sub UnprotectSheet(sheetName As String)

        Public Enum ProtectionLevel As Byte
            ''' <summary>
            ''' Select all cells, edit unlocked cells
            ''' </summary>
            Standard = 20
            ''' <summary>
            ''' Select and edit unlocked cells
            ''' </summary>
            SelectAndEditUnlockedCellsOnly = 100
            ''' <summary>
            ''' Select all cells, edit unlocked cells, insert/remove rows
            ''' </summary>
            StandardWithInsertDeleteRows = 18
            ''' <summary>
            ''' Select unlocked and locked cells, edit their formulas/cell content, but never edit formattings, objects, row set, column set, pivot/charts, insert links or anything similar
            ''' </summary>
            SelectAndEditAllCellsButNoFurtherEditing = 101
            ''' <summary>
            ''' No selection of unlocked or locked cells, and never review/edit their formulas/cell content, formattings, objects, row set, column set, pivot/charts, insert links or anything similar
            ''' </summary>
            SelectNoCellsAndNoEditing = 200
        End Enum

        ''' <summary>
        ''' Protect selected sheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Sub ProtectSheet(sheetName As String, level As ProtectionLevel)

        ''' <summary>
        ''' Is a sheet in protected state
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <returns></returns>
        Public MustOverride Function IsProtectedSheet(sheetName As String) As Boolean

        ''' <summary>
        ''' Unprotect selected sheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Sub UnhideSheet(sheetName As String)

        ''' <summary>
        ''' Protect selected sheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Sub HideSheet(sheetName As String)
            Me.HideSheet(sheetName, True)
        End Sub

        ''' <summary>
        ''' Protect selected sheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Sub HideSheet(sheetName As String, stronglyHide As Boolean)

        ''' <summary>
        ''' Is a sheet in protected state
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <returns></returns>
        Public MustOverride Function IsHiddenSheet(sheetName As String) As Boolean

        Public Enum MatrixContent As Byte
            ''' <summary>
            ''' Static values only
            ''' </summary>
            StaticValues = 0
            ''' <summary>
            ''' Formulas only
            ''' </summary>
            Formulas = 1
            ''' <summary>
            ''' Values of locked cells only
            ''' </summary>
            ValuesOfLockedCells = 2
            ''' <summary>
            ''' Values of unlocked cells only
            ''' </summary>
            ValuesOfUnlockedCells = 6
            ''' <summary>
            ''' The cell's formula or the cell's value
            ''' </summary>
            FormulaOrFormattedText = 3
            ''' <summary>
            ''' Static values or the value of the last calculation in case of a formula
            ''' </summary>
            StaticOrCalculatedValues = 4
            ''' <summary>
            ''' Static values or the value of the last calculation in case of a formula
            ''' </summary>
            FormattedText = 5
            ''' <summary>
            ''' Errors only
            ''' </summary>
            Errors = 7
        End Enum

        Public Function SheetContentMatrix(sheetName As String, contentType As MatrixContent) As TextTable
            Dim Result As New TextTable
            Result.AddColumns(Me.LookupLastContentColumnIndex(sheetName) + 1)
            Result.AddRows(Me.LookupLastContentRowIndex(sheetName) + 1)
            For MyRowCounter As Integer = 0 To Result.RowCount - 1
                For MyColCounter As Integer = 0 To Result.ColumnCount - 1
                    Select Case contentType
                        Case MatrixContent.StaticValues
                            If Me.LookupCellFormula(sheetName, MyRowCounter, MyColCounter) = Nothing Then
                                Result.Cell(MyRowCounter, MyColCounter) = Me.LookupCellFormattedText(sheetName, MyRowCounter, MyColCounter)
                            End If
                        Case MatrixContent.Formulas
                            If Me.LookupCellFormula(sheetName, MyRowCounter, MyColCounter) <> Nothing Then
                                Result.Cell(MyRowCounter, MyColCounter) = "=" & Me.LookupCellFormula(sheetName, MyRowCounter, MyColCounter)
                            End If
                        Case MatrixContent.ValuesOfLockedCells
                            If Me.LookupCellIsLocked(sheetName, MyRowCounter, MyColCounter) Then
                                Result.Cell(MyRowCounter, MyColCounter) = Me.LookupCellFormattedText(sheetName, MyRowCounter, MyColCounter)
                            End If
                        Case MatrixContent.ValuesOfUnlockedCells
                            If Not Me.LookupCellIsLocked(sheetName, MyRowCounter, MyColCounter) Then
                                Result.Cell(MyRowCounter, MyColCounter) = Me.LookupCellFormattedText(sheetName, MyRowCounter, MyColCounter)
                            End If
                        Case MatrixContent.FormulaOrFormattedText
                            If Me.LookupCellFormula(sheetName, MyRowCounter, MyColCounter) <> Nothing Then
                                Result.Cell(MyRowCounter, MyColCounter) = "=" & Me.LookupCellFormula(sheetName, MyRowCounter, MyColCounter)
                            Else
                                Result.Cell(MyRowCounter, MyColCounter) = Me.LookupCellFormattedText(sheetName, MyRowCounter, MyColCounter)
                            End If
                        Case MatrixContent.StaticOrCalculatedValues
                            Result.Cell(MyRowCounter, MyColCounter) = Me.LookupCellFormattedText(sheetName, MyRowCounter, MyColCounter)
                        Case MatrixContent.FormattedText
                            Result.Cell(MyRowCounter, MyColCounter) = Me.LookupCellFormattedText(sheetName, MyRowCounter, MyColCounter)
                        Case MatrixContent.Errors
                            Result.Cell(MyRowCounter, MyColCounter) = Me.LookupCellErrorValue(sheetName, MyRowCounter, MyColCounter)
                        Case Else
                            Throw New NotImplementedException(contentType.ToString)
                    End Select
                Next
            Next
            Result.AutoTrim()
            Return Result
        End Function

        Public Function LookupCellErrorValue(cell As ExcelCell) As String
            If cell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(cell)
            Return Me.LookupCellErrorValue(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <summary>
        ''' Check for errors in cell
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns>Null/Nothing is case of no error, otherwise error code as text, e.g. #VALUE!</returns>
        Public MustOverride Function LookupCellErrorValue(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String

        ''' <summary>
        ''' Clear cells
        ''' </summary>
        ''' <param name="range"></param>
        Public Sub ClearCells(range As ExcelRange)
            Me.ClearCells(range.SheetName, range.AddressStart, range.AddressEnd)
        End Sub

        ''' <summary>
        ''' Clear cells
        ''' </summary>
        ''' <param name="range"></param>
        Public Sub ClearCells(overrideSheetName As String, range As ExcelRange)
            Dim ClearingRange As ExcelRange = range.Clone
            ClearingRange.SheetName = overrideSheetName
            Me.ClearCells(ClearingRange)
        End Sub

        ''' <summary>
        ''' Clear cells
        ''' </summary>
        ''' <param name="overrideSheetName"></param>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public MustOverride Sub ClearCells(overrideSheetName As String, rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)

        ''' <summary>
        ''' Clear cells
        ''' </summary>
        ''' <param name="overrideSheetName"></param>
        ''' <param name="cell"></param>
        Public Sub ClearCells(overrideSheetName As String, cell As ExcelCell)
            Me.ClearCells(overrideSheetName, cell, cell)
        End Sub

        ''' <summary>
        ''' Clear cells
        ''' </summary>
        ''' <param name="cell"></param>
        Public Sub ClearCells(cell As ExcelCell)
            If cell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(cell)
            Me.ClearCells(cell.SheetName, cell, cell)
        End Sub

        ''' <summary>
        ''' Clear cells
        ''' </summary>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public Sub ClearCells(rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)
            If rangeFirstCell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(rangeFirstCell)
            If rangeLastCell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(rangeLastCell)
            If rangeFirstCell.SheetName <> rangeLastCell.SheetName Then Throw New ArgumentException("Cells must be member of the same worksheet")
            If rangeFirstCell.SheetName = Nothing Then Throw New ArgumentNullException(NameOf(rangeFirstCell))
            Me.ClearCells(rangeFirstCell.SheetName, rangeFirstCell, rangeLastCell)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="range"></param>
        Public Sub ClearCellContent(overrideSheetName As String, range As ExcelRange)
            Dim ClearingRange As ExcelRange = range.Clone
            ClearingRange.SheetName = overrideSheetName
            Me.ClearCellContent(ClearingRange)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="range"></param>
        Public Sub ClearCellContent(overrideSheetName As String, range As ExcelRange, onlyIfUnlockedCell As Boolean)
            Dim ClearingRange As ExcelRange = range.Clone
            ClearingRange.SheetName = overrideSheetName
            Me.ClearCellContent(ClearingRange, onlyIfUnlockedCell)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="cell"></param>
        Public Sub ClearCellContent(cell As ExcelCell)
            If cell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(cell)
            Me.WriteCellValue(Of Object)(cell, Nothing)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="cell"></param>
        Public Sub ClearCellContent(cell As ExcelCell, onlyIfUnlockedCell As Boolean)
            If cell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(cell)
            If onlyIfUnlockedCell = False OrElse (onlyIfUnlockedCell = True AndAlso Me.LookupCellIsLocked(cell) = False) Then
                Me.WriteCellValue(Of Object)(cell, Nothing)
            End If
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="overrideSheetName"></param>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public Sub ClearCellContent(overrideSheetName As String, rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)
            If overrideSheetName = Nothing Then Throw New ArgumentNullException(NameOf(overrideSheetName))
#Disable Warning IDE0017 ' Initialisierung von Objekten vereinfachen
            Dim TargetRange As New ExcelRange(rangeFirstCell.Clone(ExcelCell.ValueTypes.All), rangeLastCell.Clone(ExcelCell.ValueTypes.All))
#Enable Warning IDE0017 ' Initialisierung von Objekten vereinfachen
            TargetRange.SheetName = overrideSheetName
            Me.ClearCellContent(TargetRange)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="overrideSheetName"></param>
        ''' <param name="cell"></param>
        Public Sub ClearCellContent(overrideSheetName As String, cell As ExcelCell, onlyIfUnlockedCell As Boolean)
            If overrideSheetName = Nothing Then Throw New ArgumentNullException(NameOf(overrideSheetName))
#Disable Warning IDE0017 ' Initialisierung von Objekten vereinfachen
            Dim TargetRange As New ExcelRange(cell, cell)
#Enable Warning IDE0017 ' Initialisierung von Objekten vereinfachen
            TargetRange.SheetName = overrideSheetName
            Me.ClearCellContent(TargetRange, onlyIfUnlockedCell)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="overrideSheetName"></param>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public Sub ClearCellContent(overrideSheetName As String, rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell, onlyIfUnlockedCell As Boolean)
            If overrideSheetName = Nothing Then Throw New ArgumentNullException(NameOf(overrideSheetName))
#Disable Warning IDE0017 ' Initialisierung von Objekten vereinfachen
            Dim TargetRange As New ExcelRange(rangeFirstCell.Clone(ExcelCell.ValueTypes.All), rangeLastCell.Clone(ExcelCell.ValueTypes.All))
#Enable Warning IDE0017 ' Initialisierung von Objekten vereinfachen
            TargetRange.SheetName = overrideSheetName
            Me.ClearCellContent(TargetRange, onlyIfUnlockedCell)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public Sub ClearCellContent(rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)
            If rangeFirstCell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(rangeFirstCell)
            If rangeLastCell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(rangeLastCell)
            If rangeFirstCell.SheetName <> rangeLastCell.SheetName Then Throw New ArgumentException("Cells must be member of the same worksheet")
            If rangeFirstCell.SheetName = Nothing Then Throw New ArgumentNullException(NameOf(rangeFirstCell))
            Me.ClearCellContent(rangeFirstCell.SheetName, rangeFirstCell, rangeLastCell)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public Sub ClearCellContent(rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell, onlyIfUnlockedCell As Boolean)
            If rangeFirstCell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(rangeFirstCell)
            If rangeLastCell.ValidateFullCellAddressInclSheetName() = False Then Throw New ExcelOps.InvalidCellAddressException(rangeLastCell)
            If rangeFirstCell.SheetName <> rangeLastCell.SheetName Then Throw New ArgumentException("Cells must be member of the same worksheet")
            If rangeFirstCell.SheetName = Nothing Then Throw New ArgumentNullException(NameOf(rangeFirstCell))
            Me.ClearCellContent(rangeFirstCell.SheetName, rangeFirstCell, rangeLastCell, onlyIfUnlockedCell)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="range"></param>
        Public Sub ClearCellContent(range As ExcelRange)
            For Each cell In range
                Me.ClearCellContent(cell)
            Next
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="range"></param>
        Public Sub ClearCellContent(range As ExcelRange, onlyIfUnlockedCell As Boolean)
            For Each cell In range
                Me.ClearCellContent(cell, onlyIfUnlockedCell)
            Next
        End Sub

        ''' <summary>
        ''' Clear cells
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Sub ClearSheet(sheetName As String)

        ''' <summary>
        ''' The currently selected sheet name
        ''' </summary>
        ''' <returns></returns>
        Public MustOverride Function SelectedSheetName() As String

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Sub SelectSheet(sheetName As String)

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetIndex"></param>
        Public MustOverride Sub SelectSheet(sheetIndex As Integer)

        Public MustOverride Sub CopySheetContentInternal(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)

        Public Sub CopySheetContent(sheetName As String, targetWorkbook As ExcelDataOperationsBase)
            Me.CopySheetContent(sheetName, targetWorkbook, sheetName, CopySheetOption.TargetSheetMustNotExist)
        End Sub

        Public Enum CopySheetOption As Byte
            TargetSheetMustNotExist = 0
            TargetSheetMustExist = 1
            TargetSheetMightExist = 2
        End Enum

        Public Sub CopySheetContent(sheetName As String, targetWorkbook As ExcelDataOperationsBase, copyOption As CopySheetOption)
            Me.CopySheetContent(sheetName, targetWorkbook, sheetName, copyOption)
        End Sub

        Public Sub CopySheetContent(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String, copyOption As CopySheetOption)
            If sheetName Is Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If targetWorkbook Is Nothing Then Throw New ArgumentNullException(NameOf(targetWorkbook))
            If targetWorkbook Is Me Then Throw New ArgumentException("Must be another workbook", NameOf(targetWorkbook))
            Select Case copyOption
                Case CopySheetOption.TargetSheetMustNotExist
                    If targetWorkbook.SheetNames.Contains(targetSheetName) = True Then Throw New InvalidOperationException("Target workbook must not contain worksheet """ & sheetName & """")
                    targetWorkbook.AddSheet(targetSheetName)
                Case CopySheetOption.TargetSheetMustExist
                    If targetWorkbook.SheetNames.Contains(targetSheetName) = False Then Throw New InvalidOperationException("Target workbook must contain worksheet """ & sheetName & """")
                Case CopySheetOption.TargetSheetMightExist
                    If targetWorkbook.SheetNames.Contains(targetSheetName) = False Then
                        targetWorkbook.AddSheet(targetSheetName)
                    End If
                Case Else
                    Throw New ArgumentException("Invalid copyOption " & copyOption.ToString, NameOf(copyOption))
            End Select
            If Me.GetType IsNot targetWorkbook.GetType Then Throw New NotSupportedException("Excel engines must be the same for source and target workbook for copying worksheets")
            targetWorkbook.ClearSheet(sheetName)
            Me.CopySheetContentInternal(sheetName, targetWorkbook, targetSheetName)
        End Sub

        Public Function AllFormulasOfWorkbook() As List(Of TextTableCell)
            Dim Result As New List(Of TextTableCell)
            Dim Sheets As List(Of String) = Me.SheetNames
            For MyCounter As Integer = 0 To Sheets.Count - 1
                Dim FoundFormulas As IEnumerable(Of TextTableCell) = Me.SheetContentMatrix(Sheets(MyCounter), ExcelDataOperationsBase.MatrixContent.Formulas).ToCellValuesList(Sheets(MyCounter))
                Result.AddRange(FoundFormulas)
            Next
            Return Result
        End Function

        Public Sub SelectCell(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            Me.SelectCell(New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All))
        End Sub

        Public MustOverride Sub SelectCell(cell As ExcelCell)

        ''' <summary>
        ''' Select a cell (without selecting the sheet)
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="specialCell"></param>
        Public Sub SelectCell(sheetName As String, specialCell As ExcelDataOperationsBase.SpecialCells)
            Select Case specialCell
                Case ExcelDataOperationsBase.SpecialCells.FirstCell
                    Me.SelectCell(New ExcelCell(sheetName, "A1", ExcelCell.ValueTypes.All))
                Case ExcelDataOperationsBase.SpecialCells.LastContentCell
                    Me.SelectCell(Me.LookupLastContentCell(sheetName))
                Case ExcelDataOperationsBase.SpecialCells.LastCell
                    Me.SelectCell(Me.LookupLastCell(sheetName))
                Case ExcelDataOperationsBase.SpecialCells.FirstUnlockedCell
                    Me.SelectCell(Me.LookupFirstUnlockedCell(sheetName))
                Case SpecialCells.FirstUnlockedCellOrFirstCell
                    Me.SelectCell(Me.LookupFirstUnlockedCellOrFirstCell(sheetName))
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(specialCell))
            End Select
        End Sub

        ''' <summary>
        ''' In all sheets, set the active cell to the top position
        ''' </summary>
        Public Sub SelectFirstUnlockedCellOrFirstCellInAllSheets()
            For Each SheetName As String In Me.SheetNames
                Me.SelectCell(SheetName, SpecialCells.FirstUnlockedCellOrFirstCell)
            Next
        End Sub

        Public Enum SpecialCells As Byte
            FirstCell = 0
            FirstUnlockedCell = 1
            LastContentCell = 2
            LastCell = 3
            FirstUnlockedCellOrFirstCell = 4
        End Enum

        Public Overrides Function ToString() As String
            Return "FileName=" & System.IO.Path.GetFileName(Me.FilePath) & "; ExcelEngine=" & Me.EngineName
        End Function

        Public MustOverride ReadOnly Property EngineName As String

        Public Function ReadSheetToUITable(sheetName As String) As TextTable
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            Dim Result As New TextTable()
            Result.AddColumns(LastCell.ColumnIndex + 1)
            Result.AddRows(LastCell.RowIndex + 1)
            For RowCounter As Integer = 0 To LastCell.RowIndex
                For ColCounter As Integer = 0 To LastCell.ColumnIndex
                    Result.Cell(RowCounter, ColCounter) = Me.LookupCellFormattedText(sheetName, RowCounter, ColCounter)
                Next
            Next
            Return Result
        End Function

        Public MustOverride ReadOnly Property HasVbaProject As Boolean

        Public MustOverride Sub RemoveVbaProject()

        Public MustOverride Function IsMergedCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean

        Public Function IsMergedCell(cell As ExcelCell) As Boolean
            Return IsMergedCell(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        Public MustOverride Sub UnMergeCells(sheetName As String, rowIndex As Integer, columnIndex As Integer)

        Public Sub UnMergeCell(cell As ExcelCell)
            UnMergeCells(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Sub

        Public MustOverride Sub MergeCells(sheetName As String, fromRowIndex As Integer, fromColumnIndex As Integer, toRowIndex As Integer, toColumnIndex As Integer)

        Public Sub MergeCells(sheetName As String, cells As ExcelRange)
            Me.MergeCells(sheetName, cells.AddressStart.RowIndex, cells.AddressStart.ColumnIndex, cells.AddressEnd.RowIndex, cells.AddressEnd.ColumnIndex)
        End Sub

        Protected Friend MustOverride Function MergedCells(sheetName As String) As List(Of ExcelOps.ExcelRange)

        Protected Overridable Function FindLastMergedCellNonEmpty(sheetName As String) As ExcelOps.ExcelCell
            Dim Result As ExcelCell = Nothing
            Dim AllMergedCells = Me.MergedCells(sheetName)
            For MyCounter As Integer = 0 To AllMergedCells.Count - 1
                If IsEmptyCell(AllMergedCells(MyCounter).AddressStart) = False Then
                    If Result Is Nothing Then
                        Result = AllMergedCells(MyCounter).AddressEnd
                    Else
                        Result = Tools.CombineCellAddresses(Result, AllMergedCells(MyCounter).AddressEnd, Tools.CellAddressCombineMode.RightLowerCorner)
                    End If
                End If
            Next
            Return Result
        End Function

        Protected Function FindLastMergedCell(sheetName As String) As ExcelOps.ExcelCell
            Dim Result As ExcelCell = Nothing
            Dim AllMergedCells = MergedCells(sheetName)
            For MyCounter As Integer = 0 To AllMergedCells.Count - 1
                If Result Is Nothing Then
                    Result = AllMergedCells(MyCounter).AddressEnd
                Else
                    Result = Tools.CombineCellAddresses(Result, AllMergedCells(MyCounter).AddressEnd, Tools.CellAddressCombineMode.RightLowerCorner)
                End If
            Next
            Return Result
        End Function

        Public MustOverride Sub AutoFitColumns(sheetName As String)

        Public MustOverride Sub AutoFitColumns(sheetName As String, minimumWidth As Double)

        Public MustOverride Sub AutoFitColumns(sheetName As String, columnIndex As Integer)

        Public MustOverride Sub AutoFitColumns(sheetName As String, columnIndex As Integer, minimumWidth As Double)

        Public MustOverride Function ExportChartSheetImage(chartSheetName As String) As System.Drawing.Image

        Public MustOverride Function ExportChartImage(workSheetName As String) As System.Drawing.Image()

        ''' <summary>
        ''' Find all error cells of a workbook
        ''' </summary>
        ''' <param name="filterForErrorValues">The error values to find, e.g. "#REF!", "#NAME?", or null/Nothing/0-array for all types of error cells</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' #DIV/0! - This error occurs when you try to divide a number by zero. To fix this error, you can use an IF function to verify that the denominator is zero.
        ''' #NAME? - This error occurs when Excel does not recognize the name of a function or range. To fix this error, check the name and make sure it is spelled correctly.
        ''' #VALUE! - This error occurs when Excel detects an invalid data type. To fix this error, make sure that the data is in the correct form.
        ''' #REF! - This error occurs when a cell refers to an invalid range. To fix this error, check the formula and make sure that all cell references are correct.
        ''' #NUM! - This error occurs when a formula or function has an invalid numeric argument. To fix this error, check the formula and make sure that all arguments are valid.
        ''' </remarks>
        Public Function FindErrorCellsInWorkbook(ParamArray filterForErrorValues As String()) As List(Of ExcelCell)
            Dim Result As New List(Of ExcelCell)
            Dim AllSheets = Me.SheetNames
            For Each SheetName As String In AllSheets
                Result.AddRange(FindErrorCellsInSheet(SheetName, filterForErrorValues))
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Find all error cells of a sheet
        ''' </summary>
        ''' <param name="sheetName">Name of sheet</param>
        ''' <param name="filterForErrorValues">The error values to find, e.g. "#REF!", "#NAME?", or null/Nothing/0-array for all types of error cells</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' #DIV/0! - This error occurs when you try to divide a number by zero. To fix this error, you can use an IF function to verify that the denominator is zero.
        ''' #NAME? - This error occurs when Excel does not recognize the name of a function or range. To fix this error, check the name and make sure it is spelled correctly.
        ''' #VALUE! - This error occurs when Excel detects an invalid data type. To fix this error, make sure that the data is in the correct form.
        ''' #REF! - This error occurs when a cell refers to an invalid range. To fix this error, check the formula and make sure that all cell references are correct.
        ''' #NUM! - This error occurs when a formula or function has an invalid numeric argument. To fix this error, check the formula and make sure that all arguments are valid.
        ''' </remarks>
        Public Function FindErrorCellsInSheet(sheetName As String, ParamArray filterForErrorValues As String()) As List(Of ExcelCell)
            Dim Result As New List(Of ExcelCell)
            Dim LastCell = Me.LookupLastCell(sheetName)
            For RowCounter As Integer = 0 To LastCell.RowIndex
                For ColCounter As Integer = 0 To LastCell.ColumnIndex
                    Dim ErrorValue = Me.LookupCellErrorValue(sheetName, RowCounter, ColCounter)
                    If ErrorValue <> Nothing Then
                        If filterForErrorValues Is Nothing OrElse filterForErrorValues.Length = 0 Then
                            'collect all error cells
                            Result.Add(New ExcelCell(sheetName, RowCounter, ColCounter, ExcelCell.ValueTypes.All))
                        ElseIf filterForErrorValues IsNot Nothing AndAlso filterForErrorValues.Length <> 0 Then
                            'collect specified error values only
                            If Tools.IsOneOf(Of String)(ErrorValue, filterForErrorValues) = True Then
                                Result.Add(New ExcelCell(sheetName, RowCounter, ColCounter, ExcelCell.ValueTypes.All))
                            End If
                        End If
                    End If
                Next
            Next
            Return Result
        End Function

        Public Enum ExcelSheetTypes As Byte
            WorkSheet = 1
            ChartSheet = 2
        End Enum

        Public Function SheetType(sheetName As String) As ExcelSheetTypes
            If Me.ChartSheetNames.Contains(sheetName) Then
                Return ExcelSheetTypes.ChartSheet
            Else
                Return ExcelSheetTypes.WorkSheet
            End If
        End Function

        ''' <summary>
        ''' Save workbook with its sheets to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="fileName"></param>
        ''' <param name="options"></param>
        Public Sub ExportWorkbookToHtml(fileName As String, options As HtmlWorkbookExportOptions)
            Dim Html = ExportWorkbookToHtml(options)
#If NETFRAMEWORK Then
            System.IO.File.WriteAllText(fileName, Html.ToString, System.Text.Encoding.UTF8)
#Else
            Using w As New StreamWriter(fileName, append:=False, encoding:=New UTF8Encoding(encoderShouldEmitUTF8Identifier:=True))
                w.Write(Html)              ' synchron
            End Using

#End If
        End Sub

#If Not NETFRAMEWORK Then
        ''' <summary>
        ''' Save workbook with its sheets to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="fileName"></param>
        ''' <param name="options"></param>
        Public Async Sub ExportWorkbookToHtmlAsync(fileName As String, options As HtmlWorkbookExportOptions)
            Dim Html = ExportWorkbookToHtml(options)
            Using w As New StreamWriter(fileName, append:=False, encoding:=New UTF8Encoding(encoderShouldEmitUTF8Identifier:=True))
                Await w.WriteAsync(Html)  ' asynchron
            End Using
        End Sub
#End If

        ''' <summary>
        ''' Save workbook with its sheets to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="options"></param>
        Public Function ExportWorkbookToHtml(options As HtmlWorkbookExportOptions) As System.Text.StringBuilder
            Dim Result As New System.Text.StringBuilder(128 * 1024)
            Result.AppendLine(options.EffectiveHtmlDocumentHeaderAndBody)

            Dim ExportSheetNames As New List(Of String)
            For Each MySheetName As String In SheetNames()
                Dim DoExport = False
                If Me.SheetType(MySheetName) = ExcelSheetTypes.ChartSheet AndAlso options.ExportChartSheets = True Then
                    DoExport = True
                ElseIf Me.SheetType(MySheetName) = ExcelSheetTypes.WorkSheet AndAlso options.ExportWorkSheets = True Then
                    DoExport = True
                End If
                If Me.IsHiddenSheet(MySheetName) AndAlso options.ExportHiddenSheets = False Then
                    DoExport = False
                End If

                If DoExport Then
                    ExportSheetNames.Add(MySheetName)
                End If
            Next

            If options.SheetNavigationPosition = HtmlWorkbookExportOptions.SheetNavigationPositions.Top Then
                options.GenerateWorkbookSubNavigation(Result, ExportSheetNames, options)
            End If

            Dim ExportAnchorNames = HtmlWorkbookExportOptions.CreateUniqueAnchorNames(ExportSheetNames)
            Dim ActiveSheetName = Me.SelectedSheetName
            If ExportSheetNames.Contains(ActiveSheetName) = False Then
                ActiveSheetName = ExportSheetNames.First
            End If
            For SheetCounter As Integer = 0 To ExportSheetNames.Count - 1
                Dim MySheetName As String = ExportSheetNames(SheetCounter)
                Dim MyAnchorName As String = ExportAnchorNames(SheetCounter)
                Dim IsSelected As Boolean = (MySheetName = ActiveSheetName)
                ExportSheetToHtml(MySheetName, MyAnchorName, IsSelected, Result, options, HtmlDocumentExportParts.ContentOnly)
            Next

            If options.SheetNavigationPosition = HtmlWorkbookExportOptions.SheetNavigationPositions.Bottom Then
                options.GenerateWorkbookSubNavigation(Result, ExportSheetNames, options)
            End If

            Result.AppendLine().AppendLine(options.EffectiveHtmlDocumentEnd)
            Return Result
        End Function

        ''' <summary>
        ''' Save single worksheet to HTML (including HTML document header/footer, images as HTML inline data)
        ''' </summary>
        ''' <param name="worksheetName"></param>
        ''' <param name="fileName"></param>
        Public Sub ExportSheetToHtml(worksheetName As String, fileName As String, options As HtmlSheetExportOptions)
            Dim Html As New System.Text.StringBuilder(128 * 1024)
            ExportSheetToHtml(worksheetName, HtmlWorkbookExportOptions.Slugify(worksheetName), True, Html, options, HtmlDocumentExportParts.FullHtmlDocument)
#If NETFRAMEWORK Then
            System.IO.File.WriteAllText(fileName, Html.ToString, System.Text.Encoding.UTF8)
#Else
            Using w As New StreamWriter(fileName, append:=False, encoding:=New UTF8Encoding(encoderShouldEmitUTF8Identifier:=True))
                w.Write(Html)              ' synchron
            End Using

#End If
        End Sub

#If Not NETFRAMEWORK Then
        ''' <summary>
        ''' Save worksheet to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="worksheetName"></param>
        ''' <param name="fileName"></param>
        Public Async Sub ExportSheetToHtmlAsync(worksheetName As String, fileName As String, options As HtmlSheetExportOptions)
            Dim Html As New System.Text.StringBuilder(128 * 1024)
            ExportSheetToHtmlInternal(worksheetName, Html, options)
            Using w As New StreamWriter(fileName, append:=False, encoding:=New UTF8Encoding(encoderShouldEmitUTF8Identifier:=True))
                Await w.WriteAsync(Html)  ' asynchron
            End Using
        End Sub
#End If

        ''' <summary>
        ''' Save single worksheet to HTML (including HTML document header/footer, images as HTML inline data)
        ''' </summary>
        ''' <param name="worksheetName"></param>
        Public Function ExportSheetToHtml(worksheetName As String, options As HtmlSheetExportOptions) As System.Text.StringBuilder
            Dim sb As New System.Text.StringBuilder
            ExportSheetToHtml(worksheetName, HtmlWorkbookExportOptions.Slugify(worksheetName), True, sb, options, HtmlDocumentExportParts.FullHtmlDocument)
            Return sb
        End Function

        ''' <summary>
        ''' Save worksheet to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="worksheetName"></param>
        ''' <param name="sb"></param>
        Public Sub ExportSheetToHtml(worksheetName As String, anchorName As String, initiallyVisible As Boolean, sb As System.Text.StringBuilder, options As HtmlSheetExportOptions, exportedHtmlDocumentParts As HtmlDocumentExportParts)
            Select Case exportedHtmlDocumentParts
                Case HtmlDocumentExportParts.FullHtmlDocument
                    sb.AppendLine(options.EffectiveHtmlDocumentHeaderAndBody)
                    options.GenerateBeginSheetSection(sb, anchorName, initiallyVisible)
                    options.GenerateSheetSectionTitle(sb, worksheetName)
                    ExportSheetToHtmlInternal(worksheetName, sb, options)
                    options.GenerateEndSheetSection(sb)
                    sb.AppendLine().AppendLine(options.EffectiveHtmlDocumentEnd)
                Case HtmlDocumentExportParts.ContentOnly
                    options.GenerateBeginSheetSection(sb, anchorName, initiallyVisible)
                    options.GenerateSheetSectionTitle(sb, worksheetName)
                    ExportSheetToHtmlInternal(worksheetName, sb, options)
                    options.GenerateEndSheetSection(sb)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(exportedHtmlDocumentParts))
            End Select
        End Sub

        Public Enum HtmlDocumentExportParts As Byte
            ''' <summary>
            ''' Create a full HTML document with header tags, style tags, etc.
            ''' </summary>
            FullHtmlDocument = 1
            ''' <summary>
            ''' Export HTML table code only
            ''' </summary>
            ContentOnly = 2
        End Enum

        ''' <summary>
        ''' Save worksheet to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="worksheetName"></param>
        ''' <param name="sb"></param>
        Protected MustOverride Sub ExportSheetToHtmlInternal(worksheetName As String, sb As System.Text.StringBuilder, options As HtmlSheetExportOptions)

#Region "Colors and Theming (Helpers for e.g. ExcelColorToCssHex() in depending classes)"
        ''' <summary>
        ''' Default-Office-Theme (Office-Standard „Office“)
        ''' </summary>
        ''' <param name="themeIndex"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' 0=Background1, 1=Text1, 2=Background2, 3=Text2,
        ''' 4..9=Accent1..Accent6, 10=Hyperlink, 11=FollowedHyperlink
        ''' </remarks>
        Protected Shared Function DefaultOfficeTheme(themeIndex As Integer) As String
            Select Case themeIndex
                Case 0 : Return "#FFFFFF" ' lt1
                Case 1 : Return "#000000" ' dk1
                Case 2 : Return "#EEECE1" ' lt2
                Case 3 : Return "#1F497D" ' dk2
                Case 4 : Return "#4F81BD" ' Accent1
                Case 5 : Return "#C0504D" ' Accent2
                Case 6 : Return "#9BBB59" ' Accent3
                Case 7 : Return "#8064A2" ' Accent4
                Case 8 : Return "#4BACC6" ' Accent5
                Case 9 : Return "#F79646" ' Accent6
                Case 10 : Return "#0000FF" ' Hyperlink
                Case 11 : Return "#800080" ' FollowedHyperlink (ungefähr)
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Excel-Tint-Regel (heller/dunkler)
        ''' </summary>
        ''' <param name="hex"></param>
        ''' <param name="tint"></param>
        ''' <returns></returns>
        Protected Shared Function ApplyTint(hex As String, tint As Double) As String
            ' hex "#RRGGBB"
            Dim r = Convert.ToInt32(hex.Substring(1, 2), 16)
            Dim g = Convert.ToInt32(hex.Substring(3, 2), 16)
            Dim b = Convert.ToInt32(hex.Substring(5, 2), 16)

            Dim adj As Func(Of Integer, Integer) =
                Function(ch As Integer)
                    Dim v As Double
                    If tint < 0 Then
                        v = ch * (1.0 + tint)                   ' dunkler
                    Else
                        v = ch * (1.0 - tint) + 255.0 * tint    ' heller
                    End If
                    Return Math.Max(0, Math.Min(255, CInt(Math.Round(v))))
                End Function

            Return $"#{adj(r):X2}{adj(g):X2}{adj(b):X2}"
        End Function

        ''' <summary>
        ''' Vollständiges Mapping der Excel-Standardpalette (Indexed 0..63).
        ''' </summary>
        ''' <remarks>
        ''' 64=System Foreground, 65=System Background → kein Hex.
        ''' Quelle: OpenXML "indexedColors" Default-Mapping.
        ''' </remarks>
        Protected Shared Function IndexedColor(index As Integer) As String
            Select Case index
                Case 0 : Return "#000000" ' Black (dup. von 8)
                Case 1 : Return "#FFFFFF" ' White (dup. von 9)
                Case 2 : Return "#FF0000" ' Red   (dup. von 10)
                Case 3 : Return "#00FF00" ' Green (dup. von 11)
                Case 4 : Return "#0000FF" ' Blue  (dup. von 12)
                Case 5 : Return "#FFFF00" ' Yellow (dup. von 13)
                Case 6 : Return "#FF00FF" ' Magenta (dup. von 14)
                Case 7 : Return "#00FFFF" ' Cyan (dup. von 15)

                Case 8 : Return "#000000" ' Black
                Case 9 : Return "#FFFFFF" ' White
                Case 10 : Return "#FF0000" ' Red
                Case 11 : Return "#00FF00" ' Lime
                Case 12 : Return "#0000FF" ' Blue
                Case 13 : Return "#FFFF00" ' Yellow
                Case 14 : Return "#FF00FF" ' Magenta
                Case 15 : Return "#00FFFF" ' Aqua

                Case 16 : Return "#800000" ' Maroon
                Case 17 : Return "#008000" ' Green
                Case 18 : Return "#000080" ' Navy
                Case 19 : Return "#808000" ' Olive
                Case 20 : Return "#800080" ' Purple
                Case 21 : Return "#008080" ' Teal
                Case 22 : Return "#C0C0C0" ' Silver
                Case 23 : Return "#808080" ' Gray

                Case 24 : Return "#9999FF"
                Case 25 : Return "#993366"
                Case 26 : Return "#FFFFCC"
                Case 27 : Return "#CCFFFF"
                Case 28 : Return "#660066"
                Case 29 : Return "#FF8080"
                Case 30 : Return "#0066CC"
                Case 31 : Return "#CCCCFF"

                Case 32 : Return "#000080"
                Case 33 : Return "#FF00FF"
                Case 34 : Return "#FFFF00"
                Case 35 : Return "#00FFFF"
                Case 36 : Return "#800080"
                Case 37 : Return "#800000"
                Case 38 : Return "#008080"
                Case 39 : Return "#0000FF"

                Case 40 : Return "#00CCFF"
                Case 41 : Return "#CCFFFF"
                Case 42 : Return "#CCFFCC"
                Case 43 : Return "#FFFF99"
                Case 44 : Return "#99CCFF"
                Case 45 : Return "#FF99CC"
                Case 46 : Return "#CC99FF"
                Case 47 : Return "#FFCC99"

                Case 48 : Return "#3366FF"
                Case 49 : Return "#33CCCC"
                Case 50 : Return "#99CC00"
                Case 51 : Return "#FFCC00"
                Case 52 : Return "#FF9900"
                Case 53 : Return "#FF6600"
                Case 54 : Return "#666699"
                Case 55 : Return "#969696"

                Case 56 : Return "#003366"
                Case 57 : Return "#339966"
                Case 58 : Return "#003300"
                Case 59 : Return "#333300"
                Case 60 : Return "#993300"
                Case 61 : Return "#993366"
                Case 62 : Return "#333399"
                Case 63 : Return "#333333"

                Case 64 : Return Nothing   ' System Foreground
                Case 65 : Return Nothing   ' System Background
                Case Else
                    Return Nothing
            End Select
        End Function

        ' Mappt Enum-/Namen auf 0..11 (kompatibel zu Ihrer DefaultOfficeTheme-Tabelle)
        Protected Shared Function MapThemeNameToIndex(themeName As String) As Integer
            If String.IsNullOrEmpty(themeName) Then Return -1
            Dim s = themeName.Trim().ToLowerInvariant()

            Select Case s
                Case "background1", "light1", "lt1", "bg1" : Return 0
                Case "text1", "dark1", "dk1" : Return 1
                Case "background2", "light2", "lt2", "bg2" : Return 2
                Case "text2", "dark2", "dk2" : Return 3
                Case "accent1" : Return 4
                Case "accent2" : Return 5
                Case "accent3" : Return 6
                Case "accent4" : Return 7
                Case "accent5" : Return 8
                Case "accent6" : Return 9
                Case "hyperlink", "hlink" : Return 10
                Case "followedhyperlink", "folhlink" : Return 11
                Case Else : Return -1
            End Select
        End Function
#End Region

    End Class
End Namespace