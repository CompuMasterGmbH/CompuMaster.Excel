Option Explicit On
Option Strict On

Namespace ExcelOps
    Public MustInherit Class ExcelDataOperationsBase

        Public Enum OpenMode As Byte
            OpenExistingFile = 0
            CreateFile = 1
        End Enum

        Protected Sub New(file As String, mode As OpenMode, autoCalculationOnLoad As Boolean, calculationModuleDisabled As Boolean, [readOnly] As Boolean)
            If autoCalculationOnLoad AndAlso calculationModuleDisabled Then Throw New ArgumentException("Calculation engine is disabled, but AutoCalculation requested", NameOf(autoCalculationOnLoad))
            Me.AutoCalculationOnLoad = autoCalculationOnLoad
            Me.CalculationModuleDisabled = calculationModuleDisabled
            Select Case mode
                Case OpenMode.OpenExistingFile
                    Me.LoadAndInitializeWorkbookFile(file)
                Case OpenMode.CreateFile
                    Me.CreateAndInitializeWorkbookFile(file)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(mode))
            End Select
            Me.ReadOnly = [readOnly]
        End Sub

        ''' <summary>
        ''' Create a new instance for accessing Excel workbooks
        ''' </summary>
        ''' <param name="autoCalculationOnLoad">Automatically do a full recalculation after workbook has been loaded</param>
        ''' <param name="calculationModuleDisabled">Disables the Excel calculation engine</param>
        Protected Sub New(autoCalculationOnLoad As Boolean, calculationModuleDisabled As Boolean, [readOnly] As Boolean)
            If autoCalculationOnLoad AndAlso calculationModuleDisabled Then Throw New ArgumentException("Calculation engine is disabled, but AutoCalculation requested", NameOf(autoCalculationOnLoad))
            Me.AutoCalculationOnLoad = autoCalculationOnLoad
            Me.CalculationModuleDisabled = calculationModuleDisabled
            Me.ReadOnly = [readOnly]
        End Sub

        Public Sub ReloadFromFile()
            Me.LoadAndInitializeWorkbookFile(Me.FilePath)
        End Sub

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

        Protected _FilePath As String
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
        ''' Save modifications made to the workbook
        ''' </summary>
        Public MustOverride Sub Save()

        ''' <summary>
        ''' Save workbook as another file
        ''' </summary>
        ''' <param name="filePath"></param>
        <Obsolete("Use overloaded method", True)> Public Sub SaveAs(filePath As String)
            If Me.ReadOnly = True AndAlso Me._FilePath = filePath Then
                Throw New ArgumentException("File is read-only and can't be saved at same location")
            End If
            Me.SaveAsInternal(filePath, SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Me._FilePath = filePath
            Me.ReadOnly = False
        End Sub

        ''' <summary>
        ''' Save workbook as another file
        ''' </summary>
        ''' <param name="filePath"></param>
        Public Sub SaveAs(filePath As String, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            If Me.ReadOnly = True AndAlso Me._FilePath = filePath Then
                Throw New ArgumentException("File is read-only and can't be saved at same location")
            End If
            Me.SaveAsInternal(filePath, cachedCalculationsOption)
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
        ''' All available sheet names
        ''' </summary>
        ''' <returns></returns>
        Public MustOverride Function SheetNames() As List(Of String)

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

        Protected Overridable Sub LoadAndInitializeWorkbookFile(inputPath As String)
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

        Protected MustOverride Sub CreateWorkbook()

        Protected Overridable Sub CreateAndInitializeWorkbookFile(filePath As String)
            'Load the changed worksheet
            Me._FilePath = filePath
            Dim file As New System.IO.FileInfo(filePath)
            If file.Exists = True Then
                Throw New System.InvalidOperationException("File already exists: " & file.ToString)
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
        Public MustOverride Function LookupLastContentColumnIndex(sheetName As String) As Integer

        ''' <summary>
        ''' Lookup the last content row index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Function LookupLastContentRowIndex(sheetName As String) As Integer

        ''' <summary>
        ''' Lookup the last content cell (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Function LookupLastContentCell(sheetName As String) As ExcelCell

        ''' <summary>
        ''' Lookup the last column index (zero based index) (the last content cell equals to Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Function LookupLastColumnIndex(sheetName As String) As Integer

        ''' <summary>
        ''' Lookup the last row index (zero based index) (the last content cell equals to Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Function LookupLastRowIndex(sheetName As String) As Integer

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
        Public MustOverride Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer)

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
                            Throw New NotImplementedException
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
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public MustOverride Sub SelectSheet(sheetName As String)

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetIndex"></param>
        Public MustOverride Sub SelectSheet(sheetIndex As Integer)

        ''' <summary>
        ''' Lookup the (zero-based) index number of a work sheet
        ''' </summary>
        ''' <param name="worksheetName">A work sheet name</param>
        ''' <returns>-1 if the sheet name doesn't exist, otherwise its index value</returns>
        Public Function SheetIndex(ByVal worksheetName As String) As Integer
            Return Me.SheetNames.IndexOf(worksheetName)
        End Function

        Public MustOverride Sub CopySheetContent(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)

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
            Return "FileName=" & System.IO.Path.GetFileName(Me.FilePath) & "; ExcelEngine=" & Me.EngineName.ToString
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

    End Class
End Namespace