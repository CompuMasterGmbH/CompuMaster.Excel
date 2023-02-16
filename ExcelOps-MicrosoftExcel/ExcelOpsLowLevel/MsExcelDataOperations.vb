Option Explicit On
Option Strict On

Imports MsExcel = Microsoft.Office.Interop.Excel
Imports CompuMaster.Excel.MsExcelCom

Namespace Global.CompuMaster.Excel.ExcelOps

    ''' <summary>
    ''' MS Excel Interop provider (ATTENTION: watch for advised Try-Finally pattern for successful application process stop!)
    ''' </summary>
    ''' <remarks>Use with pattern
    ''' <code>
    ''' Dim MsExcelApp As New MsExcelDataOperations.MsAppInstance
    ''' Try
    '''    '...
    ''' Finally
    '''     MsExcelDataOperations.PrepareCloseExcelAppInstance(MSExcelApp)
    '''     MsExcelDataOperations.SafelyCloseExcelAppInstance(MSExcelApp)
    ''' End Try
    ''' </code>
    ''' </remarks>
    Public Class MsExcelDataOperations
        Inherits ExcelDataOperationsBase

        Public Shared Property AutoKillAllExistingMsExcelInstances As Boolean = False

        Public Shared Sub CheckForRunningMsExcelInstancesAndAskUserToKill()
            Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            If MsExcelProcesses IsNot Nothing AndAlso MsExcelProcesses.Length > 0 Then
                If AutoKillAllExistingMsExcelInstances OrElse MsgBox(MsExcelProcesses.Length & " bereits geöffenete MS Excel Instanzen wurden gefunden. Sollen diese zuvor geschlossen werden?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = vbYes Then
                    For Each ExcelInstance As System.Diagnostics.Process In MsExcelProcesses
                        Try
                            ExcelInstance.CloseMainWindow()
                        Catch
                        End Try
                        System.Threading.Thread.Sleep(200)
                        If ExcelInstance.HasExited = False Then
                            ExcelInstance.Kill()
                            ExcelInstance.Close()
                        End If
                    Next
                    System.Threading.Thread.Sleep(500) 'Process might take a few more milli-seconds to finally disappear
                End If
            End If
        End Sub

        ''' <summary>
        ''' Class for holding a reference to Excel.Application (ATTENTION: watch for advised Try-Finally pattern!)
        ''' </summary>
        ''' <remarks>Use with pattern
        ''' <code>
        ''' Dim MsExcelOps As New MsExcelDataOperations(fileName)
        ''' Try
        '''    '...
        ''' Finally
        '''     MsExcelOps.CloseExcelAppInstance()
        ''' End Try
        ''' </code>
        ''' </remarks> 
        Public Sub New(passwordForOpeningOnNextTime As String)
            Me.New(Nothing, OpenMode.CreateFile, False, True, passwordForOpeningOnNextTime)
        End Sub

        ''' <summary>
        ''' Class for holding a reference to Excel.Application (ATTENTION: watch for advised Try-Finally pattern!)
        ''' </summary>
        ''' <remarks>Use with pattern
        ''' <code>
        ''' Dim MsExcelOps As New MsExcelDataOperations(fileName)
        ''' Try
        '''    '...
        ''' Finally
        '''     MsExcelOps.CloseExcelAppInstance()
        ''' End Try
        ''' </code>
        ''' </remarks> 
        Public Sub New()
            Me.New(Nothing)
        End Sub

        ''' <summary>
        ''' Class for holding a reference to Excel.Application (ATTENTION: watch for advised Try-Finally pattern!)
        ''' </summary>
        ''' <remarks>Use with pattern
        ''' <code>
        ''' Dim MsExcelOps As New MsExcelDataOperations(fileName)
        ''' Try
        '''    '...
        ''' Finally
        '''     MsExcelOps.CloseExcelAppInstance()
        ''' End Try
        ''' </code>
        ''' </remarks>
        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String)
            Me.New(file, mode, New MsExcelApplicationWrapper, False, [readOnly], passwordForOpening)
        End Sub

        ''' <summary>
        ''' Class for holding a reference to Excel.Application (ATTENTION: watch for advised Try-Finally pattern!)
        ''' </summary>
        ''' <remarks>Use with pattern
        ''' <code>
        ''' Dim MsExcelOps As New MsExcelDataOperations(fileName)
        ''' Try
        '''    '...
        ''' Finally
        '''     MsExcelOps.CloseExcelAppInstance()
        ''' End Try
        ''' </code>
        ''' </remarks>
        Public Sub New(file As String, mode As OpenMode, unprotectWorksheets As Boolean, [readOnly] As Boolean, passwordForOpening As String)
            Me.New(file, mode, New MsExcelApplicationWrapper, unprotectWorksheets, [readOnly], passwordForOpening)
        End Sub

        ''' <summary>
        ''' MS Excel Interop provider (ATTENTION: watch for advised Try-Finally pattern for successful application process stop!) incl. unprotection of sheets
        ''' </summary>
        ''' <remarks>Use with pattern
        ''' <code>
        ''' Dim MsExcelApp As New MsExcelDataOperations.MsAppInstance
        ''' Try
        '''    '...
        ''' Finally
        '''     MsExcelDataOperations.PrepareCloseExcelAppInstance(MSExcelApp)
        '''     MsExcelDataOperations.SafelyCloseExcelAppInstance(MSExcelApp)
        ''' End Try
        ''' </code>
        ''' </remarks>
        <Obsolete("Use overload; WARNING: this overload always leads to: unprotectWorksheets = True")>
        Public Sub New(file As String, mode As OpenMode, msExcelApp As MsExcelApplicationWrapper, [readOnly] As Boolean, passwordForOpening As String)
            Me.New(file, mode, msExcelApp, True, [readOnly], passwordForOpening)
            Me._MsExcelAppInstance = msExcelApp
        End Sub

        ''' <summary>
        ''' MS Excel Interop provider (ATTENTION: watch for advised Try-Finally pattern for successful application process stop!) incl. unprotection of sheets
        ''' </summary>
        ''' <remarks>Use with pattern
        ''' <code>
        ''' Dim MsExcelApp As New MsExcelDataOperations.MsAppInstance
        ''' Try
        '''    '...
        ''' Finally
        '''     MsExcelDataOperations.PrepareCloseExcelAppInstance(MSExcelApp)
        '''     MsExcelDataOperations.SafelyCloseExcelAppInstance(MSExcelApp)
        ''' End Try
        ''' </code>
        ''' </remarks>
        Public Sub New(file As String, mode As OpenMode, msExcelApp As MsExcelApplicationWrapper, unprotectWorksheets As Boolean, [readOnly] As Boolean, passwordForOpening As String)
#Disable Warning IDE0060 ' Nicht verwendete Parameter entfernen
#Enable Warning IDE0060 ' Nicht verwendete Parameter entfernen
            MyBase.New(True, False, [readOnly], passwordForOpening)
            Me._MsExcelAppInstance = msExcelApp
            Me._Workbooks = New MsExcelWorkbooksWrapper(msExcelApp, msExcelApp.ComObjectStronglyTyped.Workbooks)
            Select Case mode
                Case OpenMode.OpenExistingFile
                    Me.LoadAndInitializeWorkbookFile(file)
                Case OpenMode.CreateFile
                    Me.CreateAndInitializeWorkbookFile(file)
                    Me.ReadOnly = [readOnly] OrElse (file = Nothing)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(mode))
            End Select
            If unprotectWorksheets = True Then
                Me.UnprotectSheets()
            End If
        End Sub

        '''' <summary>
        '''' MS Excel Interop provider (ATTENTION: watch for advised Try-Finally pattern for successful application process stop!)
        '''' </summary>
        '''' <remarks>Use with pattern
        '''' <code>
        '''' Dim MsExcelApp As New MsExcelDataOperations.MsAppInstance
        '''' Try
        ''''    '...
        '''' Finally
        ''''     MsExcelDataOperations.PrepareCloseExcelAppInstance(MSExcelApp)
        ''''     MsExcelDataOperations.SafelyCloseExcelAppInstance(MSExcelApp)
        '''' End Try
        '''' </code>
        '''' </remarks>
        'Protected Property MSExcelApp As ComMsExcelApplication

        Private _MsExcelAppInstance As MsExcelApplicationWrapper
        Public ReadOnly Property MsExcelAppInstance As MsExcelApplicationWrapper
            Get
                If _MsExcelAppInstance Is Nothing Then
                    _MsExcelAppInstance = New MsExcelApplicationWrapper()
                End If
                Return _MsExcelAppInstance
            End Get
        End Property

        Public Overrides Sub Close()
            If Me._Workbook IsNot Nothing Then
                Me.Workbook.Close(SaveChanges:=False)
                Me._Workbook.Dispose()
                Me._Workbook = Nothing
            End If
        End Sub

        Public Overrides Sub CloseExcelAppInstance()
            'Close workbook if still open
            Me.Close()
            'Close workbooks collection
            If Me._Workbooks IsNot Nothing Then
                Me._Workbooks.Dispose()
                Me._Workbooks = Nothing
            End If
            If Me._MsExcelAppInstance IsNot Nothing Then
                Me._MsExcelAppInstance.Dispose()
                Me._MsExcelAppInstance = Nothing
            End If
        End Sub

        Private _Workbooks As MsExcelWorkbooksWrapper
        Public ReadOnly Property Workbooks As MsExcel.Workbooks
            Get
                If _Workbooks Is Nothing Then
                    Return Nothing
                Else
                    Return _Workbooks.ComObjectStronglyTyped
                End If
            End Get
        End Property

        Public Overrides ReadOnly Property IsClosed As Boolean
            Get
                Return Me._Workbook Is Nothing
            End Get
        End Property

        Protected Overrides Sub SaveInternal()
            If Me.PasswordForOpening <> Nothing Then
                Me.Workbook.Protect(Me.PasswordForOpening)
            End If
            Me.Workbook.Save()
        End Sub

        Protected Overrides Sub SaveInternal_ApplyCachedCalculationOption(cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            Select Case cachedCalculationsOption
                Case SaveOptionsForDisabledCalculationEngines.DefaultBehaviour, SaveOptionsForDisabledCalculationEngines.NoReset
                Case Else
                    Throw New NotSupportedException("SaveOptionsForDisabledCalculationEngines " & cachedCalculationsOption.ToString & " not supported by MS Excel")
            End Select
            MyBase.SaveInternal_ApplyCachedCalculationOption(cachedCalculationsOption)
        End Sub

        Protected Overrides Sub SaveAsInternal(fileName As String, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            If Me.PasswordForOpening <> Nothing Then
                Me.Workbook.Protect(Me.PasswordForOpening)
            End If
            Dim Format As MsExcel.XlFileFormat?
            Select Case System.IO.Path.GetExtension(fileName.ToLowerInvariant)
                Case ".xlsx"
                    Format = MsExcel.XlFileFormat.xlOpenXMLWorkbook
                Case ".xlst"
                    Format = MsExcel.XlFileFormat.xlOpenXMLTemplate
                Case ".xlsm"
                    Format = MsExcel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled
                Case ".xlt"
                    Format = MsExcel.XlFileFormat.xlTemplate
                Case ".xls"
                    Format = MsExcel.XlFileFormat.xlWorkbookDefault
            End Select
            If Format.HasValue Then
                Me.Workbook.SaveAs(fileName, FileFormat:=Format, Password:=Me.PasswordForOpening)
            Else
                Me.Workbook.SaveAs(fileName, Password:=Me.PasswordForOpening)
            End If
        End Sub

        Private _Workbook As MsExcelWorkbookWrapper
        Public ReadOnly Property Workbook As MsExcel.Workbook
            Get
                If _Workbook Is Nothing Then
                    Return Nothing
                Else
                    Return _Workbook.ComObjectStronglyTyped
                End If
            End Get
        End Property

        Public Overrides Property AutoCalculationEnabled As Boolean
            Get
                Return (Me.MsExcelAppInstance.ComObjectStronglyTyped.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic)
            End Get
            Set(value As Boolean)
                If value Then
                    Me.MsExcelAppInstance.ComObjectStronglyTyped.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic
                Else
                    Me.MsExcelAppInstance.ComObjectStronglyTyped.Calculation = MsExcel.XlCalculation.xlCalculationManual
                End If
            End Set
        End Property

        Protected Overrides Sub CreateWorkbook()
            'If Me._Workbook IsNot Nothing Then
            '    If Me.WorkbookFilePath IsNot Nothing Then Throw New UnauthorizedAccessException("Found another workbook opened and in access by current MS Excel instance which hasn't been opened/created by (this) code")
            '    Me._Workbook.Close(MsExcel.XlSaveAction.xlDoNotSaveChanges)
            'End If
            If Me._Workbook Is Nothing Then
                Dim Wb As MsExcel.Workbook = Me.Workbooks.Add()
                Me._Workbook = New MsExcelWorkbookWrapper(Me._Workbooks, Wb)
                While Wb.Worksheets.Count > 1
                    CType(Wb.Worksheets(Wb.Worksheets.Count), MsExcel.Worksheet).Delete()
                End While
            End If
            'If Me.MSExcelApp Is Nothing AndAlso Me._Workbook IsNot Nothing Then
            '    Me.MSExcelApp = Me.Workbook.Application
            'End If
            Me.Workbook.EnableAutoRecover = False
        End Sub

        Protected Overrides Sub LoadWorkbook(file As System.IO.FileInfo)
            If file.Exists = False Then Throw New System.IO.FileNotFoundException("Workbook file must exist for loading from disk", file.FullName)
            If Me._Workbook Is Nothing Then
                If Me.Workbooks Is Nothing Then Throw New NullReferenceException("Workbooks")
                If Me.MsExcelAppInstance.ComObject Is Nothing Then Throw New NullReferenceException("MsExcelAppInstance")
                If Me.MsExcelAppInstance.IsDisposed Then Throw New NullReferenceException("MsExcelAppInstance already disposed")
                Dim Wb As MsExcel.Workbook
                If Me.PasswordForOpening <> Nothing Then
                    Wb = Me.Workbooks.Open(file.FullName, UpdateLinks:=True, [ReadOnly]:=False, Editable:=False, Notify:=False, Password:=Me.PasswordForOpening)
                Else
                    Wb = Me.Workbooks.Open(file.FullName, UpdateLinks:=True, [ReadOnly]:=False, Editable:=False, Notify:=False, Password:="")
                End If
                If Wb Is Nothing Then Throw New NullReferenceException("Null result after Workbooks.Open")
                Me._Workbook = New MsExcelWorkbookWrapper(Me._Workbooks, Wb)
            End If
            'If Me.MSExcelApp Is Nothing AndAlso Me.Workbook IsNot Nothing Then
            '    Me.MSExcelApp = Me.Workbook.Application
            'End If
            Me.Workbook.EnableAutoRecover = False
        End Sub

        Public Overrides Sub CleanupRangeNames()
            'do nothing - just needs to be done once, see Epplus implementation
        End Sub

        Public Overrides Function SheetNames() As List(Of String)
            Dim Result As New List(Of String)
            For Each ws As MsExcel.Worksheet In Me.Workbook.Worksheets
                Result.Add(ws.Name)
            Next
            Return Result
        End Function

        Public Overrides Function TryLookupCellValue(Of T As Structure)(cell As ExcelCell) As T?
            Try
                Return LookupCellValue(Of T)(cell)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Overrides Function TryLookupCellValue(Of T As Structure)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T?
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Try
                Return LookupCellValue(Of T)(sheetName, rowIndex, columnIndex)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function LookupCellValueAsObject(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Object
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return LookupCellValue(Of Object)(sheetName, rowIndex, columnIndex)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <returns></returns>
        Private Overloads Function LookupCellValueAsObject(sheet As MsExcel.Worksheet, rowIndex As Integer, columnIndex As Integer) As Object
            Return LookupCellValue(Of Object)(sheet, rowIndex, columnIndex)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Public Overrides Function LookupCellValue(Of T)(cell As ExcelCell) As T
            Return LookupCellValue(Of T)(CType(Me.Workbook.Worksheets(cell.SheetName), MsExcel.Worksheet), cell)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Private Overloads Function LookupCellValue(Of T)(sheet As MsExcel.Worksheet, cell As ExcelCell) As T
            If cell.SheetName <> Nothing AndAlso cell.SheetName <> sheet.Name Then Throw New ArgumentException("Argument sheet must match sheetname in argument cell")
            Dim Result As Object = CType(sheet.Range(cell.Address), MsExcel.Range).Value
            Return Me.ConvertCellValueObjectTo(Of T)(Result)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        Public Overrides Function LookupCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.LookupCellValue(Of T)(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet), rowIndex, columnIndex)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="sheet"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        Private Overloads Function LookupCellValue(Of T)(sheet As MsExcel.Worksheet, rowIndex As Integer, columnIndex As Integer) As T
            Dim Result As Object = CType(sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).Value
            Return Me.ConvertCellValueObjectTo(Of T)(Result)
        End Function

        ''' <summary>
        ''' Read the cell format string
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Public Overrides Function LookupCellFormat(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim Result As Object = CType(Sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).DisplayFormat.NumberFormat
            Return CType(Result, String)
        End Function

        Private Function ConvertCellValueObjectTo(Of T)(value As Object) As T
            If value Is Nothing OrElse (value.GetType Is GetType(String) AndAlso Trim(CType(value, String)) = Nothing) Then
                Return Nothing
            ElseIf value.GetType Is GetType(Double) AndAlso GetType(T).IsGenericType AndAlso Nullable.GetUnderlyingType(GetType(T)) IsNot Nothing Then
                Dim BaseT As Type = Nullable.GetUnderlyingType(GetType(T))
                Dim Result As Object = Convert.ChangeType(value, BaseT)
                Return CType(Result, T)
            Else
                Return CType(value, T)
            End If
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Public Overrides Function LookupCellFormula(cell As ExcelCell) As String
            Return Me.LookupCellFormula(CType(Me.Workbook.Worksheets(cell.SheetName), MsExcel.Worksheet), cell)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Private Overloads Function LookupCellFormula(sheet As MsExcel.Worksheet, cell As ExcelCell) As String
            If cell.SheetName <> Nothing AndAlso cell.SheetName <> sheet.Name Then Throw New ArgumentException("Argument sheet must match sheetname in argument cell")
            If CType(CType(sheet.Range(cell.Address), MsExcel.Range).HasFormula, Boolean) Then
                Return CType(CType(sheet.Range(cell.Address), MsExcel.Range).Formula, String).Substring(1)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function LookupCellFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.LookupCellFormula(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet), rowIndex, columnIndex)
        End Function

        ''' <summary>
        ''' Read a cell formula
        ''' </summary>
        ''' <param name="worksheet"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        Private Overloads Function LookupCellFormula(worksheet As MsExcel.Worksheet, rowIndex As Integer, columnIndex As Integer) As String
            If CType(CType(worksheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).HasFormula, Boolean) Then
                Return CType(CType(worksheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).Formula, String).Substring(1)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function LookupCellFormattedText(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.LookupCellFormattedText(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet), rowIndex, columnIndex)
        End Function

        Private Overloads Function LookupCellFormattedText(worksheet As MsExcel.Worksheet, rowIndex As Integer, columnIndex As Integer) As String
            Dim Cell As MsExcel.Range = CType(worksheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range)
            Dim CellValue As Object = Cell.Value
            If CellValue IsNot Nothing AndAlso CellValue.GetType() Is GetType(Boolean) Then
                Return CType(CellValue, Boolean).ToString
            Else
                Dim Result As Object = CType(worksheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).Text
                If Result Is Nothing OrElse (Result.GetType Is GetType(String) AndAlso Trim(CType(Result, String)) = Nothing) Then
                    Return Nothing
                Else
                    Return CType(Result, String)
                End If
            End If
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Public Overrides Function LookupCellIsLocked(cell As ExcelCell) As Boolean
            Return Me.LookupCellIsLocked(CType(Me.Workbook.Worksheets(cell.SheetName), MsExcel.Worksheet), cell)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Private Overloads Function LookupCellIsLocked(sheet As MsExcel.Worksheet, cell As ExcelCell) As Boolean
            If cell.SheetName <> Nothing AndAlso cell.SheetName <> sheet.Name Then Throw New ArgumentException("Argument sheet must match sheetname in argument cell")
            Return CType(CType(CType(sheet.Range(cell.Address), MsExcel.Range).Style, MsExcel.Style).Locked, Boolean)
        End Function

        Public Overrides Function LookupRowIndex(cell As ExcelCell) As Integer
            Return Me.LookupRowIndex(CType(Me.Workbook.Worksheets(cell.SheetName), MsExcel.Worksheet), cell)
        End Function

        Private Overloads Function LookupRowIndex(sheet As MsExcel.Worksheet, cell As ExcelCell) As Integer
            If cell.SheetName <> Nothing AndAlso cell.SheetName <> sheet.Name Then Throw New ArgumentException("Argument sheet must match sheetname in argument cell")
            Return CType(sheet.Range(cell.Address), MsExcel.Range).Row - 1
        End Function

        Public Overrides Function LookupColumnIndex(cell As ExcelCell) As Integer
            Return Me.LookupColumnIndex(CType(Me.Workbook.Worksheets(cell.SheetName), MsExcel.Worksheet), cell)
        End Function

        Private Overloads Function LookupColumnIndex(sheet As MsExcel.Worksheet, cell As ExcelCell) As Integer
            If cell.SheetName <> Nothing AndAlso cell.SheetName <> sheet.Name Then Throw New ArgumentException("Argument sheet must match sheetname in argument cell")
            Return CType(sheet.Range(cell.Address), MsExcel.Range).Column - 1
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        Public Overrides Function LookupCellIsLocked(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            Return Me.LookupCellIsLocked(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet), rowIndex, columnIndex)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        Private Overloads Function LookupCellIsLocked(sheet As MsExcel.Worksheet, rowIndex As Integer, columnIndex As Integer) As Boolean
            Return CType(CType(sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).Locked, Boolean)
        End Function

        Public Overrides Function IsEmptyCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.IsEmptyCell(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet), rowIndex, columnIndex)
        End Function

        Private Overloads Function IsEmptyCell(sheet As MsExcel.Worksheet, rowIndex As Integer, columnIndex As Integer) As Boolean
            Return Me.LookupCellFormula(sheet, rowIndex, columnIndex) Is Nothing AndAlso Me.LookupCellValueAsObject(sheet, rowIndex, columnIndex) Is Nothing
        End Function

        ''' <summary>
        ''' Lookup the last content column index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastContentColumnIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim LastCell As MsExcel.Range = Sheet.Cells.SpecialCells(MsExcel.XlCellType.xlCellTypeLastCell)
            Dim autoSuggestionLastRowIndex As Integer = LastCell.Row - 1
            Dim autoSuggestedResult As Integer = LastCell.Column - 1
            For colCounter As Integer = autoSuggestedResult To 0 Step -1
                For rowCounter As Integer = 0 To autoSuggestionLastRowIndex
                    If IsEmptyCell(Sheet, rowCounter, colCounter) = False Then
                        Return colCounter
                    End If
                Next
            Next
            Return 0
        End Function

        ''' <summary>
        ''' Lookup the last content row index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastContentRowIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim LastCell As MsExcel.Range = Sheet.Cells.SpecialCells(MsExcel.XlCellType.xlCellTypeLastCell)
            Dim autoSuggestionLastColumnIndex As Integer = LastCell.Column - 1
            Dim autoSuggestedResult As Integer = LastCell.Row - 1
            For rowCounter As Integer = autoSuggestedResult To 0 Step -1
                For colCounter As Integer = 0 To autoSuggestionLastColumnIndex
                    If IsEmptyCell(Sheet, rowCounter, colCounter) = False Then
                        Return rowCounter
                    End If
                Next
            Next
            Return 0
        End Function

        ''' <summary>
        ''' Lookup the last content cell (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastContentCell(sheetName As String) As ExcelOps.ExcelCell
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim LastCell As MsExcel.Range = Sheet.Cells.SpecialCells(MsExcel.XlCellType.xlCellTypeLastCell)
            Dim CellRowIndex As Integer = Me.LookupLastContentRowIndex(sheetName)
            Dim CellColIndex As Integer = Me.LookupLastContentColumnIndex(sheetName)
            Return New ExcelOps.ExcelCell(sheetName, CellRowIndex, CellColIndex, Nothing)
        End Function

        ''' <summary>
        ''' Lookup the last content column index (zero based index)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastColumnIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim LastCell As MsExcel.Range = Sheet.Cells.SpecialCells(MsExcel.XlCellType.xlCellTypeLastCell)
            Return LastCell.Column - 1
        End Function

        ''' <summary>
        ''' Lookup the last content row index (zero based index)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastRowIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim LastCell As MsExcel.Range = Sheet.Cells.SpecialCells(MsExcel.XlCellType.xlCellTypeLastCell)
            Return LastCell.Row - 1
        End Function

        ''' <summary>
        ''' Lookup the last content cell
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastCell(sheetName As String) As ExcelOps.ExcelCell
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim LastCell As MsExcel.Range = Sheet.Cells.SpecialCells(MsExcel.XlCellType.xlCellTypeLastCell)
            Dim CellRowIndex As Integer = Me.LookupLastRowIndex(sheetName)
            Dim CellColIndex As Integer = Me.LookupLastColumnIndex(sheetName)
            Return New ExcelOps.ExcelCell(sheetName, CellRowIndex, CellColIndex, Nothing)
        End Function

        ''' <summary>
        ''' Remove specified rows
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="startRowIndex">0-based row number</param>
        ''' <param name="rows">Number of rows to remove</param>
        Public Overrides Sub RemoveRows(sheetName As String, startRowIndex As Integer, rows As Integer)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Me.RemoveRows(Sheet, startRowIndex, rows)
        End Sub

        ''' <summary>
        ''' Remove specified rows
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="startRowIndex">0-based row number</param>
        ''' <param name="rows">Number of rows to remove</param>
        Public Overloads Sub RemoveRows(sheet As MsExcel.Worksheet, startRowIndex As Integer, rows As Integer)
            If rows < 0 Then Throw New ArgumentOutOfRangeException(NameOf(rows), "Row number must be a positive value or zero")
            If rows = 0 Then Return
            Dim CellOfFirstRow As MsExcel.Range = CType(sheet.Cells(startRowIndex + 1, 1), MsExcel.Range)
            Dim CellOfLastRow As MsExcel.Range = CType(sheet.Cells(startRowIndex + rows - 1, 1), MsExcel.Range)
            Dim RangeRemovalRows As MsExcel.Range = sheet.Range(CellOfFirstRow, CellOfLastRow)
            Dim RemovalRows As MsExcel.Range = RangeRemovalRows.EntireRow
            RemovalRows.Delete(MsExcel.XlDeleteShiftDirection.xlShiftUp)
        End Sub

        Public Overrides Sub WriteCellValue(Of T)(cell As ExcelCell, value As T)
            Me.WriteCellValue(Of T)(CType(Me.Workbook.Worksheets(cell.SheetName), MsExcel.Worksheet), cell, value)
        End Sub

        Private Overloads Sub WriteCellValue(Of T)(sheet As MsExcel.Worksheet, cell As ExcelCell, value As T)
            If cell.SheetName <> Nothing AndAlso cell.SheetName <> sheet.Name Then Throw New ArgumentException("Argument sheet must match sheetname in argument cell")
            CType(sheet.Range(cell.Address), MsExcel.Range).Value = value
        End Sub

        Public Overrides Sub WriteCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer, value As T)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.WriteCellValue(Of T)(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet), rowIndex, columnIndex, value)
        End Sub

        Private Overloads Sub WriteCellValue(Of T)(sheet As MsExcel.Worksheet, rowIndex As Integer, columnIndex As Integer, value As T)
            CType(sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).Value = value
        End Sub

        Public Overrides Sub WriteCellFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer, formula As String, immediatelyCalculateCellValue As Boolean)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.WriteCellFormula(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet), rowIndex, columnIndex, formula, immediatelyCalculateCellValue)
        End Sub

        Private Overloads Sub WriteCellFormula(sheet As MsExcel.Worksheet, rowIndex As Integer, columnIndex As Integer, formula As String, immediatelyCalculateCellValue As Boolean)
            If formula <> Nothing Then
                CType(sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).Formula = "=" & formula
            Else
                CType(sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).Formula = Nothing
            End If
            If immediatelyCalculateCellValue Then
                CType(sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).Calculate()
            Else
                Me.RecalculationRequired = True
            End If
        End Sub

        Public Overrides Sub UnprotectSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Unprotect()
            CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).EnableSelection = MsExcel.XlEnableSelection.xlNoRestrictions
        End Sub

        Public Overrides Sub ProtectSheet(sheetName As String, level As ProtectionLevel)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Select Case level
                Case ProtectionLevel.StandardWithInsertDeleteRows
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).EnableSelection = MsExcel.XlEnableSelection.xlNoRestrictions
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Protect(AllowInsertingRows:=True, AllowDeletingRows:=True, AllowDeletingColumns:=False, AllowFiltering:=False, AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowSorting:=False, AllowUsingPivotTables:=False, AllowInsertingHyperlinks:=False)
                Case ProtectionLevel.Standard
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).EnableSelection = MsExcel.XlEnableSelection.xlNoRestrictions
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Protect(AllowInsertingRows:=False, AllowDeletingRows:=False, AllowDeletingColumns:=False, AllowFiltering:=False, AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowSorting:=False, AllowUsingPivotTables:=False, AllowInsertingHyperlinks:=False)
                Case ProtectionLevel.SelectAndEditUnlockedCellsOnly
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).EnableSelection = MsExcel.XlEnableSelection.xlUnlockedCells
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Protect(AllowInsertingRows:=False, AllowDeletingRows:=False, AllowDeletingColumns:=False, AllowFiltering:=False, AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowSorting:=False, AllowUsingPivotTables:=False, AllowInsertingHyperlinks:=False)
                Case ProtectionLevel.SelectAndEditAllCellsButNoFurtherEditing
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).EnableSelection = MsExcel.XlEnableSelection.xlNoRestrictions
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Protect(AllowInsertingRows:=False, AllowDeletingRows:=False, AllowDeletingColumns:=False, AllowFiltering:=False, AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowSorting:=False, AllowUsingPivotTables:=False, AllowInsertingHyperlinks:=False)
                Case ProtectionLevel.SelectNoCellsAndNoEditing
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).EnableSelection = MsExcel.XlEnableSelection.xlNoSelection
                    CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Protect(AllowInsertingRows:=False, AllowDeletingRows:=False, AllowDeletingColumns:=False, AllowFiltering:=False, AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowSorting:=False, AllowUsingPivotTables:=False, AllowInsertingHyperlinks:=False)
                Case Else
                    Throw New NotImplementedException
            End Select
        End Sub

        Protected Overrides Sub RecalculateAllInternal()
            If Me.CalculationModuleDisabled Then Throw New InvalidOperationException("Calculation engine is disabled, requested recalculation failed")
            Me.Workbook.Application.CalculateFullRebuild()
        End Sub

        Public Overrides Sub RecalculateSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.RecalculateSheet(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet))
        End Sub

        Public Overloads Sub RecalculateSheet(sheet As MsExcel.Worksheet)
            If Me.CalculationModuleDisabled Then Throw New InvalidOperationException("Calculation engine is disabled, requested recalculation failed")
            sheet.Calculate()
        End Sub

        ''' <summary>
        ''' Recalculate a cell based on its formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        Public Overrides Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            CType(Sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range).Calculate()
        End Sub

        Public Overrides Function IsProtectedSheet(sheetName As String) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.IsProtectedSheet(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet))
        End Function

        Private Overloads Function IsProtectedSheet(sheet As MsExcel.Worksheet) As Boolean
            Return sheet.ProtectionMode OrElse sheet.ProtectContents OrElse sheet.ProtectScenarios OrElse sheet.ProtectDrawingObjects
        End Function

        Public Overrides Sub RemoveSheet(sheetName As String)
            Me.RemoveSheet(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet))
        End Sub

        Public Overloads Sub RemoveSheet(sheet As MsExcel.Worksheet)
            sheet.Delete()
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public Overrides Sub ClearCells(sheetName As String, rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)
            If rangeFirstCell.SheetName <> rangeLastCell.SheetName Then Throw New ArgumentException("Cells must be member of the same worksheet")
            If sheetName <> Nothing Then
                rangeFirstCell = rangeFirstCell.Clone
                rangeLastCell = rangeLastCell.Clone
                rangeFirstCell.SheetName = sheetName
                rangeLastCell.SheetName = sheetName
            End If
            If rangeFirstCell.SheetName = Nothing Then Throw New ArgumentNullException(NameOf(rangeFirstCell))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(rangeFirstCell.SheetName), MsExcel.Worksheet)
            Me.ClearCells(Sheet, rangeFirstCell, rangeLastCell)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public Overloads Sub ClearCells(sheet As MsExcel.Worksheet, rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)
            If rangeFirstCell.SheetName <> rangeLastCell.SheetName Then Throw New ArgumentException("Cells must be member of the same worksheet")
            If rangeFirstCell.SheetName = Nothing Then Throw New ArgumentNullException(NameOf(rangeFirstCell))
            Dim ClearingRange As MsExcel.Range = sheet.Range(rangeFirstCell.Address & ":" & rangeLastCell.Address)
            ClearingRange.Clear()
        End Sub

        Public Overrides Sub AddSheet(sheetName As String, beforeSheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim AddedSheet As Object
            If beforeSheetName <> Nothing Then
                Dim BeforeSheet As MsExcel.Worksheet
                BeforeSheet = CType(Me.Workbook.Worksheets.Item(beforeSheetName), MsExcel.Worksheet)
                AddedSheet = Me.Workbook.Worksheets.Add(BeforeSheet, Type.Missing, Type.Missing, Type.Missing)
            Else
                AddedSheet = Me.Workbook.Worksheets.Add(After:=Me.Workbook.Worksheets(Me.Workbook.Worksheets.Count))
            End If
            CType(AddedSheet, MsExcel.Worksheet).Name = sheetName
        End Sub

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetIndex"></param>
        Public Overrides Sub SelectSheet(sheetIndex As Integer)
            Me.SelectSheet(CType(Me.Workbook.Worksheets(sheetIndex + 1), MsExcel.Worksheet))
        End Sub

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Sub SelectSheet(sheetName As String)
            If sheetName Is Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.SelectSheet(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet))
        End Sub

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheet"></param>
        Public Overloads Sub SelectSheet(sheet As MsExcel.Worksheet)
            If sheet Is Nothing Then Throw New ArgumentNullException(NameOf(sheet))
            sheet.Select()
        End Sub

        Public Overrides Sub UnhideSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Visible = MsExcel.XlSheetVisibility.xlSheetVisible
        End Sub

        Public Overrides Sub HideSheet(sheetName As String, stronglyHide As Boolean)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If stronglyHide Then
                CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Visible = MsExcel.XlSheetVisibility.xlSheetVeryHidden
            Else
                CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Visible = MsExcel.XlSheetVisibility.xlSheetHidden
            End If
        End Sub

        Public Overrides Function IsHiddenSheet(sheetName As String) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.IsHiddenSheet(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet))
        End Function

        Private Overloads Function IsHiddenSheet(sheet As MsExcel.Worksheet) As Boolean
            Return sheet.Visible <> MsExcel.XlSheetVisibility.xlSheetVisible
        End Function

        Public Overrides Function LookupCellErrorValue(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim CellValue As Object = Me.LookupCellValueAsObject(sheetName, rowIndex, columnIndex)
            If CellValue IsNot Nothing AndAlso Me.IsXLCVErr(CellValue) Then
                Select Case CType(CellValue, CVErrEnum)
                    Case CVErrEnum.ErrDiv0
                        Return CVErrNameDiv0
                    Case CVErrEnum.ErrNA
                        Return CVErrNameNA
                    Case CVErrEnum.ErrName
                        Return CVErrNameName
                    Case CVErrEnum.ErrNull
                        Return CVErrNameNull
                    Case CVErrEnum.ErrNum
                        Return CVErrNameNum
                    Case CVErrEnum.ErrRef
                        Return CVErrNameRef
                    Case CVErrEnum.ErrValue
                        Return CVErrNameValue
                    Case Else
                        Throw New NotImplementedException("Invalid error value " & CType(CellValue, Int32))
                End Select
            Else
                Return Nothing
            End If
        End Function

        Private Function LookupCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As MsExcel.Range
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return CType(CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range)
        End Function

        Private Function LookupCell(sheet As MsExcel.Worksheet, rowIndex As Integer, columnIndex As Integer) As MsExcel.Range
            Return CType(sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range)
        End Function

        Private Const CVErrNameDiv0 As String = "#DIV/0!"
        Private Const CVErrNameNA As String = "#N/A"
        Private Const CVErrNameName As String = "#NAME?"
        Private Const CVErrNameNull As String = "#NULL!"
        Private Const CVErrNameNum As String = "#NUM!"
        Private Const CVErrNameRef As String = "#REF!"
        Private Const CVErrNameValue As String = "#VALUE!"

        Private Enum CVErrEnum As Int32
            ErrDiv0 = -2146826281
            ErrNA = -2146826246
            ErrName = -2146826259
            ErrNull = -2146826288
            ErrNum = -2146826252
            ErrRef = -2146826265
            ErrValue = -2146826273
        End Enum

        Private Function IsXLCVErr(obj As Object) As Boolean
            'Return CType(Me.MSExcelApp, Excel.WorksheetFunction).IsError(obj)
            Return TypeOf (obj) Is Int32
        End Function

        Private Function IsXLCVErr(obj As Object, whichError As CVErrEnum) As Boolean
            If TypeOf (obj) Is Int32 Then
                Return CType(obj, Int32) = whichError
            Else
                Return False
            End If
        End Function

        Public Overrides Sub ClearSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Cells.Clear()
        End Sub

        Public Overrides Sub CopySheetContentInternal(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If targetWorkbook.GetType IsNot GetType(MsExcelDataOperations) Then Throw New NotSupportedException("Excel engines must be the same for source and target workbook for copying worksheets")
            If Me.MsExcelAppInstance.ComObjectStronglyTyped IsNot CType(targetWorkbook, MsExcelDataOperations).MsExcelAppInstance.ComObjectStronglyTyped Then Throw New NotSupportedException("Excel application must be the same for source and target workbook for copying worksheets")
            Dim TargetWorkSheet As MsExcel.Worksheet = CType(CType(targetWorkbook, MsExcelDataOperations).Workbook.Worksheets(targetSheetName), MsExcel.Worksheet)
            targetWorkbook.ClearSheet(targetSheetName)
            CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Cells.Copy(TargetWorkSheet.Cells)
        End Sub

        Public Overrides Sub SelectCell(cell As ExcelCell)
            If cell.SheetName = Nothing Then Throw New ArgumentException("Sheet name required", NameOf(cell))
            CType(CType(Me.Workbook.Worksheets(cell.SheetName), MsExcel.Worksheet).Cells(cell.RowIndex + 1, cell.ColumnIndex + 1), MsExcel.Range).Select()
        End Sub

        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "Microsoft Excel (2013 or higher)"
            End Get
        End Property

        Public Overrides ReadOnly Property HasVbaProject As Boolean
            Get
                Return Me.Workbook.HasVBProject
            End Get
        End Property

        Protected Overrides ReadOnly Property WorkbookFilePath As String
            Get
                If Me.IsClosed Then
                    Return Nothing
                ElseIf Me.Workbook.FullName.Contains(".") = False Then 'e.g. Mappe1 --> is a name without file name extension --> indicates that it hasn't been saved, yet
                    Return Nothing
                Else
                    Return Me.Workbook.FullName
                End If
            End Get
        End Property

        Public Overrides Sub RemoveVbaProject()
            If Me.Workbook.HasVBProject = False Then Return 'Shortcut and circumvent following workaround

            'NOTE: Manufacturer component doesn't provide a direct way to remove the VBA project (removing from Me.Workbook.VBProject.VBComponents list typically fails because of "not trusted")
            'NOTE: VBA project will be removed automatically when saving as non-xlsm-file            

            ''0. Lookup required private field of Spire.Xls
            'Dim XlsWorkbookMembers = CompuMaster.Reflection.NonPublicInstanceMembers.GetMembers(Of System.Reflection.FieldInfo)(Me.Workbook.GetType, GetType(Spire.Xls.Core.Spreadsheet.XlsWorkbook))
            'If XlsWorkbookMembers.Count <> 1 Then
            '    Throw New NotSupportedException("Spire.Xls incompatibility, please open an issue at https://github.com/CompuMasterGmbH/CompuMaster.Excel")
            'End If

            '0. Preserve required values for later reset
            'Dim XlsWb = CompuMaster.Reflection.NonPublicInstanceMembers.InvokeFieldGet(Of Spire.Xls.Core.Spreadsheet.XlsWorkbook)(Me.Workbook, Me.Workbook.GetType, XlsWorkbookMembers(0).Name)
            Dim PreservedFileName As String = Me.Workbook.FullName
            Dim PreservedIsSavedState As Boolean = Me.Workbook.Saved

            '1. Save to temp file
            Dim TempFile As String = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx"
            Me.Workbook.SaveAs(TempFile, FileFormat:=MsExcel.XlFileFormat.xlOpenXMLWorkbook)
            Me.Close()

            '2. Reload
            Me.LoadAndInitializeWorkbookFile(TempFile)

            '3. Reset FileName property
            'Me.WorkbookFilePath = PreservedFileName or Me.SetWorkbookFilePath(PreservedFileName) or similar not available for MS Excel via COM
            '-> IGNORE this action, here!

            '4. Reset IsSaved property
            Me.Workbook.Saved = PreservedIsSavedState
        End Sub

    End Class

End Namespace