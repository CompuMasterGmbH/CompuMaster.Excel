Option Explicit On
Option Strict On

Imports System.Text
Imports CompuMaster.Excel.MsExcelCom
Imports Microsoft.Office.Interop.Excel
Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.ExcelOps

    ''' <summary>
    ''' MS Excel Interop provider (ATTENTION: watch for advised Try-Finally pattern for successful application process stop!)
    ''' </summary>
    ''' <remarks>
    ''' For proper Microsoft Excel licensing, please contact Microsoft.
    ''' PLEASE NOTE: Considerations for server-side Automation of Office https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2
    ''' </remarks>
    Public Class MsExcelDataOperations
        Inherits ExcelDataOperationsBase

        Public Shared Property AutoKillAllExistingMsExcelInstances As Boolean

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

        Protected Overrides ReadOnly Property DefaultCalculationOptions As ExcelEngineDefaultOptions
            Get
                Return New ExcelEngineDefaultOptions(False, False)
            End Get
        End Property

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
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", True)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
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
        Friend Sub New()
            MyBase.New(New ExcelDataOperationsOptions)
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
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", True)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
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
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", True)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
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
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", True)> '<Obsolete("Use overload; WARNING: this overload always leads to: unprotectWorksheets = True")>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
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
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", True)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(file As String, mode As OpenMode, msExcelApp As MsExcelApplicationWrapper, unprotectWorksheets As Boolean, [readOnly] As Boolean, passwordForOpening As String)
            Me.New(file, mode, msExcelApp, unprotectWorksheets, [readOnly], passwordForOpening, False)
        End Sub

        ''' <summary>
        ''' MS Excel Interop provider (ATTENTION: watch for advised Try-Finally pattern for successful application process stop!) incl. unprotection of sheets
        ''' </summary>
        ''' <param name="disableAutoCalculation">Disable initial and auto-calculations</param>
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
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", True)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(file As String, mode As OpenMode, msExcelApp As MsExcelApplicationWrapper, unprotectWorksheets As Boolean, [readOnly] As Boolean, passwordForOpening As String, disableAutoCalculation As Boolean)
#Disable Warning IDE0060 ' Nicht verwendete Parameter entfernen
#Enable Warning IDE0060 ' Nicht verwendete Parameter entfernen
            MyBase.New(Not disableAutoCalculation, False, [readOnly], passwordForOpening)
            Me._MsExcelAppInstance = msExcelApp
            Me._Workbooks = New MsExcelWorkbooksWrapper(msExcelApp, msExcelApp.ComObjectStronglyTyped.Workbooks)
            If disableAutoCalculation Then
                Me.AutoCalculationEnabled = False
            End If
            Select Case mode
                Case OpenMode.OpenExistingFile
                    Me.LoadAndInitializeWorkbookFile(file, ConvertToUnvalidatedOptions(Not disableAutoCalculation, False, [readOnly], passwordForOpening))
                Case OpenMode.CreateFile
                    Me.CreateAndInitializeWorkbookFile(file, ConvertToUnvalidatedOptions(Not disableAutoCalculation, False, [readOnly], passwordForOpening))
                    Me.ReadOnly = [readOnly] OrElse (file = Nothing)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(mode))
            End Select
            If unprotectWorksheets = True Then
                Me.UnprotectSheets()
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
        Public Sub New(file As String, mode As OpenMode, options As ExcelOps.ExcelDataOperationsOptions)
            Me.New(file, mode, New MsExcelApplicationWrapper, False, options)
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
        Public Sub New(file As String, mode As OpenMode, unprotectWorksheets As Boolean, options As ExcelOps.ExcelDataOperationsOptions)
            Me.New(file, mode, New MsExcelApplicationWrapper, unprotectWorksheets, options)
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
        Public Sub New(file As String, mode As OpenMode, msExcelApp As MsExcelApplicationWrapper, unprotectWorksheets As Boolean, options As ExcelOps.ExcelDataOperationsOptions)
#Disable Warning IDE0060 ' Nicht verwendete Parameter entfernen
#Enable Warning IDE0060 ' Nicht verwendete Parameter entfernen
            MyBase.New(options)
            Me._MsExcelAppInstance = msExcelApp
            Me._Workbooks = New MsExcelWorkbooksWrapper(msExcelApp, msExcelApp.ComObjectStronglyTyped.Workbooks)
            Select Case mode
                Case OpenMode.OpenExistingFile
                    Me.LoadAndInitializeWorkbookFile(file, options)
                Case OpenMode.CreateFile
                    Me.CreateAndInitializeWorkbookFile(file, options)
                    Me.ReadOnly = [ReadOnly] OrElse (file = Nothing)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(mode))
            End Select
            If unprotectWorksheets = True Then
                Me.UnprotectSheets()
            End If
        End Sub

        Protected Overrides Sub ValidateLoadOptions(options As ExcelDataOperationsOptions)
            If options.DisableCalculationEngine.Value = True Then Throw New NotSupportedException("MS Excel doesn't support disabling of calculation engine")
            MyBase.ValidateLoadOptions(options)
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

        ''' <summary>
        ''' If enabled, the calculation engine will do a full recalculation after every modification.
        ''' If disabled, the calculation engine is not allowed to automatically/continuously calculate on every change and the user has to manually force a recalculation (typically by pressing F9 key in MS Excel).
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>Please note: this property is a workbook property (not an engine property!)</remarks>
        Public Overrides Property AutoCalculationEnabledWorkbookSetting As Boolean
            Get
                If Me.MsExcelAppInstance IsNot Nothing AndAlso Me.MsExcelAppInstance.ComObjectStronglyTyped IsNot Nothing Then
                    Return (Me.MsExcelAppInstance.ComObjectStronglyTyped.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic)
                Else
                    Return MyBase.AutoCalculationEnabledWorkbookSetting
                End If
            End Get
            Set(value As Boolean)
                MyBase.AutoCalculationEnabledWorkbookSetting = value
                If Me.MsExcelAppInstance IsNot Nothing AndAlso Me.MsExcelAppInstance.ComObjectStronglyTyped IsNot Nothing Then
                    If value Then
                        Me.MsExcelAppInstance.ComObjectStronglyTyped.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic
                    Else
                        Me.MsExcelAppInstance.ComObjectStronglyTyped.Calculation = MsExcel.XlCalculation.xlCalculationManual
                    End If
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
                CType(Wb.Worksheets(1), MsExcel.Worksheet).Name = "Sheet1"
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
                Try
                    If Me.PasswordForOpening <> Nothing Then
                        Wb = Me.Workbooks.Open(file.FullName, UpdateLinks:=True, [ReadOnly]:=False, Editable:=False, Notify:=False, Password:=Me.PasswordForOpening)
                    Else
                        Wb = Me.Workbooks.Open(file.FullName, UpdateLinks:=True, [ReadOnly]:=False, Editable:=False, Notify:=False, Password:="")
                    End If
                Catch ex As System.Runtime.InteropServices.COMException
                    If ex.ErrorCode = &H800A03EC Then
                        Throw New FileCorruptedOrInvalidFileFormatException(file, ex)
                    Else
                        Throw
                    End If
                End Try
                If Wb Is Nothing Then Throw New NullReferenceException("Null result after Workbooks.Open")
                Me._Workbook = New MsExcelWorkbookWrapper(Me._Workbooks, Wb)
            End If
            'If Me.MSExcelApp Is Nothing AndAlso Me.Workbook IsNot Nothing Then
            '    Me.MSExcelApp = Me.Workbook.Application
            'End If
            Me.Workbook.EnableAutoRecover = False
        End Sub

        Protected Overrides Sub LoadWorkbook(data() As Byte)
            Throw New NotSupportedException()
        End Sub

        Protected Overrides Sub LoadWorkbook(data As IO.Stream)
            Throw New NotSupportedException()
        End Sub

        Public Overrides Sub CleanupRangeNames()
            'do nothing - just needs to be done once, see Epplus implementation
        End Sub

        Public Overrides Function SheetNames() As List(Of String)
            Dim Result As New List(Of String)
            For Each s As Object In Me.Workbook.Sheets
                If GetType(MsExcel.Worksheet).IsInstanceOfType(s) Then
                    Dim ws = CType(s, MsExcel.Worksheet)
                    Result.Add(ws.Name)
                ElseIf GetType(MsExcel.Chart).IsInstanceOfType(s) Then
                    Dim ws = CType(s, MsExcel.Chart)
                    Result.Add(ws.Name)
                Else
                    Throw New NotImplementedException
                End If
            Next
            Return Result
        End Function

        Public Overrides Function WorkSheetNames() As List(Of String)
            Dim Result As New List(Of String)
            For Each ws As MsExcel.Worksheet In Me.Workbook.Worksheets
                Result.Add(ws.Name)
            Next
            Return Result
        End Function

        Public Overrides Function ChartSheetNames() As List(Of String)
            Dim Result As New List(Of String)
            For Each ws As MsExcel.Chart In Me.Workbook.Charts
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
        Protected Overrides Function LookupLastContentColumnIndex(sheetName As String, lastMergedCell As ExcelCell) As Integer
            'INTERNAL NOTE: method override (only) for performance reasons
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim LastCell As MsExcel.Range = Sheet.Cells.SpecialCells(MsExcel.XlCellType.xlCellTypeLastCell)
            Dim autoSuggestionLastRowIndex As Integer = LastCell.Row - 1
            Dim autoSuggestedResult As Integer = LastCell.Column - 1
            Dim lastMergeCellRowIndex As Integer
            Dim lastMergeCellColumnIndex As Integer
            If LastMergedCell IsNot Nothing Then
                lastMergeCellRowIndex = LastMergedCell.RowIndex
                lastMergeCellColumnIndex = LastMergedCell.ColumnIndex
            Else
                lastMergeCellRowIndex = 0
                lastMergeCellColumnIndex = 0
            End If
            'Find last content cell
            For colCounter As Integer = autoSuggestedResult To lastMergeCellColumnIndex Step -1
                For rowCounter As Integer = lastMergeCellRowIndex To autoSuggestionLastRowIndex
                    If IsEmptyCell(Sheet, rowCounter, colCounter) = False Then
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
        Protected Overrides Function LookupLastContentRowIndex(sheetName As String, lastMergedCell As ExcelCell) As Integer
            'INTERNAL NOTE: method override (only) for performance reasons
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim LastCell As MsExcel.Range = Sheet.Cells.SpecialCells(MsExcel.XlCellType.xlCellTypeLastCell)
            Dim autoSuggestionLastColumnIndex As Integer = LastCell.Column - 1
            Dim autoSuggestedResult As Integer = LastCell.Row - 1
            Dim lastMergeCellRowIndex As Integer
            Dim lastMergeCellColumnIndex As Integer
            If LastMergedCell IsNot Nothing Then
                lastMergeCellRowIndex = LastMergedCell.RowIndex
                lastMergeCellColumnIndex = LastMergedCell.ColumnIndex
            Else
                lastMergeCellRowIndex = 0
                lastMergeCellColumnIndex = 0
            End If
            'Find last content cell
            For rowCounter As Integer = autoSuggestedResult To lastMergeCellRowIndex Step -1
                For colCounter As Integer = lastMergeCellColumnIndex To autoSuggestionLastColumnIndex
                    If IsEmptyCell(Sheet, rowCounter, colCounter) = False Then
                        Return System.Math.Max(lastMergeCellRowIndex, rowCounter)
                    End If
                Next
            Next
            Return System.Math.Max(lastMergeCellRowIndex, 0)
        End Function

        ''' <summary>
        ''' Lookup the last content cell
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastCell(sheetName As String) As ExcelOps.ExcelCell
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim UsedRange As New ExcelOps.ExcelRange(sheetName, Sheet.UsedRange.AddressLocal(False, False))
            Dim LastCell As New ExcelOps.ExcelCell(sheetName, Tools.LookupCellAddresFromRange(Sheet.Cells.SpecialCells(MsExcel.XlCellType.xlCellTypeLastCell).AddressLocal(False, False), 1), ExcelCell.ValueTypes.All)
            Return Tools.CombineCellAddresses(LastCell, UsedRange.AddressEnd, Tools.CellAddressCombineMode.RightLowerCorner)
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
        Public Overrides Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer, throwExceptionOnCalculationError As Boolean)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim ECell = CType(Sheet.Cells(rowIndex + 1, columnIndex + 1), MsExcel.Range)
            ECell.Calculate()
            If throwExceptionOnCalculationError AndAlso ECell.Value?.GetType Is GetType(MsExcel.XlCVError) Then
                Dim Cell As New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
                Throw New NotSupportedException("Epplus calculation at " & Cell.Address(True) & " resulted in #" & UCase(CType(ECell.Value, MsExcel.XlCVError).ToString) & "!" & " for formula =" & Me.LookupCellFormula(Cell))
            End If
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
        ''' <param name="overrideSheetName"></param>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public Overrides Sub ClearCells(overrideSheetName As String, rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)
            If rangeFirstCell.SheetName <> rangeLastCell.SheetName Then Throw New ArgumentException("Cells must be member of the same worksheet")
            If overrideSheetName <> Nothing Then
                rangeFirstCell = rangeFirstCell.Clone
                rangeLastCell = rangeLastCell.Clone
                rangeFirstCell.SheetName = overrideSheetName
                rangeLastCell.SheetName = overrideSheetName
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

        Public Overrides Function SelectedSheetName() As String
            Return CType(Me.Workbook.ActiveSheet, MsExcel.Worksheet).Name
        End Function

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

        Private Function LookupRange(sheetName As String, fromRowIndex As Integer, fromColumnIndex As Integer, toRowIndex As Integer, toColumnIndex As Integer) As MsExcel.Range
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Range(
                CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Cells(fromRowIndex + 1, fromColumnIndex + 1),
                CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Cells(toRowIndex + 1, toColumnIndex + 1)
                )
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
                ElseIf Me.Workbook.FullName.Contains("."c) = False Then 'e.g. Mappe1 --> is a name without file name extension --> indicates that it hasn't been saved, yet
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
            Me.LoadAndInitializeWorkbookFile(TempFile, Me.LoadOptions)

            '3. Reset FileName property
            'Me.WorkbookFilePath = PreservedFileName or Me.SetWorkbookFilePath(PreservedFileName) or similar not available for MS Excel via COM
            '-> IGNORE this action, here!

            '4. Reset IsSaved property
            Me.Workbook.Saved = PreservedIsSavedState
        End Sub

        Protected Overrides Function MergedCells(sheetName As String) As List(Of ExcelOps.ExcelRange)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Dim FoundRanges As New Collections.Specialized.StringCollection()
            For Each x As MsExcel.Range In Sheet.UsedRange
                If CType(x.MergeCells, Boolean) Then
                    Dim RangeAddress As String = x.MergeArea.AddressLocal(False, False)
                    If FoundRanges.Contains(RangeAddress) = False Then
                        FoundRanges.Add(RangeAddress)
                    End If
                End If
            Next
            Dim Result As New List(Of ExcelOps.ExcelRange)
            For MyCounter As Integer = 0 To FoundRanges.Count - 1
                Result.Add(New ExcelOps.ExcelRange(sheetName, FoundRanges(MyCounter)))
            Next
            Return Result
        End Function

        Public Overrides Function IsMergedCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return CType(LookupCell(sheetName, rowIndex, columnIndex).MergeCells, Boolean)
        End Function

        Public Overrides Sub UnMergeCells(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            LookupCell(sheetName, rowIndex, columnIndex).MergeArea.UnMerge()
        End Sub

        Public Overrides Sub MergeCells(sheetName As String, fromRowIndex As Integer, fromColumnIndex As Integer, toRowIndex As Integer, toColumnIndex As Integer)
            LookupRange(sheetName, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex).Merge()
        End Sub

        Public Overrides Sub AutoFitColumns(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Sheet.Columns.AutoFit()
        End Sub

        Public Overrides Sub AutoFitColumns(sheetName As String, minimumWidth As Double)
            Dim LastColIndex As Integer = Me.LookupLastColumnIndex(sheetName)
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            Sheet.Columns.AutoFit()
            For ColumnIndex As Integer = 0 To LastColIndex
                If CType(CType(Sheet.Columns(ColumnIndex + 1), Range).ColumnWidth, Double) < minimumWidth Then
                    CType(Sheet.Columns(ColumnIndex + 1), Range).ColumnWidth = minimumWidth
                End If
            Next
        End Sub

        Public Overrides Sub AutoFitColumns(sheetName As String, columnIndex As Integer)
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            CType(Sheet.Columns(columnIndex + 1), Range).AutoFit()
        End Sub

        Public Overrides Sub AutoFitColumns(sheetName As String, columnIndex As Integer, minimumWidth As Double)
            Dim Sheet As MsExcel.Worksheet = CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet)
            CType(Sheet.Columns(columnIndex + 1), Range).AutoFit()
            If CType(CType(Sheet.Columns(columnIndex + 1), Range).ColumnWidth, Double) < minimumWidth Then
                CType(Sheet.Columns(columnIndex + 1), Range).ColumnWidth = minimumWidth
            End If
        End Sub

        Public Overrides Function ExportChartSheetImage(chartSheetName As String) As System.Drawing.Image
            Throw New NotImplementedException()
        End Function

        Public Overrides Function ExportChartImage(workSheetName As String) As System.Drawing.Image()
            Throw New NotImplementedException()
        End Function

        ''' <summary>
        ''' Save worksheet to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="worksheetName"></param>
        ''' <param name="sb"></param>
        Protected Overrides Sub ExportSheetToHtmlInternal(worksheetName As String, sb As StringBuilder, options As HtmlSheetExportOptions)
            Throw New NotImplementedException
        End Sub

    End Class

End Namespace