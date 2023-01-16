Option Explicit On
Option Strict On
Imports MsExcel = Microsoft.Office.Interop.Excel

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
        Implements IDisposable

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
        ''' Are there any running MS Excel instances on the current system (owned by any user)
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function HasRunningMsExcelInstances() As Boolean
            Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Return MsExcelProcesses IsNot Nothing AndAlso MsExcelProcesses.Length > 0
        End Function

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
        Public Sub New(file As String, mode As OpenMode, unprotectWorksheets As Boolean, [readOnly] As Boolean)
            Me.New(file, mode, CreateMsExcelAppInstance, unprotectWorksheets, [readOnly])
        End Sub

#Disable Warning CA1034 ' Nested types should not be visible
        ''' <summary>
        ''' Class for holding a reference to Excel.Application (ATTENTION: watch for advised Try-Finally pattern!)
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
        Public Class MsAppInstance
#Enable Warning CA1034 ' Nested types should not be visible
            Public Sub New()
                Me.AppInstance = CreateMsExcelAppInstance()
            End Sub
            Friend Sub New(instance As MsExcel.Application)
                Me.AppInstance = instance
            End Sub
            Friend Property AppInstance As MsExcel.Application
            Public Sub Close()
                If Me.AppInstance IsNot Nothing Then
                    MsExcelDataOperations.PrepareCloseExcelAppInstance(Me)
                    MsExcelDataOperations.SafelyCloseExcelAppInstance(Me)
                End If
            End Sub
        End Class

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
        Private Shared Function CreateMsExcelAppInstance() As MsExcel.Application
            Dim MSExcelApp As New MsExcel.Application()
            Try
                MSExcelApp.Interactive = False
                MSExcelApp.ScreenUpdating = False
                MSExcelApp.DisplayAlerts = False
                MSExcelApp.Visible = False
            Catch ex As Exception
                Throw New PlatformNotSupportedException("App and installed MS Office must both 64 bit or both 32 bit processed")
            End Try
            Return MSExcelApp
        End Function

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
        Public Sub New(file As String, mode As OpenMode, msExcelApp As MsAppInstance, [readOnly] As Boolean)
            Me.New(file, mode, msExcelApp.AppInstance, True, [readOnly])
            Me._MsExcelAppInstance = msExcelApp
        End Sub

#Disable Warning IDE0060 ' Nicht verwendete Parameter entfernen
        Private Sub New(file As String, mode As OpenMode, msExcelApp As MsExcel.Application, unprotectWorksheets As Boolean, [readOnly] As Boolean)
#Enable Warning IDE0060 ' Nicht verwendete Parameter entfernen
            MyBase.New(True, False, [readOnly])
            Me.MSExcelApp = msExcelApp
            Me._Workbooks = msExcelApp.Workbooks
            Select Case mode
                Case OpenMode.OpenExistingFile
                    Me.LoadAndInitializeWorkbookFile(file)
                Case OpenMode.CreateFile
                    Me.CreateAndInitializeWorkbookFile(file)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(mode))
            End Select
            If unprotectWorksheets = True Then
                Me.UnprotectSheets()
            End If
        End Sub

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
        Protected Property MSExcelApp As MsExcel.Application

        Private _MsExcelAppInstance As MsAppInstance
        Public ReadOnly Property MsExcelAppInstance As MsAppInstance
            Get
                If _MsExcelAppInstance Is Nothing Then
                    _MsExcelAppInstance = New MsAppInstance(MSExcelApp)
                End If
                Return _MsExcelAppInstance
            End Get
        End Property

        Private Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer

        Public Shared Sub PrepareCloseExcelAppInstance(msExcelApp As MsAppInstance)
            If msExcelApp IsNot Nothing AndAlso msExcelApp.AppInstance IsNot Nothing Then
                PrepareCloseExcelAppInstanceInternal(msExcelApp.AppInstance)
            End If
        End Sub

        Private Shared Sub PrepareCloseExcelAppInstanceInternal(ByRef msExcelApp As MsExcel.Application)
            Try
                If msExcelApp IsNot Nothing Then
                    msExcelApp.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic 'reset value from manual to automatic (=expected default setting of user in 99% of all situations)
                End If
            Catch
            End Try
        End Sub

        Public Shared Sub SafelyCloseExcelAppInstance(ByRef msExcelApp As MsAppInstance)
            If msExcelApp Is Nothing Then Return
            SafelyCloseExcelAppInstanceInternal(msExcelApp.AppInstance)
            If msExcelApp IsNot Nothing Then
                msExcelApp.AppInstance = Nothing
            End If
            msExcelApp = Nothing
        End Sub

        Private Shared Sub SafelyCloseExcelAppInstanceInternal(ByRef msExcelApp As MsExcel.Application)
            If msExcelApp Is Nothing Then Return
            Dim ExcelProcess As System.Diagnostics.Process = Nothing
            Try
                Dim ExcelProcessID As Integer = Nothing
                GetWindowThreadProcessId(msExcelApp.Hwnd, ExcelProcessID)
                ExcelProcess = System.Diagnostics.Process.GetProcessById(ExcelProcessID)
            Catch
            End Try
            'Do Until msExcelApp.Workbooks.Count = 0
            '    Dim wb As MsExcel.Workbook = Nothing
            '    Try
            '        wb = msExcelApp.Workbooks(0)
            '    Catch
            '    End Try
            '    Try
            '        wb = msExcelApp.Workbooks(1)
            '    Catch
            '    End Try
            '    wb.Close(SaveChanges:=False)
            '    Do While System.Runtime.InteropServices.Marshal.ReleaseComObject(wb) > 0
            '    Loop
            '    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb)
            '    GC.Collect()
            '    GC.WaitForPendingFinalizers()
            'Loop
            Try
                msExcelApp.Quit()
            Catch
            End Try
            Do While System.Runtime.InteropServices.Marshal.ReleaseComObject(msExcelApp) > 0
            Loop
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(msExcelApp)
            msExcelApp = Nothing
            GC.Collect(0, GCCollectionMode.Forced, True, False)
            GC.WaitForPendingFinalizers()
            GC.Collect(2, GCCollectionMode.Forced, True, False)
            GC.WaitForPendingFinalizers()
            GC.Collect(0, GCCollectionMode.Forced, True, False)
            GC.WaitForPendingFinalizers()
            If ExcelProcess IsNot Nothing AndAlso ExcelProcess.HasExited = False Then
                ExcelProcess.Kill()
                System.Threading.Thread.Sleep(1000)
            End If
        End Sub

        Public Overrides Sub CloseExcelAppInstance()
            PrepareCloseExcelAppInstanceInternal(Me.MSExcelApp)
            'Close workbook if still open
            Me.Close()
            'Close workbooks collection
            If Me._Workbooks IsNot Nothing Then
                Me._Workbooks.Close()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Workbooks)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_Workbooks)
                Me._Workbooks = Nothing
            End If
            SafelyCloseExcelAppInstanceInternal(Me.MSExcelApp)
        End Sub

        Public Overrides Sub Close()
            If Me._Workbook IsNot Nothing Then
                Me._Workbook.Close(SaveChanges:=False)
                Do While System.Runtime.InteropServices.Marshal.ReleaseComObject(_Workbook) > 0
                Loop
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_Workbook)
                Me._Workbook = Nothing
            End If
        End Sub

        Private _Workbooks As MsExcel.Workbooks
        Public ReadOnly Property Workbooks As MsExcel.Workbooks
            Get
                Return _Workbooks
            End Get
        End Property

        Public Overrides ReadOnly Property IsClosed As Boolean
            Get
                Return Me._Workbook Is Nothing
            End Get
        End Property

        Public Overrides Sub Save()
            If Me.ReadOnly Then Throw New InvalidOperationException("Saving of read-only file forbidden")
            If Me.RecalculationRequired Then Me.RecalculateAll()
            If Me.FilePath <> Nothing AndAlso CType(Me.Workbook.Path, String) = Nothing Then
                'Created workbook, initial save must provide a file path, so use SaveAs method instead
                Me.SaveAs(Me.FilePath, SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Else
                Dim AutoCalcBuffer As Boolean = Me.AutoCalculationEnabled
                Try
                    Me.AutoCalculationEnabled = True
                    Me.Workbook.Save()
                Finally
                    Me.AutoCalculationEnabled = AutoCalcBuffer
                End Try
            End If
        End Sub

        Protected Overrides Sub SaveAsInternal(fileName As String, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            Select Case cachedCalculationsOption
                Case SaveOptionsForDisabledCalculationEngines.DefaultBehaviour, SaveOptionsForDisabledCalculationEngines.NoReset
                Case Else
                    Throw New NotSupportedException("SaveOptionsForDisabledCalculationEngines " & cachedCalculationsOption.ToString & " not supported by MS Excel")
            End Select
            If Me.RecalculationRequired Then Me.RecalculateAll()
            Dim AutoCalcBuffer As Boolean = Me.AutoCalculationEnabled
            Try
                Me.AutoCalculationEnabled = True
                Me.Workbook.SaveAs(fileName) 'calculation module is not disabled - just ignore 2nd argument
            Finally
                Me.AutoCalculationEnabled = AutoCalcBuffer
            End Try
        End Sub

        Private _Workbook As MsExcel.Workbook
        Public ReadOnly Property Workbook As MsExcel.Workbook
            Get
                Return _Workbook
            End Get
        End Property

        Public Overrides Property AutoCalculationEnabled As Boolean
            Get
                Return (Me.MSExcelApp.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic)
            End Get
            Set(value As Boolean)
                If value Then
                    Me.MSExcelApp.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic
                Else
                    Me.MSExcelApp.Calculation = MsExcel.XlCalculation.xlCalculationManual
                End If
            End Set
        End Property

        Protected Overrides Sub CreateWorkbook()
            If Me.FilePath <> Nothing AndAlso System.IO.File.Exists(Me.FilePath) = True Then Throw New System.InvalidOperationException("File already exists: " & Me.FilePath)
            If Me._Workbook Is Nothing Then
                Dim Wb As MsExcel.Workbook = Me._Workbooks.Add()
                Me._Workbook = Wb
                While Wb.Worksheets.Count > 1
                    CType(Wb.Worksheets(Wb.Worksheets.Count), MsExcel.Worksheet).Delete()
                End While
            End If
            If Me.MSExcelApp Is Nothing AndAlso Me.Workbook IsNot Nothing Then
                Me.MSExcelApp = Me.Workbook.Application
            End If
            Me.Workbook.EnableAutoRecover = False
        End Sub

        Protected Overrides Sub LoadWorkbook(file As System.IO.FileInfo)
            If file.Exists = False Then Throw New System.IO.FileNotFoundException("Workbook file must exist for loading from disk", file.FullName)
            If Me._Workbook Is Nothing Then
                Dim Wb As MsExcel.Workbook = Me._Workbooks.Open(file.FullName, UpdateLinks:=True, [ReadOnly]:=False, Editable:=False, Notify:=False)
                Me._Workbook = Wb
            End If
            If Me.MSExcelApp Is Nothing AndAlso Me.Workbook IsNot Nothing Then
                Me.MSExcelApp = Me.Workbook.Application
            End If
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

#Region "IDisposable Support"
        Private disposedValue As Boolean ' Dient zur Erkennung redundanter Aufrufe.

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: verwalteten Zustand (verwaltete Objekte) entsorgen.
                End If

                ' TODO: nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalize() weiter unten überschreiben.
                ' TODO: große Felder auf Null setzen.
                Me.Close()
                If Me.MSExcelApp IsNot Nothing Then
                    Me.CloseExcelAppInstance()
                End If
            End If
            disposedValue = True
        End Sub

        ' TODO: Finalize() nur überschreiben, wenn Dispose(disposing As Boolean) weiter oben Code zur Bereinigung nicht verwalteter Ressourcen enthält.
        'Protected Overrides Sub Finalize()
        '    ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in Dispose(disposing As Boolean) weiter oben ein.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in Dispose(disposing As Boolean) weiter oben ein.
            Dispose(True)
            ' TODO: Auskommentierung der folgenden Zeile aufheben, wenn Finalize() oben überschrieben wird.
            ' GC.SuppressFinalize(Me)
        End Sub

#End Region

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

        Public Overrides Sub CopySheetContent(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If targetWorkbook.GetType IsNot GetType(MsExcelDataOperations) Then Throw New NotSupportedException("Excel engines must be the same for source and target workbook for copying worksheets")
            If Me.MSExcelApp IsNot CType(targetWorkbook, MsExcelDataOperations).MSExcelApp Then Throw New NotSupportedException("Excel application must be the same for source and target workbook for copying worksheets")
            Dim TargetWorkSheet As MsExcel.Worksheet = CType(CType(targetWorkbook, MsExcelDataOperations).Workbook.Worksheets(targetSheetName), MsExcel.Worksheet)
            targetWorkbook.ClearSheet(targetSheetName)
            CType(Me.Workbook.Worksheets(sheetName), MsExcel.Worksheet).Cells.Copy(TargetWorkSheet.Cells)
        End Sub

        Public Shared Sub RecalculateFile(filePath As String)
            RecalculateFile(filePath, Nothing)
        End Sub

        Public Shared Sub RecalculateFile(filePath As String, msAppInstance As MsExcelDataOperations.MsAppInstance)
            Dim MsExcelApp As MsAppInstance = msAppInstance
            If MsExcelApp Is Nothing Then
                MsExcelApp = New MsAppInstance()
            End If
            Dim wb As New MsExcelDataOperations(filePath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelApp, False)
            Try
                wb.RecalculateAll()
                wb.Save()
            Finally
                wb.Close()
                If msAppInstance Is Nothing Then wb.CloseExcelAppInstance()
            End Try
        End Sub

        Public Overrides Sub SelectCell(cell As ExcelCell)
            If cell.SheetName = Nothing Then Throw New ArgumentException("Sheet name required", NameOf(cell))
            CType(CType(Me.Workbook.Worksheets(cell.SheetName), MsExcel.Worksheet).Cells(cell.RowIndex + 1, cell.ColumnIndex + 1), MsExcel.Range).Select()
        End Sub

        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "MS Excel"
            End Get
        End Property

    End Class

End Namespace