Imports Microsoft.Office.Interop.Excel
Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.MsExcelCom

    ''' <summary>
    ''' An MS Excel wrapper for safe COM object handling and release
    ''' </summary>
    Public Class MsExcelApplicationWrapper
        Inherits CompuMaster.ComInterop.ComRootObject(Of MsExcel.Application)

        ''' <summary>
        ''' Create a new MS Excel instance within its wrapper instance
        ''' </summary>
        Public Sub New()
            MyBase.New(CreateMsExcelApplication, Nothing)
            Me.ComObjectStronglyTyped.Visible = False
            Me.ComObjectStronglyTyped.Interactive = False
            Me.ComObjectStronglyTyped.ScreenUpdating = False
            Me.ComObjectStronglyTyped.DisplayAlerts = False
            Try
                Dim ExcelProcessID As Integer = Nothing
                GetWindowThreadProcessId(Me.ComObjectStronglyTyped.Hwnd, ExcelProcessID)
                Me.ProcessId = ExcelProcessID
            Catch
            End Try
            Me.Workbooks.CloseAllWorkbooks() 'Close initial empty workbook which is always there after 
        End Sub

        Private Shared Function CreateMsExcelApplication() As MsExcel.Application
            Try
                Return New MsExcel.Application()
            Catch ex As PlatformNotSupportedException
                Throw
            Catch ex As Exception
                Throw New PlatformNotSupportedException(ex.Message, ex)
            End Try
        End Function

        Private Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer

        ''' <summary>
        ''' The process ID of the COM server
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ProcessId As Integer

        ''' <summary>
        ''' The process of the COM server
        ''' </summary>
        ''' <returns></returns>
        Public Function Process() As System.Diagnostics.Process
            If Me.ProcessId = 0 Then
                Return Nothing
            Else
                Return System.Diagnostics.Process.GetProcessById(Me.ProcessId)
            End If
        End Function

        ''' <summary>
        ''' Required close commands for the COM object like App.Quit() or Document.Close()
        ''' </summary>
        Protected Overrides Sub OnClosing()
            If Not Me.IsDisposedComObject Then
                Try
                    Me.ComObjectStronglyTyped.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic 'reset value from manual to automatic (=expected default setting of user in 99% of all situations)
                Catch
                End Try
            End If
            Me.ComObjectStronglyTyped.Quit()
            MyBase.OnClosing()
        End Sub

        ''' <summary>
        ''' Required actions after the COM object has been closed, e.g. removing from a list of open documents
        ''' </summary>
        Protected Overrides Sub OnClosed()
            MyBase.OnClosed()
            CompuMaster.ComInterop.ComTools.ReleaseComObject(Me.ComObject)
            SafelyCloseExcelAppInstanceInternal()
        End Sub

        ''' <summary>
        ''' A timeout value for closing MS Excel regulary, default to 15 seconds
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>After timeout, process will be killed if the process hasn't exited</remarks>
        Public Property Timeout1AfterAppClosing As New TimeSpan(0, 0, 15)

        ''' <summary>
        ''' A timeout value for process exiting after MS Excel process has been killed, defaults to 1 second
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>After timeout, code waits for disappearance of process in process list</remarks>
        Public Property Timeout2ProcessExitAfterAppKill As New TimeSpan(0, 0, 1)

        ''' <summary>
        ''' A timeout value for watching process list for disappeared MS Excel process, defaults to 1 second
        ''' </summary>
        ''' <returns>After timeout, process is expected to be closed "with chance of 99.99%" (not guaranteed)</returns>
        Public Property Timeout3ProcessListDisappearanceAfterAppKill As New TimeSpan(0, 0, 1)

        ''' <summary>
        ''' At some unkown circumstances, MS Excel process wasn't closed sometimes and required a forced process killing
        ''' </summary>
        Private Sub SafelyCloseExcelAppInstanceInternal()
            If ProcessId <> Nothing AndAlso Process() IsNot Nothing AndAlso Process.HasExited = False Then
                Try
                    MsExcelTools.WaitUntilTrueOrTimeout(Function() Me.Process.HasExited = True, Timeout1AfterAppClosing) 'Sometimes it takes time to close MS Excel...
                    Me.Process.Refresh()
                    If Me.Process.HasExited = False Then
                        'Force kill on Excel 
                        Me.Process.Kill()
                        Try
                            MsExcelTools.WaitUntilTrueOrTimeout(Function() Me.Process.HasExited = True, Timeout2ProcessExitAfterAppKill)
                        Catch 'ex As ArgumentException
                            'expected for invalid processId after kill
                        End Try
                        Try
                            MsExcelTools.WaitUntilTrueOrTimeout(Function()
                                                                    Dim ExcelProcesses() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                                                                    For Each ExcelProcess In ExcelProcesses
                                                                        If ExcelProcess.Id = Me.ProcessId Then Return False
                                                                    Next
                                                                    Return True
                                                                End Function, Timeout3ProcessListDisappearanceAfterAppKill)
                        Catch 'ex As Exception
                            'ignore any exceptions on getting process list
                        End Try
                    End If
                Catch 'ex As Exception
                    'ignore any exceptions of watching/handling process close/kill
                End Try
            End If
        End Sub

        Public ReadOnly Property IsDisposed As Boolean
            Get
                Return MyBase.IsDisposedComObject
            End Get
        End Property

        Private _Workbooks As MsExcelWorkbooksWrapper
        ''' <summary>
        ''' The Excel workbooks collection
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Workbooks As MsExcelWorkbooksWrapper
            Get
                If _Workbooks Is Nothing Then
                    _Workbooks = New MsExcelWorkbooksWrapper(Me, Me.ComObjectStronglyTyped.Workbooks)
                End If
                Return _Workbooks
            End Get
        End Property

    End Class

End Namespace