Imports Microsoft.Office.Interop.Excel
Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.MsExcelCom

    ''' <summary>
    ''' An MS Excel wrapper for safe COM object handling and release
    ''' </summary>
    Public Class MsExcelApplicationWrapper
        Inherits CompuMaster.ComInterop.ComApplication(Of MsExcel.Application)

        Const ExpectedProcessName As String = "EXCEL"

        ''' <summary>
        ''' Create a new MS Excel instance within its wrapper instance
        ''' </summary>
        Public Sub New()
            MyBase.New(CreateMsExcelApplication, Function(x) x.ComObjectStronglyTyped.Hwnd, ExpectedProcessName)
            Me.ComObjectStronglyTyped.Visible = False
            Me.ComObjectStronglyTyped.Interactive = False
            Me.ComObjectStronglyTyped.ScreenUpdating = False
            Me.ComObjectStronglyTyped.DisplayAlerts = False
            Me.SetCultureContext(System.Threading.Thread.CurrentThread.CurrentCulture) 'Always set MS Excel culture context to current thread's culture
            Me.Workbooks.CloseAllWorkbooks() 'Close initial empty workbook which is always there after 
        End Sub

        Private Shared Function CreateMsExcelApplication() As MsExcel.Application
            Try
                Return New MsExcel.Application()
            Catch ex As PlatformNotSupportedException
                Throw
            Catch ex As System.Runtime.InteropServices.COMException
                Throw New CompuMaster.ComInterop.ComApplicationNotAvailableException("Microsoft Excel application not available", ex)
            Catch ex As Exception
                Throw New PlatformNotSupportedException(ex.Message, ex)
            End Try
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
                Me.ComObjectStronglyTyped.Quit()
            End If
            MyBase.OnClosing()
        End Sub

        Public ReadOnly Property ExcelProcessId As Integer
            Get
                Return MyBase.ProcessId
            End Get
        End Property

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

        Public Sub SetCultureContext(culture As System.Globalization.CultureInfo)
            Me.ComObjectStronglyTyped.UseSystemSeparators = (culture Is Nothing) 'True allows customization in lines below, False uses system settings
            Me.ComObjectStronglyTyped.DecimalSeparator = culture.NumberFormat.NumberDecimalSeparator
            Me.ComObjectStronglyTyped.ThousandsSeparator = culture.NumberFormat.NumberGroupSeparator
        End Sub

    End Class

End Namespace