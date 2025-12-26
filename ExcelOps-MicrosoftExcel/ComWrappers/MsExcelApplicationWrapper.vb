Imports Microsoft.Office.Interop.Excel
Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.MsExcelCom

    ''' <summary>
    ''' An MS Excel wrapper for safe COM object handling and release
    ''' </summary>
    ''' <remarks>
    ''' For proper Microsoft Excel licensing, please contact Microsoft.
    ''' PLEASE NOTE: Considerations for server-side Automation of Office https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2
    ''' </remarks>
    Public Class MsExcelApplicationWrapper
        Implements IDisposable

        Private _ComApp As CompuMaster.ComInterop.ComApplication(Of MsExcel.Application)

        Friend Const ExpectedProcessName As String = "EXCEL"

        ''' <summary>
        ''' Create a new MS Excel instance within its wrapper instance
        ''' </summary>
        Public Sub New()
            Me._ComApp = New CompuMaster.ComInterop.ComApplication(Of MsExcel.Application)(CreateMsExcelApplication, Function(x) New IntPtr(x.ComObjectStronglyTyped.Hwnd), AddressOf OnClosing, ExpectedProcessName)
            _ComApp.ComObjectStronglyTyped.Visible = False
            _ComApp.ComObjectStronglyTyped.Interactive = False
            _ComApp.ComObjectStronglyTyped.ScreenUpdating = False
            _ComApp.ComObjectStronglyTyped.DisplayAlerts = False
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
        Private Shared Sub OnClosing(comApplicationObject As CompuMaster.ComInterop.ComApplication(Of MsExcel.Application))
            If Not comApplicationObject.IsDisposedComObject Then
                comApplicationObject.ComObjectStronglyTyped.Quit()
            End If
        End Sub

        Public ReadOnly Property ComApp As CompuMaster.ComInterop.ComApplication(Of MsExcel.Application)
            Get
                Return _ComApp
            End Get
        End Property

        Public ReadOnly Property ComObject As Object
            Get
                Return _ComApp.ComObject
            End Get
        End Property

        Public ReadOnly Property ComObjectStronglyTyped As Application
            Get
                Return _ComApp.ComObjectStronglyTyped
            End Get
        End Property

        Public ReadOnly Property ExcelProcessId As Integer
            Get
                Return _ComApp.ProcessId
            End Get
        End Property

        Public Function ExcelProcess() As System.Diagnostics.Process
            Return _ComApp.Process
        End Function

        Private _Workbooks As MsExcelWorkbooksWrapper
        ''' <summary>
        ''' The Excel workbooks collection
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Workbooks As MsExcelWorkbooksWrapper
            Get
                If _Workbooks Is Nothing Then
                    _Workbooks = New MsExcelWorkbooksWrapper(Me, _ComApp.ComObjectStronglyTyped.Workbooks)
                End If
                Return _Workbooks
            End Get
        End Property

        Public Sub SetCultureContext(culture As System.Globalization.CultureInfo)
            _ComApp.ComObjectStronglyTyped.UseSystemSeparators = (culture Is Nothing) 'True allows customization in lines below, False uses system settings
            _ComApp.ComObjectStronglyTyped.DecimalSeparator = culture.NumberFormat.NumberDecimalSeparator
            _ComApp.ComObjectStronglyTyped.ThousandsSeparator = culture.NumberFormat.NumberGroupSeparator
        End Sub

        ''' <summary>
        ''' Is MS Excel application closed
        ''' </summary>
        Public ReadOnly Property IsClosed() As Boolean
            Get
                Return _ComApp Is Nothing OrElse _ComApp.IsClosed
            End Get
        End Property

        ''' <summary>
        ''' Close MS Excel application
        ''' </summary>
        Public Sub Close()
            If _ComApp IsNot Nothing Then
                _ComApp.Close()
                _ComApp.Dispose()
                _ComApp = Nothing
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return NameOf(MsExcelApplicationWrapper) & " (" & _ComApp.ToString() & ")"
        End Function

#Region "Invoke methods"
        ''' <summary>
        ''' Invoke function member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="values">Arguments for the called member, remember to use System.Reflection.Missing.Value where required</param>
        ''' <returns></returns>
        Public Function InvokeFunction(Of T)(name As String, ParamArray values As Object()) As T
            Return _ComApp.InvokeFunction(Of T)(name, values)
        End Function

        ''' <summary>
        ''' Invoke method member
        ''' </summary>
        ''' <param name="name"></param>
        ''' <param name="values">Arguments for the called member, remember to use System.Reflection.Missing.Value where required</param>
        Public Sub InvokeMethod(name As String, ParamArray values As Object())
            _ComApp.InvokeMethod(name, values)
        End Sub

        ''' <summary>
        ''' Invoke property-get member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <returns></returns>
        Public Function InvokePropertyGet(Of T)(name As String) As T
            Return _ComApp.InvokePropertyGet(Of T)(name)
        End Function

        ''' <summary>
        ''' Invoke property-get member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="propertyArrayItem">Arguments for the called member, remember to use System.Reflection.Missing.Value where required</param>
        ''' <returns></returns>
        Public Function InvokePropertyGet(Of T)(name As String, propertyArrayItem As Object) As T
            Return _ComApp.InvokePropertyGet(Of T)(name, propertyArrayItem)
        End Function

        ''' <summary>
        ''' Invoke property-set member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="value"></param>
        Public Sub InvokePropertySet(Of T)(name As String, value As T)
            _ComApp.InvokePropertySet(Of T)(name, value)
        End Sub

        ''' <summary>
        ''' Invoke property-set member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="values">Arguments for the called member, remember to use System.Reflection.Missing.Value where required</param>
        Public Sub InvokePropertySet(Of T)(name As String, values As T())
            _ComApp.InvokePropertySet(Of T)(name, values)
        End Sub

        ''' <summary>
        ''' Invoke field-get member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <returns></returns>
        Public Function InvokeFieldGet(Of T)(name As String) As T
            Return _ComApp.InvokeFieldGet(Of T)(name)
        End Function

        ''' <summary>
        ''' Invoke field-set member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="value"></param>
        Public Sub InvokeFieldSet(Of T)(name As String, value As T)
            _ComApp.InvokeFieldSet(Of T)(name, value)
        End Sub

        ''' <summary>
        ''' Invoke field-set member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="values"></param>
        Public Sub InvokeFieldSet(Of T)(name As String, values As T())
            _ComApp.InvokeFieldSet(Of T)(name, values)
        End Sub
#End Region

#Region "IDisposable Support"
        Private disposedValue As Boolean ' Dient zur Erkennung redundanter Aufrufe.

        Public ReadOnly Property IsDisposed As Boolean
            Get
                Return disposedValue
            End Get
        End Property

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    Me.Close()
                End If
            End If
            disposedValue = True
        End Sub

        ' ' TODO: Finalizer nur überschreiben, wenn "Dispose(disposing As Boolean)" Code für die Freigabe nicht verwalteter Ressourcen enthält
        ' Protected Overrides Sub Finalize()
        '     ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
        '     Dispose(disposing:=False)
        '     MyBase.Finalize()
        ' End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

End Namespace