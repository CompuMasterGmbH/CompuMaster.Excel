Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.ExcelOps

    Public Class MsExcelApplicationWrapper
        Inherits CompuMaster.ComInterop.ComRootObject(Of MsExcel.Application)

        Public Sub New()
            MyBase.New(New MsExcel.Application(), Nothing)
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
        End Sub

        Private Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer

        Public ReadOnly Property ProcessId As Integer

        Public Function Process() As System.Diagnostics.Process
            Return System.Diagnostics.Process.GetProcessById(Me.ProcessId)
        End Function

        Protected Overrides Sub OnClosing()
            Me.ComObjectStronglyTyped.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic 'reset value from manual to automatic (=expected default setting of user in 99% of all situations)
            Me.ComObjectStronglyTyped.Quit()
            MyBase.OnClosing()
            SafelyCloseExcelAppInstanceInternal()
        End Sub

        Private Sub SafelyCloseExcelAppInstanceInternal()
            If ProcessId <> Nothing AndAlso Process() IsNot Nothing AndAlso Process.HasExited = False Then
                MsExcelTools.WaitUntilTrueOrTimeout(Function() Me.Process.HasExited = True, New TimeSpan(0, 0, 15)) 'Sometimes it takes time to close MS Excel...
                If Me.Process.HasExited = False Then
                    'Force kill on Excel 
                    Me.Process.Kill()
                    MsExcelTools.WaitUntilTrueOrTimeout(Function() Me.Process.HasExited = True, New TimeSpan(0, 0, 1))
                End If
            End If
        End Sub

    End Class

End Namespace