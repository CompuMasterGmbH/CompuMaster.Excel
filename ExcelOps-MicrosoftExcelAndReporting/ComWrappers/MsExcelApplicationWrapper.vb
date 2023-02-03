﻿Imports MsExcel = Microsoft.Office.Interop.Excel

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
            If Not Me.IsDisposedComObject Then
                Try
                    Me.ComObjectStronglyTyped.Calculation = MsExcel.XlCalculation.xlCalculationAutomatic 'reset value from manual to automatic (=expected default setting of user in 99% of all situations)
                Catch
                End Try
            End If
            Me.ComObjectStronglyTyped.Quit()
            MyBase.OnClosing()
            SafelyCloseExcelAppInstanceInternal()
        End Sub

        Private Sub SafelyCloseExcelAppInstanceInternal()
            If ProcessId <> Nothing AndAlso Process() IsNot Nothing AndAlso Process.HasExited = False Then
                Try
                    MsExcelTools.WaitUntilTrueOrTimeout(Function() Me.Process.HasExited = True, New TimeSpan(0, 0, 15)) 'Sometimes it takes time to close MS Excel...
                    Me.Process.Refresh()
                    If Me.Process.HasExited = False Then
                        'Force kill on Excel 
                        Me.Process.Kill()
                        Try
                            MsExcelTools.WaitUntilTrueOrTimeout(Function() Me.Process.HasExited = True, New TimeSpan(0, 0, 1))
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
                                                                End Function, New TimeSpan(0, 0, 1))
                        Catch 'ex As Exception
                            'ignore any exceptions on getting process list
                        End Try
                    End If
                Catch 'ex As Exception
                    'ignore any exceptions of watching/handling process close/kill
                End Try
            End If
        End Sub

    End Class

End Namespace