Imports CompuMaster.Excel.MsExcelCom

Namespace Global.CompuMaster.Excel.MsExcelCom

    ''' <summary>
    ''' Tools for Microsoft Excel automation
    ''' </summary>
    ''' <remarks>
    ''' PLEASE NOTE: Considerations for server-side Automation of Office https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2
    ''' </remarks>
    Public NotInheritable Class MsExcelTools

        ''' <summary>
        ''' Are there any running MS Excel instances on the current system (owned by any user)
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function HasRunningMsExcelInstances() As Boolean
            Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Return MsExcelProcesses IsNot Nothing AndAlso MsExcelProcesses.Length > 0
        End Function

        Public Shared Sub RecalculateFile(filePath As String)
            RecalculateFile(filePath, Nothing)
        End Sub

        Public Shared Sub RecalculateFile(filePath As String, msAppInstance As MsExcelApplicationWrapper)
            RecalculateFile(filePath, msAppInstance, Nothing)
        End Sub

        Public Shared Sub RecalculateFile(filePath As String, msAppInstance As MsExcelApplicationWrapper, passwordForOpening As String)
            Dim MsExcelApp As MsExcelApplicationWrapper = msAppInstance
            If MsExcelApp Is Nothing Then
                MsExcelApp = New MsExcelApplicationWrapper()
            End If
            Dim wb As New ExcelOps.MsExcelDataOperations(filePath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelApp, False, False, passwordForOpening)
            Try
                wb.RecalculateAll()
                wb.Save()
            Finally
                wb.Close()
                If msAppInstance Is Nothing Then wb.CloseExcelAppInstance()
            End Try
        End Sub

        Public Shared Function IsPlatformSupportingComInteropAndMsExcelAppInstalled() As Boolean
            Return CompuMaster.ComInterop.ComTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled("Excel.Application")
        End Function

    End Class

End Namespace