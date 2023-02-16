Imports CompuMaster.Excel.MsExcelCom

Namespace Global.CompuMaster.Excel.MsExcelCom

    Public NotInheritable Class MsExcelTools

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
        Friend Shared Function CreateMsExcelAppInstance() As MsExcelApplicationWrapper
            Try
                Return New MsExcelApplicationWrapper
            Catch ex As Exception
                Throw New PlatformNotSupportedException("App and installed MS Office must both 64 bit or both 32 bit processed")
            End Try
        End Function

        Friend Delegate Function WaitUntilTrueConditionTest() As Boolean

        Friend Shared Function WaitUntilTrueOrTimeout(test As WaitUntilTrueConditionTest, maxTimeout As TimeSpan) As Boolean
            Dim Start As DateTime = Now
            Do While Now.Subtract(Start) < maxTimeout
                If test() = True Then
                    Return True
                End If
                If maxTimeout.TotalDays > 0 OrElse maxTimeout.Hours > 0 Then 'prevent exceeded range when calling maxTimeout.TotalMilliseconds
                    'Check at least twice per second
                    System.Threading.Thread.Sleep(500)
                Else
                    'Check at least 10 times and minimum twice per second
                    System.Threading.Thread.Sleep(System.Math.Min(CType(maxTimeout.TotalMilliseconds / 10, Integer), 500))
                End If
            Loop
            Return False
        End Function

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

        Public Shared Function IsPlatformSupportingComInterop() As Boolean
            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Public Shared Function IsPlatformSupportingComInteropAndMsExcelAppInstalled() As Boolean
            If IsPlatformSupportingComInterop() = False Then
                Return False
            Else
                'Windows platform ok - MS Excel installed?
                Try
#Disable Warning CA1416
                    Dim MsExcelType As Type = Type.GetTypeFromProgID("Excel.Application")
#Enable Warning
                    Return MsExcelType IsNot Nothing
                Catch
                    Return False
                End Try
            End If
        End Function

    End Class

End Namespace