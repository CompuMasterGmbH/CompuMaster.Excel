Imports NUnit.Framework

Namespace MsExcelComInteropTest

    Public Class Tests

        <SetUp>
        Public Sub Setup()
        End Sub

        <OneTimeTearDown>
        Public Sub TearDown()
            'CloseDisposeFinalizeExcelAppInstance
            ExcelApp.Dispose()
            GC.Collect(2, GCCollectionMode.Forced)
            AssertNoExcelProcessesAvailable()
        End Sub

        <OneTimeSetUp>
        Public Sub OneTimeSetUp()
            AssertNoExcelProcessesAvailable()
            Try
                ExcelApp = New CompuMaster.Excel.MsExcelComInterop.ExcelApplication()
            Catch
                Assert.Ignore("Platform not supported or MS Excel application not installed")
            End Try
        End Sub

        Private Sub AssertNoExcelProcessesAvailable()
            If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
                Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                If MsExcelProcesses.Length <> 0 Then
                    Assert.Fail("Found " & MsExcelProcesses.Length & " EXCEL processes, but no excel processes allowed")
                End If
            End If
        End Sub

        Private ExcelApp As CompuMaster.Excel.MsExcelComInterop.ExcelApplication

        Protected Function OpenExcelAppAndWorkbook(path As String) As CompuMaster.Excel.MsExcelComInterop.ExcelWorkbook
            Return ExcelApp.Workbooks.Open(path)
        End Function

        <Test>
        Public Sub ExportAsFixedFormat()
            Dim TargetTestFile As String = TestEnvironment.FullPathOfDynTestFile(
                                       "pdf-export",
                                       System.IO.Path.GetFileNameWithoutExtension(TestFiles.TestFileExcelOpsTestCollection.FullName) & ".pdf"
                                       )
            If System.IO.File.Exists(TargetTestFile) Then System.IO.File.Delete(TargetTestFile)
            Dim Wb = OpenExcelAppAndWorkbook(TestFiles.TestFileExcelOpsTestCollection.FullName)
            Wb.ExportAsFixedFormat(MsExcelComInterop.Enumerations.XlFixedFormatType.XlTypePDF, TargetTestFile)
            Assert.True(System.IO.File.Exists(TargetTestFile))
            Wb.Close()
        End Sub

        <Test, Explicit("Known2Fail But Less Important")> Public Sub ManualRunOnly_KillAllMsExcelAppProcesses()
            Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Console.WriteLine("Found " & MsExcelProcessesBefore.Length & " EXCEL processes")
            CheckForRunningMsExcelInstancesAndAskUserToKill()
            Assert.Pass()
        End Sub

        Private Shared Sub CheckForRunningMsExcelInstancesAndAskUserToKill()
            Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            If MsExcelProcesses IsNot Nothing AndAlso MsExcelProcesses.Length > 0 Then
                If MsgBox(MsExcelProcesses.Length & " bereits geöffenete MS Excel Instanzen wurden gefunden. Sollen diese zuvor geschlossen werden?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = vbYes Then
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
    End Class

End Namespace