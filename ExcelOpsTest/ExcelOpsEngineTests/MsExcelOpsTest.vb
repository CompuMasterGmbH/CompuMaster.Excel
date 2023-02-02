Imports NUnit.Framework

Public Class MsExcelOpsTest
    Inherits ExcelOpsTestBase(Of ExcelOps.MsExcelDataOperations)

    Protected Overrides Function CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String) As ExcelOps.MsExcelDataOperations
        Return New ExcelOps.MsExcelDataOperations(file, mode, False, [readOnly], passwordForOpening)
    End Function

    Protected Overrides Function CreateInstance() As ExcelOps.MsExcelDataOperations
        Return New ExcelOps.MsExcelDataOperations()
    End Function

    <SetUp>
    Public Sub AssertNoExcelProcessesAvailable()
        If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
            Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            If MsExcelProcesses.Length <> 0 Then
                Assert.Fail("Found " & MsExcelProcesses.Length & " EXCEL processes, but no excel processes allowed")
            End If
        End If
    End Sub

    <Test, Explicit("Known2Fail But Less Important")> Public Sub ManualRunOnly_KillAllMsExcelAppProcesses()
        Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        Console.WriteLine("Found " & MsExcelProcessesBefore.Length & " EXCEL processes")
        ExcelOps.MsExcelDataOperations.CheckForRunningMsExcelInstancesAndAskUserToKill()
        Assert.Pass()
    End Sub

End Class