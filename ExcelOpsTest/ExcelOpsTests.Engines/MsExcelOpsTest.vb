Imports NUnit.Framework

<NonParallelizable>
Public Class MsExcelOpsTest
    Inherits ExcelOpsTestBase(Of ExcelOps.MsExcelDataOperations)

    Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String) As ExcelOps.MsExcelDataOperations
        If MsExcelInstance Is Nothing OrElse MsExcelInstance.IsDisposed Then
            'recreate excel instance
            MsExcelInstance = New CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper
        End If
        Return New ExcelOps.MsExcelDataOperations(file, mode, MsExcelInstance, False, [readOnly], passwordForOpening)
    End Function

    Protected Overrides Function _CreateInstance() As ExcelOps.MsExcelDataOperations
        If MsExcelInstance Is Nothing OrElse MsExcelInstance.IsDisposed Then
            'recreate excel instance
            MsExcelInstance = New CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper
        End If
        Return New ExcelOps.MsExcelDataOperations(Nothing, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, MsExcelInstance, False, True, Nothing)
    End Function

    <OneTimeSetUp>
    Public Sub OneTimeSetup()
        If Global.CompuMaster.Excel.MsExcelCom.MsExcelTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled = False Then
            AssertNoExcelProcessesAvailable()
        End If
    End Sub

    <SetUp>
    Public Sub Setup()
        AssertExactlyOur1ExcelProcessAvailableInProcessList()
    End Sub

    <TearDown>
    Public Sub TearDown()
        'CloseDisposeFinalizeExcelAppInstance
        If MsExcelInstance IsNot Nothing Then MsExcelInstance.Dispose()
        GC.Collect(2, GCCollectionMode.Forced)
        AssertNoExcelProcessesAvailable()
    End Sub

    <OneTimeTearDown>
    Public Sub OneTimeTearDown()
        If MsExcelInstance IsNot Nothing Then MsExcelInstance.Dispose()
        GC.Collect(2, GCCollectionMode.Forced)
        AssertNoExcelProcessesAvailable()
    End Sub

    Private Sub AssertExactlyOur1ExcelProcessAvailableInProcessList()
        If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
            Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            If MsExcelProcesses.Length <> 1 Then
                Assert.Fail("Found " & MsExcelProcesses.Length & " EXCEL processes, but 1 excel processes allowed/required")
            Else
                Assert.AreEqual(MsExcelInstance.ProcessId, MsExcelProcesses(0).Id)
            End If
        End If
    End Sub

    Private Sub AssertNoExcelProcessesAvailable()
        If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
            Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            If MsExcelProcesses.Length <> 0 Then
                Assert.Fail("Found " & MsExcelProcesses.Length & " EXCEL processes, but no excel processes allowed")
            End If
        End If
    End Sub

    Private MsExcelInstance As CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper

    <Test, Explicit("Known2Fail But Less Important")> Public Sub ManualRunOnly_KillAllMsExcelAppProcesses()
        Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        Console.WriteLine("Found " & MsExcelProcessesBefore.Length & " EXCEL processes")
        ExcelOps.MsExcelDataOperations.CheckForRunningMsExcelInstancesAndAskUserToKill()
        Assert.Pass()
    End Sub

End Class