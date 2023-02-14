Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.MsExcelCom

Namespace ExcelOpsEngineTests
    <TestFixture(Explicit:=True, IgnoreReason:="Only for explicit calling")> Public Class MsExcelAppReleaseTest

        <OneTimeSetUp()> Public Sub OneTimeInitConfig()
        End Sub

        <SetUp> Public Sub Setup()
            Test.Console.ResetConsoleForTestOutput()
        End Sub

#If Not CI_CD Then
        <Test, Explicit("Known2Fail But Less Important")> Public Sub ManualRunOnly_KillAllMsExcelAppProcesses()
            Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Console.WriteLine("Found " & MsExcelProcessesBefore.Length & " EXCEL processes")
            MsExcelDataOperations.CheckForRunningMsExcelInstancesAndAskUserToKill()
            Assert.Pass()
        End Sub

        <Test> Public Sub OpenAnCloseMsExcelWithPropertProcessCleanup_SeparateMsExcelAppWithExplicitQuit()
            Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Assert.Zero(MsExcelProcessesBefore.Length)
            Dim DummyCTWb As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, True, Nothing)
            DummyCTWb.CloseExcelAppInstance()
            Dim MsExcelProcessesAfterExplicitQuit As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Assert.AreEqual(MsExcelProcessesBefore.Length, MsExcelProcessesAfterExplicitQuit.Length, "Process count after ExcelApp.Quit")
        End Sub

        <Test, NUnit.Framework.Ignore("Known2Fail But Less Important"), Explicit> Public Sub OpenAnCloseMsExcelWithPropertProcessCleanup_SeparateMsExcelApp()
            Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Assert.Zero(MsExcelProcessesBefore.Length)
            Dim DummyCTWb As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, True, Nothing)
#Disable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
            DummyCTWb = Nothing
#Enable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
            GC.Collect(0, GCCollectionMode.Forced, True, False)
            GC.Collect(2, GCCollectionMode.Forced, True, False)
            GC.Collect(0, GCCollectionMode.Forced, True, False)
            Dim MsExcelProcessesAfter As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Assert.AreEqual(MsExcelProcessesBefore.Length, MsExcelProcessesAfter.Length, "Process count after GC.Collect")
        End Sub

        <Test, NUnit.Framework.Ignore("Known2Fail But Less Important"), Explicit> Public Sub OpenAnCloseMsExcelWithPropertProcessCleanup_ReusedMsExcelApp()
            Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Assert.Zero(MsExcelProcessesBefore.Length)
            Dim MsExcelApp As New MsExcelApplicationWrapper()
            Dim DummyCTWb As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelApp, False, True, Nothing)
            Dim DummyCTWb2 As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelApp, False, True, Nothing)
#Disable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
            DummyCTWb = Nothing
            DummyCTWb2 = Nothing
            MsExcelApp = Nothing
#Enable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
            GC.Collect(0, GCCollectionMode.Forced, True, False)
            GC.Collect(2, GCCollectionMode.Forced, True, False)
            GC.Collect(0, GCCollectionMode.Forced, True, False)
            Dim MsExcelProcessesAfter As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Assert.AreEqual(MsExcelProcessesBefore.Length, MsExcelProcessesAfter.Length, "Process count after GC.Collect")
        End Sub
#End If

    End Class
End Namespace