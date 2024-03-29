﻿Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.MsExcelCom

Namespace ExcelOpsTests.MsExcelSpecials

    <NonParallelizable>
    Public Class MsExcelAppReleaseTest

        <OneTimeSetUp()> Public Sub OneTimeInitConfig()
            If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
                If MsExcelTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled = False Then
                    Assert.Ignore("Platform not supported or MS Excel not installed")
                End If
            End If
        End Sub

        <SetUp> Public Sub Setup()
            Test.Console.ResetConsoleForTestOutput()
            If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
                Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                Assert.Zero(MsExcelProcessesBefore.Length, "There are already Excel processes before test started")
            End If
        End Sub

        <TearDown> Public Sub TearDown()
            If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
                CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
                Dim MsExcelProcessesAfter As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                Assert.Zero(MsExcelProcessesAfter.Length, "There are Excel processes after test completed (TearDown)")
            End If
        End Sub

        <OneTimeTearDown> Public Sub OneTimeTearDown()
            If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
                Dim MsExcelProcessesAfter As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                Assert.Zero(MsExcelProcessesAfter.Length, "There are Excel processes after test completed (OneTimeTearDown)")
            End If
        End Sub

        <Test, Explicit("Known2Fail But Less Important")> Public Sub ManualRunOnly_KillAllMsExcelAppProcesses()
            Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Console.WriteLine("Found " & MsExcelProcessesBefore.Length & " EXCEL processes")
            MsExcelDataOperations.CheckForRunningMsExcelInstancesAndAskUserToKill()
            Assert.Pass()
        End Sub

        <Test> Public Sub OpenAnCloseMsExcelWithProperProcessCleanup_SeparateMsExcelAppWithExplicitQuit()
            Dim DummyCTWb As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, True, Nothing)
            DummyCTWb.CloseExcelAppInstance()
            Dim MsExcelProcessesAfterExplicitQuit As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Assert.AreEqual(MsExcelProcessesAfterExplicitQuit.Length, MsExcelProcessesAfterExplicitQuit.Length, "Process count after ExcelApp.Quit")
        End Sub

        '<NUnit.Framework.Ignore("Known2Fail But Less Important"), Explicit>
        <Test> Public Sub OpenAnCloseMsExcelWithPropertProcessCleanup_SeparateMsExcelApp(<Values(True, False)> explicitlyCloseMsExcelAppInstance As Boolean)
            Dim Dummy = Sub()
                            Dim DummyCTWb As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, True, Nothing)
#Disable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
                            If explicitlyCloseMsExcelAppInstance Then DummyCTWb.CloseExcelAppInstance()
                            DummyCTWb = Nothing
#Enable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
                        End Sub
            Dummy()
            CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
            'If System.Diagnostics.Process.GetProcessesByName("EXCEL").Length <> 0 Then System.Console.WriteLine(Found)
            'MsExcelTools.WaitUntilTrueOrTimeout(Function() System.Diagnostics.Process.GetProcessesByName("EXCEL").Length = 0, New TimeSpan(0, 0, 15))
            Assert.AreEqual(0, System.Diagnostics.Process.GetProcessesByName("EXCEL").Length, "Process count after GC.Collect")
        End Sub

        '<NUnit.Framework.Ignore("Known2Fail But Less Important"), Explicit>
        <Test> Public Sub OpenAnCloseMsExcelWithPropertProcessCleanup_ReusedMsExcelApp(<Values(True, False)> explicitlyCloseMsExcelAppInstance As Boolean)
            Dim Dummy = Sub()
                            Dim MsExcelApp As New MsExcelApplicationWrapper()
                            Dim DummyCTWb As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelApp, False, True, Nothing)
                            Dim DummyCTWb2 As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelApp, False, True, Nothing)
#Disable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
                            DummyCTWb = Nothing
                            DummyCTWb2 = Nothing
                            If explicitlyCloseMsExcelAppInstance Then MsExcelApp.Dispose()
                            MsExcelApp = Nothing
#Enable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
                        End Sub
            Dummy()
            CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
            Assert.AreEqual(0, System.Diagnostics.Process.GetProcessesByName("EXCEL").Length, "Process count after GC.Collect")
        End Sub

    End Class

End Namespace