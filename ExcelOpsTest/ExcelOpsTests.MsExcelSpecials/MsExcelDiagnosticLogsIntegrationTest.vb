Imports System.Diagnostics
Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.MsExcelCom
Imports NUnit.Framework

Namespace ExcelOpsTests.MsExcelSpecials
    <Platform("Win"), NonParallelizable, Apartment(Threading.ApartmentState.STA)>
    Public Class MsExcelDiagnosticLogsIntegrationTest
        <SetUp>
        Public Sub RequireExcel()
            If Not MsExcelTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled Then Assert.Ignore("Microsoft Excel is required.")
        End Sub

        <Test>
        Public Sub ConcurrentProcessesAndSharedWorkbooks()
            ' Close only our own instances; the concurrent foreign process must keep its logs.
            Using foreign As New MsExcelApplicationWrapper(), owned As New MsExcelApplicationWrapper()
                foreign.PreserveDiagnosticLogFiles = True
                Assert.That(owned.ExcelProcessId, [Is].Not.EqualTo(foreign.ExcelProcessId))
                Dim first As New MsExcelDataOperations(TestFiles.TestFileGrund01.FullName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, owned, False, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Dim second As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, owned, False, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                first.PreserveDiagnosticLogFiles = True
                Assert.That(second.PreserveDiagnosticLogFiles, [Is].True)
                Assert.That(owned.PreserveDiagnosticLogFiles, [Is].True)
                second.PreserveDiagnosticLogFiles = False
                Assert.That(first.PreserveDiagnosticLogFiles, [Is].False)
                Dim ownedFiles = WaitForLogs(owned)
                Dim foreignFiles = WaitForLogs(foreign)
                Dim ownedIdentities = ownedFiles.ToDictionary(Function(path) path, AddressOf ReadIdentity)
                Dim foreignIdentities = foreignFiles.ToDictionary(Function(path) path, AddressOf ReadIdentity)
                Assert.That(ownedFiles.Intersect(foreignFiles), [Is].Empty)
                first.Close()
                Assert.That(second.DiagnosticLogFiles.Select(Function(file) file.FullName), [Is].EquivalentTo(ownedFiles))
                second.Close()
                ' Release temporary RCWs in the test before requesting a regular application exit.
                CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
                Dim regular As Boolean
                Using process = owned.ExcelProcess()
                    Dim pinned = process.Handle
                    first.CloseExcelAppInstance()
                    TestContext.Out.WriteLine("Owned Excel exit code: {0}; remaining logs: {1}.", process.ExitCode, owned.DiagnosticLogFiles.Count)
                    regular = process.ExitCode = 0
                End Using
                ' Some Office configurations require the existing force-termination fallback.
                Assert.That(ownedFiles.All(Function(path) File.Exists(path) <> regular), [Is].True)
                AssertSurvivingPaths(owned, ownedIdentities)
                Assert.That(first.DiagnosticLogFiles.Select(Function(file) file.FullName), [Is].EquivalentTo(owned.DiagnosticLogFiles.Select(Function(file) file.FullName)))
                Using process = foreign.ExcelProcess()
                    Assert.That(process.HasExited, [Is].False)
                End Using
                Assert.That(foreignFiles.All(Function(path) File.Exists(path)), [Is].True)
                foreign.Close()
                Assert.That(foreignFiles.All(Function(path) File.Exists(path)), [Is].True)
                AssertSurvivingPaths(foreign, foreignIdentities)
                TestContext.Out.WriteLine("Concurrent sessions: {0} owned logs handled; {1} foreign logs retained; regular exit = {2}.", ownedFiles.Length, foreignFiles.Length, regular)
            End Using
        End Sub

        <Test>
        Public Sub ApplicationCloseHandlesDiagnosticLogs(<Values(True, False)> preserve As Boolean)
            Using application As New MsExcelApplicationWrapper With {.PreserveDiagnosticLogFiles = preserve}
                Dim paths = WaitForLogs(application)
                Dim identities = paths.ToDictionary(Function(path) path, AddressOf ReadIdentity)
                Dim regular As Boolean
                Using process = application.ExcelProcess()
                    Dim pinned = process.Handle
                    application.Close()
                    Assert.That(process.HasExited, [Is].True)
                    regular = process.ExitCode = 0
                End Using
                Dim retained = preserve OrElse Not regular
                Assert.That(paths.All(Function(path) File.Exists(path) = retained), [Is].True)
                AssertSurvivingPaths(application, identities)
                TestContext.Out.WriteLine("Application close: {0} logs verified; preservation = {1}; regular exit = {2}.", paths.Length, preserve, regular)
            End Using
        End Sub

        <Test>
        Public Sub ForcedExitRetainsDiagnostics(<Values(True, False)> preserve As Boolean)
            Dim application As New MsExcelApplicationWrapper With {.PreserveDiagnosticLogFiles = preserve}
            Try
                Dim paths = WaitForLogs(application)
                Using process = application.ExcelProcess()
                    ' This test intentionally terminates only the Excel process it just created.
                    process.Kill()
                    Assert.That(process.WaitForExit(10000), [Is].True)
                End Using
                Dim original = paths.ToDictionary(Function(path) path, Function(path) File.ReadAllBytes(path))
                Try
                    application.Close()
                Catch ex As Exception
                    ' The existing COM lifecycle may fail after an external process termination.
                    TestContext.Out.WriteLine("Original close failure retained: " & ex.GetType().FullName)
                End Try
                For Each path In paths
                    Dim expected = original(path)
                    If Not preserve Then
                        Dim last = Array.FindLastIndex(expected, Function(value) value <> 0)
                        expected = expected.Take(last + 1).ToArray()
                    End If
                    Assert.That(File.ReadAllBytes(path), [Is].EqualTo(expected), path)
                Next
                Assert.That(application.DiagnosticLogFiles.Select(Function(file) file.FullName), [Is].EquivalentTo(paths))
                TestContext.Out.WriteLine("Forced exit: {0} logs verified; preservation = {1}.", paths.Length, preserve)
            Finally
                Try
                    application.Dispose()
                Catch ex As Exception
                    TestContext.Out.WriteLine("Original dispose failure retained: " & ex.GetType().FullName)
                End Try
            End Try
        End Sub

        Private Shared Function ReadIdentity(path As String) As DiagnosticLogNative.FileInformation
            Using handle = DiagnosticLogNative.OpenMetadata(path)
                Return DiagnosticLogNative.Information(handle)
            End Using
        End Function

        Private Shared Sub AssertSurvivingPaths(application As MsExcelApplicationWrapper, identities As Dictionary(Of String, DiagnosticLogNative.FileInformation))
            Dim expected As New List(Of String)
            For Each entry In identities
                If File.Exists(entry.Key) Then
                    If entry.Value.SameFile(ReadIdentity(entry.Key)) Then
                        expected.Add(entry.Key)
                    Else
                        TestContext.Out.WriteLine("Office changed a file identity during exit; the replacement is conservatively omitted: " & Path.GetFileName(entry.Key))
                    End If
                End If
            Next
            Dim actual = application.DiagnosticLogFiles.Select(Function(file) file.FullName).Intersect(identities.Keys)
            Assert.That(actual, [Is].EquivalentTo(expected), "Every surviving original file remains accessible; replacements are omitted.")
        End Sub

        Private Shared Function WaitForLogs(application As MsExcelApplicationWrapper) As String()
            Dim watch = Stopwatch.StartNew()
            Do
                Dim files = application.DiagnosticLogFiles.Select(Function(file) file.FullName).ToArray()
                If files.Any(Function(path) IO.Path.GetFileName(path).StartsWith("Primary", StringComparison.Ordinal)) AndAlso files.Any(Function(path) IO.Path.GetFileName(path).StartsWith("Additional", StringComparison.Ordinal)) Then Return files
                Threading.Thread.Sleep(100)
            Loop While watch.Elapsed < TimeSpan.FromSeconds(5)
            Assert.Fail("Expected positively associated Primary and Additional logs while Excel holds its locks.")
            Return Array.Empty(Of String)()
        End Function
    End Class
End Namespace
