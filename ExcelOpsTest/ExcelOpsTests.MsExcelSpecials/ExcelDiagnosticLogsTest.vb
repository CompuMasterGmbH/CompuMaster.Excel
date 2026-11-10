Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.MsExcelCom
Imports NUnit.Framework

Namespace ExcelOpsTests.MsExcelSpecials
    Public Class ExcelDiagnosticLogsTest
        <Test>
        Public Sub PaddingPreservesEveryNonPaddingByte(<Values(0, 1, 8191, 8192, 8193, 1048576, 20971520)> padding As Integer)
            ' Internal NULs, BOM, whitespace, and the final newline must survive exactly.
            Dim original = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes("header" & ChrW(0) & "body  " & vbCrLf)).ToArray()
            AssertTrimmed(original, padding)
        End Sub

        <Test>
        Public Sub EmptyAndAllNulLogs(<Values(0, 1, 8192, 8193, 1048576)> size As Integer)
            AssertTrimmed(Array.Empty(Of Byte)(), size)
        End Sub

        <Test>
        Public Sub UnicodeTerminalCodeUnits(<Values("utf16le", "utf16be", "utf32le", "utf32be")> encodingName As String, <Values(True, False)> bom As Boolean)
            Dim encoding As Encoding
            Select Case encodingName
                Case "utf16le" : encoding = New UnicodeEncoding(False, bom)
                Case "utf16be" : encoding = New UnicodeEncoding(True, bom)
                Case "utf32le" : encoding = New UTF32Encoding(False, bom)
                Case Else : encoding = New UTF32Encoding(True, bom)
            End Select
            Dim original = encoding.GetPreamble().Concat(encoding.GetBytes("Timestamp" & vbTab & "Unicode ü😀 " & vbCrLf)).ToArray()
            AssertTrimmed(original, 8193)
            ' FF FE 00 00 is also a UTF-32 BOM: ambiguous BOM-only logs retain four bytes.
            If bom AndAlso encodingName <> "utf16le" Then AssertTrimmed(encoding.GetPreamble(), 8193)
        End Sub

        Private Shared Sub AssertTrimmed(original As Byte(), padding As Integer)
            Using stream As New MemoryStream()
                stream.Write(original, 0, original.Length)
                stream.Write(New Byte(padding - 1) {}, 0, padding)
                ExcelDiagnosticLogs.TrimPadding(stream)
                Assert.That(stream.ToArray(), [Is].EqualTo(original))
            End Using
        End Sub

        <Test>
        Public Sub AmbiguousBomlessUnicodeCannotLoseTerminalBytes(<Values("A", "你好A", "ü", "😀", "ഊA")> content As String, <Values(2, 4)> width As Integer, <Values(True, False)> bigEndian As Boolean)
            Dim encoding As Encoding = If(width = 2, DirectCast(New UnicodeEncoding(bigEndian, False), Encoding), New UTF32Encoding(bigEndian, False))
            Dim original = encoding.GetBytes(content)
            Using stream As New MemoryStream()
                stream.Write(original, 0, original.Length)
                stream.Write(New Byte(10000) {}, 0, 10001)
                ExcelDiagnosticLogs.TrimPadding(stream)
                Assert.That(stream.ToArray().Take(original.Length), [Is].EqualTo(original))
                Assert.That(stream.Length, [Is].InRange(CLng(original.Length), CLng(original.Length + 3)))
            End Using
        End Sub

        <Test>
        Public Sub UninitializedApiDoesNotStartExcel()
            Dim workbook As New MsExcelDataOperations(ExcelDataOperationsBase.OpenMode.Uninitialized)
            Assert.That(workbook.DiagnosticLogFiles, [Is].Empty)
            Assert.That(workbook.PreserveDiagnosticLogFiles, [Is].False)
            workbook.PreserveDiagnosticLogFiles = True
            Assert.That(workbook.PreserveDiagnosticLogFiles, [Is].True)
            Assert.That(workbook.DiagnosticLogFiles, [Is].Empty)
        End Sub

        <Test>
        Public Sub AssociationExcludesForeignOldAndUnknownSessions()
            If Environment.OSVersion.Platform <> PlatformID.Win32NT Then Assert.Ignore("Windows file identities required.")
            WithDirectory(Sub(root)
                              Using process = System.Diagnostics.Process.GetCurrentProcess()
                                  Dim logs As New ExcelDiagnosticLogs(process, root)
                                  Dim primary = WritePrimary(root, process.Id)
                                  Dim companion = WriteCompanion(root, primary, {0, 0, 0})
                                  Dim foreign = WritePrimary(root, process.Id + 1)
                                  WriteCompanion(root, foreign, {0})
                                  Dim old = WritePrimary(root, process.Id)
                                  File.SetCreationTimeUtc(old, process.StartTime.ToUniversalTime().AddDays(-1))
                                  WriteCompanion(root, old, {0})
                                  Dim unknown = Path.Combine(root, "Additional", "Additional123_" & Guid.NewGuid().ToString() & ".log")
                                  File.WriteAllBytes(unknown, {0})
                                  Dim snapshot = logs.Snapshot()
                                  Assert.That(snapshot.Select(Function(file) file.FullName), [Is].EquivalentTo({primary, companion}))
                                  Assert.Throws(Of NotSupportedException)(Sub() DirectCast(snapshot, IList(Of FileInfo)).Clear())
                                  logs.BeforeQuit()
                                  logs.AfterQuit()
                                  logs.Finish(True)
                                  Assert.That(File.Exists(primary), [Is].True, "A live process must never lose logs.")
                                  File.Delete(companion)
                                  Assert.That(snapshot.Count, [Is].EqualTo(2), "Snapshots do not change retrospectively.")
                                  Assert.That(logs.Snapshot().Count, [Is].EqualTo(1))
                              End Using
                          End Sub)
        End Sub

        <Test>
        Public Sub LockedPrimaryUsesOnlyExactRestartManagerOwner()
            If Environment.OSVersion.Platform <> PlatformID.Win32NT Then Assert.Ignore("Windows Restart Manager required.")
            WithDirectory(Sub(root)
                              Using process = System.Diagnostics.Process.GetCurrentProcess()
                                  Dim primary = WritePrimary(root, process.Id)
                                  WriteCompanion(root, primary, {0})
                                  Using locked As New FileStream(primary, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                                      Dim logs As New ExcelDiagnosticLogs(process, root)
                                      Assert.That(logs.Snapshot().Count, [Is].EqualTo(2))
                                      Assert.That(DiagnosticLogNative.IsOnlyOwner(primary, process.Id, process.StartTime.ToUniversalTime()), [Is].True)
                                      Assert.That(DiagnosticLogNative.IsOnlyOwner(primary, process.Id, process.StartTime.ToUniversalTime().AddSeconds(-1)), [Is].False)
                                      Assert.That(DiagnosticLogNative.IsOnlyOwner(primary, process.Id + 1, process.StartTime.ToUniversalTime()), [Is].False)
                                  End Using
                              End Using
                          End Sub)
        End Sub

        <Test>
        Public Sub NormalErrorAndPreservedExit(<Values(True, False)> regular As Boolean, <Values(True, False)> preserve As Boolean)
            If Environment.OSVersion.Platform <> PlatformID.Win32NT Then Assert.Ignore("Windows file identities required.")
            WithDirectory(Sub(root)
                              Using process = StartShortProcess()
                                  Dim logs As New ExcelDiagnosticLogs(process, root) With {.Preserve = preserve}
                                  Dim primary = WritePrimary(root, process.Id)
                                  Dim body As Byte() = {65, 0, 66, 32, 32, 32, 13, 10}
                                  Dim companion = WriteCompanion(root, primary, body.Concat(New Byte(8192) {}).ToArray())
                                  Dim before = File.ReadAllBytes(primary)
                                  Assert.That(logs.Snapshot().Count, [Is].EqualTo(2))
                                  logs.BeforeQuit()
                                  If regular Then logs.AfterQuit()
                                  Assert.That(process.WaitForExit(10000), [Is].True)
                                  logs.Finish(regular)
                                  If preserve Then
                                      Assert.That(File.ReadAllBytes(primary), [Is].EqualTo(before))
                                      Assert.That(New FileInfo(companion).Length, [Is].EqualTo(body.Length + 8193))
                                  ElseIf regular Then
                                      Assert.That(File.Exists(primary), [Is].False)
                                      Assert.That(File.Exists(companion), [Is].False)
                                  Else
                                      Assert.That(File.ReadAllBytes(companion), [Is].EqualTo(body))
                                      Assert.That(logs.Snapshot().Count, [Is].EqualTo(2), "Surviving paths remain available after exit.")
                                  End If
                              End Using
                          End Sub)
        End Sub

        <Test, Platform("Win")>
        Public Sub SameProcessSharesRetentionState()
            Using firstProcess = Process.GetCurrentProcess(), secondProcess = Process.GetCurrentProcess()
                Dim first = ExcelDiagnosticLogs.ForProcess(firstProcess)
                Dim second = ExcelDiagnosticLogs.ForProcess(secondProcess)
                Assert.That(second, [Is].SameAs(first))
                Try
                    first.Preserve = True
                    Assert.That(second.Preserve, [Is].True)
                    second.Preserve = False
                    Assert.That(first.Preserve, [Is].False)
                Finally
                    first.Preserve = False
                End Try
            End Using
        End Sub

        <Test>
        Public Sub UnconfirmedExitCannotMaskOriginalException()
            If Environment.OSVersion.Platform <> PlatformID.Win32NT Then Assert.Ignore("Windows file identities required.")
            WithDirectory(Sub(root)
                              Dim process = System.Diagnostics.Process.GetCurrentProcess()
                              Dim logs As New ExcelDiagnosticLogs(process, root)
                              Dim primary = WritePrimary(root, process.Id)
                              Dim before = File.ReadAllBytes(primary)
                              Assert.That(logs.Snapshot().Count, [Is].EqualTo(1))
                              process.Dispose()
                              Dim original As New InvalidOperationException("Original operation failed.")
                              Dim caught = Assert.Throws(Of InvalidOperationException)(Sub()
                                                                                          Try
                                                                                              Throw original
                                                                                          Finally
                                                                                              logs.Finish(False)
                                                                                          End Try
                                                                                      End Sub)
                              Assert.That(caught, [Is].SameAs(original))
                              Assert.That(File.ReadAllBytes(primary), [Is].EqualTo(before))
                          End Sub)
        End Sub

        <Test>
        Public Sub LockedLogsAreNotChangedAfterExit()
            If Environment.OSVersion.Platform <> PlatformID.Win32NT Then Assert.Ignore("Windows file identities required.")
            WithDirectory(Sub(root)
                              Using process = StartShortProcess()
                                  Dim logs As New ExcelDiagnosticLogs(process, root)
                                  Dim primary = WritePrimary(root, process.Id)
                                  Dim companion = WriteCompanion(root, primary, {65, 0, 0, 0})
                                  Assert.That(logs.Snapshot().Count, [Is].EqualTo(2))
                                  logs.BeforeQuit()
                                  logs.AfterQuit()
                                  Using locked As New FileStream(companion, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                                      Assert.That(process.WaitForExit(10000), [Is].True)
                                      logs.Finish(True)
                                      Assert.That(locked.Length, [Is].EqualTo(4))
                                  End Using
                                  Assert.That(File.Exists(primary), [Is].False)
                                  Assert.That(File.ReadAllBytes(companion), [Is].EqualTo(New Byte() {65, 0, 0, 0}))
                              End Using
                          End Sub)
        End Sub

        <Test>
        Public Sub ReparseDirectoriesAreNotTraversed()
            If Environment.OSVersion.Platform <> PlatformID.Win32NT Then Assert.Ignore("Windows directory junctions required.")
            WithDirectory(Sub(root)
                              Dim additional = Path.Combine(root, "Additional")
                              Dim target = Path.Combine(root, "Foreign")
                              Directory.Delete(additional)
                              Directory.CreateDirectory(target)
                              Using command = Process.Start(New ProcessStartInfo("cmd.exe", "/c mklink /J """ & additional & """ """ & target & """") With {.UseShellExecute = False, .CreateNoWindow = True, .RedirectStandardOutput = True})
                                  Assert.That(command.WaitForExit(5000), [Is].True)
                                  Assert.That(command.ExitCode, [Is].Zero)
                              End Using
                              Try
                                  Using process = StartShortProcess()
                                      Dim logs As New ExcelDiagnosticLogs(process, root)
                                      Dim primary = WritePrimary(root, process.Id)
                                      Dim foreign = Path.Combine(target, Path.GetFileName(primary).Replace("Primary", "Additional"))
                                      File.WriteAllBytes(foreign, {65, 0, 0, 0})
                                      Assert.That(logs.Snapshot().Select(Function(file) file.FullName), [Is].EquivalentTo({primary}))
                                      logs.BeforeQuit()
                                      logs.AfterQuit()
                                      Assert.That(process.WaitForExit(10000), [Is].True)
                                      logs.Finish(True)
                                      Assert.That(File.ReadAllBytes(foreign), [Is].EqualTo(New Byte() {65, 0, 0, 0}))
                                  End Using
                              Finally
                                  Directory.Delete(additional)
                              End Try
                          End Sub)
        End Sub

        <Test>
        Public Sub ReplacedAndReadOnlyFilesAreNotChanged()
            If Environment.OSVersion.Platform <> PlatformID.Win32NT Then Assert.Ignore("Windows file identities required.")
            WithDirectory(Sub(root)
                              Using process = StartShortProcess()
                                  Dim logs As New ExcelDiagnosticLogs(process, root)
                                  Dim primary = WritePrimary(root, process.Id)
                                  Dim companion = WriteCompanion(root, primary, {65, 0, 0})
                                  Assert.That(logs.Snapshot().Count, [Is].EqualTo(2))
                                  File.Move(companion, companion & ".original")
                                  File.WriteAllBytes(companion, {66, 0, 0})
                                  File.SetAttributes(primary, FileAttributes.ReadOnly)
                                  Try
                                      Assert.That(process.WaitForExit(10000), [Is].True)
                                      logs.Finish(True)
                                      Assert.That(File.ReadAllBytes(companion), [Is].EqualTo(New Byte() {66, 0, 0}))
                                      Assert.That(File.Exists(primary), [Is].True)
                                      Assert.That(logs.Snapshot().Select(Function(file) file.FullName), [Is].EquivalentTo({primary}))
                                  Finally
                                      File.SetAttributes(primary, FileAttributes.Normal)
                                  End Try
                              End Using
                          End Sub)
        End Sub

        Private Shared Function StartShortProcess() As Process
            Dim result = Process.Start(New ProcessStartInfo("cmd.exe", "/c ping -n 3 127.0.0.1 > nul") With {.UseShellExecute = False, .CreateNoWindow = True})
            ' Allow a clock tick before creating files and millisecond log timestamps on Windows runners.
            ' Otherwise a rounded record timestamp can precede the more precise process start time.
            Threading.Thread.Sleep(50)
            Return result
        End Function

        Private Shared Function WritePrimary(root As String, pid As Integer) As String
            Dim path = IO.Path.Combine(root, "Primary123_" & Guid.NewGuid().ToString() & ".log")
            Dim record = DateTime.UtcNow.ToString("MM/dd/yyyy HH:mm:ss.fff", CultureInfo.InvariantCulture) & vbTab & "EXCEL (0x" & pid.ToString("X", CultureInfo.InvariantCulture) & ")" & vbTab & "message " & vbCrLf
            File.WriteAllBytes(path, Encoding.UTF8.GetBytes(record).Concat(New Byte(100) {}).ToArray())
            Return path
        End Function

        Private Shared Function WriteCompanion(root As String, primary As String, bytes As Byte()) As String
            Dim path = IO.Path.Combine(root, "Additional", IO.Path.GetFileName(primary).Replace("Primary", "Additional"))
            File.WriteAllBytes(path, bytes)
            Return path
        End Function

        Private Shared Sub WithDirectory(action As Action(Of String))
            Dim root = Path.Combine(Path.GetTempPath(), "ExcelDiagnosticLogsTest-" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(Path.Combine(root, "Additional"))
            Try
                action(root)
            Finally
                Directory.Delete(root, True)
            End Try
        End Sub
    End Class
End Namespace
