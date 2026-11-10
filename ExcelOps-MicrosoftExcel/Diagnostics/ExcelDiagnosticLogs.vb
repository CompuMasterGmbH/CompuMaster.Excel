Imports System.Collections.ObjectModel
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Global.CompuMaster.Excel.MsExcelCom

    ' State belongs to the exact process, not a workbook or a directory difference.
    Friend NotInheritable Class ExcelDiagnosticLogs
        Private Shared ReadOnly Registry As New Dictionary(Of String, WeakReference)
        Private ReadOnly Gate As New Object
        Private ReadOnly OwnedProcess As Process
        Private ReadOnly Pid As Integer
        Private ReadOnly Started As DateTime
        Private ReadOnly Root As String
        Private ReadOnly Files As New Dictionary(Of String, DiagnosticLogNative.FileInformation)(StringComparer.OrdinalIgnoreCase)
        Private Keep As Boolean
        Private QuitSucceeded As Boolean
        Private ReadOnly SessionPattern As New Regex("^(?:Primary|Additional)[0-9]+_(?<id>[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12})\.log$", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

        Friend Sub New(process As Process, rootDirectory As String)
            OwnedProcess = process
            Pid = process.Id
            Started = process.StartTime.ToUniversalTime()
            ' Pin the process handle before it can exit or its PID can be reused.
            Dim pinnedHandle = process.Handle
            Root = Path.GetFullPath(rootDirectory)
        End Sub

        Friend Shared Function ForProcess(process As Process) As ExcelDiagnosticLogs
            Try
                Dim key = process.Id.ToString(CultureInfo.InvariantCulture) & ":" & process.StartTime.ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture)
                SyncLock Registry
                    Dim reference As WeakReference = Nothing
                    If Registry.TryGetValue(key, reference) Then
                        Dim existing = TryCast(reference.Target, ExcelDiagnosticLogs)
                        If existing IsNot Nothing Then Return existing
                    End If
                    For Each dead In Registry.Where(Function(entry) Not entry.Value.IsAlive).Select(Function(entry) entry.Key).ToArray()
                        Registry.Remove(dead)
                    Next
                    Dim result = New ExcelDiagnosticLogs(process, Path.Combine(Path.GetTempPath(), "Diagnostics", "EXCEL"))
                    Registry(key) = New WeakReference(result)
                    Return result
                End SyncLock
            Catch
                Return Nothing
            End Try
        End Function

        Friend Property Preserve As Boolean
            Get
                SyncLock Gate
                    Return Keep
                End SyncLock
            End Get
            Set(value As Boolean)
                SyncLock Gate
                    Keep = value
                End SyncLock
            End Set
        End Property

        Friend Function Snapshot() As ReadOnlyCollection(Of FileInfo)
            SyncLock Gate
                Discover()
                Dim result As New List(Of FileInfo)
                For Each entry In Files
                    Try
                        If Not SafePath(entry.Key) Then Continue For
                        Using handle = DiagnosticLogNative.OpenMetadata(entry.Key)
                            If entry.Value.SameFile(DiagnosticLogNative.Information(handle)) Then result.Add(New FileInfo(entry.Key))
                        End Using
                    Catch
                        ' Missing, inaccessible, or replaced files are not part of the snapshot.
                    End Try
                Next
                Return result.AsReadOnly()
            End SyncLock
        End Function

        Friend Sub BeforeQuit()
            SyncLock Gate
                QuitSucceeded = False
                Discover()
            End SyncLock
        End Sub

        Friend Sub AfterQuit()
            SyncLock Gate
                QuitSucceeded = True
            End SyncLock
        End Sub

        Friend Sub Finish(lifecycleSucceeded As Boolean)
            ' The existing Quit/release/wait/kill sequence has already finished.
            Try
                SyncLock Gate
                    For attempt = 0 To 2
                        If OwnedProcess.HasExited Then
                            Discover()
                            If Keep Then Return
                            Dim regular = lifecycleSucceeded AndAlso QuitSucceeded AndAlso OwnedProcess.ExitCode = 0
                            For Each entry In Files.ToArray()
                                Clean(entry.Key, entry.Value, regular)
                            Next
                        End If
                        If attempt < 2 Then Threading.Thread.Sleep(350)
                    Next
                End SyncLock
            Catch
                ' Housekeeping must never replace an exception from the original lifecycle.
            End Try
        End Sub

        Private Function SafePath(file As String) As Boolean
            Dim parent = Path.GetDirectoryName(file)
            If Not String.Equals(parent, Root, StringComparison.OrdinalIgnoreCase) AndAlso Not String.Equals(parent, Path.Combine(Root, "Additional"), StringComparison.OrdinalIgnoreCase) Then Return False
            Dim directory As New DirectoryInfo(parent)
            While directory IsNot Nothing
                If Not directory.Exists OrElse (directory.Attributes And FileAttributes.ReparsePoint) <> 0 Then Return False
                directory = directory.Parent
            End While
            Return True
        End Function

        Private Function ValidTimes(info As DiagnosticLogNative.FileInformation) As Boolean
            Dim created = DateTime.FromFileTimeUtc(info.Creation)
            Dim written = DateTime.FromFileTimeUtc(info.Write)
            Dim latest = If(OwnedProcess.HasExited, OwnedProcess.ExitTime.ToUniversalTime(), DateTime.UtcNow.AddSeconds(1))
            Return created >= Started AndAlso created <= latest AndAlso written >= Started AndAlso written <= latest
        End Function

        Private Sub Discover()
            Try
                If Not Directory.Exists(Root) Then Return
                Dim sessions As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim candidates As New List(Of String)
                For Each directory In {Root, Path.Combine(Root, "Additional")}
                    If SafePath(Path.Combine(directory, "probe.log")) Then candidates.AddRange(IO.Directory.GetFiles(directory, "*.log", SearchOption.TopDirectoryOnly))
                Next
                ' A Primary must prove ownership before any same-session companions qualify.
                For Each file In candidates.Where(Function(value) Path.GetFileName(value).StartsWith("Primary", StringComparison.OrdinalIgnoreCase) AndAlso String.Equals(Path.GetDirectoryName(value), Root, StringComparison.OrdinalIgnoreCase))
                    Try
                        Dim session = SessionPattern.Match(Path.GetFileName(file))
                        If Not session.Success OrElse Not SafePath(file) Then Continue For
                        Dim info As DiagnosticLogNative.FileInformation
                        Using handle = DiagnosticLogNative.OpenMetadata(file)
                            info = DiagnosticLogNative.Information(handle)
                        End Using
                        If Not ValidTimes(info) Then Continue For
                        Dim previous As DiagnosticLogNative.FileInformation
                        Dim known = Files.TryGetValue(file, previous)
                        If known AndAlso Not previous.SameFile(info) Then Continue For
                        Dim proven = known
                        If Not proven Then
                            Dim proof = PrimaryProvesOwnership(file, info)
                            If proof.HasValue Then
                                proven = proof.Value
                            ElseIf Not OwnedProcess.HasExited Then
                                proven = DiagnosticLogNative.IsOnlyOwner(file, Pid, Started)
                            End If
                        End If
                        If Not proven Then Continue For
                        Using current = DiagnosticLogNative.OpenMetadata(file)
                            If Not info.SameFile(DiagnosticLogNative.Information(current)) Then Continue For
                        End Using
                        Files(file) = info
                        sessions.Add(session.Groups("id").Value)
                    Catch
                    End Try
                Next
                For Each file In candidates
                    Try
                        Dim session = SessionPattern.Match(Path.GetFileName(file))
                        If Not session.Success OrElse Not sessions.Contains(session.Groups("id").Value) OrElse Not Path.GetFileName(file).StartsWith("Additional", StringComparison.OrdinalIgnoreCase) OrElse Not SafePath(file) Then Continue For
                        Using handle = DiagnosticLogNative.OpenMetadata(file)
                            Dim info = DiagnosticLogNative.Information(handle)
                            If Not ValidTimes(info) Then Continue For
                            Dim previous As DiagnosticLogNative.FileInformation
                            If Files.TryGetValue(file, previous) AndAlso Not previous.SameFile(info) Then Continue For
                            ' The GUID and process-era timestamps associate even empty companions.
                            Files(file) = info
                        End Using
                    Catch
                    End Try
                Next
            Catch
                ' Unavailable directories and unknown formats deliberately produce no new ownership.
            End Try
        End Sub

        Private Function PrimaryProvesOwnership(file As String, expected As DiagnosticLogNative.FileInformation) As Boolean?
            Try
                Using handle = DiagnosticLogNative.OpenExclusive(file, False)
                    Dim info = DiagnosticLogNative.Information(handle)
                    If Not expected.SameFile(info) OrElse Not ValidTimes(info) Then Return False
                    Using stream As New FileStream(handle, FileAccess.Read)
                        Dim bytes(65535) As Byte
                        Dim count = stream.Read(bytes, 0, bytes.Length)
                        Dim encoding As Encoding = Encoding.UTF8
                        If count >= 4 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE AndAlso bytes(2) = 0 AndAlso bytes(3) = 0 Then
                            encoding = New UTF32Encoding(False, False)
                        ElseIf count >= 4 AndAlso bytes(0) = 0 AndAlso bytes(1) = 0 AndAlso bytes(2) = &HFE AndAlso bytes(3) = &HFF Then
                            encoding = New UTF32Encoding(True, False)
                        ElseIf count >= 2 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE Then
                            encoding = Encoding.Unicode
                        ElseIf count >= 2 AndAlso bytes(0) = &HFE AndAlso bytes(1) = &HFF Then
                            encoding = Encoding.BigEndianUnicode
                        End If
                        Dim content = encoding.GetString(bytes, 0, count).TrimStart(ChrW(&HFEFF))
                        Dim records = Regex.Matches(content, "(?m)^(?<time>[0-9]{2}/[0-9]{2}/[0-9]{4} [0-9]{2}:[0-9]{2}:[0-9]{2}\.[0-9]{3})\tEXCEL \(0x(?<pid>[0-9a-fA-F]+)\)\t")
                        If records.Count = 0 Then Return False
                        For Each record As Match In records
                            If Long.Parse(record.Groups("pid").Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture) <> Pid Then Return False
                            Dim timestamp As DateTime
                            Dim latest = If(OwnedProcess.HasExited, OwnedProcess.ExitTime.ToUniversalTime(), DateTime.UtcNow.AddSeconds(1))
                            If Not DateTime.TryParseExact(record.Groups("time").Value, "MM/dd/yyyy HH:mm:ss.fff", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal, timestamp) OrElse timestamp < Started OrElse timestamp > latest Then Return False
                        Next
                        Return True
                    End Using
                End Using
            Catch
                ' Exclusive Office locks may instead be verified by the query-only Restart Manager.
                Return Nothing
            End Try
        End Function

        Private Sub Clean(path As String, expected As DiagnosticLogNative.FileInformation, regular As Boolean)
            Try
                If Not OwnedProcess.HasExited OrElse Not SafePath(path) Then Return
                Using directories = DiagnosticLogNative.LockParents(path)
                    Using handle = DiagnosticLogNative.OpenExclusive(path, True)
                        Dim actual = DiagnosticLogNative.Information(handle)
                        If (actual.Attributes And CUInt(FileAttributes.ReadOnly)) <> 0 Then Return
                        If Not expected.SameFile(actual) OrElse expected.Write <> actual.Write OrElse expected.SizeHigh <> actual.SizeHigh OrElse expected.SizeLow <> actual.SizeLow Then Return
                        If regular Then
                            DiagnosticLogNative.Remove(handle)
                        Else
                            Using stream As New FileStream(handle, FileAccess.ReadWrite)
                                TrimPadding(stream)
                            End Using
                        End If
                    End Using
                End Using
            Catch
                ' Locked/read-only/replaced files remain untouched; later attempts may succeed.
            End Try
        End Sub

        Friend Shared Sub TrimPadding(stream As Stream)
            Dim header(511) As Byte
            stream.Position = 0
            Dim headerCount = 0
            While headerCount < header.Length
                Dim count = stream.Read(header, headerCount, header.Length - headerCount)
                If count = 0 Then Exit While
                headerCount += count
            End While
            ' Unknown BOM-less encodings retain up to three ambiguous padding bytes.
            ' This protects terminal UTF-16/UTF-32 code units even for very short logs.
            Dim unit As Integer = 4
            If headerCount >= 4 AndAlso ((header(0) = &HFF AndAlso header(1) = &HFE AndAlso header(2) = 0 AndAlso header(3) = 0) OrElse (header(0) = 0 AndAlso header(1) = 0 AndAlso header(2) = &HFE AndAlso header(3) = &HFF)) Then
                unit = 4
            ElseIf headerCount >= 2 AndAlso ((header(0) = &HFF AndAlso header(1) = &HFE) OrElse (header(0) = &HFE AndAlso header(1) = &HFF)) Then
                unit = 2
            ElseIf headerCount >= 3 AndAlso header(0) = &HEF AndAlso header(1) = &HBB AndAlso header(2) = &HBF Then
                unit = 1
            Else
                Dim prefix = Encoding.ASCII.GetString(header, 0, headerCount)
                If prefix.StartsWith("Timestamp", StringComparison.Ordinal) OrElse Regex.IsMatch(prefix, "^[0-9]{2}/[0-9]{2}/[0-9]{4} ") Then unit = 1
                ' Office text without a BOM: recognize the ASCII header in Unicode code units.
                If headerCount >= 32 Then
                    For Each width In {4, 2}
                        For Each bigEndian In {False, True}
                            Dim matches As Boolean = True
                            For offset = 0 To 7
                                For index = 0 To width - 1
                                    Dim value = header(offset * width + index)
                                    If index = If(bigEndian, width - 1, 0) Then
                                        If value < 9 OrElse value > 126 Then matches = False
                                    ElseIf value <> 0 Then
                                        matches = False
                                    End If
                                Next
                            Next
                            If matches Then unit = width
                        Next
                    Next
                End If
            End If
            Dim buffer(8191) As Byte
            Dim boundary = stream.Length
            While boundary > 0
                Dim size = CInt(Math.Min(buffer.Length, boundary))
                stream.Position = boundary - size
                Dim read = 0
                While read < size
                    Dim count = stream.Read(buffer, read, size - read)
                    If count = 0 Then Throw New EndOfStreamException()
                    read += count
                End While
                For index = size - 1 To 0 Step -1
                    If buffer(index) <> 0 Then
                        Dim length = boundary - size + index + 1
                        ' Keep the rest of the terminal UTF-16/UTF-32 code unit, including BOMs.
                        length = Math.Min(stream.Length, ((length + unit - 1) \ unit) * unit)
                        If length < stream.Length Then stream.SetLength(length)
                        Return
                    End If
                Next
                boundary -= size
            End While
            If stream.Length <> 0 Then stream.SetLength(0)
        End Sub
    End Class
End Namespace
