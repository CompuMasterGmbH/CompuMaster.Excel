Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Win32.SafeHandles

Namespace Global.CompuMaster.Excel.MsExcelCom

    ' Only query Restart Manager. Never request shutdown, restart, or lock removal.
    Friend NotInheritable Class DiagnosticLogNative
        <StructLayout(LayoutKind.Sequential, Pack:=4)>
        Friend Structure FileInformation
            Friend Attributes As UInteger
            Friend Creation As Long
            Friend Access As Long
            Friend Write As Long
            Friend Volume As UInteger
            Friend SizeHigh As UInteger
            Friend SizeLow As UInteger
            Friend Links As UInteger
            Friend IndexHigh As UInteger
            Friend IndexLow As UInteger

            Friend Function SameFile(other As FileInformation) As Boolean
                Return Volume = other.Volume AndAlso IndexHigh = other.IndexHigh AndAlso IndexLow = other.IndexLow AndAlso Creation = other.Creation
            End Function
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure UniqueProcess
            Friend Pid As UInteger
            Friend Started As System.Runtime.InteropServices.ComTypes.FILETIME
        End Structure

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure ProcessInformation
            Friend Process As UniqueProcess
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Friend Name As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)> Friend Service As String
            Friend ApplicationType As UInteger
            Friend Status As UInteger
            Friend Session As UInteger
            <MarshalAs(UnmanagedType.Bool)> Friend Restartable As Boolean
        End Structure

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function CreateFileW(path As String, access As UInteger, share As UInteger, security As IntPtr, creation As UInteger, flags As UInteger, template As IntPtr) As SafeFileHandle
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function GetFileInformationByHandle(handle As SafeFileHandle, ByRef information As FileInformation) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function SetFileInformationByHandle(handle As SafeFileHandle, informationClass As Integer, ByRef deleteFile As Integer, size As UInteger) As Boolean
        End Function

        <DllImport("rstrtmgr.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function RmStartSession(ByRef session As UInteger, flags As UInteger, <Out, MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U2, SizeConst:=33)> key As Char()) As Integer
        End Function

        <DllImport("rstrtmgr.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function RmRegisterResources(session As UInteger, count As UInteger, <MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.LPWStr)> files As String(), appCount As UInteger, apps As IntPtr, serviceCount As UInteger, services As IntPtr) As Integer
        End Function

        <DllImport("rstrtmgr.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function RmGetList(session As UInteger, ByRef needed As UInteger, ByRef count As UInteger, <[In], Out> apps As ProcessInformation(), ByRef reasons As UInteger) As Integer
        End Function

        <DllImport("rstrtmgr.dll")>
        Private Shared Function RmEndSession(session As UInteger) As Integer
        End Function

        Friend Shared Function OpenMetadata(path As String) As SafeFileHandle
            Return CreateFileW(path, 0UI, 7UI, IntPtr.Zero, 3UI, &H2200000UI, IntPtr.Zero)
        End Function

        Friend Shared Function OpenExclusive(path As String, writable As Boolean) As SafeFileHandle
            ' OPEN_EXISTING + OPEN_REPARSE_POINT; DELETE access allows deletion by handle.
            Return CreateFileW(path, If(writable, &HC0010000UI, &H80000000UI), 0UI, IntPtr.Zero, 3UI, &H200000UI, IntPtr.Zero)
        End Function

        Friend Shared Function Information(handle As SafeFileHandle) As FileInformation
            Dim result As FileInformation
            If handle.IsInvalid OrElse Not GetFileInformationByHandle(handle, result) Then Throw New IOException("Cannot identify diagnostic log.")
            If (result.Attributes And CUInt(FileAttributes.ReparsePoint Or FileAttributes.Directory)) <> 0 OrElse result.Links <> 1 Then Throw New IOException("Unsafe diagnostic log.")
            Return result
        End Function

        Friend Shared Function LockParents(path As String) As IDisposable
            Dim lease As New DirectoryLease()
            Try
                Dim paths As New Stack(Of String)
                Dim directory As New DirectoryInfo(IO.Path.GetDirectoryName(path))
                While directory IsNot Nothing
                    paths.Push(directory.FullName)
                    directory = directory.Parent
                End While
                For Each parent In paths
                    ' Deny directory deletion/renaming while resolving and operating on its child.
                    Dim handle = CreateFileW(parent, 0UI, 3UI, IntPtr.Zero, 3UI, &H2200000UI, IntPtr.Zero)
                    lease.DirectoryHandles.Add(handle)
                    Dim info As FileInformation
                    If handle.IsInvalid OrElse Not GetFileInformationByHandle(handle, info) OrElse (info.Attributes And CUInt(FileAttributes.ReparsePoint)) <> 0 OrElse (info.Attributes And CUInt(FileAttributes.Directory)) = 0 Then Throw New IOException("Unsafe diagnostic directory.")
                Next
                Return lease
            Catch
                lease.Dispose()
                Throw
            End Try
        End Function

        Private NotInheritable Class DirectoryLease
            Implements IDisposable

            Friend ReadOnly DirectoryHandles As New List(Of SafeFileHandle)

            ''' <inheritdoc/>
            Public Sub Dispose() Implements IDisposable.Dispose
                For Each handle In DirectoryHandles
                    handle.Dispose()
                Next
            End Sub
        End Class

        Friend Shared Sub Remove(handle As SafeFileHandle)
            Dim disposition As Integer = 1
            If Not SetFileInformationByHandle(handle, 4, disposition, 4UI) Then Throw New IOException("Cannot remove diagnostic log.")
        End Sub

        Friend Shared Function IsOnlyOwner(path As String, pid As Integer, started As DateTime) As Boolean
            Dim session As UInteger
            Dim key(32) As Char
            If RmStartSession(session, 0UI, key) <> 0 Then Return False
            Try
                If RmRegisterResources(session, 1UI, {path}, 0UI, IntPtr.Zero, 0UI, IntPtr.Zero) <> 0 Then Return False
                Dim needed, count, reasons As UInteger
                Dim code = RmGetList(session, needed, count, Nothing, reasons)
                If code <> 234 OrElse needed <> 1UI Then Return False
                Dim apps(0) As ProcessInformation
                count = 1UI
                If RmGetList(session, needed, count, apps, reasons) <> 0 OrElse count <> 1UI Then Return False
                Dim time = (CLng(apps(0).Process.Started.dwHighDateTime) << 32) Or (CLng(apps(0).Process.Started.dwLowDateTime) And &HFFFFFFFFL)
                Return apps(0).Process.Pid = CUInt(pid) AndAlso time = started.ToFileTimeUtc()
            Finally
                Dim endResult = RmEndSession(session)
            End Try
        End Function
    End Class
End Namespace
