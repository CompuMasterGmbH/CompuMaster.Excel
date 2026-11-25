Imports System.Collections.Concurrent
Imports System.IO

Namespace ExcelOps

    ''' <summary>Coordinates structural reference updates between open Open XML workbooks.</summary>
    ''' <remarks>
    ''' Attach workbooks that should participate in updates. The session identifies links by saved path or
    ''' by a temporary name assigned to each unsaved workbook. Save dependent workbooks while they are
    ''' attached to persist changes to their link targets. Closed workbooks do not participate.
    ''' Disposing the session detaches its workbooks.
    ''' </remarks>
    Public NotInheritable Class OpenXmlWorkbookSession
        Implements IDisposable

        Private ReadOnly _workbooks As New List(Of OpenXmlExcelDataOperations)
        Private ReadOnly _temporaryLinks As New ConcurrentDictionary(Of String, Uri)(StringComparer.Ordinal)
        Private _disposed As Boolean

        ''' <summary>Attaches an open workbook to this session.</summary>
        ''' <param name="workbook">The open workbook to attach.</param>
        ''' <exception cref="ArgumentNullException"><paramref name="workbook"/> is null.</exception>
        ''' <exception cref="InvalidOperationException">The workbook is closed or attached to another session.</exception>
        ''' <exception cref="ObjectDisposedException">The session is disposed.</exception>
        Public Sub Attach(workbook As OpenXmlExcelDataOperations)
            If workbook Is Nothing Then Throw New ArgumentNullException(NameOf(workbook))
            SyncLock OpenXmlExcelDataOperations.WorkbookSessionLock
                If _disposed Then Throw New ObjectDisposedException(NameOf(OpenXmlWorkbookSession))
                If workbook.IsClosed Then Throw New InvalidOperationException("The workbook is closed.")
                If workbook.WorkbookSession Is Me Then Return
                If workbook.WorkbookSession IsNot Nothing Then
                    Throw New InvalidOperationException("The workbook is already attached to another session.")
                End If
                _workbooks.Add(workbook)
                RegisterTemporaryLink(workbook.TemporaryWorkbookName)
                workbook.WorkbookSession = Me
            End SyncLock
        End Sub

        ''' <summary>Detaches a workbook from this session.</summary>
        ''' <param name="workbook">The workbook to detach.</param>
        ''' <returns>True if the workbook was attached to this session.</returns>
        Public Function Detach(workbook As OpenXmlExcelDataOperations) As Boolean
            SyncLock OpenXmlExcelDataOperations.WorkbookSessionLock
                If Not _workbooks.Remove(workbook) Then Return False
                workbook.WorkbookSession = Nothing
                Return True
            End SyncLock
        End Function

        ''' <inheritdoc/>
        Public Sub Dispose() Implements IDisposable.Dispose
            SyncLock OpenXmlExcelDataOperations.WorkbookSessionLock
                For Each workbook In _workbooks
                    workbook.WorkbookSession = Nothing
                Next
                _workbooks.Clear()
                _disposed = True
            End SyncLock
        End Sub

        ' The caller holds WorkbookSessionLock and locks each workbook before using its document or path.
        Friend Function Snapshot() As OpenXmlExcelDataOperations()
            Return _workbooks.ToArray()
        End Function

        Friend Function ResolveTemporaryLink(link As Uri) As Uri
            If link.IsAbsoluteUri Then Return Nothing
            Dim target As Uri = Nothing
            _temporaryLinks.TryGetValue(link.OriginalString, target)
            Return target
        End Function

        Friend Sub RegisterTemporaryLink(name As String)
            _temporaryLinks.TryAdd(name, New Uri(name, UriKind.Relative))
        End Sub

        Friend Sub BindTemporaryLink(name As String, destination As String)
            Dim previous As Uri = Nothing
            If _temporaryLinks.TryGetValue(name, previous) Then
                _temporaryLinks.TryUpdate(name, New Uri(Path.GetFullPath(destination)), previous)
            End If
        End Sub

    End Class

End Namespace
