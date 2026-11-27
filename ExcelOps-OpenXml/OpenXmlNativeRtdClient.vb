Imports System.Collections.Concurrent
Imports System.Diagnostics
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports IRtdServer = Microsoft.Office.Interop.Excel.IRtdServer
Imports IRTDUpdateEvent = Microsoft.Office.Interop.Excel.IRTDUpdateEvent

Namespace ExcelOps

    ' A workbook owns the subscriptions. All COM calls and RCW releases run on one pumping STA.
    Friend NotInheritable Class OpenXmlNativeRtdClient
        Implements IDisposable

        Private ReadOnly _requests As New BlockingCollection(Of Action)
        Private ReadOnly _connections As New Dictionary(Of String, Connection)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _updated As Action(Of OpenXmlNativeRtdClient)
        Private ReadOnly _thread As Thread
        Private ReadOnly _activate As Func(Of String, Object)
        Private _disposed As Integer

        Friend Sub New(updated As Action(Of OpenXmlNativeRtdClient), Optional activate As Func(Of String, Object) = Nothing)
            If Not RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then Throw New PlatformNotSupportedException("Native RTD requires Windows COM.")
            _updated = updated
            _activate = If(activate, AddressOf ActivateServer)
            _thread = New Thread(AddressOf Run) With {.IsBackground = True, .Name = "Open XML RTD"}
            _thread.SetApartmentState(ApartmentState.STA)
            _thread.Start()
        End Sub

        Friend Function Read(progId As String, topics As String()) As ExcelFormulaValue
            Dim result As New TaskCompletionSource(Of ExcelFormulaValue)(TaskCreationOptions.RunContinuationsAsynchronously)
            _requests.Add(
                Sub()
                    Try
                        Dim current As Connection = Nothing
                        If Not _connections.TryGetValue(progId, current) Then
                            current = Connection.Open(Function() _activate(progId))
                            If current Is Nothing Then
                                result.SetResult(ExcelFormulaValue.FromError("#N/A"))
                                Return
                            End If
                            _connections.Add(progId, current)
                        End If
                        result.SetResult(current.Read(topics))
                    Catch ex As Exception
                        result.SetException(ex)
                    End Try
                End Sub)
            Return result.Task.GetAwaiter().GetResult()
        End Function

        Private Shared Function ActivateServer(progId As String) As Object
            If Not RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then Return Nothing
            Dim serverType = Type.GetTypeFromProgID(progId)
            Return If(serverType Is Nothing, Nothing, Activator.CreateInstance(serverType))
        End Function

        Private Sub Run()
            Try
                While Not _requests.IsCompleted
                    Dim message As NativeMessage
                    While PeekMessage(message, IntPtr.Zero, 0UI, 0UI, 1UI)
                        TranslateMessage(message)
                        DispatchMessage(message)
                    End While
                    Dim request As Action = Nothing
                    If _requests.TryTake(request, 25) Then request()
                    If _requests.IsAddingCompleted Then Continue While
                    Dim changed = False
                    For Each current In _connections.Values
                        changed = current.Poll() Or changed
                    Next
                    If changed Then _updated(Me)
                End While
            Finally
                For Each current In _connections.Values
                    current.Dispose()
                Next
                _connections.Clear()
            End Try
        End Sub

        ''' <inheritdoc/>
        Public Sub Dispose() Implements IDisposable.Dispose
            If Interlocked.Exchange(_disposed, 1) <> 0 Then Return
            _requests.CompleteAdding()
            _thread.Join()
            _requests.Dispose()
        End Sub

        Private Shared Function IsServerFailure(exception As Exception) As Boolean
            Return TypeOf exception Is COMException OrElse TypeOf exception Is InvalidComObjectException OrElse
                TypeOf exception Is InvalidCastException OrElse TypeOf exception Is SafeArrayRankMismatchException OrElse
                TypeOf exception Is SafeArrayTypeMismatchException OrElse TypeOf exception Is TargetInvocationException
        End Function

        Private Shared Sub ReleaseServer(instance As Object)
            If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) AndAlso instance IsNot Nothing AndAlso Marshal.IsComObject(instance) Then
                Marshal.ReleaseComObject(instance)
            End If
        End Sub

        Private NotInheritable Class Connection
            Implements IDisposable
            Private ReadOnly _instance As Object
            Private ReadOnly _server As IRtdServer
            Private ReadOnly _notifications As New RtdNotifications()
            Private ReadOnly _lastActivity As Stopwatch = Stopwatch.StartNew()
            Private ReadOnly _topicIds As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Private ReadOnly _values As New Dictionary(Of Integer, ExcelFormulaValue)
            Private _nextTopicId As Integer
            Private _disconnected As Boolean

            Private Sub New(instance As Object, server As IRtdServer)
                _instance = instance
                _server = server
            End Sub

            Friend Shared Function Open(activate As Func(Of Object)) As Connection
                Dim instance As Object = Nothing
                Dim result As Connection = Nothing
                Try
                    instance = activate()
                    Dim server = TryCast(instance, IRtdServer)
                    If server Is Nothing Then Return Nothing
                    Dim candidate As New Connection(instance, server)
                    If server.ServerStart(candidate._notifications) <= 0 Then Return Nothing
                    result = candidate
                    Return result
                Catch ex As Exception When IsServerFailure(ex)
                    Return Nothing
                Finally
                    If result Is Nothing Then ReleaseServer(instance)
                End Try
            End Function

            Friend Function Read(topics As String()) As ExcelFormulaValue
                If _disconnected Then Return ExcelFormulaValue.FromError("#N/A")
                Dim key = String.Concat(topics.Select(Function(topic) topic.Length.ToString(CultureInfo.InvariantCulture) & ":" & topic))
                Dim id As Integer
                If _topicIds.TryGetValue(key, id) Then Return _values(id)
                _nextTopicId += 1
                id = _nextTopicId
                Try
                    Dim topicData As Array = topics.Cast(Of Object)().ToArray()
                    Dim requestFresh = True
                    Dim value = ReadValue(_server.ConnectData(id, topicData, requestFresh))
                    _topicIds.Add(key, id)
                    _values.Add(id, value)
                    Return value
                Catch ex As Exception When IsServerFailure(ex)
                    Return ExcelFormulaValue.FromError("#N/A")
                End Try
            End Function

            Friend Function Poll() As Boolean
                If _disconnected Then Return False
                Try
                    If _notifications.HasDisconnected() Then Return MarkDisconnected()
                    If _notifications.TakeUpdate() Then
                        _lastActivity.Restart()
                        Return Refresh()
                    End If
                    Dim heartbeatInterval = _notifications.HeartbeatInterval
                    If heartbeatInterval >= 0 AndAlso _lastActivity.ElapsedMilliseconds >= Math.Max(15000, heartbeatInterval) Then
                        _lastActivity.Restart()
                        If _server.Heartbeat() <= 0 Then Return MarkDisconnected()
                    End If
                    Return False
                Catch ex As Exception When IsServerFailure(ex)
                    Return MarkDisconnected()
                End Try
            End Function

            Private Function Refresh() As Boolean
                Dim count = 0
                Dim updates = _server.RefreshData(count)
                If count <= 0 OrElse updates Is Nothing OrElse updates.Rank <> 2 OrElse updates.GetLength(0) <> 2 Then Return False
                Dim kind = updates.GetLowerBound(0)
                Dim first = updates.GetLowerBound(1)
                Dim changed = False
                For offset = 0 To Math.Min(count, updates.GetLength(1)) - 1
                    Dim id = updates.GetValue(kind, first + offset)
                    If Not TypeOf id Is Integer Then Continue For
                    Dim topic = DirectCast(id, Integer)
                    Dim previous As ExcelFormulaValue
                    If Not _values.TryGetValue(topic, previous) Then Continue For
                    Dim current = ReadValue(updates.GetValue(kind + 1, first + offset))
                    If previous.Kind = current.Kind AndAlso Object.Equals(previous.Value, current.Value) Then Continue For
                    _values(topic) = current
                    changed = True
                Next
                Return changed
            End Function

            Private Function MarkDisconnected() As Boolean
                _disconnected = True
                For Each id In _values.Keys.ToArray()
                    _values(id) = ExcelFormulaValue.FromError("#N/A")
                Next
                Return _values.Count > 0
            End Function

            ''' <inheritdoc/>
            Public Sub Dispose() Implements IDisposable.Dispose
                Try
                    For Each id In _values.Keys
                        Try
                            _server.DisconnectData(id)
                        Catch ex As Exception When IsServerFailure(ex)
                            Trace.TraceWarning("RTD DisconnectData failed: {0}", ex.Message)
                        End Try
                    Next
                    _server.ServerTerminate()
                Catch ex As Exception When IsServerFailure(ex)
                    Trace.TraceWarning("RTD ServerTerminate failed: {0}", ex.Message)
                Finally
                    ReleaseServer(_instance)
                End Try
            End Sub

            Private Shared Function ReadValue(value As Object) As ExcelFormulaValue
                If value Is Nothing OrElse value Is DBNull.Value Then Return Nothing
                If TypeOf value Is String Then Return ExcelFormulaValue.FromText(DirectCast(value, String))
                If TypeOf value Is Boolean Then Return ExcelFormulaValue.FromBoolean(DirectCast(value, Boolean))
                If TypeOf value Is DateTime Then Return ExcelFormulaValue.FromNumber(DirectCast(value, DateTime).ToOADate())
                Dim errorValue = TryCast(value, ErrorWrapper)
                If errorValue IsNot Nothing Then Return ReadError(errorValue.ErrorCode)
                If TypeOf value Is Integer AndAlso (DirectCast(value, Integer) And &HFFFF0000) = &H800A0000 Then
                    Return ReadError(DirectCast(value, Integer))
                End If
                Select Case Type.GetTypeCode(value.GetType())
                    Case TypeCode.Byte, TypeCode.SByte, TypeCode.Int16, TypeCode.UInt16, TypeCode.Int32, TypeCode.UInt32,
                         TypeCode.Int64, TypeCode.UInt64, TypeCode.Single, TypeCode.Double, TypeCode.Decimal
                        Return ExcelFormulaValue.FromNumber(Convert.ToDouble(value, CultureInfo.InvariantCulture))
                    Case Else
                        Return ExcelFormulaValue.FromError("#VALUE!")
                End Select
            End Function

            Private Shared Function ReadError(code As Integer) As ExcelFormulaValue
                Select Case code And &HFFFF
                    Case 2000 : Return ExcelFormulaValue.FromError("#NULL!")
                    Case 2007 : Return ExcelFormulaValue.FromError("#DIV/0!")
                    Case 2015 : Return ExcelFormulaValue.FromError("#VALUE!")
                    Case 2023 : Return ExcelFormulaValue.FromError("#REF!")
                    Case 2029 : Return ExcelFormulaValue.FromError("#NAME?")
                    Case 2036 : Return ExcelFormulaValue.FromError("#NUM!")
                    Case 2043 : Return ExcelFormulaValue.FromError("#GETTING_DATA")
                    Case Else : Return ExcelFormulaValue.FromError("#N/A")
                End Select
            End Function
        End Class

        <ComVisible(True), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(GetType(IRTDUpdateEvent))>
        Private NotInheritable Class RtdNotifications
            Implements IRTDUpdateEvent

            Private _heartbeatInterval As Integer = 15000
            Private _updated As Integer
            Private _disconnected As Integer

            ''' <inheritdoc/>
            Public Property HeartbeatInterval As Integer Implements IRTDUpdateEvent.HeartbeatInterval
                Get
                    Return Volatile.Read(_heartbeatInterval)
                End Get
                Set(value As Integer)
                    Volatile.Write(_heartbeatInterval, value)
                End Set
            End Property

            ''' <inheritdoc/>
            Public Sub UpdateNotify() Implements IRTDUpdateEvent.UpdateNotify
                Interlocked.Exchange(_updated, 1)
            End Sub

            ''' <inheritdoc/>
            Public Sub Disconnect() Implements IRTDUpdateEvent.Disconnect
                Interlocked.Exchange(_disconnected, 1)
            End Sub

            Friend Function TakeUpdate() As Boolean
                Return Interlocked.Exchange(_updated, 0) <> 0
            End Function

            Friend Function HasDisconnected() As Boolean
                Return Volatile.Read(_disconnected) <> 0
            End Function
        End Class

        <StructLayout(LayoutKind.Sequential)>
        Private Structure NativeMessage
            Private Window As IntPtr
            Private Message As UInteger
            Private WParam As UIntPtr
            Private LParam As IntPtr
            Private Time As UInteger
            Private CursorX As Integer
            Private CursorY As Integer
            Private PrivateData As UInteger
        End Structure

        <DllImport("user32.dll", EntryPoint:="PeekMessageW", ExactSpelling:=True)>
        Private Shared Function PeekMessage(ByRef message As NativeMessage, window As IntPtr, minimum As UInteger,
                                            maximum As UInteger, remove As UInteger) As Boolean
        End Function

        <DllImport("user32.dll", ExactSpelling:=True)>
        Private Shared Function TranslateMessage(ByRef message As NativeMessage) As Boolean
        End Function

        <DllImport("user32.dll", EntryPoint:="DispatchMessageW", ExactSpelling:=True)>
        Private Shared Function DispatchMessage(ByRef message As NativeMessage) As IntPtr
        End Function
    End Class

End Namespace
