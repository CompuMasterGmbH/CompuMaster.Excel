Imports System.Runtime.InteropServices
Imports System.Threading
Imports CompuMaster.Excel.ExcelOps
Imports Microsoft.Office.Interop.Excel
Imports NUnit.Framework

<TestFixture>
<Platform("Win")>
Public Class OpenXmlNativeRtdTest

    <Test>
    Public Sub ReusesSubscriptionsAndReceivesChangesOnTheComApartment()
        Dim provider As New TestServer()
        Using updated As New ManualResetEventSlim()
            Using client As New OpenXmlNativeRtdClient(Sub(sender) updated.Set(), Function(progId) provider)
                Dim topics = {"Hall1", "Temperature"}
                Assert.That(client.Read("Example.Rtd", topics).Value, [Is].EqualTo(21.0))
                Assert.That(client.Read("example.rtd", topics).Value, [Is].EqualTo(21.0))
                Assert.That(provider.Subscriptions, [Is].EqualTo(1))
                Assert.That(provider.Apartment, [Is].EqualTo(ApartmentState.STA))

                provider.Notify(22)
                Assert.That(updated.Wait(TimeSpan.FromSeconds(5)), [Is].True)
                Assert.That(client.Read("Example.Rtd", topics).Value, [Is].EqualTo(22.0))
            End Using
        End Using
        Assert.That(provider.Unsubscriptions, [Is].EqualTo(1))
        Assert.That(provider.Terminated, [Is].True)
    End Sub

    <Test>
    Public Sub ReportsUnavailableNativeServersAsExcelErrors()
        Using client As New OpenXmlNativeRtdClient(Sub(sender) Return, Function(progId) Nothing)
            Assert.That(client.Read("Example.Rtd", {"Temperature"}).Value, [Is].EqualTo("#N/A"))
        End Using
    End Sub

    Private NotInheritable Class TestServer
        Implements IRtdServer

        Private _callback As IRTDUpdateEvent
        Private _topicId As Integer
        Private _value As Integer = 21

        Friend Subscriptions As Integer
        Friend Unsubscriptions As Integer
        Friend Apartment As ApartmentState
        Friend Terminated As Boolean

        Public Function ServerStart(callback As IRTDUpdateEvent) As Integer Implements IRtdServer.ServerStart
            Apartment = Thread.CurrentThread.GetApartmentState()
            _callback = callback
            Dim unknown = Marshal.GetIUnknownForObject(callback)
            Try
                Dim eventInterface = New Guid("A43788C1-D91B-11D3-8F39-00C04F3651B8")
                Dim pointer As IntPtr
                Marshal.ThrowExceptionForHR(Marshal.QueryInterface(unknown, eventInterface, pointer))
                Marshal.Release(pointer)
            Finally
                Marshal.Release(unknown)
            End Try
            Return 1
        End Function

        Public Function ConnectData(topicId As Integer, ByRef topics As Array, ByRef requestFresh As Boolean) As Object Implements IRtdServer.ConnectData
            _topicId = topicId
            Subscriptions += 1
            Assert.That(topics.Cast(Of String)(), [Is].EqualTo({"Hall1", "Temperature"}))
            Assert.That(requestFresh, [Is].True)
            Return Volatile.Read(_value)
        End Function

        Public Function RefreshData(ByRef topicCount As Integer) As Array Implements IRtdServer.RefreshData
            topicCount = 1
            Return New Object(,) {{_topicId}, {Volatile.Read(_value)}}
        End Function

        Public Sub DisconnectData(topicId As Integer) Implements IRtdServer.DisconnectData
            Assert.That(topicId, [Is].EqualTo(_topicId))
            Unsubscriptions += 1
        End Sub

        Public Function Heartbeat() As Integer Implements IRtdServer.Heartbeat
            Return 1
        End Function

        Public Sub ServerTerminate() Implements IRtdServer.ServerTerminate
            Terminated = True
        End Sub

        Friend Sub Notify(value As Integer)
            Volatile.Write(_value, value)
            _callback.UpdateNotify()
        End Sub
    End Class
End Class
