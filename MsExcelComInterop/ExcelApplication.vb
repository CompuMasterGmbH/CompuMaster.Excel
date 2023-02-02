Public Class ExcelApplication
    Inherits ComObjectBase

    Public Sub New()
        MyBase.New(Nothing, CreateObject("Excel.Application"))
        Me.Workbooks = New ExcelWorkbooksCollection(Me, Me)
    End Sub

    Public ReadOnly Property Workbooks As ExcelWorkbooksCollection

    Public Property UserControl As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("UserControl")
        End Get
        Set(value As Boolean)
            InvokePropertySet("UserControl", value)
        End Set
    End Property

    Public Property DisplayAlerts As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("DisplayAlerts")
        End Get
        Set(value As Boolean)
            InvokePropertySet("DisplayAlerts", value)
        End Set
    End Property

    Public Property Visible As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("Visible")
        End Get
        Set(value As Boolean)
            InvokePropertySet("Visible", value)
        End Set
    End Property

    Public Function Dialogs(type As Enumerations.XlBuiltInDialog) As ExcelDialog
        Return New ExcelDialog(Me, InvokePropertyGet("Dialogs", CType(type, Integer)))
    End Function

    Public Function Run(vbaMethodNameInclWorkbookName As String) As Object
        Return InvokeFunction(Of Object)("Run", New Object() {vbaMethodNameInclWorkbookName})
    End Function

    Public Function Run(workbookName As String, vbaMethod As String) As Object
        Return InvokeFunction(Of Object)("Run", New Object() {"'" & workbookName & "'!" & vbaMethod})
    End Function

    Public ReadOnly Property IsClosed As Boolean
        Get
            Return MyBase.IsDisposedComObject
        End Get
    End Property

    Public Sub Close()
        Me.Quit()
    End Sub

    Public Sub Quit()
        If Not IsDisposedComObject Then
            UserControl = True
            MyBase.CloseAndDisposeChildrenAndComObject()
        End If
    End Sub

    Private AdditionalDisposeChildrenList As New List(Of ComObjectBase)

    Protected Overrides Sub OnDisposeChildren()
        If Me.Workbooks IsNot Nothing Then Me.Workbooks.Dispose()
    End Sub

    Protected Overrides Sub OnClosing()
        InvokeMethod("Quit")
    End Sub

    Protected Overrides Sub OnClosed()
        GC.Collect(2, GCCollectionMode.Forced, True)
    End Sub

End Class
