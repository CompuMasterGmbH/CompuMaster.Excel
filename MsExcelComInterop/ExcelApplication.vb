Public Class ExcelApplication
    Inherits ComRootObject(Of Object)

    Public Sub New()
#Disable Warning CA1416
        MyBase.New(CreateObject("Excel.Application"), Sub(instance) instance.InvokeMethod("Quit"))
#Enable Warning CA1416
        Me.Workbooks = New ExcelWorkbooksCollection(Me)
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

    Public Property Interactive As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("Interactive")
        End Get
        Set(value As Boolean)
            InvokePropertySet("Interactive", value)
        End Set
    End Property

    Public Property ScreenUpdating As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("ScreenUpdating")
        End Get
        Set(value As Boolean)
            InvokePropertySet("ScreenUpdating", value)
        End Set
    End Property

    Public Function Dialogs(type As Enumerations.XlBuiltInDialog) As ExcelDialog
        Return New ExcelDialog(Me, InvokePropertyGet(Of Object)("Dialogs", CType(type, Integer)))
    End Function

    Public Function Run(vbaMethodNameInclWorkbookName As String) As Object
        Return InvokeFunction(Of Object)("Run", New Object() {vbaMethodNameInclWorkbookName})
    End Function

    Public Function Run(workbookName As String, vbaMethod As String) As Object
        Return InvokeFunction(Of Object)("Run", New Object() {"'" & workbookName & "'!" & vbaMethod})
    End Function

    Private AdditionalDisposeChildrenList As New List(Of ComObjectBase)

    Protected Overrides Sub OnDisposeChildren()
        If Me.Workbooks IsNot Nothing Then Me.Workbooks.Dispose()
    End Sub

    Protected Overrides Sub onBeforeClosing()
        UserControl = True
    End Sub

End Class
