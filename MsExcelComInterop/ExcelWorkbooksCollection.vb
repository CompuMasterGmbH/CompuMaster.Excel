<CodeAnalysis.SuppressMessage("Naming", "CA1711:Bezeichner dürfen kein falsches Suffix aufweisen", Justification:="<Ausstehend>")>
Public Class ExcelWorkbooksCollection
    Inherits ComChildObject(Of ExcelApplication, Object)

    Friend Sub New(app As ExcelApplication)
        MyBase.New(app, app.InvokePropertyGet(Of Object)("Workbooks"))
    End Sub

    Public Property Workbooks As New List(Of ExcelWorkbook)

    Public Function Open(path As String) As ExcelWorkbook
        Dim wb As New ExcelWorkbook(Me, path)
        Me.Workbooks.Add(wb)
        Return wb
    End Function

    Protected Overrides Sub OnDisposeChildren()
        For MyCounter As Integer = Workbooks.Count - 1 To 0 Step -1
            Workbooks(MyCounter).Dispose()
        Next
    End Sub

End Class
