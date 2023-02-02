Public Class ExcelWorkbooksCollection
    Inherits ComObjectBase

    Friend Sub New(parentItemResponsibleForDisposal As ComObjectBase, app As ExcelApplication)
        MyBase.New(parentItemResponsibleForDisposal, app.InvokePropertyGet(Of Object)("Workbooks"))
        Me.Parent = app
    End Sub

    Friend ReadOnly Parent As ExcelApplication

    Public Workbooks As New List(Of ExcelWorkbook)

    Public Function Open(path As String) As ExcelWorkbook
        Dim wb As New ExcelWorkbook(Me, Me, path)
        Me.Workbooks.Add(wb)
        Return wb
    End Function

    Protected Overrides Sub OnDisposeChildren()
        For MyCounter As Integer = Workbooks.Count - 1 To 0 Step -1
            Workbooks(MyCounter).Dispose()
        Next
    End Sub

    Protected Overrides Sub OnClosing()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub

End Class
