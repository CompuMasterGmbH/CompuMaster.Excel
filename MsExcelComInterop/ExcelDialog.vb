Public Class ExcelDialog
    Inherits ComObjectBase

    Public Sub New(parentItemResponsibleForDisposal As ComObjectBase, comObject As Object)
        MyBase.New(parentItemResponsibleForDisposal, comObject)
    End Sub

    ''' <summary>
    ''' Show a dialog
    ''' </summary>
    ''' <returns>True if the dialog was confirmed, false if the dialog was cancelled</returns>
    Public Function Show() As Boolean
        Return InvokeFunction(Of Boolean)("Show")
    End Function

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub
End Class
