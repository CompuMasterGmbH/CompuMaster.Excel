Public Class ExcelDialog
    Inherits ComChildObject(Of ExcelApplication, Object)

    Public Sub New(parentItemResponsibleForDisposal As ExcelApplication, comObject As Object)
        MyBase.New(parentItemResponsibleForDisposal, comObject)
    End Sub

    ''' <summary>
    ''' Show a dialog
    ''' </summary>
    ''' <returns>True if the dialog was confirmed, false if the dialog was cancelled</returns>
    Public Function Show() As Boolean
        Return InvokeFunction(Of Boolean)("Show")
    End Function

End Class
