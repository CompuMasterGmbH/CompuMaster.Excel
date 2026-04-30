''' <summary>
''' A wrapper for an Excel dialog.
''' </summary>
Public Class ExcelDialog
    Inherits ComChildObject(Of ExcelApplication, Object)

    ''' <summary>
    ''' Creates a wrapper for an Excel dialog COM object.
    ''' </summary>
    ''' <param name="parentItemResponsibleForDisposal">Parent application responsible for disposal.</param>
    ''' <param name="comObject">Excel dialog COM object.</param>
    Public Sub New(parentItemResponsibleForDisposal As ExcelApplication, comObject As Object)
        MyBase.New(parentItemResponsibleForDisposal, comObject)
    End Sub

    ''' <summary>
    ''' Show a dialog.
    ''' </summary>
    ''' <returns>True if the dialog was confirmed, false if the dialog was cancelled</returns>
    Public Function Show() As Boolean
        Return InvokeFunction(Of Boolean)("Show")
    End Function

End Class
