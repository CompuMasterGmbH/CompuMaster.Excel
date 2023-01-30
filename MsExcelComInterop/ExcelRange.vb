Public Class ExcelRange
    Inherits ComObjectBase

    Friend Sub New(parentItemResponsibleForDisposal As ComObjectBase, sheet As ExcelSheet, rangeName As String)
        MyBase.New(parentItemResponsibleForDisposal, sheet.InvokeFunction("Range", rangeName))
    End Sub

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub

End Class