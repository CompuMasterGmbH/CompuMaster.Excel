Public Class ExcelRange
    Inherits ComChildObject(Of ExcelSheet, Object)

    Friend Sub New(parentItemResponsibleForDisposal As ExcelSheet, sheet As ExcelSheet, rangeName As String)
        MyBase.New(parentItemResponsibleForDisposal, sheet.InvokeFunction(Of Object)("Range", rangeName))
    End Sub

End Class