Public Class ExcelSheetCollection
    Inherits ComObjectBase

    Public Sub New(parentItemResponsibleForDisposal As ComObjectBase, workbook As ExcelWorkbook)
        MyBase.New(parentItemResponsibleForDisposal, workbook.InvokePropertyGet("Sheets"))
        Parent = workbook
    End Sub

    Public ReadOnly Property Parent As ExcelWorkbook

    Public Function Add() As ExcelSheet
        Return New ExcelSheet(Me, Me, InvokeFunction("Add"))
    End Function

    Public Function Item(sheetName As String) As ExcelSheet
        If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
        Return New ExcelSheet(Me, Me, InvokePropertyGet("Item", sheetName))
    End Function

    Public Function Item(index As Integer) As ExcelSheet
        Return New ExcelSheet(Me, Me, InvokePropertyGet("Item", index + 1))
    End Function

    Public Function Count() As Integer
        Return InvokePropertyGet(Of Integer)("Count")
    End Function

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub
End Class
