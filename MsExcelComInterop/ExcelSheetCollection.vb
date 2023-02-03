Public Class ExcelSheetCollection
    Inherits ComChildObject(Of ExcelWorkbook, Object)

    Public Sub New(parent As ExcelWorkbook)
        MyBase.New(parent, parent.InvokePropertyGet(Of Object)("Sheets"))
    End Sub

    Public Function Add() As ExcelSheet
        Return New ExcelSheet(Me, InvokeFunction(Of Object)("Add"))
    End Function

    Public Function Item(sheetName As String) As ExcelSheet
        If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
        Return New ExcelSheet(Me, InvokePropertyGet(Of Object)("Item", sheetName))
    End Function

    Public Function Item(index As Integer) As ExcelSheet
        Return New ExcelSheet(Me, InvokePropertyGet(Of Object)("Item", index + 1))
    End Function

    Public Function Count() As Integer
        Return InvokePropertyGet(Of Integer)("Count")
    End Function

End Class
