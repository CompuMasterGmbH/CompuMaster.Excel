Namespace ExcelOps
    Public Class TextTableCell
        Public Sub New(sheetName As String, address As String, formula As String)
            Me.SheetName = sheetName
            Me.Address = address
            Me.CellContent = formula
        End Sub
        Public CellContent As String
        Public SheetName As String
        Public Address As String
        Public Overrides Function ToString() As String
            Return "'" & Me.SheetName & "'!" & Me.Address & ":" & Me.CellContent
        End Function
    End Class
End Namespace