Namespace ExcelOps

    ''' <summary>
    ''' Cell data of a text table
    ''' </summary>
    Public Class TextTableCell

        Public Sub New(sheetName As String, address As String, formula As String)
            Me.SheetName = sheetName
            Me.Address = address
            Me.CellContent = formula
        End Sub

        ''' <summary>
        ''' The formula (or cell content as value)
        ''' </summary>
        Public Property CellContent As String

        ''' <summary>
        ''' Sheet name containing the formula
        ''' </summary>
        Public Property SheetName As String

        ''' <summary>
        ''' Cell address containing the formula
        ''' </summary>
        Public Property Address As String

        ''' <summary>
        ''' A text representation of sheet name and address and it's formula/cell content
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function ToString() As String
            Return "'" & Me.SheetName & "'!" & Me.Address & ":" & Me.CellContent
        End Function

        ''' <summary>
        ''' The excel cell address representation of this text table cell
        ''' </summary>
        ''' <returns></returns>
        Public Function ToExcelCellAddress() As Excel.ExcelOps.ExcelCell
            Return New Excel.ExcelOps.ExcelCell(Me.SheetName, Me.Address, Excel.ExcelOps.ExcelCell.ValueTypes.All)
        End Function

    End Class

End Namespace