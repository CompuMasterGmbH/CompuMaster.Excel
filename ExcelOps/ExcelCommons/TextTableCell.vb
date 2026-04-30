Namespace ExcelOps

    ''' <summary>
    ''' Cell data of a text table
    ''' </summary>
    Public Class TextTableCell

        ''' <summary>
        ''' Creates a text table cell reference.
        ''' </summary>
        ''' <param name="sheetName">Name of the worksheet containing the cell.</param>
        ''' <param name="address">Cell address without the worksheet name.</param>
        ''' <param name="formula">Formula or cell content stored for the cell.</param>
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

        ''' <inheritdoc/>
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
