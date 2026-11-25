Namespace ExcelOps

    ''' <summary>Reports that a formula depends on its own calculated value.</summary>
    Public Class CircularReferenceException
        Inherits InvalidOperationException

        ''' <summary>Creates an exception for a circular formula reference.</summary>
        ''' <param name="cell">The cell whose calculated value was requested again.</param>
        ''' <param name="formula">The stored formula in that cell.</param>
        Public Sub New(cell As ExcelCell, formula As String)
            MyBase.New("Circular formula reference at " & cell.Address(True) & ".")
            Me.Cell = cell
            Me.Formula = formula
        End Sub

        ''' <summary>Gets the cell whose value caused the circular reference.</summary>
        ''' <value>The worksheet and cell address.</value>
        Public ReadOnly Property Cell As ExcelCell

        ''' <summary>Gets the formula in the cell that caused the circular reference.</summary>
        ''' <value>The stored formula text.</value>
        Public ReadOnly Property Formula As String

    End Class

End Namespace
