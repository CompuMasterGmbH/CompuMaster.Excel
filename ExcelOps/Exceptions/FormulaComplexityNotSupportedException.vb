Namespace ExcelOps

    ''' <summary>Reports that an engine cannot evaluate a formula with the requested calculation scope.</summary>
    Public Class FormulaComplexityNotSupportedException
        Inherits NotSupportedException

        ''' <summary>Creates an exception for a known formula calculation limitation.</summary>
        ''' <param name="cell">The affected formula cell.</param>
        ''' <param name="formula">The stored formula.</param>
        ''' <param name="message">The reason the requested calculation is unsupported.</param>
        ''' <param name="innerException">The original engine exception, if available.</param>
        Public Sub New(cell As ExcelCell, formula As String, message As String, Optional innerException As Exception = Nothing)
            MyBase.New(message, innerException)
            Me.Cell = cell
            Me.Formula = formula
        End Sub

        ''' <summary>Gets the affected cell.</summary>
        ''' <value>The worksheet and cell coordinates.</value>
        Public ReadOnly Property Cell As ExcelCell

        ''' <summary>Gets the stored formula.</summary>
        ''' <value>The formula text.</value>
        Public ReadOnly Property Formula As String

    End Class

End Namespace
