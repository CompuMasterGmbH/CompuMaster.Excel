Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type.
    ''' </summary>
    Public Class InvalidCellTypeException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        ''' <summary>
        ''' Creates an exception for a cell value with an unexpected type.
        ''' </summary>
        ''' <param name="targetSheetName">Name of the worksheet containing the invalid value.</param>
        ''' <param name="rowIndex">Zero-based row index of the cell.</param>
        ''' <param name="columnIndex">Zero-based column index of the cell.</param>
        ''' <param name="foundFormattedText">Formatted display text found in the cell.</param>
        ''' <param name="foundFormula">Formula found in the cell, or <see langword="Nothing"/>.</param>
        ''' <param name="expectedDataType">Expected data type.</param>
        ''' <param name="innerException">Original exception that caused this exception.</param>
        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer, foundFormattedText As String, foundFormula As String, expectedDataType As Type, innerException As Exception)
            MyBase.New("Expected type """ & expectedDataType.Name & """ in cell " & CellAddress(targetSheetName, rowIndex, columnIndex) &
            ", but " & If(foundFormula <> Nothing, "found formula & """ & foundFormula & """ resulting in display value """, "found display value """) & foundFormattedText & """", innerException)
            Me.TargetSheetName = targetSheetName
            Me.CellRowIndex = rowIndex
            Me.CellColumnIndex = columnIndex
            Me.ExpectedDataType = expectedDataType
            Me.FoundFormattedText = foundFormattedText
            Me.FoundFormula = foundFormula
        End Sub

        ''' <summary>
        ''' Gets or sets the worksheet name containing the invalid value.
        ''' </summary>
        Public Property TargetSheetName As String
        ''' <summary>
        ''' Gets or sets the zero-based row index of the cell.
        ''' </summary>
        Public Property CellRowIndex As Integer
        ''' <summary>
        ''' Gets or sets the zero-based column index of the cell.
        ''' </summary>
        Public Property CellColumnIndex As Integer
        ''' <summary>
        ''' Gets or sets the expected data type.
        ''' </summary>
        Public Property ExpectedDataType As Type
        ''' <summary>
        ''' Gets or sets the formatted display text found in the cell.
        ''' </summary>
        Public Property FoundFormattedText As String
        ''' <summary>
        ''' Gets or sets the formula found in the cell.
        ''' </summary>
        Public Property FoundFormula As String

        Private Shared Function CellAddress(targetSheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Dim Cell As New ExcelCell(targetSheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
            Return Cell.Address(True)
        End Function

    End Class
End Namespace
