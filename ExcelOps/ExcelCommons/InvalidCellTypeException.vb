Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type
    ''' </summary>
    Public Class InvalidCellTypeException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

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

        Public Property TargetSheetName As String
        Public Property CellRowIndex As Integer
        Public Property CellColumnIndex As Integer
        Public Property ExpectedDataType As Type
        Public Property FoundFormattedText As String
        Public Property FoundFormula As String

        Private Shared Function CellAddress(targetSheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Dim Cell As New ExcelCell(targetSheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
            Return Cell.Address(True)
        End Function

    End Class
End Namespace