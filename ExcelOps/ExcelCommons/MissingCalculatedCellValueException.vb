Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type
    ''' </summary>
    Public Class MissingCalculatedCellValueException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        Public Sub New(filePath As String, cell As ExcelCell, foundFormula As String)
            MyBase.New("Lookup of cell value failed for formula cell " & cell.Address(True) & " without calculated value" &
            If(foundFormula <> Nothing, "; found formula=""" & foundFormula & """", ""))
            Me.Cell = cell
            Me.FoundFormula = foundFormula
            Me.FilePath = filePath
        End Sub

        Public Sub New(filePath As String, targetSheetName As String, rowIndex As Integer, columnIndex As Integer, foundFormula As String)
            Me.New(filePath, New ExcelCell(targetSheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All), foundFormula)
        End Sub

        Public Property Cell As ExcelCell

        Public Property FoundFormula As String

        Public Property FilePath As String

    End Class
End Namespace