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

        ''' <summary>
        ''' Creates an exception for a formula cell without a cached calculated value.
        ''' </summary>
        ''' <param name="filePath">Workbook file path.</param>
        ''' <param name="cell">Formula cell without a cached calculated value.</param>
        ''' <param name="foundFormula">Formula found in the cell.</param>
        Public Sub New(filePath As String, cell As ExcelCell, foundFormula As String)
            MyBase.New("Lookup of cell value failed for formula cell " & cell.Address(True) & " without calculated value" &
            If(foundFormula <> Nothing, "; found formula=""" & foundFormula & """", ""))
            Me.Cell = cell
            Me.FoundFormula = foundFormula
            Me.FilePath = filePath
        End Sub

        ''' <summary>
        ''' Creates an exception for a formula cell without a cached calculated value.
        ''' </summary>
        ''' <param name="filePath">Workbook file path.</param>
        ''' <param name="targetSheetName">Name of the worksheet containing the formula cell.</param>
        ''' <param name="rowIndex">Zero-based row index of the formula cell.</param>
        ''' <param name="columnIndex">Zero-based column index of the formula cell.</param>
        ''' <param name="foundFormula">Formula found in the cell.</param>
        Public Sub New(filePath As String, targetSheetName As String, rowIndex As Integer, columnIndex As Integer, foundFormula As String)
            Me.New(filePath, New ExcelCell(targetSheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All), foundFormula)
        End Sub

        ''' <summary>
        ''' Gets or sets the formula cell without a cached calculated value.
        ''' </summary>
        Public Property Cell As ExcelCell

        ''' <summary>
        ''' Gets or sets the formula found in the cell.
        ''' </summary>
        Public Property FoundFormula As String

        ''' <summary>
        ''' Gets or sets the workbook file path.
        ''' </summary>
        Public Property FilePath As String

    End Class
End Namespace
