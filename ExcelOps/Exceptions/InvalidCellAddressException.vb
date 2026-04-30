Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type.
    ''' </summary>
    Public Class InvalidCellAddressException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        ''' <summary>
        ''' Creates an exception for an invalid cell address.
        ''' </summary>
        ''' <param name="cellAddress">Invalid cell address.</param>
        Public Sub New(cellAddress As ExcelCell)
            MyBase.New(ErrorMessage(cellAddress))
            Me.CellAddress = cellAddress
        End Sub

        ''' <summary>
        ''' Creates an exception for an invalid cell address.
        ''' </summary>
        ''' <param name="cellAddress">Invalid cell address.</param>
        ''' <param name="innerException">Original exception that caused this exception.</param>
        Public Sub New(cellAddress As ExcelCell, innerException As Exception)
            MyBase.New(ErrorMessage(cellAddress), innerException)
            Me.CellAddress = cellAddress
        End Sub

        ''' <summary>
        ''' Creates an exception for an invalid cell address.
        ''' </summary>
        ''' <param name="targetSheetName">Worksheet name of the invalid address.</param>
        ''' <param name="rowIndex">Zero-based row index of the invalid address.</param>
        ''' <param name="columnIndex">Zero-based column index of the invalid address.</param>
        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer)
            Me.New(CalculatedCellAddress(targetSheetName, rowIndex, columnIndex))
        End Sub

        ''' <summary>
        ''' Creates an exception for an invalid cell address.
        ''' </summary>
        ''' <param name="targetSheetName">Worksheet name of the invalid address.</param>
        ''' <param name="rowIndex">Zero-based row index of the invalid address.</param>
        ''' <param name="columnIndex">Zero-based column index of the invalid address.</param>
        ''' <param name="innerException">Original exception that caused this exception.</param>
        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer, innerException As Exception)
            Me.New(CalculatedCellAddress(targetSheetName, rowIndex, columnIndex), innerException)
        End Sub

        ''' <summary>
        ''' Gets or sets the invalid cell address.
        ''' </summary>
        Public Property CellAddress As ExcelCell

        Private Shared Function CalculatedCellAddress(targetSheetName As String, rowIndex As Integer, columnIndex As Integer) As ExcelCell
            Return New ExcelCell(targetSheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
        End Function

        Private Shared Function ErrorMessage(cellAddress As ExcelCell) As String
            If cellAddress Is Nothing Then
                Return "Missing cell address (NullReference at instance)"
            ElseIf cellAddress.SheetName = Nothing Then
                Return "Missing sheet name"
            ElseIf cellAddress.Address = Nothing Then
                Return "Missing cell address (NullReference at property Address)"
            ElseIf cellAddress.ValidateFullCellAddressInclSheetName = False Then
                Return "Invalid address: " & cellAddress.Address(True)
            Else
                Return "Unknown reason for invalid address: " & cellAddress.Address(True)
            End If
        End Function
    End Class
End Namespace
