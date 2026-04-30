Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data doesn't match with expected data
    ''' </summary>
    Public Class ExpectedDataMismatchException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        ''' <summary>
        ''' Creates an exception for a mismatching cell value.
        ''' </summary>
        ''' <param name="cellAddress">Cell address where the mismatch was found.</param>
        ''' <param name="expectedValue">Expected cell value.</param>
        ''' <param name="foundValue">Found cell value.</param>
        Public Sub New(cellAddress As ExcelCell, expectedValue As String, foundValue As String)
            Me.New(cellAddress.SheetName, cellAddress.RowIndex, cellAddress.ColumnIndex, expectedValue, foundValue, False)
        End Sub

        ''' <summary>
        ''' Creates an exception for a mismatching cell value.
        ''' </summary>
        ''' <param name="cellAddress">Cell address where the mismatch was found.</param>
        ''' <param name="expectedValue">Expected cell value.</param>
        ''' <param name="foundValue">Found cell value.</param>
        ''' <param name="foundValueIsAlreadySlightlyModified">Whether the found value was already slightly modified before this exception was created.</param>
        Public Sub New(cellAddress As ExcelCell, expectedValue As String, foundValue As String, foundValueIsAlreadySlightlyModified As Boolean)
            Me.New(cellAddress.SheetName, cellAddress.RowIndex, cellAddress.ColumnIndex, expectedValue, foundValue, foundValueIsAlreadySlightlyModified)
        End Sub

        ''' <summary>
        ''' Creates an exception for a mismatching cell value.
        ''' </summary>
        ''' <param name="targetSheetName">Name of the worksheet containing the mismatch.</param>
        ''' <param name="rowIndex">Zero-based row index of the mismatching cell.</param>
        ''' <param name="columnIndex">Zero-based column index of the mismatching cell.</param>
        ''' <param name="expectedValue">Expected cell value.</param>
        ''' <param name="foundValue">Found cell value.</param>
        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer, expectedValue As String, foundValue As String)
            Me.New(targetSheetName, rowIndex, columnIndex, expectedValue, foundValue, False)
        End Sub

        ''' <summary>
        ''' Creates an exception for a mismatching cell value.
        ''' </summary>
        ''' <param name="targetSheetName">Name of the worksheet containing the mismatch.</param>
        ''' <param name="rowIndex">Zero-based row index of the mismatching cell.</param>
        ''' <param name="columnIndex">Zero-based column index of the mismatching cell.</param>
        ''' <param name="expectedValue">Expected cell value.</param>
        ''' <param name="foundValue">Found cell value.</param>
        ''' <param name="foundValueIsAlreadySlightlyModified">Whether the found value was already slightly modified before this exception was created.</param>
        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer, expectedValue As String, foundValue As String, foundValueIsAlreadySlightlyModified As Boolean)
            MyBase.New("Expected value """ & expectedValue & """ in cell " & CellAddress(targetSheetName, rowIndex, columnIndex) &
          If(foundValueIsAlreadySlightlyModified, ", but found ", ", but found (and corrected to) ") & """" & foundValue & """")
            Me.TargetSheetName = targetSheetName
            Me.CellRowIndex = rowIndex
            Me.CellColumnIndex = columnIndex
            Me.ExpectedValue = expectedValue
            Me.FoundValue = foundValue
        End Sub

        ''' <summary>
        ''' Creates an exception for a mismatching cell value and includes a table snapshot.
        ''' </summary>
        ''' <param name="targetSheetName">Name of the worksheet containing the mismatch.</param>
        ''' <param name="rowIndex">Zero-based row index of the mismatching cell.</param>
        ''' <param name="columnIndex">Zero-based column index of the mismatching cell.</param>
        ''' <param name="expectedValue">Expected cell value.</param>
        ''' <param name="foundValue">Found cell value.</param>
        ''' <param name="foundValueIsAlreadySlightlyModified">Whether the found value was already slightly modified before this exception was created.</param>
        ''' <param name="table">Table snapshot appended to the exception message.</param>
        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer, expectedValue As String, foundValue As String, foundValueIsAlreadySlightlyModified As Boolean, table As TextTable)
            MyBase.New("Expected value """ & expectedValue & """ in cell " & CellAddress(targetSheetName, rowIndex, columnIndex) &
            If(foundValueIsAlreadySlightlyModified, ", but found ", ", but found (and corrected to) ") & """" & foundValue & """")
            Me.TargetSheetName = targetSheetName
            Me.CellRowIndex = rowIndex
            Me.CellColumnIndex = columnIndex
            Me.ExpectedValue = expectedValue
            Me.FoundValue = foundValue
            Me.Table = table
        End Sub

        ''' <summary>
        ''' Gets or sets the optional table snapshot appended to the exception message.
        ''' </summary>
        Public Property Table As TextTable
        ''' <summary>
        ''' Gets or sets the worksheet name containing the mismatch.
        ''' </summary>
        Public Property TargetSheetName As String
        ''' <summary>
        ''' Gets or sets the zero-based row index of the mismatching cell.
        ''' </summary>
        Public Property CellRowIndex As Integer
        ''' <summary>
        ''' Gets or sets the zero-based column index of the mismatching cell.
        ''' </summary>
        Public Property CellColumnIndex As Integer
        ''' <summary>
        ''' Gets or sets the expected cell value.
        ''' </summary>
        Public Property ExpectedValue As String
        ''' <summary>
        ''' Gets or sets the found cell value.
        ''' </summary>
        Public Property FoundValue As String

        Private Shared Function CellAddress(targetSheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Dim Cell As New ExcelCell(targetSheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
            Return Cell.Address(True)
        End Function

        ''' <inheritdoc/>
        Public Overrides ReadOnly Property Message As String
            Get
                If Me.Table IsNot Nothing Then
                    Return MyBase.Message & System.Environment.NewLine & Me.Table.ToUIExcelTable
                Else
                    Return MyBase.Message
                End If
            End Get
        End Property

    End Class
End Namespace
