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

        Public Sub New(cellAddress As ExcelCell, expectedValue As String, foundValue As String)
            Me.New(cellAddress.SheetName, cellAddress.RowIndex, cellAddress.ColumnIndex, expectedValue, foundValue, False)
        End Sub

        Public Sub New(cellAddress As ExcelCell, expectedValue As String, foundValue As String, foundValueIsAlreadySlightlyModified As Boolean)
            Me.New(cellAddress.SheetName, cellAddress.RowIndex, cellAddress.ColumnIndex, expectedValue, foundValue, foundValueIsAlreadySlightlyModified)
        End Sub

        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer, expectedValue As String, foundValue As String)
            Me.New(targetSheetName, rowIndex, columnIndex, expectedValue, foundValue, False)
        End Sub

        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer, expectedValue As String, foundValue As String, foundValueIsAlreadySlightlyModified As Boolean)
            MyBase.New("Expected value """ & expectedValue & """ in cell " & CellAddress(targetSheetName, rowIndex, columnIndex) &
          If(foundValueIsAlreadySlightlyModified, ", but found ", ", but found (and corrected to) ") & """" & foundValue & """")
            Me.TargetSheetName = targetSheetName
            Me.CellRowIndex = rowIndex
            Me.CellColumnIndex = columnIndex
            Me.ExpectedValue = expectedValue
            Me.FoundValue = foundValue
        End Sub

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

        Public Property Table As TextTable
        Public Property TargetSheetName As String
        Public Property CellRowIndex As Integer
        Public Property CellColumnIndex As Integer
        Public Property ExpectedValue As String
        Public Property FoundValue As String

        Private Shared Function CellAddress(targetSheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Dim Cell As New ExcelCell(targetSheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
            Return Cell.Address(True)
        End Function

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