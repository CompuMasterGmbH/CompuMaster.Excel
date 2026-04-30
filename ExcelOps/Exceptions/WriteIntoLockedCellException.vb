Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data doesn't match with expected data.
    ''' </summary>
    Public Class WriteIntoLockedCellException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        ''' <summary>
        ''' Creates an exception for a write attempt into a locked cell.
        ''' </summary>
        ''' <param name="targetSheetName">Name of the worksheet containing the locked cell.</param>
        ''' <param name="rowIndex">Zero-based row index of the locked cell.</param>
        ''' <param name="columnIndex">Zero-based column index of the locked cell.</param>
        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer)
            MyBase.New("Writing into locked cell " & CellAddress(targetSheetName, rowIndex, columnIndex) & " not allowed")
            Me.TargetSheetName = targetSheetName
            Me.CellRowIndex = rowIndex
            Me.CellColumnIndex = columnIndex
            Me.TemplateVersion = -1
        End Sub

        ''' <summary>
        ''' Creates an exception for a write attempt into a locked cell.
        ''' </summary>
        ''' <param name="targetSheetName">Name of the worksheet containing the locked cell.</param>
        ''' <param name="rowIndex">Zero-based row index of the locked cell.</param>
        ''' <param name="columnIndex">Zero-based column index of the locked cell.</param>
        ''' <param name="templateVersion">Template version that caused the cell to be treated as locked.</param>
        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer, templateVersion As Integer)
            MyBase.New("Writing into locked cell " & CellAddress(targetSheetName, rowIndex, columnIndex) & " not allowed for template V" & templateVersion)
            Me.TargetSheetName = targetSheetName
            Me.CellRowIndex = rowIndex
            Me.CellColumnIndex = columnIndex
            Me.TemplateVersion = templateVersion
        End Sub

        ''' <summary>
        ''' Gets or sets the worksheet name containing the locked cell.
        ''' </summary>
        Public Property TargetSheetName As String
        ''' <summary>
        ''' Gets or sets the zero-based row index of the locked cell.
        ''' </summary>
        Public Property CellRowIndex As Integer
        ''' <summary>
        ''' Gets or sets the zero-based column index of the locked cell.
        ''' </summary>
        Public Property CellColumnIndex As Integer
        ''' <summary>
        ''' Gets or sets the template version that caused the cell to be treated as locked.
        ''' </summary>
        Public Property TemplateVersion As Integer

        Private Shared Function CellAddress(targetSheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Dim Cell As New ExcelCell(targetSheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
            Return Cell.Address(True)
        End Function

    End Class

End Namespace
