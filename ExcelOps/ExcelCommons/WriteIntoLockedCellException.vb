Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data doesn't match with expected data
    ''' </summary>
    Public Class WriteIntoLockedCellException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer)
            MyBase.New("Writing into locked cell " & CellAddress(targetSheetName, rowIndex, columnIndex) & " not allowed")
            Me.TargetSheetName = targetSheetName
            Me.CellRowIndex = rowIndex
            Me.CellColumnIndex = columnIndex
            Me.TemplateVersion = -1
        End Sub

        Public Sub New(targetSheetName As String, rowIndex As Integer, columnIndex As Integer, templateVersion As Integer)
            MyBase.New("Writing into locked cell " & CellAddress(targetSheetName, rowIndex, columnIndex) & " not allowed for template V" & templateVersion)
            Me.TargetSheetName = targetSheetName
            Me.CellRowIndex = rowIndex
            Me.CellColumnIndex = columnIndex
            Me.TemplateVersion = templateVersion
        End Sub

        Public Property TargetSheetName As String
        Public Property CellRowIndex As Integer
        Public Property CellColumnIndex As Integer
        Public Property TemplateVersion As Integer

        Private Shared Function CellAddress(targetSheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Dim Cell As New ExcelCell(targetSheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
            Return Cell.Address(True)
        End Function

    End Class

End Namespace