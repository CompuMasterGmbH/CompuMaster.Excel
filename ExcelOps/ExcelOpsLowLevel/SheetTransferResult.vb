Namespace ExcelOps

    ''' <summary>Describes the destination of a copied or moved sheet.</summary>
    Public NotInheritable Class SheetTransferResult
        Friend Sub New(targetSheetName As String, targetSheetIndex As Integer)
            Me.TargetSheetName = targetSheetName
            Me.TargetSheetIndex = targetSheetIndex
        End Sub

        ''' <summary>Gets the name assigned to the sheet in the destination workbook.</summary>
        ''' <remarks>The name can differ from the source name when the destination already contains a sheet with that name.</remarks>
        Public ReadOnly Property TargetSheetName As String

        ''' <summary>Gets the zero-based index in the combined destination tab order of worksheets and chart sheets.</summary>
        ''' <remarks>
        ''' The index is the position after the transfer in <see cref="ExcelDataOperationsBase.SheetNames"/>.
        ''' For example, if a worksheet follows a chart sheet at index 1, the worksheet is at index 2.
        ''' This index uses the same numbering as <see cref="ExcelDataOperationsBase.SheetIndex"/>.
        ''' </remarks>
        Public ReadOnly Property TargetSheetIndex As Integer
    End Class

End Namespace
