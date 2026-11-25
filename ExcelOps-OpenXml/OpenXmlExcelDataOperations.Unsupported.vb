Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Function UnsupportedOperation(operation As String) As NotSupportedException
            Return New NotSupportedException(operation & " is not supported by the " & EngineName & " engine.")
        End Function

        ''' <inheritdoc/>
        Public Overrides Sub CleanupRangeNames()
            Throw UnsupportedOperation("Cleaning up range names")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub RemoveSheet(sheetName As String)
            Throw UnsupportedOperation("Removing sheets")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub AddSheet(sheetName As String, beforeSheetName As String)
            Throw UnsupportedOperation("Adding sheets")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub ClearCells(overrideSheetName As String, rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)
            Throw UnsupportedOperation("Clearing cell ranges")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub ClearSheet(sheetName As String)
            Throw UnsupportedOperation("Clearing sheets")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub SelectSheet(sheetName As String)
            Throw UnsupportedOperation("Selecting sheets")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub SelectSheet(sheetIndex As Integer)
            Throw UnsupportedOperation("Selecting sheets")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub CopySheetContentInternal(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)
            Throw UnsupportedOperation("Copying sheet content")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub SelectCell(cell As ExcelCell)
            Throw UnsupportedOperation("Selecting cells")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub UnMergeCells(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            Throw UnsupportedOperation("Unmerging cells")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub MergeCells(sheetName As String, fromRowIndex As Integer, fromColumnIndex As Integer, toRowIndex As Integer, toColumnIndex As Integer)
            Throw UnsupportedOperation("Merging cells")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub AutoFitColumns(sheetName As String)
            Throw UnsupportedOperation("Automatically sizing columns")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub AutoFitColumns(sheetName As String, minimumWidth As Double)
            Throw UnsupportedOperation("Automatically sizing columns")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub AutoFitColumns(sheetName As String, columnIndex As Integer)
            Throw UnsupportedOperation("Automatically sizing columns")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub AutoFitColumns(sheetName As String, columnIndex As Integer, minimumWidth As Double)
            Throw UnsupportedOperation("Automatically sizing columns")
        End Sub

        ''' <inheritdoc/>
        Public Overrides Function ExportChartSheetImage(chartSheetName As String) As System.Drawing.Image
            Throw UnsupportedOperation("Rendering chart sheets")
        End Function

        ''' <inheritdoc/>
        Public Overrides Function ExportChartImage(workSheetName As String) As System.Drawing.Image()
            Throw UnsupportedOperation("Rendering chart images")
        End Function

    End Class

End Namespace
