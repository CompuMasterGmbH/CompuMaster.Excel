Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Function UnsupportedOperation(operation As String) As NotSupportedException
            Return New NotSupportedException(operation & " is not supported by the " & EngineName & " engine.")
        End Function

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
