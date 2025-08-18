Option Explicit On
Option Strict On

Namespace ExcelOps

    Public Class HtmlSheetExportOptions

        Public Sub New()
        End Sub

        Public Enum SheetTitleStyles As Integer
            None = 0
            H1 = 1
            H2 = 2
            H3 = 3
            H4 = 4
            H5 = 5
            H6 = 6
            P = 10
        End Enum

        ''' <summary>
        ''' Add a header before the sheet data
        ''' </summary>
        ''' <returns></returns>
        Public Property ExportSheetNameAsTitle As SheetTitleStyles

    End Class

End Namespace