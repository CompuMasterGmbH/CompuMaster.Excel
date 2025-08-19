Option Explicit On
Option Strict On
Imports System.Text

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

        ''' <summary>
        ''' Cells of these row indexes will be converted to TH cells instead of TD cells
        ''' </summary>
        ''' <returns></returns>
        Public Property ConsiderRowIndexesAsTableHeader As List(Of Integer)

        ''' <summary>
        ''' The class name for the table in HTML
        ''' </summary>
        ''' <returns></returns>
        Public Property TableCssClassName As String = "xlTable"

#Disable Warning CA1805 ' Keine unnötige Initialisierung
        Public Property FirstRowIndex As Integer = 0
        Public Property FirstColumnIndex As Integer = 0
        Public Property LastRowIndex As Integer?
        Public Property LastColumnIndex As Integer?
#Enable Warning CA1805 ' Keine unnötige Initialisierung

        Public Property HtmlForEmptySheet As String

        ''' <summary>
        ''' HTML code on top of everything (including html and head tags)
        ''' </summary>
        ''' <returns></returns>
        Public Property HtmlDocumentHeader As String

        ''' <summary>
        ''' HTML code on top of exported sheets (usually &lt;/head&gt;&lt;body&gt;)
        ''' </summary>
        ''' <returns></returns>
        Public Property HtmlDocumentHeaderEndAndBeginOfBody As String

        ''' <summary>
        ''' HTML code on bottom of everything (usually &lt;/body&gt;&lt;/html&gt;)
        ''' </summary>
        ''' <returns></returns>
        Public Property HtmlDocumentEnd As String

        ''' <summary>
        ''' Typically html+head tags
        ''' </summary>
        ''' <returns></returns>
        Protected Friend ReadOnly Property DefaultHtmlDocumentHeader As String =
            "<!doctype html><html><head>" & ControlChars.CrLf &
            "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & ControlChars.CrLf &
            "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">"

        ''' <summary>
        ''' Typically /head+body tags
        ''' </summary>
        ''' <returns></returns>
        Protected Friend ReadOnly Property DefaultHtmlDocumentHeaderEndAndBeginOfBody As String = "</head><body>"

        ''' <summary>
        ''' Typically /body+/html tags
        ''' </summary>
        ''' <returns></returns>
        Protected Friend ReadOnly Property DefaultHtmlDocumentEnd As String = "</body></html>"

        ''' <summary>
        ''' Typically some HTML indicating there was no content in Excel worksheet
        ''' </summary>
        ''' <returns></returns>
        Protected Friend ReadOnly Property DefaultHtmlForEmptySheet As String = "-/-"

        ''' <summary>
        ''' Typically some HTML indicating there was no content in Excel worksheet
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Function EffectiveHtmlForEmptySheet() As String
            If HtmlForEmptySheet Is Nothing Then
                Return DefaultHtmlForEmptySheet
            Else
                Return HtmlForEmptySheet
            End If
        End Function

        ''' <summary>
        ''' Typically /head+body tags
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Function EffectiveHtmlDocumentHeaderEndAndBeginOfBody() As String
            If HtmlDocumentHeaderEndAndBeginOfBody Is Nothing Then
                Return DefaultHtmlDocumentHeaderEndAndBeginOfBody
            Else
                Return HtmlDocumentHeaderEndAndBeginOfBody
            End If
        End Function

        ''' <summary>
        ''' Typically html+head+styles+body tags
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Function EffectiveHtmlDocumentHeaderAndBody() As String
            Dim sb As New System.Text.StringBuilder(1024)
            sb.AppendLine(Me.EffectiveHtmlDocumentHeader)
            sb.AppendLine(Me.EffectiveTableCssClassStyleHtml)
            sb.AppendLine(Me.EffectiveHtmlDocumentHeaderEndAndBeginOfBody)
            Return sb.ToString
        End Function

        ''' <summary>
        ''' Typically html+head tags
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Function EffectiveHtmlDocumentHeader() As String
            If HtmlDocumentHeader Is Nothing Then
                Return DefaultHtmlDocumentHeader
            Else
                Return HtmlDocumentHeader
            End If
        End Function

        ''' <summary>
        ''' Typically /body+/html tags
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Function EffectiveHtmlDocumentEnd() As String
            If HtmlDocumentEnd Is Nothing Then
                Return DefaultHtmlDocumentEnd
            Else
                Return HtmlDocumentEnd
            End If
        End Function

        ''' <summary>
        ''' Style HTML for the table (&lt;style&gt;...&lt;/style&gt;)
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Function EffectiveTableCssClassStyleHtml() As String
            Dim sb As New System.Text.StringBuilder(1024)
            sb.AppendLine("<style>")
            sb.AppendLine("table." & Me.TableCssClassName & " {border-collapse:collapse; border:1px solid #ccc; font-family:Segoe UI,Arial,sans-serif;}")
            sb.AppendLine("table." & Me.TableCssClassName & " th, table." & Me.TableCssClassName & " td {border:1px solid #ddd; padding:4px 6px; vertical-align:top;}")
            sb.AppendLine("table." & Me.TableCssClassName & " td {white-space:pre-wrap;}")
            sb.AppendLine("</style>")
            Return sb.ToString
        End Function

        ''' <summary>
        ''' Öffnenden Wrapper für ein Sheet ausgeben.
        ''' </summary>
        ''' <param name="anchorName"></param>
        ''' <param name="sb"></param>
        ''' <param name="visible">Sichtbar: True -> dieses Sheet ist initial sichtbar.</param>
        Public Overridable Sub GenerateBeginSheetSection(sb As StringBuilder, anchorName As String, Optional visible As Boolean = False)
            If String.IsNullOrWhiteSpace(anchorName) Then Throw New ArgumentNullException(NameOf(anchorName))
            Dim visibleClass As String = If(visible, " cm-visible", String.Empty)
            sb.Append("<section id=""").
               Append(System.Net.WebUtility.HtmlEncode(anchorName)).
               Append(""" class=""cm-wb-sheet").
               Append(visibleClass).
               Append(""" data-sheet=""").
               Append(System.Net.WebUtility.HtmlEncode(anchorName)).
               Append(""">")
            ' Kompatibel zu älteren Ankernamen
            sb.Append("<a name=""").Append(System.Net.WebUtility.HtmlEncode(anchorName)).Append("""></a>")
            sb.AppendLine()
        End Sub

        Public Property WorksheetTitleCssClassName As String = "cm-wb-sheet-title"

        ''' <summary>
        ''' Öffnenden Wrapper für ein Sheet ausgeben.
        ''' </summary>
        ''' <param name="sb"></param>
        Public Overridable Sub GenerateSheetSectionTitle(sb As StringBuilder, title As String)
            If Me.ExportSheetNameAsTitle = SheetTitleStyles.None Then Return
            Dim TagName As String = Me.ExportSheetNameAsTitle.ToString
            sb.Append("<"c).Append(TagName).
               Append(" class=""" & WorksheetTitleCssClassName & """").
               Append(""">").
               Append(System.Net.WebUtility.HtmlEncode(title)).
               Append("</").Append(TagName).Append(">"c)
            sb.AppendLine()
        End Sub

        ''' <summary>
        ''' Schließenden Wrapper für ein Sheet ausgeben.
        ''' </summary>
        Public Overridable Sub GenerateEndSheetSection(sb As StringBuilder)
            sb.Append("</section>").AppendLine.AppendLine()
        End Sub

    End Class

End Namespace