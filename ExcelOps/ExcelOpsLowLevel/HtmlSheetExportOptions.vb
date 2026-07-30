Option Explicit On
Option Strict On
Imports System.Text

Namespace ExcelOps

    ''' <summary>
    ''' Defines options for exporting a worksheet to HTML.
    ''' </summary>
    Public Class HtmlSheetExportOptions

        ''' <summary>
        ''' Creates a new HTML worksheet export options instance.
        ''' </summary>
        Public Sub New()
        End Sub

        ''' <summary>
        ''' Defines how a worksheet title is rendered before exported worksheet data.
        ''' </summary>
        Public Enum SheetTitleStyles As Integer
            ''' <summary>
            ''' Does not render a worksheet title.
            ''' </summary>
            None = 0
            ''' <summary>
            ''' Renders the worksheet title as an H1 element.
            ''' </summary>
            H1 = 1
            ''' <summary>
            ''' Renders the worksheet title as an H2 element.
            ''' </summary>
            H2 = 2
            ''' <summary>
            ''' Renders the worksheet title as an H3 element.
            ''' </summary>
            H3 = 3
            ''' <summary>
            ''' Renders the worksheet title as an H4 element.
            ''' </summary>
            H4 = 4
            ''' <summary>
            ''' Renders the worksheet title as an H5 element.
            ''' </summary>
            H5 = 5
            ''' <summary>
            ''' Renders the worksheet title as an H6 element.
            ''' </summary>
            H6 = 6
            ''' <summary>
            ''' Renders the worksheet title as a paragraph element.
            ''' </summary>
            P = 10
        End Enum

        ''' <summary>
        ''' Gets or sets how the worksheet name is rendered before worksheet data.
        ''' </summary>
        Public Property ExportSheetNameAsTitle As SheetTitleStyles

        ''' <summary>
        ''' Gets or sets row indexes whose cells are rendered as TH elements instead of TD elements.
        ''' </summary>
        Public Property ConsiderRowIndexesAsTableHeader As List(Of Integer)

        ''' <summary>
        ''' Gets or sets the CSS class name used for generated worksheet tables.
        ''' </summary>
        Public Property TableCssClassName As String = "xlTable"

#Disable Warning CA1805 ' Keine unnötige Initialisierung
        ''' <summary>
        ''' Gets or sets the zero-based index of the first worksheet row to export.
        ''' </summary>
        Public Property FirstRowIndex As Integer = 0
        ''' <summary>
        ''' Gets or sets the zero-based index of the first worksheet column to export.
        ''' </summary>
        Public Property FirstColumnIndex As Integer = 0
        ''' <summary>
        ''' Gets or sets the zero-based index of the last worksheet row to export.
        ''' </summary>
        Public Property LastRowIndex As Integer?
        ''' <summary>
        ''' Gets or sets the zero-based index of the last worksheet column to export.
        ''' </summary>
        Public Property LastColumnIndex As Integer?
#Enable Warning CA1805 ' Keine unnötige Initialisierung

        ''' <summary>
        ''' Gets or sets the HTML emitted when a worksheet has no exportable content.
        ''' </summary>
        Public Property HtmlForEmptySheet As String

        ''' <summary>
        ''' Gets or sets the HTML emitted before generated worksheet content, including HTML and HEAD tags.
        ''' </summary>
        Public Property HtmlDocumentHeader As String

        ''' <summary>
        ''' Gets or sets the HTML emitted between the document header and exported worksheets.
        ''' </summary>
        Public Property HtmlDocumentHeaderEndAndBeginOfBody As String

        ''' <summary>
        ''' Gets or sets the HTML emitted after generated worksheet content.
        ''' </summary>
        Public Property HtmlDocumentEnd As String

        ''' <summary>
        ''' Gets the default HTML document header.
        ''' </summary>
        Protected Friend ReadOnly Property DefaultHtmlDocumentHeader As String =
            "<!doctype html><html><head>" & ControlChars.CrLf &
            "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & ControlChars.CrLf &
            "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">"

        ''' <summary>
        ''' Gets the default HTML fragment that closes HEAD and opens BODY.
        ''' </summary>
        Protected Friend ReadOnly Property DefaultHtmlDocumentHeaderEndAndBeginOfBody As String = "</head><body>"

        ''' <summary>
        ''' Gets the default HTML document ending.
        ''' </summary>
        Protected Friend ReadOnly Property DefaultHtmlDocumentEnd As String = "</body></html>"

        ''' <summary>
        ''' Gets the default HTML emitted for an empty worksheet.
        ''' </summary>
        Protected Friend ReadOnly Property DefaultHtmlForEmptySheet As String = "-/-"

        ''' <summary>
        ''' Gets the effective HTML emitted for an empty worksheet.
        ''' </summary>
        ''' <returns>Configured empty-sheet HTML, or the default empty-sheet HTML when no value is configured.</returns>
        Public Overridable Function EffectiveHtmlForEmptySheet() As String
            If HtmlForEmptySheet Is Nothing Then
                Return DefaultHtmlForEmptySheet
            Else
                Return HtmlForEmptySheet
            End If
        End Function

        ''' <summary>
        ''' Gets the effective HTML fragment that closes HEAD and opens BODY.
        ''' </summary>
        ''' <returns>Configured header/body transition HTML, or the default transition HTML when no value is configured.</returns>
        Public Overridable Function EffectiveHtmlDocumentHeaderEndAndBeginOfBody() As String
            If HtmlDocumentHeaderEndAndBeginOfBody Is Nothing Then
                Return DefaultHtmlDocumentHeaderEndAndBeginOfBody
            Else
                Return HtmlDocumentHeaderEndAndBeginOfBody
            End If
        End Function

        ''' <summary>
        ''' Gets the effective HTML document header including styles and BODY start.
        ''' </summary>
        ''' <returns>HTML emitted before worksheet content.</returns>
        Public Overridable Function EffectiveHtmlDocumentHeaderAndBody() As String
            Dim sb As New System.Text.StringBuilder(1024)
            sb.AppendLine(Me.EffectiveHtmlDocumentHeader)
            sb.AppendLine(Me.EffectiveTableCssClassStyleHtml)
            sb.AppendLine(Me.EffectiveHtmlDocumentHeaderEndAndBeginOfBody)
            Return sb.ToString
        End Function

        ''' <summary>
        ''' Gets the effective HTML document header.
        ''' </summary>
        ''' <returns>Configured document header HTML, or the default document header HTML when no value is configured.</returns>
        Public Overridable Function EffectiveHtmlDocumentHeader() As String
            If HtmlDocumentHeader Is Nothing Then
                Return DefaultHtmlDocumentHeader
            Else
                Return HtmlDocumentHeader
            End If
        End Function

        ''' <summary>
        ''' Gets the effective HTML document ending.
        ''' </summary>
        ''' <returns>Configured document end HTML, or the default document end HTML when no value is configured.</returns>
        Public Overridable Function EffectiveHtmlDocumentEnd() As String
            If HtmlDocumentEnd Is Nothing Then
                Return DefaultHtmlDocumentEnd
            Else
                Return HtmlDocumentEnd
            End If
        End Function

        ''' <summary>
        ''' Gets the effective STYLE element for generated worksheet tables.
        ''' </summary>
        ''' <returns>HTML STYLE element for generated worksheet tables.</returns>
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
        ''' Generates the opening HTML section wrapper for a worksheet.
        ''' </summary>
        ''' <param name="sb">String builder receiving generated HTML.</param>
        ''' <param name="anchorName">Section ID and anchor name for the worksheet.</param>
        ''' <param name="visible"><see langword="True"/> to mark the worksheet section as initially visible; otherwise, <see langword="False"/>.</param>
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

        ''' <summary>
        ''' Gets or sets the CSS class name used for generated worksheet titles.
        ''' </summary>
        Public Property WorksheetTitleCssClassName As String = "cm-wb-sheet-title"

        ''' <summary>
        ''' Generates the worksheet title HTML.
        ''' </summary>
        ''' <param name="sb">String builder receiving generated HTML.</param>
        ''' <param name="title">Worksheet title to render.</param>
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
        ''' Generates the closing HTML section wrapper for a worksheet.
        ''' </summary>
        ''' <param name="sb">String builder receiving generated HTML.</param>
        Public Overridable Sub GenerateEndSheetSection(sb As StringBuilder)
            sb.Append("</section>").AppendLine.AppendLine()
        End Sub

    End Class

End Namespace
