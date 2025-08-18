Imports CompuMaster.Epplus4
Imports CompuMaster.Epplus4.Style
Imports System.Net
Imports System.Text

Friend Module EpplusHtmlRenderer

    ''' <summary>
    ''' Rendert ein einzelnes Worksheet als vollständiges HTML-Dokument (UTF-8).
    ''' Kompatibel zu EPPlus 4.5.x (kein Lizenz-Setup nötig).
    ''' </summary>
    Public Function XlsxToHtmlString(xlsxPath As String,
                                     Optional sheetIndex As Integer = 1,
                                     Optional tableCssClass As String = "xl",
                                     Optional treatFirstRowAsHeader As Boolean = False) As String
        Using pkg As New ExcelPackage(New IO.FileInfo(xlsxPath))
            Return XlsxToHtmlString(pkg, sheetIndex, tableCssClass, treatFirstRowAsHeader)
        End Using
    End Function

    ''' <summary>
    ''' Rendert ein einzelnes Worksheet als vollständiges HTML-Dokument (UTF-8).
    ''' Kompatibel zu EPPlus 4.5.x (kein Lizenz-Setup nötig).
    ''' </summary>
    Public Function XlsxToHtmlString(pkg As ExcelPackage,
                                     Optional sheetIndex As Integer = 1,
                                     Optional tableCssClass As String = "xl",
                                     Optional treatFirstRowAsHeader As Boolean = False) As String
        Dim ws = pkg.Workbook.Worksheets(sheetIndex)
        Return XlsxToHtmlString(ws, tableCssClass, treatFirstRowAsHeader)
    End Function

    ''' <summary>
    ''' Rendert ein einzelnes Worksheet als vollständiges HTML-Dokument (UTF-8).
    ''' Kompatibel zu EPPlus 4.5.x (kein Lizenz-Setup nötig).
    ''' </summary>
    Public Function XlsxToHtmlString(ws As ExcelWorksheet,
                                     Optional tableCssClass As String = "xl",
                                     Optional treatFirstRowAsHeader As Boolean = False) As String

        If ws Is Nothing OrElse ws.Dimension Is Nothing Then
            Return "<!doctype html><html><meta charset=""utf-8""><body>(kein Inhalt)</body></html>"
        End If

        Dim firstRow = ws.Dimension.Start.Row
        Dim lastRow = ws.Dimension.End.Row
        Dim firstCol = ws.Dimension.Start.Column
        Dim lastCol = ws.Dimension.End.Column

        ' --- Merge-Map vorbereiten ---
        Dim mergeTopLeft As New Dictionary(Of String, (RowSpan As Integer, ColSpan As Integer))()
        Dim covered As New HashSet(Of String)()
        For Each addr In ws.MergedCells
            Dim range = ws.Cells(addr)
            Dim r1 = range.Start.Row, c1 = range.Start.Column
            Dim r2 = range.End.Row, c2 = range.End.Column
            Dim keyTL = Key(r1, c1)
            mergeTopLeft(keyTL) = ((r2 - r1 + 1), (c2 - c1 + 1))
            For rr = r1 To r2
                For cc = c1 To c2
                    If Not (rr = r1 AndAlso cc = c1) Then covered.Add(Key(rr, cc))
                Next
            Next
        Next

        Dim sb As New StringBuilder(128 * 1024)

        ' --- Minimal-CSS (neutral, dunkelgraue Rahmen) ---
        sb.AppendLine("<!doctype html>")
        sb.AppendLine("<html><head><meta charset=""utf-8"">")
        sb.AppendLine("<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">")
        sb.AppendLine("<meta name=""viewport"" content=""width=device-width,initial-scale=1"">")
        sb.AppendLine("<style>")
        sb.AppendLine("table." & tableCssClass & " {border-collapse:collapse; border:1px solid #ccc; font-family:Segoe UI,Arial,sans-serif; font-size:14px;}")
        sb.AppendLine("table." & tableCssClass & " th, table." & tableCssClass & " td {border:1px solid #ddd; padding:4px 6px; vertical-align:top;}")
        sb.AppendLine("table." & tableCssClass & " td {white-space:pre-wrap;}")
        sb.AppendLine("</style></head><body>")

        ' Optional: Blattname als Überschrift
        sb.AppendLine("<h3>" & Html(ws.Name) & "</h3>")

        sb.Append("<table class=""" & tableCssClass & """>")

        For r = firstRow To lastRow
            sb.Append("<tr>")
            For c = firstCol To lastCol
                Dim k = Key(r, c)
                If covered.Contains(k) Then
                    ' Teil eines Merge-Bereichs, aber nicht Top-Left => überspringen
                    Continue For
                End If

                Dim cell = ws.Cells(r, c)
                Dim tag As String = If(treatFirstRowAsHeader AndAlso r = firstRow, "th", "td")

                ' --- Merge-Attribute ---
                Dim rowspan As Integer = 1, colspan As Integer = 1
                If mergeTopLeft.ContainsKey(k) Then
                    rowspan = mergeTopLeft(k).RowSpan
                    colspan = mergeTopLeft(k).ColSpan
                End If

                ' --- Styles extrahieren ---
                Dim styles As New List(Of String)

                ' Horizontal-Align
                Select Case cell.Style.HorizontalAlignment
                    Case ExcelHorizontalAlignment.Center : styles.Add("text-align:center")
                    Case ExcelHorizontalAlignment.Right : styles.Add("text-align:right")
                    Case ExcelHorizontalAlignment.Justify : styles.Add("text-align:justify")
                    Case Else : styles.Add("text-align:left")
                End Select

                ' Font
                If cell.Style.Font IsNot Nothing Then
                    If cell.Style.Font.Bold Then styles.Add("font-weight:bold")
                    If cell.Style.Font.Italic Then styles.Add("font-style:italic")
                    If cell.Style.Font.UnderLine Then styles.Add("text-decoration:underline")
                    Dim fc = SafeRgb(cell.Style.Font.Color)
                    If fc IsNot Nothing Then styles.Add("color:" & fc)
                End If

                ' Hintergrund
                If cell.Style.Fill IsNot Nothing AndAlso cell.Style.Fill.PatternType <> ExcelFillStyle.None Then
                    Dim bg = SafeRgb(cell.Style.Fill.BackgroundColor)
                    If bg IsNot Nothing Then styles.Add("background-color:" & bg)
                End If

                ' Zeilenumbruch
                If cell.Style.WrapText Then styles.Add("white-space:pre-wrap")

                Dim styleAttr As String = If(styles.Count > 0, $" style=""{String.Join(";", styles)}""", "")

                ' Inhalt (Text = bereits nach NumberFormat formatiert)
                Dim content As String = cell.Text
                If String.IsNullOrEmpty(content) Then content = " "
                content = Html(content)

                ' Zelle schreiben
                sb.Append("<" & tag)
                If rowspan > 1 Then sb.Append(" rowspan=""" & rowspan & """")
                If colspan > 1 Then sb.Append(" colspan=""" & colspan & """")
                sb.Append(styleAttr)
                sb.Append(">"c)
                sb.Append(content)
                sb.Append("</" & tag & ">")
            Next
            sb.AppendLine("</tr>")
        Next

        sb.AppendLine("</table></body></html>")
        Return sb.ToString()
    End Function

    ' ---------- Helpers ----------

    Private Function Key(r As Integer, c As Integer) As String
        Return r.ToString() & "|" & c.ToString()
    End Function

    Private Function Html(s As String) As String
        Return WebUtility.HtmlEncode(s)
    End Function

    ''' <summary>
    ''' Liefert ein CSS-#RRGGBB aus einer EPPlus-Farbquelle (ARGB → RRGGBB).
    ''' </summary>
    Private Function SafeRgb(color As CompuMaster.Epplus4.Style.ExcelColor) As String
        If color Is Nothing Then Return Nothing
        ' EPPlus 4.5 nutzt bevorzugt .Rgb (8-stellig ARGB). Theme/Indexed werden hier ausgelassen.
        If Not String.IsNullOrEmpty(color.Rgb) AndAlso color.Rgb.Length = 8 Then
            Dim rgb = "#" & color.Rgb.Substring(2) ' ARGB → RRGGBB
            Return rgb
        End If
        Return Nothing
    End Function

End Module