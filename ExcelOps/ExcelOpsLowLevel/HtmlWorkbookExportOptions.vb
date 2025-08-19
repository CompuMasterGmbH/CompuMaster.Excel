Option Explicit On
Option Strict On

Imports System.Globalization
Imports System.Text

Namespace ExcelOps

#Disable Warning CA1805 ' Keine unnötige Initialisierung
    Public Class HtmlWorkbookExportOptions
        Inherits HtmlSheetExportOptions

        Public Sub New()
            MyBase.New
        End Sub

        Public Enum SheetNavigationPositions As Byte
            ''' <summary>
            ''' No code for navigation embedded (must be added afterwards)
            ''' </summary>
            None = 0
            ''' <summary>
            ''' Code for navigation at top
            ''' </summary>
            Top = 1
            ''' <summary>
            ''' Code for navigation as bottom
            ''' </summary>
            Bottom = 2
        End Enum

        ''' <summary>
        ''' Display style of navigation
        ''' </summary>
        ''' <returns></returns>
        Public Property SheetNavigationPosition As SheetNavigationPositions = SheetNavigationPositions.Top

        Public Enum SheetNavigationActionStyles As Byte
            ''' <summary>
            ''' All sheets are always visible (ignores selected sheet of Excel workbook)
            ''' </summary>
            JumpToAnchor = 0
            ''' <summary>
            ''' Only the selected sheet is visible
            ''' </summary>
            SwitchVisibleSheet = 1
        End Enum

        ''' <summary>
        ''' Display style of navigation
        ''' </summary>
        ''' <returns></returns>
        Public Property SheetNavigationActionStyle As SheetNavigationActionStyles = SheetNavigationActionStyles.JumpToAnchor

        ''' <summary>
        ''' Always show navigation while scrolling
        ''' </summary>
        ''' <returns></returns>
        Public Property SheetNavigationAlwaysVisible As Boolean = False

        ''' <summary>
        ''' Export all sheets or visible sheets only
        ''' </summary>
        ''' <returns></returns>
        Public Property ExportHiddenSheets As Boolean = False

        ''' <summary>
        ''' Export chart sheets
        ''' </summary>
        ''' <returns></returns>
        Public Property ExportChartSheets As Boolean = False

        ''' <summary>
        ''' Export work sheets
        ''' </summary>
        ''' <returns></returns>
        Public Property ExportWorkSheets As Boolean = True


        ''' <summary>
        ''' The CSS class name for the worksheets navigation
        ''' </summary>
        ''' <returns></returns>
        Public Property WorksheetsNavigationCssClassName As String = "cm-wb-subnav"

        ''' <summary>
        ''' Style HTML for the table (&lt;style&gt;...&lt;/style&gt;)
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Function EffectiveWorksheetsNavigationCssStyleHtml() As String
            Dim sb As New System.Text.StringBuilder(1024)
            sb.AppendLine("<style>")
            sb.AppendLine("." & Me.WorksheetsNavigationCssClassName & " {display:flex;align-items:center;gap:.5rem;flex-wrap:wrap;padding:.4rem .8rem;background:#f8f9fa;border-bottom:1px solid #ddd;z-index:1000;font-family:Segoe UI,Arial,sans-serif;font-size:1.25rem;}")
            sb.AppendLine("." & Me.WorksheetsNavigationCssClassName & ".sticky-top{position:sticky;top:0;}")
            sb.AppendLine("." & Me.WorksheetsNavigationCssClassName & ".sticky-bottom{position:sticky;bottom:0;border-top:1px solid #ddd;border-bottom:none;}")
            sb.AppendLine("." & Me.WorksheetsNavigationCssClassName & " ul{list-style:none;margin:0;padding:0;display:flex;gap:.4rem;flex-wrap:wrap;}")
            sb.AppendLine("." & Me.WorksheetsNavigationCssClassName & " a{text-decoration:none;padding:.28rem .5rem;border-radius:.4rem;}")
            sb.AppendLine("." & Me.WorksheetsNavigationCssClassName & " a:hover,." & Me.WorksheetsNavigationCssClassName & " a:focus{background:rgba(0,0,0,.06);outline:none;}")
            sb.AppendLine("</style>")
            If Me.SheetNavigationActionStyle = SheetNavigationActionStyles.SwitchVisibleSheet Then
                sb.AppendLine("<style>")
                sb.AppendLine(".cm-wb-sheet{display:none;}")
                sb.AppendLine(".cm-wb-sheet.cm-visible{display:block;}")
                sb.AppendLine("</style>")
                sb.AppendLine("<noscript><style>.cm-wb-sheet{display:block !important;}</style></noscript>")
            End If
            Return sb.ToString
        End Function

        ''' <summary>
        ''' Typically html+head+styles+body tags
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function EffectiveHtmlDocumentHeaderAndBody() As String
            Dim sb As New System.Text.StringBuilder(1024)
            sb.AppendLine(Me.EffectiveHtmlDocumentHeader)
            sb.AppendLine(Me.EffectiveTableCssClassStyleHtml)
            sb.AppendLine(Me.EffectiveWorksheetsNavigationCssStyleHtml)
            sb.AppendLine(Me.EffectiveHtmlDocumentHeaderEndAndBeginOfBody)
            Return sb.ToString
        End Function

        ''' <summary>
        ''' The ID attribute value for the NAV tag
        ''' </summary>
        ''' <returns></returns>
        Public Property WorksheetsNavigationTagId As String = "workbook-nav"

        ''' <summary>
        ''' Generate the HTML code for a worksheet navigation
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="items">Labels for the several worksheets (typically worksheet names)</param>
        ''' <param name="options"></param>
        Public Overridable Sub GenerateWorkbookSubNavigation(sb As StringBuilder,
                                             items As IEnumerable(Of String),
                                             options As HtmlWorkbookExportOptions)
            GenerateWorkbookSubNavigation(sb, items, Nothing, options)
        End Sub

        ''' <summary>
        ''' Generate the HTML code for a worksheet navigation
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="navigationItemTitles">Labels for the several worksheets (typically worksheet names)</param>
        ''' <param name="options"></param>
        ''' <param name="anchorNames">Section/anchor names for the several worksheets</param>
        Public Overridable Sub GenerateWorkbookSubNavigation(sb As StringBuilder,
                                             navigationItemTitles As IEnumerable(Of String),
                                             anchorNames As IEnumerable(Of String),
                                             options As HtmlWorkbookExportOptions
                                             )

            If sb Is Nothing Then Throw New ArgumentNullException(NameOf(sb))
            If navigationItemTitles Is Nothing Then Throw New ArgumentNullException(NameOf(navigationItemTitles))
            If options Is Nothing Then Throw New ArgumentNullException(NameOf(options))

            Dim labels = navigationItemTitles.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToList()
            If labels.Count = 0 Then Return

            Dim anchors As List(Of String)
            If anchorNames IsNot Nothing Then
                anchors = anchorNames.ToList()
                If anchors.Count <> labels.Count Then
                    Throw New ArgumentException("Anzahl von anchorNames muss der Anzahl von items entsprechen.")
                End If
            Else
                anchors = CreateUniqueAnchorNames(labels)
            End If

            Dim stickyClass As String = String.Empty
            Dim restoreMarginTop As String = String.Empty
            If options.SheetNavigationAlwaysVisible Then
                Select Case options.SheetNavigationPosition
                    Case HtmlWorkbookExportOptions.SheetNavigationPositions.Top
                        stickyClass = " sticky-top"
                        restoreMarginTop = "<br /><br />"
                    Case HtmlWorkbookExportOptions.SheetNavigationPositions.Bottom
                        stickyClass = " sticky-bottom"
                        restoreMarginTop = "<br /><br />"
                End Select
            End If

            sb.Append("<nav")
            If Not String.IsNullOrWhiteSpace(options.WorksheetsNavigationTagId) Then
                sb.Append(" id=""").Append(System.Net.WebUtility.HtmlEncode(options.WorksheetsNavigationTagId)).Append(""""c)
            End If
            sb.Append(" class=""")
            sb.Append(System.Net.WebUtility.HtmlEncode(options.WorksheetsNavigationCssClassName))
            sb.Append(System.Net.WebUtility.HtmlEncode(stickyClass))
            sb.Append(""" role=""navigation"">")

            sb.Append("<ul>")
            For i = 0 To labels.Count - 1
                Dim label = labels(i)
                Dim anchor = anchors(i)

                Select Case options.SheetNavigationActionStyle
                    Case HtmlWorkbookExportOptions.SheetNavigationActionStyles.SwitchVisibleSheet
                        ' Link mit data-target -> JS zeigt entsprechendes Sheet
                        sb.Append("<li><a href=""#").
                           Append(System.Net.WebUtility.HtmlEncode(anchor)).
                           Append(""" data-target=""").
                           Append(System.Net.WebUtility.HtmlEncode(anchor)).
                           Append(""">").
                           Append(System.Net.WebUtility.HtmlEncode(label)).
                           Append("</a></li>")
                    Case HtmlWorkbookExportOptions.SheetNavigationActionStyles.JumpToAnchor
                        ' Klassischer Anker-Sprung
                        sb.Append("<li><a href=""#").
                           Append(System.Net.WebUtility.HtmlEncode(anchor)).
                           Append(""">").
                           Append(System.Net.WebUtility.HtmlEncode(label)).
                           Append("</a></li>")
                    Case Else
                        Throw New NotImplementedException(options.SheetNavigationActionStyle.ToString)
                End Select
            Next
            sb.Append("</ul></nav>")
            sb.AppendLine(restoreMarginTop)
            sb.AppendLine()
            sb.AppendLine()

            ' Bei SwitchVisibleSheet: Sheet-CSS + JS einbetten (einmalig pro Seite ist ausreichend)
            If options.SheetNavigationActionStyle = HtmlWorkbookExportOptions.SheetNavigationActionStyles.SwitchVisibleSheet Then
                'Styles/Script für SwitchVisibleSheet-Modus emitten (einmal pro Seite)
                sb.AppendLine("<script>(function(){")
                sb.AppendLine("function showSheet(id){var s=document.querySelectorAll('.cm-wb-sheet');for(var i=0;i<s.length;i++){s[i].classList.remove('cm-visible');}var el=document.getElementById(id);if(el){el.classList.add('cm-visible');}}")
                sb.AppendLine("function handleHash(){var id=location.hash?location.hash.substring(1):null;if(id){showSheet(id);}}")
                sb.AppendLine("document.addEventListener('click',function(e){var a=e.target.closest('." & WorksheetsNavigationCssClassName & " a[data-target]');if(a){e.preventDefault();var id=a.getAttribute('data-target');history.replaceState(null,'','#'+id);showSheet(id);}});")
                sb.AppendLine("window.addEventListener('hashchange',handleHash);handleHash();})();</script>")
            End If
        End Sub

        ' "Über uns" -> "uber-uns"
        Friend Shared Function Slugify(input As String) As String
            Dim normalized As String = input.Normalize(NormalizationForm.FormD)
            Dim b As New StringBuilder(normalized.Length)
            For Each ch As Char In normalized
                If CharUnicodeInfo.GetUnicodeCategory(ch) <> UnicodeCategory.NonSpacingMark Then b.Append(ch)
            Next
            Dim noDiacritics As String = b.ToString().Normalize(NormalizationForm.FormC)
            Dim lowered As String = noDiacritics.ToLowerInvariant().Replace("ß", "ss")
            Dim replaced As String = System.Text.RegularExpressions.Regex.Replace(lowered, "[^\w\-]+", "-")   ' Nicht-Wortzeichen -> "-"
            replaced = System.Text.RegularExpressions.Regex.Replace(replaced, "[-_]+", "-")                    ' Mehrfachtrennstriche verdichten
            Return replaced.Trim("-"c, "_"c)
        End Function

        Private Shared Function MakeUnique(baseName As String, used As HashSet(Of String)) As String
            Dim candidate = baseName
            Dim i = 2
            While used.Contains(candidate)
                candidate = baseName & "-" & i.ToString(CultureInfo.InvariantCulture)
                i += 1
            End While
            used.Add(candidate)
            Return candidate
        End Function

        Friend Shared Function CreateUniqueAnchorNames(workbookNames As IEnumerable(Of String)) As List(Of String)
            Dim used As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Return workbookNames.Select(Function(lbl) MakeUnique(Slugify(lbl), used)).ToList()
        End Function

    End Class
#Enable Warning CA1805 ' Keine unnötige Initialisierung

End Namespace