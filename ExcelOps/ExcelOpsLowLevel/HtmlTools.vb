Option Strict On
Option Explicit On

Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Net

Namespace ExcelOps

    Friend NotInheritable Class HtmlTools

        Private Sub New()
        End Sub

        ' In Namespace ExcelOps, Klasse HtmlTools (bestehende Importe beibehalten)

        Public Shared Function GenerateWorkbookSubnav(sb As StringBuilder,
                                             items As IEnumerable(Of String),
                                             sticky As HtmlWorkbookExportOptions.SheetNavigationPositions,
                                             actionStyle As HtmlWorkbookExportOptions.SheetNavigationActionStyles,
                                             alwaysVisible As Boolean,
                                             Optional navId As String = "workbook-nav",
                                             Optional cssClass As String = "cm-wb-subnav",
                                             Optional includeStyleBlock As Boolean = True,
                                             Optional anchorNames As IEnumerable(Of String) = Nothing) As String
            If items Is Nothing Then Throw New ArgumentNullException(NameOf(items))

            Dim labels = items.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToList()
            If labels.Count = 0 Then Return sb.ToString()

            Dim anchors As List(Of String)
            If anchorNames IsNot Nothing Then
                anchors = anchorNames.ToList()
                If anchors.Count <> labels.Count Then
                    Throw New ArgumentException("Anzahl von anchorNames muss der Anzahl von items entsprechen.")
                End If
            Else
                Dim used As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                anchors = labels.Select(Function(lbl) MakeUnique(Slugify(lbl), used)).ToList()
            End If

            ' Basis-Styles der Subnav (wie gehabt)
            If includeStyleBlock Then
                sb.AppendLine("<style>")
                sb.AppendLine(".cm-wb-subnav{display:flex;align-items:center;gap:.5rem;flex-wrap:wrap;padding:.4rem .8rem;background:#f8f9fa;border-bottom:1px solid #ddd;z-index:1000;font-size:.95rem;}")
                sb.AppendLine(".cm-wb-subnav.sticky-top{position:sticky;top:0;}")
                sb.AppendLine(".cm-wb-subnav.sticky-bottom{position:sticky;bottom:0;border-top:1px solid #ddd;border-bottom:none;}")
                sb.AppendLine(".cm-wb-subnav ul{list-style:none;margin:0;padding:0;display:flex;gap:.4rem;flex-wrap:wrap;}")
                sb.AppendLine(".cm-wb-subnav a{text-decoration:none;padding:.28rem .5rem;border-radius:.4rem;}")
                sb.AppendLine(".cm-wb-subnav a:hover,.cm-wb-subnav a:focus{background:rgba(0,0,0,.06);outline:none;}")
                sb.AppendLine("</style>")
            End If

            ' Bei SwitchVisibleSheet: Sheet-CSS + JS einbetten (einmalig pro Seite ist ausreichend)
            If actionStyle = HtmlWorkbookExportOptions.SheetNavigationActionStyles.SwitchVisibleSheet Then
                EmitSwitchVisibleSupport(sb, includeStyleBlock:=includeStyleBlock)
            End If

            Dim stickyClass As String = String.Empty
            If alwaysVisible Then
                Select Case sticky
                    Case HtmlWorkbookExportOptions.SheetNavigationPositions.Top
                        stickyClass = " sticky-top"
                    Case HtmlWorkbookExportOptions.SheetNavigationPositions.Bottom
                        stickyClass = " sticky-bottom"
                End Select
            End If

            sb.Append("<nav")
            If Not String.IsNullOrWhiteSpace(navId) Then
                sb.Append(" id=""").Append(WebUtility.HtmlEncode(navId)).Append(""""c)
            End If
            sb.Append(" class=""")
            sb.Append(WebUtility.HtmlEncode(cssClass))
            sb.Append(WebUtility.HtmlEncode(stickyClass))
            sb.Append(""" role=""navigation"" aria-label=""Unter-Navigation (Arbeitsmappe)"">")

            sb.Append("<ul>")
            For i = 0 To labels.Count - 1
                Dim label = labels(i)
                Dim anchor = anchors(i)

                If actionStyle = HtmlWorkbookExportOptions.SheetNavigationActionStyles.SwitchVisibleSheet Then
                    ' Link mit data-target -> JS zeigt entsprechendes Sheet
                    sb.Append("<li><a href=""#").
                       Append(WebUtility.HtmlEncode(anchor)).
                       Append(""" data-target=""").
                       Append(WebUtility.HtmlEncode(anchor)).
                       Append(""">").
                       Append(WebUtility.HtmlEncode(label)).
                       Append("</a></li>")
                Else
                    ' Klassischer Anker-Sprung
                    sb.Append("<li><a href=""#").
                       Append(WebUtility.HtmlEncode(anchor)).
                       Append(""">").
                       Append(WebUtility.HtmlEncode(label)).
                       Append("</a></li>")
                End If
            Next
            sb.Append("</ul></nav>")

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Styles/Script für SwitchVisibleSheet-Modus emitten (einmal pro Seite).
        ''' </summary>
        Public Shared Sub EmitSwitchVisibleSupport(sb As StringBuilder, Optional includeStyleBlock As Boolean = True)
            If includeStyleBlock Then
                sb.AppendLine("<style>")
                sb.AppendLine(".cm-wb-sheet{display:none;}")
                sb.AppendLine(".cm-wb-sheet.cm-visible{display:block;}")
                sb.AppendLine("</style>")
                sb.AppendLine("<noscript><style>.cm-wb-sheet{display:block !important;}</style></noscript>")
            End If

            sb.AppendLine("<script>(function(){")
            sb.AppendLine("function showSheet(id){var s=document.querySelectorAll('.cm-wb-sheet');for(var i=0;i<s.length;i++){s[i].classList.remove('cm-visible');}var el=document.getElementById(id);if(el){el.classList.add('cm-visible');}}")
            sb.AppendLine("function handleHash(){var id=location.hash?location.hash.substring(1):null;if(id){showSheet(id);}}")
            sb.AppendLine("document.addEventListener('click',function(e){var a=e.target.closest('.cm-wb-subnav a[data-target]');if(a){e.preventDefault();var id=a.getAttribute('data-target');history.replaceState(null,'','#'+id);showSheet(id);}});")
            sb.AppendLine("window.addEventListener('hashchange',handleHash);handleHash();})();</script>")
        End Sub

        ''' <summary>
        ''' Öffnenden Wrapper für ein Sheet ausgeben. Sichtbar: True -> dieses Sheet ist initial sichtbar.
        ''' </summary>
        Public Shared Sub BeginSheetSection(sb As StringBuilder, anchorName As String, Optional visible As Boolean = False)
            Dim visibleClass As String = If(visible, " cm-visible", String.Empty)
            sb.Append("<section id=""").
       Append(WebUtility.HtmlEncode(anchorName)).
       Append(""" class=""cm-wb-sheet").
       Append(visibleClass).
       Append(""" data-sheet=""").
       Append(WebUtility.HtmlEncode(anchorName)).
       Append(""">")
            ' Kompatibel zu älteren Ankernamen
            sb.Append("<a name=""").Append(WebUtility.HtmlEncode(anchorName)).Append("""></a>")
        End Sub

        ''' <summary>
        ''' Schließenden Wrapper für ein Sheet ausgeben.
        ''' </summary>
        Public Shared Sub EndSheetSection(sb As StringBuilder)
            sb.Append("</section>")
        End Sub

        ''' <summary>
        ''' Hilfsfunktion: erzeugt einen (leeren) Anker, kompatibel zu Excel-Export (<a name="xy"></a>).
        ''' </summary>
        Public Shared Function BuildNamedAnchorTag(nameValue As String) As String
            If String.IsNullOrWhiteSpace(nameValue) Then Throw New ArgumentNullException(NameOf(nameValue))
            Return $"<a name=""{WebUtility.HtmlEncode(nameValue)}""></a>"
        End Function

        ' "Über uns" -> "uber-uns"
        Private Shared Function Slugify(input As String) As String
            Dim normalized As String = input.Normalize(NormalizationForm.FormD)
            Dim b As New StringBuilder(normalized.Length)
            For Each ch As Char In normalized
                If CharUnicodeInfo.GetUnicodeCategory(ch) <> UnicodeCategory.NonSpacingMark Then b.Append(ch)
            Next
            Dim noDiacritics As String = b.ToString().Normalize(NormalizationForm.FormC)
            Dim lowered As String = noDiacritics.ToLowerInvariant().Replace("ß", "ss")
            Dim replaced As String = Regex.Replace(lowered, "[^\w\-]+", "-")   ' Nicht-Wortzeichen -> "-"
            replaced = Regex.Replace(replaced, "[-_]+", "-")                    ' Mehrfachtrennstriche verdichten
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

    End Class

End Namespace