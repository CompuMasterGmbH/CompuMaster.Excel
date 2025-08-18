Option Explicit On
Option Strict On

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
        ''' HTML code on top of everything (including html and body tags)
        ''' </summary>
        ''' <returns></returns>
        Public Property HtmlBefore As String

        ''' <summary>
        ''' HTML code on bottom of everything (including html and body tags)
        ''' </summary>
        ''' <returns></returns>
        Public Property HtmlBehind As String

        Protected Friend ReadOnly Property DefaultHtmlBefore As String = "<html><body>"
        Protected Friend ReadOnly Property DefaultHtmlBehind As String = "</body></html>"

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

    End Class
#Enable Warning CA1805 ' Keine unnötige Initialisierung

End Namespace