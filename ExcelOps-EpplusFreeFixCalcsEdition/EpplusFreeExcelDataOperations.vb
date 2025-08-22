Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Data
Imports System.Text
Imports CompuMaster.Epplus4
Imports CompuMaster.Epplus4.FormulaParsing
Imports CompuMaster.Epplus4.FormulaParsing.Logging

Namespace ExcelOps

    ''' <summary>
    ''' An Excel operations engine based on Epplus 4 with its LGPL license
    ''' </summary>
    ''' <remarks>
    ''' For licensing issues of origin Epplus 4 project, please see https://github.com/JanKallman/EPPlus
    ''' </remarks>
    Public Class EpplusFreeExcelDataOperations
        Inherits ExcelDataOperationsBase

        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String)
            MyBase.New(file, mode, False, True, [readOnly], passwordForOpening)
        End Sub

        Public Sub New(data As Byte(), passwordForOpening As String)
            MyBase.New(data, False, True, passwordForOpening)
        End Sub

        Public Sub New(data As Byte(), passwordForOpening As String, disableInitialCalculation As Boolean)
            MyBase.New(data, Not disableInitialCalculation, True, passwordForOpening)
        End Sub

        Public Sub New(data As System.IO.Stream, passwordForOpening As String)
            MyBase.New(data, False, True, passwordForOpening)
        End Sub

        Public Sub New(data As System.IO.Stream, passwordForOpening As String, disableInitialCalculation As Boolean)
            MyBase.New(data, Not disableInitialCalculation, True, passwordForOpening)
        End Sub

        Public Sub New()
            Me.New(Nothing)
        End Sub

        Public Sub New(passwordForOpeningOnNextTime As String)
            MyBase.New(False, True, True, passwordForOpeningOnNextTime)
        End Sub

        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "Epplus 4 (LGPL)"
            End Get
        End Property

        Private Const FULL_CALC_ON_LOAD As Boolean = True

        Private _WorkbookPackage As CompuMaster.Epplus4.ExcelPackage
        Public ReadOnly Property WorkbookPackage As CompuMaster.Epplus4.ExcelPackage
            Get
                If Me._WorkbookPackage Is Nothing Then
                    Throw New InvalidOperationException("Workbook has already been closed")
                End If
                Return Me._WorkbookPackage
            End Get
        End Property

        Public ReadOnly Property Workbook As CompuMaster.Epplus4.ExcelWorkbook
            Get
                If Me._WorkbookPackage Is Nothing Then
                    Throw New InvalidOperationException("Workbook has already been closed")
                End If
                Return Me._WorkbookPackage.Workbook
            End Get
        End Property

        ''' <summary>
        ''' Reset the calculated cell value from a cell with formula to force MS Excel calculation engine to recalculate the cell value
        ''' </summary>
        Public Sub ResetCellValueFromFormulaCellInWholeWorkbook()
            Dim AllSheetNames As List(Of String) = Me.SheetNames
            For Each SheetName As String In AllSheetNames
                Me.ResetCellValueFromFormulaCell(SheetName)
            Next
        End Sub

        ''' <summary>
        ''' Reset the calculated cell value from a cell with formula to force MS Excel calculation engine to recalculate the cell value
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Sub ResetCellValueFromFormulaCell(sheetName As String)
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            For MyRowIndexCounter As Integer = 0 To LastCell.RowIndex
                For MyColIndexCounter As Integer = 0 To LastCell.ColumnIndex
                    If Me.Workbook.Worksheets(sheetName).Cells(MyRowIndexCounter + 1, MyColIndexCounter + 1).Formula <> Nothing Then
                        Me.ResetCellValueFromFormulaCell(sheetName, MyRowIndexCounter, MyColIndexCounter)
                    End If
                Next
            Next
        End Sub

        ''' <summary>
        ''' Reset the calculated cell value from a cell with formula to force MS Excel calculation engine to recalculate the cell value
        ''' </summary>
        Public Sub ResetCellValueFromFormulaCell(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim CurrentCellFormula As String = Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Formula
            If CurrentCellFormula = Nothing Then
                Throw New ArgumentException("Cell " & New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All).Address(True) & " doesn't contain a formula")
            End If
            Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).ClearCachedCalculatedFormulaValue()
            Me.RecalculationRequired = True
        End Sub

        ''' <summary>
        ''' Reset the calculated cell value from a cell with formula to force MS Excel calculation engine to recalculate the cell value
        ''' </summary>
        ''' <param name="cell"></param>
        Public Sub ResetCellValueFromFormulaCell(cell As ExcelCell)
            Dim MyExcelCellAddress As ExcelCellAddress = New ExcelAddress(cell.Address).Start
            Dim CurrentCellFormula As String = Me.Workbook.Worksheets(cell.SheetName).Cells(MyExcelCellAddress.Row, MyExcelCellAddress.Column).Formula
            If CurrentCellFormula = Nothing Then
                Throw New ArgumentException("Cell " & cell.Address(True) & " doesn't contain a formula", NameOf(cell))
            End If
            Me.Workbook.Worksheets(cell.SheetName).Cells(MyExcelCellAddress.Row, MyExcelCellAddress.Column).ClearCachedCalculatedFormulaValue()
            Me.RecalculationRequired = True
        End Sub

        ''' <summary>
        ''' Has the workbook some cells which got a formula without a calculated value
        ''' </summary>
        Public Function FindMissingCalculatedCellValueFromFormulaCell() As List(Of MissingCalculatedCellValueException)
            Dim Result As New List(Of MissingCalculatedCellValueException)
            Dim AllSheetNames As List(Of String) = Me.SheetNames
            For Each SheetName As String In AllSheetNames
                Result.AddRange(Me.FindMissingCalculatedCellValueFromFormulaCell(SheetName))
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Has the specified sheet some cells which got a formula without a calculated value
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Function FindMissingCalculatedCellValueFromFormulaCell(sheetName As String) As List(Of MissingCalculatedCellValueException)
            Dim Result As New List(Of MissingCalculatedCellValueException)
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            For MyRowIndexCounter As Integer = 0 To LastCell.RowIndex
                For MyColIndexCounter As Integer = 0 To LastCell.ColumnIndex
                    Dim CellFormula As String = Me.Workbook.Worksheets(sheetName).Cells(MyRowIndexCounter + 1, MyColIndexCounter + 1).Formula
                    If CellFormula <> Nothing Then
                        If Me.IsMissingCalculatedCellValueFromFormulaCell(sheetName, MyRowIndexCounter, MyColIndexCounter) Then
                            Result.Add(New MissingCalculatedCellValueException(Me.FilePath, sheetName, MyRowIndexCounter, MyColIndexCounter, CellFormula))
                        End If
                    End If
                Next
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Has the workbook some cells which got a formula without a calculated value
        ''' </summary>
        Public Function HasMissingCalculatedCellValueFromFormulaCell() As Boolean
            Dim AllSheetNames As List(Of String) = Me.SheetNames
            For Each SheetName As String In AllSheetNames
                Dim Result As Boolean = Me.HasMissingCalculatedCellValueFromFormulaCell(SheetName)
                If Result = True Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Has the specified sheet some cells which got a formula without a calculated value
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Function HasMissingCalculatedCellValueFromFormulaCell(sheetName As String) As Boolean
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            For MyRowIndexCounter As Integer = 0 To LastCell.RowIndex
                For MyColIndexCounter As Integer = 0 To LastCell.ColumnIndex
                    Dim CellFormula As String = Me.Workbook.Worksheets(sheetName).Cells(MyRowIndexCounter + 1, MyColIndexCounter + 1).Formula
                    If CellFormula <> Nothing Then
                        If Me.IsMissingCalculatedCellValueFromFormulaCell(sheetName, MyRowIndexCounter, MyColIndexCounter) Then
                            Return True
                        End If
                    End If
                Next
            Next
            Return False
        End Function

        ''' <summary>
        ''' Has the specified cell got a formula without a calculated value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Public Function IsMissingCalculatedCellValueFromFormulaCell(cell As ExcelCell) As Boolean
            Dim MyExcelCellAddress As ExcelCellAddress = New ExcelAddress(cell.Address).Start
            If Me.Workbook.Worksheets(cell.SheetName) Is Nothing Then Throw New ArgumentOutOfRangeException(NameOf(cell), "Sheet not found: " & cell.SheetName)
            Dim CheckResult As Boolean = Me.Workbook.Worksheets(cell.SheetName).Cells(MyExcelCellAddress.Row, MyExcelCellAddress.Column).HasMissingCachedCalculatedFormulaValue()
            If CheckResult = True AndAlso Tools.IsFormulaWithoutCellReferences(Me.Workbook.Worksheets(cell.SheetName).Cells(MyExcelCellAddress.Row, MyExcelCellAddress.Column).Formula) Then
                Me.Workbook.Worksheets(cell.SheetName).Cells(MyExcelCellAddress.Row, MyExcelCellAddress.Column).Calculate
                CheckResult = False
            End If
            Return CheckResult
        End Function

        ''' <summary>
        ''' Has the specified cell got a formula without a calculated value
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Public Function IsMissingCalculatedCellValueFromFormulaCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If Me.Workbook.Worksheets(sheetName) Is Nothing Then Throw New ArgumentOutOfRangeException(NameOf(sheetName), "Sheet not found: " & sheetName)
            Dim CheckResult As Boolean = Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).HasMissingCachedCalculatedFormulaValue()
            If CheckResult = True AndAlso Tools.IsFormulaWithoutCellReferences(Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Formula) Then
                Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Calculate
                CheckResult = False
            End If
            Return CheckResult
        End Function

        Public Overrides Function SelectedSheetName() As String
            Return Me.Workbook.Worksheets(Me.Workbook.View.ActiveTab).Name
        End Function

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Sub SelectSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim CurrentSheetNames = Me.SheetNames
            If CurrentSheetNames.Contains(sheetName) = False Then
                Throw New ArgumentOutOfRangeException(NameOf(sheetName), "Sheet name not found: " & sheetName)
            End If
            Dim RequestedSheetIndex As Integer = CurrentSheetNames.IndexOf(sheetName)
            'Me.SelectSheet(RequestedSheetIndex)
            Me.Workbook.Worksheets(sheetName).Select()
            'Me.Workbook.Worksheets(RequestedSheetIndex).Select()
            'If Me.SelectedSheetName <> sheetName Then
            '    Throw New InvalidOperationException("Sheet selection requested for """ & sheetName & """, but after selection the active tab was """ & Workbook.Worksheets(Me.Workbook.View.ActiveTab).Name & """")
            'End If
        End Sub

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetIndex"></param>
        Public Overrides Sub SelectSheet(sheetIndex As Integer)
            Me.SelectSheet(Me.SheetNames(sheetIndex))
            Return
            Me.Workbook.Worksheets(sheetIndex + 1).Select() 'Worksheets collection always uses 1-based index (independent of Epplus compatibility setting)
            If Me.Workbook.View.ActiveTab <> sheetIndex Then
                Throw New InvalidOperationException("Sheet selection requested for index """ & sheetIndex & """, but after selection the active tab was """ & Me.Workbook.View.ActiveTab & """")
            End If
        End Sub

#Region "Colors and Theming (Helpers for e.g. ExportSheetToHtmlInternal)"
        ''' <summary>
        ''' Zentrale Farbauflösung inkl. Cache
        ''' </summary>
        ''' <remarks>
        ''' Die Farbauflösung erfolgt inkl. Cache für:
        ''' <list type="bullet">
        ''' <item>Rgb (ARGB oder RGB)</item>
        ''' <item>Theme (0..11) mit Tint</item>
        ''' <item>Indexed (kleine Palette, 64=auto)</item>
        ''' </list>
        ''' </remarks>
        ''' <returns>Gibt "#RRGGBB" oder Nothing zurück.</returns>
        Private Shared Function ExcelColorToCssHex(color As CompuMaster.Epplus4.Style.ExcelColor,
                                                   cache As IDictionary(Of String, String)) As String
            If color Is Nothing Then Return Nothing

            Dim ColorCacheKeyName As String = BuildColorCacheKey(color)
            Dim hex As String = Nothing

            If Not String.IsNullOrEmpty(ColorCacheKeyName) AndAlso cache.TryGetValue(ColorCacheKeyName, hex) Then
                Return hex
            End If

            ' 1) Direkter RGB/ARGB?
            If Not String.IsNullOrEmpty(color.Rgb) Then
                Dim v = color.Rgb.Trim()
                If v.Length = 8 Then
                    hex = "#" & v.Substring(2) ' ARGB -> RRGGBB
                ElseIf v.Length = 6 Then
                    hex = "#" & v
                End If
                'Add to cache if valid and return hex
                If Not String.IsNullOrEmpty(ColorCacheKeyName) AndAlso Not String.IsNullOrEmpty(hex) AndAlso hex.Length = 7 AndAlso hex(0) = "#"c Then
                    cache(ColorCacheKeyName) = hex
                End If
                Return hex
            End If

            ' 2) Theme-Farbe (0..11) + Tint
            Dim themeIdx As Integer
            If Not String.IsNullOrEmpty(color.Theme) AndAlso Integer.TryParse(color.Theme, themeIdx) Then
                hex = DefaultOfficeTheme(themeIdx)
                If Not String.IsNullOrEmpty(hex) Then
                    Dim tint As Double = 0
                    ' Tint kann je nach EPPlus-Version Double sein; defensiv parsen:
                    Double.TryParse(Convert.ToString(color.Tint, Globalization.CultureInfo.InvariantCulture), tint)
                    If Math.Abs(tint) > Double.Epsilon Then
                        hex = ApplyTint(hex, tint)
                    End If
                    'Add to cache if valid and return hex
                    If Not String.IsNullOrEmpty(ColorCacheKeyName) AndAlso Not String.IsNullOrEmpty(hex) AndAlso hex.Length = 7 AndAlso hex(0) = "#"c Then
                        cache(ColorCacheKeyName) = hex
                    End If
                    Return hex
                End If
            End If

            ' 3) Indexed-Farben (kleine, praxisnahe Palette)
            Dim idx As Integer = color.Indexed
            If idx > 0 AndAlso idx <> 64 Then
                hex = IndexedColor(idx)
                If Not String.IsNullOrEmpty(hex) Then
                    'Add to cache if valid and return hex
                    If Not String.IsNullOrEmpty(ColorCacheKeyName) AndAlso Not String.IsNullOrEmpty(hex) AndAlso hex.Length = 7 AndAlso hex(0) = "#"c Then
                        cache(ColorCacheKeyName) = hex
                    End If
                    Return hex
                End If
            End If

            ' nichts gefunden
            Return Nothing

        End Function

        ''' <summary>
        ''' Konstruiert einen stabilen Schlüssel für die Cache-Map
        ''' </summary>
        ''' <param name="color"></param>
        ''' <returns></returns>
        Private Shared Function BuildColorCacheKey(color As CompuMaster.Epplus4.Style.ExcelColor) As String
            If color Is Nothing Then Return Nothing

            If Not String.IsNullOrEmpty(color.Rgb) Then
                Return "rgb:" & color.Rgb
            End If

            If Not String.IsNullOrEmpty(color.Theme) Then
                Dim tintStr As String = Convert.ToString(color.Tint, Globalization.CultureInfo.InvariantCulture)
                Return "theme:" & color.Theme & ":" & tintStr
            End If

            If color.Indexed <> 0 Then
                Return "idx:" & color.Indexed
            End If

            ' Kein stabiler Key
            Return Nothing
        End Function
#End Region

#Region "HTML Export"

        ''' <summary>
        ''' Rendert ein einzelnes Worksheet als vollständiges HTML-Dokument (UTF-8).
        ''' Kompatibel zu EPPlus 4.5.x (kein Lizenz-Setup nötig).
        ''' </summary>
        Friend Overloads Shared Sub ExportSheetToHtmlInternal(ws As ExcelWorksheet, sb As StringBuilder, options As HtmlSheetExportOptions)

            If ws Is Nothing OrElse ws.Dimension Is Nothing Then
                sb.AppendLine(options.HtmlForEmptySheet)
                Return
            End If

            Dim firstRow = ws.Dimension.Start.Row
            Dim lastRow = ws.Dimension.End.Row
            Dim firstCol = ws.Dimension.Start.Column
            Dim lastCol = ws.Dimension.End.Column

            ' --- Cache für Farbauflösungen (pro Methodenaufruf) ---
            Dim colorCache As New Dictionary(Of String, String)(StringComparer.Ordinal)

            ' --- Merge-Map vorbereiten ---
            Dim mergeTopLeft As New Dictionary(Of String, (RowSpan As Integer, ColSpan As Integer))()
            Dim coveredByMergedCellsMasterCell As New HashSet(Of String)()
            For Each addr In ws.MergedCells
                Dim range = ws.Cells(addr)
                Dim RowStartIndex As Integer = range.Start.Row, ColumnStartIndex = range.Start.Column
                Dim RowEndIndex As Integer = range.End.Row, ColumnEndIndex = range.End.Column
                Dim keyTL = ExportSheetToHtmlInternal_CellAddressKey(RowStartIndex, ColumnStartIndex)
                mergeTopLeft(keyTL) = ((RowEndIndex - RowStartIndex + 1), (ColumnEndIndex - ColumnStartIndex + 1))
                For RowIndex As Integer = RowStartIndex To RowEndIndex
                    For ColumnIndex = ColumnStartIndex To ColumnEndIndex
                        If Not (RowIndex = RowStartIndex AndAlso ColumnIndex = ColumnStartIndex) Then coveredByMergedCellsMasterCell.Add(ExportSheetToHtmlInternal_CellAddressKey(RowIndex, ColumnIndex))
                    Next
                Next
            Next

            Dim tableCssClass As String = options.TableCssClassName
            sb.Append("<table class=""" & tableCssClass & """>")

            For RowIndex = firstRow To lastRow
                sb.Append("<tr>")
                For ColumnIndex = firstCol To lastCol
                    Dim CellAddressKey = ExportSheetToHtmlInternal_CellAddressKey(RowIndex, ColumnIndex)
                    If coveredByMergedCellsMasterCell.Contains(CellAddressKey) Then Continue For

                    Dim cell = ws.Cells(RowIndex, ColumnIndex)
                    Dim tag As String = If(options.ConsiderRowIndexesAsTableHeader?.Contains(RowIndex), "th", "td")

                    ' --- Merge-Attribute ---
                    Dim rowspan As Integer = 1, colspan As Integer = 1
                    If mergeTopLeft.ContainsKey(CellAddressKey) Then
                        rowspan = mergeTopLeft(CellAddressKey).RowSpan
                        colspan = mergeTopLeft(CellAddressKey).ColSpan
                    End If

                    ' --- Styles extrahieren ---
                    Dim styles As New List(Of String)

                    ' Horizontal-Align
                    Select Case cell.Style.HorizontalAlignment
                        Case Style.ExcelHorizontalAlignment.Center : styles.Add("text-align:center")
                        Case Style.ExcelHorizontalAlignment.Right : styles.Add("text-align:right")
                        Case Style.ExcelHorizontalAlignment.Justify : styles.Add("text-align:justify")
                        Case Else : styles.Add("text-align:left")
                    End Select

                    ' Font
                    If cell.Style.Font IsNot Nothing Then
                        If cell.Style.Font.Bold Then styles.Add("font-weight:bold")
                        If cell.Style.Font.Italic Then styles.Add("font-style:italic")
                        If cell.Style.Font.UnderLine Then styles.Add("text-decoration:underline")
                        Dim fc = ExcelColorToCssHex(cell.Style.Font.Color, colorCache)
                        If Not String.IsNullOrEmpty(fc) Then styles.Add("color:" & fc)
                    End If

                    ' Hintergrund (PatternColor bevorzugen, dann BackgroundColor)
                    If cell.Style.Fill IsNot Nothing AndAlso cell.Style.Fill.PatternType <> Style.ExcelFillStyle.None Then
                        Dim bg As String = Nothing
                        If cell.Style.Fill.PatternType = Style.ExcelFillStyle.Solid Then
                            bg = ExcelColorToCssHex(cell.Style.Fill.PatternColor, colorCache)
                            If String.IsNullOrEmpty(bg) Then
                                bg = ExcelColorToCssHex(cell.Style.Fill.BackgroundColor, colorCache)
                            End If
                        Else
                            ' bei anderen Pattern-Typen beide prüfen
                            bg = ExcelColorToCssHex(cell.Style.Fill.BackgroundColor, colorCache)
                            If String.IsNullOrEmpty(bg) Then
                                bg = ExcelColorToCssHex(cell.Style.Fill.PatternColor, colorCache)
                            End If
                        End If
                        If Not String.IsNullOrEmpty(bg) Then styles.Add("background-color:" & bg)
                    End If

                    ' Zeilenumbruch
                    If cell.Style.WrapText Then styles.Add("white-space:pre-wrap")

                    Dim styleAttr As String = If(styles.Count > 0, $" style=""{String.Join(";", styles)}""", "")

                    ' Inhalt (Text = bereits nach NumberFormat formatiert)
                    Dim content As String = cell.Text
                    If String.IsNullOrEmpty(content) Then content = " "
                    content = System.Net.WebUtility.HtmlEncode(content)

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

            sb.AppendLine("</table>")
        End Sub

        ' ---------- Helpers ----------

        Private Shared Function ExportSheetToHtmlInternal_CellAddressKey(r As Integer, c As Integer) As String
            Return r.ToString() & "|" & c.ToString()
        End Function

#End Region

    End Class

End Namespace