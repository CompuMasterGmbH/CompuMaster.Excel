Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Data
Imports System.Globalization
Imports System.Text
Imports System.Windows.Input
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.FormulaParsing
Imports OfficeOpenXml.FormulaParsing.Logging

Namespace ExcelOps

    ''' <summary>
    ''' An Excel operations engine based on Epplus with its Polyform license
    ''' </summary>
    ''' <remarks>
    ''' Correct licensing required, see <see cref="LicenseContext"/> and https://www.epplussoftware.com/en/LicenseOverview
    ''' </remarks>
    Public Class EpplusPolyformExcelDataOperations
        Inherits ExcelDataOperationsBase

        Protected Overrides ReadOnly Property DefaultCalculationOptions As ExcelEngineDefaultOptions
            Get
                Return New ExcelEngineDefaultOptions(False, False)
            End Get
        End Property

        ''' <summary>
        ''' Create or open a workbook (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <param name="file">Path to a file which shall be loaded or null if a new workbook shall be created</param>
        ''' <param name="mode">Open an existing file or (re)create a new file</param>
        ''' <param name="options">File and engine options</param>
        Public Sub New(file As String, mode As OpenMode, options As ExcelDataOperationsOptions)
            MyBase.New(file, mode, options)
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        ''' <param name="data"></param>
        ''' <param name="options">File and engine options</param>
        Public Sub New(data As Byte(), options As ExcelDataOperationsOptions)
            MyBase.New(data, options)
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        ''' <param name="data"></param>
        ''' <param name="options">File and engine options</param>
        Public Sub New(data As System.IO.Stream, options As ExcelDataOperationsOptions)
            MyBase.New(data, options)
        End Sub

        ''' <summary>
        ''' Create or open a workbook
        ''' </summary>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String, disableInitialCalculation As Boolean, disableCalculationEngine As Boolean)
            MyBase.New(file, mode, Not disableInitialCalculation, disableCalculationEngine, [readOnly], passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        ''' <summary>
        ''' Create or open a workbook
        ''' </summary>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String, disableInitialCalculation As Boolean)
            MyBase.New(file, mode, Not disableInitialCalculation, False, [readOnly], passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        ''' <summary>
        ''' Create or open a workbook
        ''' </summary>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String)
            MyBase.New(file, mode, True, False, [readOnly], passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As Byte(), passwordForOpening As String)
            MyBase.New(data, True, False, passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As Byte(), passwordForOpening As String, disableInitialCalculation As Boolean, disableCalculationEngine As Boolean)
            MyBase.New(data, Not disableInitialCalculation, disableCalculationEngine, passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As Byte(), passwordForOpening As String, disableInitialCalculation As Boolean)
            MyBase.New(data, Not disableInitialCalculation, False, passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As System.IO.Stream, passwordForOpening As String)
            MyBase.New(data, True, False, passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As System.IO.Stream, passwordForOpening As String, disableInitialCalculation As Boolean, disableCalculationEngine As Boolean)
            MyBase.New(data, Not disableInitialCalculation, disableCalculationEngine, passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        ''' <summary>
        ''' Open a workbook
        ''' </summary>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As System.IO.Stream, passwordForOpening As String, disableInitialCalculation As Boolean)
            MyBase.New(data, Not disableInitialCalculation, False, passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        ''' <summary>
        ''' Create a new workbook or just create an uninitialized instance of this Excel engine
        ''' </summary>
        ''' <param name="mode"></param>
        Public Sub New(mode As OpenMode)
            MyBase.New(mode)
        End Sub

        ''' <summary>
        ''' Create a new workbook or just create an uninitialized instance of this Excel engine
        ''' </summary>
        ''' <param name="mode"></param>
        ''' <param name="options"></param>
        Public Sub New(mode As OpenMode, options As ExcelDataOperationsOptions)
            MyBase.New(mode, options)
        End Sub

        ''' <summary>
        ''' Create a new instance for accessing Excel workbooks (still requires creating or loading of a workbook)
        ''' </summary>
        ''' <param name="passwordForOpeningOnNextTime">Pre-define encryption password on future save actions</param>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(passwordForOpeningOnNextTime As String)
            MyBase.New(False, True, True, passwordForOpeningOnNextTime)
            ValidateLicenseContext(Me)
        End Sub

        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "Epplus (Polyform license edition)"
            End Get
        End Property

        ''' <summary>
        ''' The license context for Epplus (see its polyform license)
        ''' </summary>
        ''' <remarks>https://epplussoftware.com/en/LicenseOverview/LicenseFAQ</remarks>
        ''' <returns></returns>
        Public Shared Property LicenseContext As EpplusLicenseActivator?
            Get
                If OfficeOpenXml.ExcelPackage.License.LicenseType Is Nothing Then
                    Return Nothing
                End If
                Select Case OfficeOpenXml.ExcelPackage.License.LicenseType.Value
                    Case OfficeOpenXml.EPPlusLicenseType.Commercial
                        Return New EpplusLicenseActivator(OfficeOpenXml.ExcelPackage.License.LicenseType.Value, OfficeOpenXml.ExcelPackage.License.LicenseKey)
                    Case OfficeOpenXml.EPPlusLicenseType.NonCommercialOrganization, OfficeOpenXml.EPPlusLicenseType.NonCommercialPersonal
                        Return New EpplusLicenseActivator(OfficeOpenXml.ExcelPackage.License.LicenseType.Value, OfficeOpenXml.ExcelPackage.License.LegalName)
                    Case Else
                        Throw New NotImplementedException("Unsupported license type: " & OfficeOpenXml.ExcelPackage.License.LicenseType.Value.ToString())
                End Select
            End Get
            Set(value As EpplusLicenseActivator?)
                If value.HasValue = False Then Throw New ArgumentNullException(NameOf(value), "License must be specified")
                If value.Value.KeyOrName = Nothing Then Throw New ArgumentNullException(NameOf(value), "License must be specified with license key or licensor name")
                Select Case value.Value.LicenseType
                    Case EPPlusLicenseType.Commercial
                        OfficeOpenXml.ExcelPackage.License.SetCommercial(value.Value.KeyOrName)
                    Case EPPlusLicenseType.NonCommercialOrganization
                        OfficeOpenXml.ExcelPackage.License.SetNonCommercialOrganization(value.Value.KeyOrName)
                    Case EPPlusLicenseType.NonCommercialPersonal
                        OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal(value.Value.KeyOrName)
                    Case Else
                        Throw New NotImplementedException("Unsupported license type: " & value.Value.LicenseType.ToString())
                End Select
            End Set
        End Property

        Public Structure EpplusLicenseActivator

            Public Sub New(licenseType As OfficeOpenXml.EPPlusLicenseType, licenseKeyOrLegalName As String)
                Me.LicenseType = licenseType
                Me.KeyOrName = licenseKeyOrLegalName
            End Sub

            Public Property LicenseType As OfficeOpenXml.EPPlusLicenseType

            ''' <summary>
            ''' License key for commercial use or personal/organisation name for non-commercial use
            ''' </summary>
            ''' <returns></returns>
            Public Property KeyOrName As String

        End Structure

        Private Shared Sub ValidateLicenseContext(instance As EpplusPolyformExcelDataOperations)
            If LicenseContext.HasValue = False Then
                Throw New System.ComponentModel.LicenseException(GetType(EpplusPolyformExcelDataOperations), instance, NameOf(LicenseContext) & " must be assigned before creating instances")
            End If
        End Sub

        Private _WorkbookPackage As OfficeOpenXml.ExcelPackage
        Public ReadOnly Property WorkbookPackage As OfficeOpenXml.ExcelPackage
            Get
                ValidateLicenseContext(Me)
                If Me._WorkbookPackage Is Nothing Then
                    Throw New InvalidOperationException("Workbook has already been closed")
                End If
                Return Me._WorkbookPackage
            End Get
        End Property

        Public ReadOnly Property Workbook As OfficeOpenXml.ExcelWorkbook
            Get
                ValidateLicenseContext(Me)
                If Me._WorkbookPackage Is Nothing Then
                    Throw New InvalidOperationException("Workbook has already been closed")
                End If
                Return Me._WorkbookPackage.Workbook
            End Get
        End Property

        'Public ReadOnly Property DrawingsCount As Integer
        '    Get
        '        Return Me.Workbook.Worksheets
        '        Return OfficeOpenXml.Drawing.ExcelPicture
        '    End Get
        'End Property
        '
        'Public ReadOnly Property Drawings As OfficeOpenXml.Drawing.ExcelPicture

#Disable Warning CA1822 ' Member als statisch markieren
        ''' <summary>
        ''' NOT AVAILABLE, but implemented as stub method for SharedCode compatibility: Reset the calculated cell value from a cell with formula to force MS Excel calculation engine to recalculate the cell value
        ''' </summary>
        Private Sub ResetCellValueFromFormulaCellInWholeWorkbook()
        End Sub

        ''' <summary>
        ''' NOT AVAILABLE, but implemented as stub method for SharedCode compatibility: Has the specified cell got a formula without a calculated value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Public Function IsMissingCalculatedCellValueFromFormulaCell(cell As ExcelCell) As Boolean
            Return False
        End Function

        ''' <summary>
        ''' NOT AVAILABLE, but implemented as stub method for SharedCode compatibility: Has the specified cell got a formula without a calculated value
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Public Function IsMissingCalculatedCellValueFromFormulaCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            Return False
        End Function
#Enable Warning CA1822 ' Member als statisch markieren

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
            'Me.Workbook.Worksheets(sheetName).Select()
            Me.Workbook.Worksheets(RequestedSheetIndex).Select()
            If Me.SelectedSheetName <> sheetName Then
                Throw New InvalidOperationException("Sheet selection requested for """ & sheetName & """, but after selection the active tab was """ & Workbook.Worksheets(Me.Workbook.View.ActiveTab).Name & """")
            End If
        End Sub

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetIndex"></param>
        Public Overrides Sub SelectSheet(sheetIndex As Integer)
            Me.SelectSheet(Me.SheetNames(sheetIndex))
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
        Private Shared Function ExcelColorToCssHex(color As OfficeOpenXml.Style.ExcelColor,
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

            ' 2) Theme + Tint
            If color.Theme.HasValue Then
                Dim ThemeMappedToIndex = MapThemeToIndex(color.Theme.Value)
                If ThemeMappedToIndex >= 0 Then
                    hex = DefaultOfficeTheme(ThemeMappedToIndex)
                    If Not String.IsNullOrEmpty(hex) Then
                        If Math.Abs(color.Tint) > Double.Epsilon Then
                            hex = ApplyTint(hex, color.Tint)
                        End If
                        'Add to cache if valid and return hex
                        If Not String.IsNullOrEmpty(ColorCacheKeyName) AndAlso Not String.IsNullOrEmpty(hex) AndAlso hex.Length = 7 AndAlso hex(0) = "#"c Then
                            cache(ColorCacheKeyName) = hex
                        End If
                        Return hex
                    End If
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
        Private Shared Function BuildColorCacheKey(color As OfficeOpenXml.Style.ExcelColor) As String
            If color Is Nothing Then Return Nothing

            If Not String.IsNullOrEmpty(color.Rgb) Then
                Return "rgb:" & color.Rgb
            End If

            If color.Theme.HasValue Then
                Dim idx = MapThemeToIndex(color.Theme.Value)
                If idx >= 0 Then
                    Dim tintStr = color.Tint.ToString(CultureInfo.InvariantCulture)
                    Return "theme:" & idx.ToString(CultureInfo.InvariantCulture) & ":" & tintStr
                End If
            End If

            If color.Indexed <> 0 Then
                Return "idx:" & color.Indexed.ToString(CultureInfo.InvariantCulture)
            End If

            ' Kein stabiler Key
            Return Nothing
        End Function

        ''' <summary>
        ''' Enum→Index-Mapping für Default-Theme (0..11) ---
        ''' </summary>
        ''' <param name="theme"></param>
        ''' <returns></returns>
        Private Shared Function MapThemeToIndex(theme As eThemeSchemeColor) As Integer
            Select Case theme
                Case eThemeSchemeColor.Background1 : Return 0
                Case eThemeSchemeColor.Text1 : Return 1
                Case eThemeSchemeColor.Background2 : Return 2
                Case eThemeSchemeColor.Text2 : Return 3
                Case eThemeSchemeColor.Accent1 : Return 4
                Case eThemeSchemeColor.Accent2 : Return 5
                Case eThemeSchemeColor.Accent3 : Return 6
                Case eThemeSchemeColor.Accent4 : Return 7
                Case eThemeSchemeColor.Accent5 : Return 8
                Case eThemeSchemeColor.Accent6 : Return 9
                Case eThemeSchemeColor.Hyperlink : Return 10
                Case eThemeSchemeColor.FollowedHyperlink : Return 11
                Case Else : Return -1
            End Select
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

#Disable Warning CA1034 ' Nested types should not be visible
        Public Class FormulaParserLogger
#Enable Warning CA1034 ' Nested types should not be visible
            Implements OfficeOpenXml.FormulaParsing.Logging.IFormulaParserLogger

            Public ReadOnly Property FullLog As New System.Text.StringBuilder
            Public ReadOnly Property ExceptionsLog As New System.Text.StringBuilder

            Public Sub Log(context As ParsingContext, ex As Exception) Implements IFormulaParserLogger.Log
                Me.FullLog.AppendLine("ERROR at " & context.ToString & ": " & ex.ToString)
                Me.ExceptionsLog.AppendLine("ERROR at " & context.ToString & ": " & ex.Message)
            End Sub

            Public Sub Log(context As ParsingContext, message As String) Implements IFormulaParserLogger.Log
                Me.FullLog.AppendLine("WARNING at " & context.Parser.ToString & ": " & message)
            End Sub

            Public Sub Log(message As String) Implements IFormulaParserLogger.Log
                Me.FullLog.AppendLine("WARNING: " & message)
            End Sub

            Public Sub LogCellCounted() Implements IFormulaParserLogger.LogCellCounted
                Me.FullLog.AppendLine("INFO: CellCounted")
            End Sub

            Public Sub LogFunction(func As String) Implements IFormulaParserLogger.LogFunction
                Me.FullLog.AppendLine("FUNC: " & func)
            End Sub

            Public Sub LogFunction(func As String, milliseconds As Long) Implements IFormulaParserLogger.LogFunction
                Me.FullLog.AppendLine("FUNC: " & func & " (required " & milliseconds & " ms)")
            End Sub

#Region "IDisposable Support"
            Private disposedValue As Boolean ' Dient zur Erkennung redundanter Aufrufe.

            ' IDisposable
            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not disposedValue Then
                    If disposing Then
                        ' TODO: verwalteten Zustand (verwaltete Objekte) entsorgen.
                    End If

                    ' TODO: nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalize() weiter unten überschreiben.
                    ' TODO: große Felder auf Null setzen.
                End If
                disposedValue = True
            End Sub

            ' TODO: Finalize() nur überschreiben, wenn Dispose(disposing As Boolean) weiter oben Code zur Bereinigung nicht verwalteter Ressourcen enthält.
            'Protected Overrides Sub Finalize()
            '    ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in Dispose(disposing As Boolean) weiter oben ein.
            '    Dispose(False)
            '    MyBase.Finalize()
            'End Sub

            ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
            Public Sub Dispose() Implements IDisposable.Dispose
                ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in Dispose(disposing As Boolean) weiter oben ein.
                Dispose(True)
                ' TODO: Auskommentierung der folgenden Zeile aufheben, wenn Finalize() oben überschrieben wird.
                GC.SuppressFinalize(Me)
            End Sub
#End Region
        End Class

    End Class

End Namespace