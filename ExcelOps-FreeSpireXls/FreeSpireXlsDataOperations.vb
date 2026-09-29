Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Data
Imports System.Text
Imports Spire
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports Spire.Xls.Collections

Namespace ExcelOps

    ''' <summary>
    ''' An Excel operations engine based on FreeSpire.Xls.
    ''' </summary>
    ''' <remarks>
    ''' Just as a reminder for usage of FreeSpire.Xls: the manufacturer has limited the feature set for this component. Free version is limited to 5 sheets per workbook and 150 rows per sheet. 
    ''' See https://www.e-iceblue.com/ for more details on limitations and licensing.
    ''' </remarks>
    Public Class FreeSpireXlsDataOperations
        Inherits ExcelDataOperationsBase

        ''' <inheritdoc/>
        Protected Overrides ReadOnly Property DefaultCalculationOptions As ExcelEngineDefaultOptions
            Get
                Return New ExcelEngineDefaultOptions(False, False)
            End Get
        End Property

        ''' <inheritdoc/>
        Protected Overrides ReadOnly Property AutomaticallyUpdatesFormulasAndReferencesForStructuralChanges As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' Creates or opens a workbook (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on).
        ''' </summary>
        ''' <param name="file">Path to a file which shall be loaded or null if a new workbook shall be created</param>
        ''' <param name="mode">Open an existing file or (re)create a new file</param>
        ''' <param name="options">File and engine options</param>
        ''' <remarks>
        ''' Just as a reminder for usage of FreeSpire.Xls: the manufacturer has limited the feature set for this component. Free version is limited to 5 sheets per workbook and 150 rows per sheet. 
        ''' See https://www.e-iceblue.com/ for more details on limitations and licensing.
        ''' </remarks>
        Public Sub New(file As String, mode As OpenMode, options As ExcelDataOperationsOptions)
            MyBase.New(file, mode, options)
        End Sub

        ''' <summary>
        ''' Opens a workbook.
        ''' </summary>
        ''' <param name="data">Workbook data.</param>
        ''' <param name="options">File and engine options</param>
        Public Sub New(data As Byte(), options As ExcelDataOperationsOptions)
            MyBase.New(data, options)
        End Sub

        ''' <summary>
        ''' Opens a workbook.
        ''' </summary>
        ''' <param name="data">Workbook data.</param>
        ''' <param name="options">File and engine options</param>
        Public Sub New(data As System.IO.Stream, options As ExcelDataOperationsOptions)
            MyBase.New(data, options)
        End Sub

        ''' <summary>
        ''' Creates a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on).
        ''' </summary>
        ''' <param name="file">Workbook file.</param>
        ''' <param name="mode">Mode used to open the workbook.</param>
        ''' <param name="readOnly">Whether the workbook is opened read-only.</param>
        ''' <param name="passwordForOpening">Password required to open the workbook, or <see langword="Nothing"/> when no password is required.</param>
        ''' <param name="disableInitialCalculation">Whether calculation is disabled while the workbook is loaded.</param>
        ''' <param name="disableCalculationEngine">Whether the calculation engine is disabled.</param>
        ''' <remarks>
        ''' Just as a reminder for usage of FreeSpire.Xls: the manufacturer has limited the feature set for this component. Free version is limited to 5 sheets per workbook and 150 rows per sheet. 
        ''' See https://www.e-iceblue.com/ for more details on limitations and licensing.
        ''' </remarks>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String, disableInitialCalculation As Boolean, disableCalculationEngine As Boolean)
            MyBase.New(file, mode, Not disableInitialCalculation, disableCalculationEngine, [readOnly], passwordForOpening)
        End Sub

        ''' <summary>
        ''' Creates a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on).
        ''' </summary>
        ''' <param name="file">Workbook file.</param>
        ''' <param name="mode">Mode used to open the workbook.</param>
        ''' <param name="readOnly">Whether the workbook is opened read-only.</param>
        ''' <param name="passwordForOpening">Password required to open the workbook, or <see langword="Nothing"/> when no password is required.</param>
        ''' <remarks>
        ''' Just as a reminder for usage of FreeSpire.Xls: the manufacturer has limited the feature set for this component. Free version is limited to 5 sheets per workbook and 150 rows per sheet. 
        ''' See https://www.e-iceblue.com/ for more details on limitations and licensing.
        ''' </remarks>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String)
            MyBase.New(file, mode, True, False, [readOnly], passwordForOpening)
        End Sub

        ''' <inheritdoc cref="New(Byte(), ExcelDataOperationsOptions)"/>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As Byte(), passwordForOpening As String)
            MyBase.New(data, True, False, passwordForOpening)
        End Sub

        ''' <inheritdoc cref="New(Byte(), ExcelDataOperationsOptions)"/>
        ''' <param name="disableInitialCalculation">If set to true, no initial calculation of formulas is performed when opening/loading an Excel file</param>
        ''' <param name="disableCalculationEngine">If set to true, the calculation engine is disabled and no formula calculations are performed</param>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As Byte(), passwordForOpening As String, disableInitialCalculation As Boolean, disableCalculationEngine As Boolean)
            MyBase.New(data, Not disableInitialCalculation, disableCalculationEngine, passwordForOpening)
        End Sub

        ''' <inheritdoc cref="New(System.IO.Stream, ExcelDataOperationsOptions)"/>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As System.IO.Stream, passwordForOpening As String)
            MyBase.New(data, True, False, passwordForOpening)
        End Sub

        ''' <inheritdoc cref="New(System.IO.Stream, ExcelDataOperationsOptions)"/>
        ''' <param name="disableInitialCalculation">If set to true, no initial calculation of formulas is performed when opening/loading an Excel file</param>
        ''' <param name="disableCalculationEngine">If set to true, the calculation engine is disabled and no formula calculations are performed</param>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(data As System.IO.Stream, passwordForOpening As String, disableInitialCalculation As Boolean, disableCalculationEngine As Boolean)
            MyBase.New(data, Not disableInitialCalculation, disableCalculationEngine, passwordForOpening)
        End Sub

        ''' <summary>
        ''' Creates a new workbook or creates an uninitialized instance of this Excel engine.
        ''' </summary>
        ''' <param name="mode">Mode used to open the workbook.</param>
        ''' <remarks>
        ''' Just as a reminder for usage of FreeSpire.Xls: the manufacturer has limited the feature set for this component. Free version is limited to 5 sheets per workbook and 150 rows per sheet. 
        ''' See https://www.e-iceblue.com/ for more details on limitations and licensing.
        ''' </remarks>
        Public Sub New(mode As OpenMode)
            MyBase.New(mode)
        End Sub

        ''' <summary>
        ''' Creates a new workbook or creates an uninitialized instance of this Excel engine.
        ''' </summary>
        ''' <param name="mode">Mode used to open the workbook.</param>
        ''' <param name="options">Options controlling the operation.</param>
        ''' <remarks>
        ''' Just as a reminder for usage of FreeSpire.Xls: the manufacturer has limited the feature set for this component. Free version is limited to 5 sheets per workbook and 150 rows per sheet. 
        ''' See https://www.e-iceblue.com/ for more details on limitations and licensing.
        ''' </remarks>
        Public Sub New(mode As OpenMode, options As ExcelDataOperationsOptions)
            MyBase.New(mode, options)
        End Sub

        ''' <summary>
        ''' Creates a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on).
        ''' </summary>
        ''' <param name="passwordForOpeningOnNextTime">Password required the next time the workbook is opened, or <see langword="Nothing"/> to remove the password.</param>
        ''' <remarks>
        ''' Just as a reminder for usage of FreeSpire.Xls: the manufacturer has limited the feature set for this component. Free version is limited to 5 sheets per workbook and 150 rows per sheet. 
        ''' See https://www.e-iceblue.com/ for more details on limitations and licensing.
        ''' </remarks>
        <Obsolete("Use overloaded method with ExcelDataOperationsOptions", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub New(passwordForOpeningOnNextTime As String)
            MyBase.New(True, False, True, passwordForOpeningOnNextTime)
        End Sub

        ''' <inheritdoc/>
        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "FreeSpire.Xls"
            End Get
        End Property

        ''' <inheritdoc/>
        Public Overrides Sub CopySheetContentInternal(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If targetWorkbook.GetType IsNot GetType(Spire.Xls.Workbook) Then Throw New NotSupportedException("Excel engines must be the same for source and target workbook for copying worksheets")
            'Me.Workbook.Worksheets.Copy(sheetName, targetSheetName)
            Throw New NotSupportedException("Epplus doesn't support copying of sheets with data + formats + locks")
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            targetWorkbook.ClearSheet(targetSheetName)
            Dim CopyRange As New ExcelRange(New ExcelCell(sheetName, 1, 1, ExcelCell.ValueTypes.All), New ExcelCell(sheetName, LastCell.RowIndex + 1, LastCell.ColumnIndex + 1, ExcelCell.ValueTypes.All))
            Me.Workbook.Worksheets.Item(sheetName).Range(CopyRange.LocalAddress).Copy(CType(targetWorkbook, FreeSpireXlsDataOperations).Workbook.Worksheets.Item(targetSheetName).Range(CopyRange.LocalAddress))
            'Me.Workbook.Worksheets.Item(sheetName).Cells.Copy(CType(targetWorkbook, EpplusExcelDataOperations).Workbook.Worksheets.Item(targetSheetName).Cells)
        End Sub

        ''' <summary>
        ''' Saves workbook with its sheets to HTML (including images as HTML inline data).
        ''' </summary>
        ''' <param name="fileName">Path of the target HTML file.</param>
        ''' <param name="skipHiddenSheets">Whether hidden worksheets are excluded.</param>
        ''' <remarks>Supported on Windows platforms only, e.g. Linux is known to throw TypeInitializationExceptions</remarks>
        Public Sub SaveToHtml(fileName As String, skipHiddenSheets As Boolean)
            Me._Workbook.SaveToHtml(fileName, skipHiddenSheets)
        End Sub

        ''' <summary>
        ''' Saves worksheet to HTML (including images as HTML inline data).
        ''' </summary>
        ''' <param name="worksheetName">Name of the worksheet.</param>
        ''' <param name="fileName">Path of the target HTML file.</param>
        ''' <remarks>Supported on Windows platforms only, e.g. Linux is known to throw TypeInitializationExceptions</remarks>
        Public Sub SaveWorksheetToHtml(worksheetName As String, fileName As String)
            Dim Options As New Core.Spreadsheet.HTMLOptions With {
                .ImageEmbedded = True,
                .StyleDefine = Core.Spreadsheet.HTMLOptions.StyleDefineType.Head,
                .TextMode = Core.Spreadsheet.HTMLOptions.GetText.NumberText
            }
            Dim Worksheet As Worksheet = Me._Workbook.Worksheets()(worksheetName)
            Worksheet.SaveToHtml(fileName, Options)
        End Sub

        ''' <summary>
        ''' Saves worksheet to HTML (including images as HTML inline data).
        ''' </summary>
        ''' <param name="worksheetName">Name of the worksheet.</param>
        ''' <param name="stream">Stream receiving the generated HTML.</param>
        ''' <remarks>Supported on Windows platforms only, e.g. Linux is known to throw TypeInitializationExceptions</remarks>
        Public Sub SaveWorksheetToHtml(worksheetName As String, stream As System.IO.Stream)
            Dim Options As New Core.Spreadsheet.HTMLOptions With {
                .ImageEmbedded = True,
                .StyleDefine = Core.Spreadsheet.HTMLOptions.StyleDefineType.Head,
                .TextMode = Core.Spreadsheet.HTMLOptions.GetText.NumberText
            }
            Dim Worksheet As Worksheet = Me._Workbook.Worksheets()(worksheetName)
            Worksheet.SaveToHtml(stream, Options)
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub ExportSheetToHtmlInternal(worksheetName As String, sb As StringBuilder, options As HtmlSheetExportOptions)
            Throw New NotImplementedException
        End Sub

    End Class

End Namespace
