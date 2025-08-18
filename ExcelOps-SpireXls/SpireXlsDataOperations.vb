Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls

Namespace ExcelOps

    ''' <summary>
    ''' An Excel operations engine based on Spire.Xls
    ''' </summary>
    ''' <remarks>
    ''' Correct licensing required, see Spire.License.LicenseProvider and https://www.e-iceblue.com/ 
    ''' </remarks>
    Public Class SpireXlsDataOperations
        Inherits ExcelDataOperationsBase

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <param name="file"></param>
        ''' <param name="mode"></param>
        ''' <param name="[readOnly]"></param>
        ''' <param name="passwordForOpening"></param>
        ''' <remarks>Correct licensing required, see Spire.License.LicenseProvider and https://www.e-iceblue.com/</remarks>
        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String, disableInitialCalculation As Boolean)
            MyBase.New(file, mode, Not disableInitialCalculation, False, [readOnly], passwordForOpening)
            If AllowInstancingForNonLicencedContextForTestingPurposesOnly = False AndAlso IsLicensedContext = False Then Throw New LicenseException(GetType(Spire.License.LicenseProvider), Nothing, "Correct licensing required, see Spire.License.LicenseProvider and https://www.e-iceblue.com/")
        End Sub

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <param name="file"></param>
        ''' <param name="mode"></param>
        ''' <param name="[readOnly]"></param>
        ''' <param name="passwordForOpening"></param>
        ''' <remarks>Correct licensing required, see Spire.License.LicenseProvider and https://www.e-iceblue.com/</remarks>
        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String)
            MyBase.New(file, mode, True, False, [readOnly], passwordForOpening)
            If AllowInstancingForNonLicencedContextForTestingPurposesOnly = False AndAlso IsLicensedContext = False Then Throw New LicenseException(GetType(Spire.License.LicenseProvider), Nothing, "Correct licensing required, see Spire.License.LicenseProvider and https://www.e-iceblue.com/")
        End Sub

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <remarks>Correct licensing required, see Spire.License.LicenseProvider and https://www.e-iceblue.com/</remarks>
        Public Sub New()
            Me.New(Nothing)
        End Sub

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <param name="passwordForOpeningOnNextTime"></param>
        ''' <remarks>Correct licensing required, see Spire.License.LicenseProvider and https://www.e-iceblue.com/</remarks>
        Public Sub New(passwordForOpeningOnNextTime As String)
            MyBase.New(True, False, True, passwordForOpeningOnNextTime)
            If AllowInstancingForNonLicencedContextForTestingPurposesOnly = False AndAlso IsLicensedContext = False Then Throw New LicenseException(GetType(Spire.License.LicenseProvider), Nothing, "Correct licensing required, see Spire.License.LicenseProvider and https://www.e-iceblue.com/")
        End Sub

        ''' <summary>
        ''' Allow instancing of engine in (OneTime)Setup methods of unit tests
        ''' </summary>
        ''' <returns></returns>
        Friend Shared Property AllowInstancingForNonLicencedContextForTestingPurposesOnly As Boolean = False

#Disable Warning CA1822 ' Mark members as static
        ''' <summary>
        ''' Is a valid Spire.Xls license assigned
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsLicensedContext As Boolean
#Enable Warning CA1822 ' Mark members as static
            Get
                Return Utils.IsLicensedContext
            End Get
        End Property

        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "Spire.Xls"
            End Get
        End Property

        Public Overrides Sub CopySheetContentInternal(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If targetWorkbook.GetType IsNot GetType(Spire.Xls.Workbook) Then Throw New NotSupportedException("Excel engines must be the same for source and target workbook for copying worksheets")
            'Me.Workbook.Worksheets.Copy(sheetName, targetSheetName)
            Throw New NotSupportedException("Epplus doesn't support copying of sheets with data + formats + locks")
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            targetWorkbook.ClearSheet(targetSheetName)
            Dim CopyRange As New ExcelRange(New ExcelCell(sheetName, 1, 1, ExcelCell.ValueTypes.All), New ExcelCell(sheetName, LastCell.RowIndex + 1, LastCell.ColumnIndex + 1, ExcelCell.ValueTypes.All))
            Me.Workbook.Worksheets.Item(sheetName).Range(CopyRange.LocalAddress).Copy(CType(targetWorkbook, SpireXlsDataOperations).Workbook.Worksheets.Item(targetSheetName).Range(CopyRange.LocalAddress))
            'Me.Workbook.Worksheets.Item(sheetName).Cells.Copy(CType(targetWorkbook, EpplusExcelDataOperations).Workbook.Worksheets.Item(targetSheetName).Cells)
        End Sub

        ''' <summary>
        ''' Save workbook with its sheets to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="fileName"></param>
        ''' <param name="skipHiddenSheets"></param>
        ''' <remarks>Supported on Windows platforms only, e.g. Linux is known to throw TypeInitializationExceptions</remarks>
        Public Sub SaveToHtml(fileName As String, skipHiddenSheets As Boolean)
            Me._Workbook.SaveToHtml(fileName, skipHiddenSheets)
        End Sub

        ''' <summary>
        ''' Save worksheet to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="worksheetName"></param>
        ''' <param name="fileName"></param>
        ''' <remarks>Supported on Windows platforms only, e.g. Linux is known to throw TypeInitializationExceptions</remarks>
        Public Sub SaveWorksheetToHtml(worksheetName As String, fileName As String)
            Dim Options As New Core.Spreadsheet.HTMLOptions With {
                .ImageEmbedded = True,
                .IsFixedTableColWidth = False,
                .StyleDefine = Core.Spreadsheet.HTMLOptions.StyleDefineType.Head,
                .TextMode = Core.Spreadsheet.HTMLOptions.GetText.NumberText
            }
            Dim Worksheet As Worksheet = Me._Workbook.Worksheets()(worksheetName)
            Worksheet.SaveToHtml(fileName, Options)
        End Sub

        ''' <summary>
        ''' Save worksheet to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="worksheetName"></param>
        ''' <param name="stream"></param>
        ''' <remarks>Supported on Windows platforms only, e.g. Linux is known to throw TypeInitializationExceptions</remarks>
        Public Sub SaveWorksheetToHtml(worksheetName As String, stream As System.IO.Stream)
            Dim Options As New Core.Spreadsheet.HTMLOptions With {
                .ImageEmbedded = True,
                .IsFixedTableColWidth = False,
                .StyleDefine = Core.Spreadsheet.HTMLOptions.StyleDefineType.Head,
                .TextMode = Core.Spreadsheet.HTMLOptions.GetText.NumberText
            }
            Dim Worksheet As Worksheet = Me._Workbook.Worksheets()(worksheetName)
            Worksheet.SaveToHtml(stream, Options)
        End Sub

        ''' <summary>
        ''' Save worksheet to HTML (including images as HTML inline data)
        ''' </summary>
        ''' <param name="worksheetName"></param>
        ''' <param name="sb"></param>
        Protected Overrides Sub ExportSheetToHtmlInternal(worksheetName As String, sb As StringBuilder, options As HtmlSheetExportOptions)
            Throw New NotImplementedException
        End Sub

    End Class

End Namespace