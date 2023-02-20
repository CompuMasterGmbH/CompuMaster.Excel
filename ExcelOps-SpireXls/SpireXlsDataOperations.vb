Option Strict On
Option Explicit On

Imports System.ComponentModel

Namespace ExcelOps
    Public Class SpireXlsDataOperations
        Inherits ExcelDataOperationsBase

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <param name="file"></param>
        ''' <param name="mode"></param>
        ''' <param name="[readOnly]"></param>
        ''' <param name="passwordForOpening"></param>
        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String)
            MyBase.New(file, mode, True, False, [readOnly], passwordForOpening)
            If AllowInstancingForNonLicencedContextForTestingPurposesOnly = False AndAlso IsLicensedContext = False Then Throw New LicenseException(GetType(Spire.License.LicenseProvider))
        End Sub

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        Public Sub New()
            Me.New(Nothing)
        End Sub

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <param name="passwordForOpeningOnNextTime"></param>
        Public Sub New(passwordForOpeningOnNextTime As String)
            MyBase.New(True, False, True, passwordForOpeningOnNextTime)
            If AllowInstancingForNonLicencedContextForTestingPurposesOnly = False AndAlso IsLicensedContext = False Then Throw New LicenseException(GetType(Spire.License.LicenseProvider))
        End Sub

        ''' <summary>
        ''' Allow instancing of engine in (OneTime)Setup methods of unit tests
        ''' </summary>
        ''' <returns></returns>
        Friend Shared Property AllowInstancingForNonLicencedContextForTestingPurposesOnly As Boolean = False

        ''' <summary>
        ''' Is a valid Spire.Xls license assigned
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsLicensedContext As Boolean
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

    End Class

End Namespace