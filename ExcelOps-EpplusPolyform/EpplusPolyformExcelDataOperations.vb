Option Strict On
Option Explicit On

Imports System.Data
Imports System.ComponentModel
Imports OfficeOpenXml
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

        ''' <summary>
        ''' The license context for Epplus (see its polyform license)
        ''' </summary>
        ''' <remarks>https://epplussoftware.com/en/LicenseOverview/LicenseFAQ</remarks>
        ''' <returns></returns>
        Public Shared Property LicenseContext As OfficeOpenXml.LicenseContext?
            Get
                Return OfficeOpenXml.ExcelPackage.LicenseContext
            End Get
            Set(value As OfficeOpenXml.LicenseContext?)
                OfficeOpenXml.ExcelPackage.LicenseContext = value
            End Set
        End Property

        Private Shared Sub ValidateLicenseContext(instance As EpplusPolyformExcelDataOperations)
            If LicenseContext.HasValue = False Then
                Throw New System.ComponentModel.LicenseException(GetType(EpplusPolyformExcelDataOperations), instance, NameOf(LicenseContext) & " must be assigned before creating instances")
            End If
        End Sub

        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String, disableInitialCalculation As Boolean)
            MyBase.New(file, mode, Not disableInitialCalculation, False, [readOnly], passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String)
            MyBase.New(file, mode, True, False, [readOnly], passwordForOpening)
            ValidateLicenseContext(Me)
        End Sub

        Public Sub New()
            Me.New(Nothing)
        End Sub

        Public Sub New(passwordForOpeningOnNextTime As String)
            MyBase.New(False, True, True, passwordForOpeningOnNextTime)
            ValidateLicenseContext(Me)
        End Sub

        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "Epplus (Polyform license edition)"
            End Get
        End Property

        Private Const FULL_CALC_ON_LOAD As Boolean = True

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

    End Class

End Namespace