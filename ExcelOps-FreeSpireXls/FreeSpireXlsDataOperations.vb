Option Strict On
Option Explicit On

Imports System.Data
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Collections
Imports Spire.Xls.Charts
Imports Spire

Namespace ExcelOps
    Public Class FreeSpireXlsDataOperations
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
        End Sub

        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "FreeSpire.Xls"
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
            Me.Workbook.Worksheets.Item(sheetName).Range(CopyRange.LocalAddress).Copy(CType(targetWorkbook, FreeSpireXlsDataOperations).Workbook.Worksheets.Item(targetSheetName).Range(CopyRange.LocalAddress))
            'Me.Workbook.Worksheets.Item(sheetName).Cells.Copy(CType(targetWorkbook, EpplusExcelDataOperations).Workbook.Worksheets.Item(targetSheetName).Cells)
        End Sub

    End Class

End Namespace