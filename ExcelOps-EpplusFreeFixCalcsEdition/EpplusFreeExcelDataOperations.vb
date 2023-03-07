Option Strict On
Option Explicit On

Imports System.Data
Imports System.ComponentModel
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
            If Me.Workbook.Worksheets(cell.SheetName) Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & cell.SheetName, NameOf(cell))
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
            If Me.Workbook.Worksheets(sheetName) Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & sheetName, NameOf(sheetName))
            Dim CheckResult As Boolean = Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).HasMissingCachedCalculatedFormulaValue()
            If CheckResult = True AndAlso Tools.IsFormulaWithoutCellReferences(Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Formula) Then
                Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Calculate
                CheckResult = False
            End If
            Return CheckResult
        End Function

    End Class

End Namespace