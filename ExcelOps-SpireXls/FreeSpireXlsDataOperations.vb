﻿Option Strict On
Option Explicit On

'TODO: FIX COMPILE ERRORS
'TODO: FIX ExcelCell.LocalCellAddress(x,y) -> ExcelCell.LocalCellAddress(x-1,y-1)
'TODO: ADD NUNIT TESTS
'TODO: BACKPORT: ExcelCell.LocalAddress -> Public!
'TODO: BACKPORT: ExcelRange.LocalAddress 
'TODO: NUGET-PACKAGE: ExcelOps-Projects

Imports System.Data
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Collections
Imports Spire.Xls.Charts
Imports Spire

Namespace ExcelOps
    Public Class FreeSpireXlsDataOperations
        Inherits ExcelDataOperationsBase

        Public Sub New(file As String, mode As OpenMode, [readOnly] As Boolean, passwordForOpening As String)
            MyBase.New(file, mode, True, False, [readOnly], passwordForOpening)
        End Sub

        Public Sub New()
            Me.New(Nothing)
        End Sub

        Public Sub New(passwordForOpeningOnNextTime As String)
            MyBase.New(True, False, True, passwordForOpeningOnNextTime)
        End Sub

        Private _Workbook As Spire.Xls.Workbook
        Public ReadOnly Property Workbook As Spire.Xls.Workbook
            Get
                If Me._Workbook Is Nothing Then
                    Throw New InvalidOperationException("Workbook has already been closed")
                End If
                Return Me._Workbook
            End Get
        End Property

        Public Overrides Sub Close()
            If Me.IsClosed = False Then Me._Workbook.Dispose()
            Me._Workbook = Nothing
        End Sub

        Public Overrides ReadOnly Property IsClosed As Boolean
            Get
                Return Me._Workbook Is Nothing
            End Get
        End Property

        Protected Overrides Sub SaveInternal()
            Me._Workbook.SaveToFile(Me.FilePath) 'NOTE: _Workbook.Save is forbidden since the file path might have changed in background due to a workaround required for RemoveVbaProject
        End Sub

        Protected Overrides Sub SaveInternal_ApplyCachedCalculationOption(cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            MyBase.SaveInternal_ApplyCachedCalculationOption(cachedCalculationsOption)
        End Sub

        Protected Overrides Sub SaveAsInternal(fileName As String, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            Me._Workbook.SaveToFile(fileName)
            Me.SetWorkbookFilePath(New System.IO.FileInfo(fileName).FullName)
        End Sub

        ''' <summary>
        ''' Setter for workbook filename in SpireXls
        ''' </summary>
        ''' <param name="fileName"></param>
        ''' <remarks>WORKAROUND FEATURE: required for methods
        ''' <list type="bullet">
        ''' <item>RemoveVbaProject() to reset filepath in workbook</item>
        ''' <item>SaveAs() to set the filepath in workbook also for 2nd and following times (original manufacturer logic sets the file path once and never updates it any more!)</item>
        ''' </list></remarks>
        Private Sub SetWorkbookFilePath(fileName As String)
            Dim XlsWorkbookMembers = CompuMaster.Reflection.NonPublicInstanceMembers.GetMembers(Of System.Reflection.FieldInfo)(Me.Workbook.GetType, GetType(Spire.Xls.Core.Spreadsheet.XlsWorkbook))
            If XlsWorkbookMembers.Count <> 1 Then
                Throw New NotSupportedException("Spire.Xls incompatibility, please open an issue at https://github.com/CompuMasterGmbH/CompuMaster.Excel")
            End If
            Dim XlsWb = CompuMaster.Reflection.NonPublicInstanceMembers.InvokeFieldGet(Of Spire.Xls.Core.Spreadsheet.XlsWorkbook)(Me.Workbook, Me.Workbook.GetType, XlsWorkbookMembers(0).Name)
            Dim pi = CompuMaster.Reflection.PublicInstanceMembers.GetMembers(Of System.Reflection.PropertyInfo)(XlsWb.GetType, "FullFileName")
            Dim p = pi.GetSetMethod(True) 'NOTE: property-setter is non-public, while the property and its getter is public!
            p.Invoke(XlsWb, New Object() {CompuMaster.Data.Utils.StringNotEmptyOrAlternativeValue(fileName, "/")}) 'WORKAROUND: Setter indirectly calls System.IO.Path.GetDirectoryName which crashed on empty string -> requires property WorkbookFileName to consider "/" as null/Nothing
        End Sub

        ''' <summary>
        ''' The current workbook filename as reported by the Spire.Xls engine
        ''' </summary>
        ''' <remarks>WARNING: The file path might not reflect the expected value because it changed in background due to a workaround required for <see cref="RemoveVbaProject"/></remarks>
        ''' <returns></returns>
        Protected Overrides ReadOnly Property WorkbookFilePath As String
            Get
                If Me.IsClosed Then
                    Return Nothing
                ElseIf Me.Workbook.FileName = "/" Then 'WORKAROUND DEPENDENCY: required for RemoveVbaProject to reset workbook filename to blank value
                    Return Nothing
                Else
                    Return CompuMaster.Data.Utils.StringNotEmptyOrNothing(Me.Workbook.FileName)
                End If
            End Get
        End Property

        Public Overrides Function SheetNames() As List(Of String)
            Dim Result As New List(Of String)
            For MyCounter As Integer = 0 To Me.Workbook.Worksheets.Count - 1
                Result.Add(Me.Workbook.Worksheets(0).Name)
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Remove all named ranges in Excel Workbook since this feature is involved but not actively used in Master Template V15; but this feature might throw exceptions in EPPlus when removing rows
        ''' </summary>
        Public Overrides Sub CleanupRangeNames()
            Dim NamesToRemove As New List(Of String)
            For Each namedRange As Spire.Xls.Core.INamedRange In Me.Workbook.NameRanges
                NamesToRemove.Add(namedRange.Name)
            Next
            For Each Name As String In NamesToRemove
                Me.Workbook.NameRanges.Remove(Name)
            Next

            For Each worksheet In Me.Workbook.Worksheets
                Dim NamesInWorksheetToRemove As New List(Of String)
                For Each namedRange As Spire.Xls.Core.INamedRange In worksheet.Names
                    NamesInWorksheetToRemove.Add(namedRange.Name)
                Next
                For Each Name As String In NamesInWorksheetToRemove
                    worksheet.Names.Remove(Name)
                Next
            Next
        End Sub

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public Overrides Function LookupCellValue(Of T)(cell As ExcelCell) As T
            Try
                Select Case GetType(T)
                    Case GetType(Integer), GetType(Long), GetType(Byte), GetType(Double), GetType(Decimal), GetType(Boolean)
                        If Trim(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).DisplayedText) = Nothing Then
                            Return Nothing
                        Else
                            Return CType(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).Value2, T)
                        End If
                    Case GetType(String)
                        If Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).HasError Then
                            Return CType(CType(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).ErrorValue, Object), T)
                        Else
                            Return CType(CType(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).DisplayedText, Object), T)
                        End If
                    Case GetType(Object)
                        Return CType(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).Value2, T)
                    Case Else
                        Return CType(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).Value2, T)
                End Select
            Catch ex As InvalidCastException
                Throw New System.FormatException("Value """ & Me.LookupCellFormattedText(cell) & """ in cell """ & cell.Address(True) & """ can't be converted to " & GetType(T).Name, ex)
            Catch ex As FormatException
                Throw New System.FormatException("Value """ & Me.LookupCellFormattedText(cell) & """ in cell """ & cell.Address(True) & """ can't be converted to " & GetType(T).Name, ex)
            End Try
        End Function

        ''' <summary>
        ''' Read the cell format string
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Public Overrides Function LookupCellFormat(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Utils.StringNotEmptyOrNothing(Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Style.NumberFormat)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public Overrides Function LookupCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.LookupCellValue(Of T)(New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All))
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public Overrides Function TryLookupCellValue(Of T As Structure)(cell As ExcelCell) As T?
            Try
                Select Case GetType(T)
                    Case GetType(Integer), GetType(Long), GetType(Byte), GetType(Double), GetType(Decimal), GetType(Boolean)
                        If Trim(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).DisplayedText) = Nothing Then
                            Return Nothing
                        Else
                            Return CType(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).Value2, T)
                        End If
                    Case Else
                        Return CType(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).Value2, T)
                End Select
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public Overrides Function TryLookupCellValue(Of T As Structure)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T?
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.TryLookupCellValue(Of T)(New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All))
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public Overrides Function LookupCellValueAsObject(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Object
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If Me.Workbook.Worksheets.Item(sheetName) Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & sheetName, NameOf(sheetName))
            Return Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Value2
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Public Overrides Function LookupCellFormula(cell As ExcelCell) As String
            If Me.Workbook.Worksheets.Item(cell.SheetName) Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & cell.SheetName, NameOf(cell))
            Return Utils.StringNotEmptyOrNothing(Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).Formula)
        End Function

        ''' <summary>
        ''' Read a cell formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <returns></returns>
        Public Overrides Function LookupCellFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If Me.Workbook.Worksheets.Item(sheetName) Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & sheetName, NameOf(sheetName))
            If rowIndex < 0 Then Throw New ArgumentOutOfRangeException("RowIndex " & rowIndex & " must be equal or bigger than 0", NameOf(rowIndex))
            If columnIndex < 0 Then Throw New ArgumentOutOfRangeException("ColumnIndex " & columnIndex & " must be equal or bigger than 0", NameOf(columnIndex))
            Return Utils.StringNotEmptyOrNothing(Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Formula)
        End Function

        Public Overrides Function LookupCellIsLocked(cell As ExcelCell) As Boolean
            Return Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.RowNumber, cell.ColumnNumber).Style.Locked
        End Function

        Public Overrides Function LookupCellIsLocked(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Style.Locked
        End Function

        ''' <summary>
        ''' Write a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <param name="value"></param>
        Public Overrides Sub WriteCellValue(Of T)(cell As ExcelCell, value As T)
            Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.Address).Formula = Nothing
            Me.Workbook.Worksheets.Item(cell.SheetName).Range(cell.Address).Value2 = value
        End Sub

        ''' <summary>
        ''' Write a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <param name="value"></param>
        Public Overrides Sub WriteCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer, value As T)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Formula = Nothing
            Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Value2 = value
        End Sub

        ''' <summary>
        ''' Write a cell formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        ''' <param name="formula"></param>
        Public Overrides Sub WriteCellFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer, formula As String, immediatelyCalculateCellValue As Boolean)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Formula = formula
            If immediatelyCalculateCellValue Then
                Me.RecalculateCell(sheetName, rowIndex, columnIndex)
            Else
                Me.RecalculationRequired = True
            End If
        End Sub

        ''' <summary>
        ''' Recalculate a cell based on its formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        Public Overrides Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.RecalculateCell(sheetName, rowIndex, columnIndex, True)
        End Sub

        ''' <summary>
        ''' Recalculate a cell based on its formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        Public Overloads Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer, throwExceptionOnCalculationError As Boolean)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).CalculateAllValue()
            If throwExceptionOnCalculationError AndAlso Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).HasError Then
                Dim Cell As New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
                Throw New NotSupportedException("Epplus calculation at " & Cell.Address(True) & " resulted in #" & UCase(Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).ErrorValue) & "!" & " for formula =" & Me.LookupCellFormula(Cell))
            End If
        End Sub

        ''' <summary>
        ''' Recalculate all cells of a sheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Sub RecalculateSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If Me.CalculationModuleDisabled Then Throw New InvalidOperationException("Calculation engine is disabled, requested recalculation failed")
            Me.Workbook.Worksheets.Item(sheetName).CalculateAllValue()
        End Sub

        Protected Overrides Sub CreateWorkbook()
            Me._Workbook = New Spire.Xls.Workbook
            Me.Workbook.Worksheets(0).Name = "Sheet1"
        End Sub

        Protected Overrides Sub LoadWorkbook(file As System.IO.FileInfo)
            Me._Workbook = New Spire.Xls.Workbook
            Me.Workbook.LoadFromFile(file.FullName)
        End Sub

        ''' <summary>
        ''' Lookup the last content column index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastColumnIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As Worksheet = Me.Workbook.Worksheets.Item(sheetName)
            If Sheet Is Nothing Then Throw New NullReferenceException("Sheet """ & sheetName & """ doesn't exist in workbook")
            Return Sheet.LastColumn - 1
        End Function

        ''' <summary>
        ''' Lookup the last content column index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastContentColumnIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As Worksheet = Me.Workbook.Worksheets.Item(sheetName)
            If Sheet Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & sheetName, NameOf(sheetName))
            Dim autoSuggestionLastRowIndex As Integer = Sheet.LastRow - 1
            Dim autoSuggestedResult As Integer = Sheet.LastColumn - 1
            For colCounter As Integer = autoSuggestedResult To 0 Step -1
                For rowCounter As Integer = 0 To autoSuggestionLastRowIndex
                    If IsEmptyCell(Sheet, rowCounter, colCounter) = False Then
                        Return colCounter
                    End If
                Next
            Next
            Return 0
        End Function

        ''' <summary>
        ''' Lookup the last content row index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastRowIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As Worksheet = Me.Workbook.Worksheets.Item(sheetName)
            If Sheet Is Nothing Then Throw New NullReferenceException("Sheet """ & sheetName & """ doesn't exist in workbook")
            Return Sheet.LastRow - 1
        End Function

        ''' <summary>
        ''' Lookup the last content row index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastContentRowIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As Worksheet = Me.Workbook.Worksheets.Item(sheetName)
            If Sheet Is Nothing Then Throw New ArgumentException("Must be an existing sheet name: """ & sheetName & """", NameOf(sheetName))
            Dim autoSuggestionLastColumnIndex As Integer = Sheet.LastColumn - 1
            Dim autoSuggestedResult As Integer = Sheet.LastRow - 1
            For rowCounter As Integer = autoSuggestedResult To 0 Step -1
                For colCounter As Integer = 0 To autoSuggestionLastColumnIndex
                    If IsEmptyCell(Sheet, rowCounter, colCounter) = False Then
                        Return rowCounter
                    End If
                Next
            Next
            Return 0
        End Function

        ''' <summary>
        ''' Lookup the last content cell (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastCell(sheetName As String) As ExcelOps.ExcelCell
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim CellRowIndex As Integer = Me.LookupLastRowIndex(sheetName)
            Dim CellColIndex As Integer = Me.LookupLastColumnIndex(sheetName)
            Return New ExcelOps.ExcelCell(sheetName, CellRowIndex, CellColIndex, Nothing)
        End Function

        ''' <summary>
        ''' Lookup the last content cell (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastContentCell(sheetName As String) As ExcelOps.ExcelCell
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim CellRowIndex As Integer = Me.LookupLastContentRowIndex(sheetName)
            Dim CellColIndex As Integer = Me.LookupLastContentColumnIndex(sheetName)
            Return New ExcelOps.ExcelCell(sheetName, CellRowIndex, CellColIndex, Nothing)
        End Function

        ''' <summary>
        ''' Lookup the row index (zero based index)
        ''' </summary>
        ''' <param name="cell"></param>
        Public Overrides Function LookupRowIndex(cell As ExcelOps.ExcelCell) As Integer
            Return cell.RowIndex
        End Function

        ''' <summary>
        ''' Lookup the column index (zero based index)
        ''' </summary>
        ''' <param name="cell"></param>
        Public Overrides Function LookupColumnIndex(cell As ExcelOps.ExcelCell) As Integer
            Return cell.ColumnIndex
        End Function

        ''' <summary>
        ''' Remove specified rows
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="startrowIndex">0-based row number</param>
        ''' <param name="rows">Number of rows to remove</param>
        Public Overrides Sub RemoveRows(sheetName As String, startRowIndex As Integer, rows As Integer)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If rows < 0 Then Throw New ArgumentOutOfRangeException(NameOf(rows), "Row number must be a positive value or zero")
            If rows = 0 Then Return
            Dim Sheet As Worksheet = Me.Workbook.Worksheets.Item(sheetName)
            If Sheet.Type <> ExcelSheetType.NormalWorksheet Then Throw New Exception("Specified worksheet is not a data worksheet")
            Sheet.DeleteRow(startRowIndex + 1, rows)
        End Sub

        ''' <summary>
        ''' Clear cell content
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rangeFirstCell"></param>
        ''' <param name="rangeLastCell"></param>
        Public Overrides Sub ClearCells(sheetName As String, rangeFirstCell As ExcelCell, rangeLastCell As ExcelCell)
            If rangeFirstCell.SheetName <> rangeLastCell.SheetName Then Throw New ArgumentException("Cells must be member of the same worksheet")
            If sheetName <> Nothing Then
                rangeFirstCell = rangeFirstCell.Clone
                rangeLastCell = rangeLastCell.Clone
                rangeFirstCell.SheetName = sheetName
                rangeLastCell.SheetName = sheetName
            End If
            If rangeFirstCell.SheetName = Nothing Then Throw New ArgumentNullException(NameOf(rangeFirstCell))
            Dim ws = Me.Workbook.Worksheets.Item(rangeFirstCell.SheetName)
            ws.Range(rangeFirstCell.Address & ":" & rangeLastCell.Address).ClearAll()
        End Sub

        ''' <summary>
        ''' Determine if a cell contains empty content
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">Zero-based index</param>
        ''' <param name="columnIndex">Zero-based index</param>
        Public Overrides Function IsEmptyCell(sheetName As String, ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet = Me.Workbook.Worksheets.Item(sheetName)
            Return IsEmptyCell(Sheet, rowIndex, columnIndex)
        End Function

        ''' <summary>
        ''' Determine if a cell contains empty content (cells with formulas are always considered as filled cells)
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="rowIndex">Zero-based index</param>
        ''' <param name="columnIndex">Zero-based index</param>
        Private Overloads Function IsEmptyCell(ByVal sheet As Worksheet, ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Boolean
            If sheet.Range(rowIndex + 1, columnIndex + 1).Formula <> Nothing Then
                Return False
            End If
            Dim value As Object = sheet.Range(rowIndex + 1, columnIndex + 1).Value
            If value Is Nothing Then
                Return True
            ElseIf value.GetType Is GetType(String) AndAlso CType(value, String) = Nothing Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Try to lookup the cell's value to a string anyhow
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function LookupCellFormattedText(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Try
                Dim CellValue As Object = Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Value2
                If CellValue IsNot Nothing AndAlso CellValue.GetType() Is GetType(Boolean) AndAlso Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Style.NumberFormat = "General" Then
                    Return Me.LookupCellValue(Of String)(sheetName, rowIndex, columnIndex)
                Else
                    Return Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).Text
                End If
            Catch ex As InvalidCastException
                If Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).HasError Then
                    Return Me.Workbook.Worksheets.Item(sheetName).Range(rowIndex + 1, columnIndex + 1).ErrorValue
                Else
                    Throw
                End If
            Catch ex As Exception
                Return "#ERROR: " & ex.Message
            End Try
        End Function

        Public Overrides Sub UnprotectSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Item(sheetName).Unprotect()
        End Sub

        Public Overrides Sub ProtectSheet(sheetName As String, level As ProtectionLevel)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Select Case level
                Case ProtectionLevel.StandardWithInsertDeleteRows
                    Me.Workbook.Worksheets.Item(sheetName).Protect(Nothing, SheetProtectionType.All)
                Case ProtectionLevel.Standard
                    Me.Workbook.Worksheets.Item(sheetName).Protect(Nothing, SheetProtectionType.LockedCells Or SheetProtectionType.UnLockedCells)
                Case ProtectionLevel.SelectAndEditUnlockedCellsOnly
                    Me.Workbook.Worksheets.Item(sheetName).Protect(Nothing, SheetProtectionType.UnLockedCells)
                Case ProtectionLevel.SelectAndEditAllCellsButNoFurtherEditing
                    Me.Workbook.Worksheets.Item(sheetName).Protect(Nothing, SheetProtectionType.LockedCells Or SheetProtectionType.UnLockedCells)
                Case ProtectionLevel.SelectNoCellsAndNoEditing
                    Me.Workbook.Worksheets.Item(sheetName).Protect(Nothing, SheetProtectionType.None)
                Case Else
                    Throw New NotImplementedException
            End Select
        End Sub

        Protected Overrides Sub RecalculateAllInternal()
            'Throw New NotSupportedException("Epplus can't successfully calculate all formulas")
            If Me.CalculationModuleDisabled Then Throw New InvalidOperationException("Calculation engine is disabled, requested recalculation failed")
            Me.Workbook.CalculateAllValue()
        End Sub

        ''' <summary>
        ''' Is the Excel engine allowed to automatically/continuously calculate on every change or does the user has to manually force a recalculation (typically by pressing F9 key in MS Excel)
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Property AutoCalculationEnabled As Boolean
            Get
                Return (Me.Workbook.CalculationMode = ExcelCalculationMode.Auto)
            End Get
            Set(value As Boolean)
                If value Then
                    Me.Workbook.CalculationMode = ExcelCalculationMode.Auto
                Else
                    Me.Workbook.CalculationMode = ExcelCalculationMode.Manual
                End If
            End Set
        End Property

        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "Epplus"
            End Get
        End Property

        Public Overrides Function IsProtectedSheet(sheetName As String) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.Workbook.Worksheets.Item(sheetName).ProtectContents
        End Function

        Public Overrides Sub RemoveSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Remove(sheetName)
        End Sub

        Public Overrides Sub AddSheet(sheetName As String, beforeSheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim ws As Worksheet = Me.Workbook.Worksheets.Add(sheetName)
            If beforeSheetName <> Nothing Then
                Dim OldIndex As Integer = Me.Workbook.Worksheets(sheetName).Index
                Dim NewIndex As Integer = Me.Workbook.Worksheets(beforeSheetName).Index
                Me.Workbook.Worksheets.Move(oldIndex, newIndex)
            End If
        End Sub

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Sub SelectSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Item(sheetName).Select()
        End Sub

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetIndex"></param>
        Public Overrides Sub SelectSheet(sheetIndex As Integer)
            Me.SelectSheet(Me.SheetNames(sheetIndex))
        End Sub

        Public Overrides Sub CloseExcelAppInstance()
            Me.Close()
            'No external excel engine application to close
        End Sub

        Public Overrides Sub UnhideSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Item(sheetName).Visibility = WorksheetVisibility.Visible
        End Sub

        Public Overrides Sub HideSheet(sheetName As String, stronglyHide As Boolean)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If stronglyHide Then
                Me.Workbook.Worksheets.Item(sheetName).Visibility = WorksheetVisibility.StrongHidden
            Else
                Me.Workbook.Worksheets.Item(sheetName).Visibility = WorksheetVisibility.Hidden
            End If
        End Sub

        Public Overrides Function IsHiddenSheet(sheetName As String) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.Workbook.Worksheets.Item(sheetName).Visibility <> WorksheetVisibility.Visible
        End Function

        Public Sub WriteTableToSheet(dataTable As DataTable, sheetName As String)

            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim WorkSheet = Me.Workbook.Worksheets.Item(sheetName)
            WorkSheet.Clear()

            'Paste the column headers
            For ColCounter As Integer = 0 To dataTable.Columns.Count - 1
                Dim headline As String = dataTable.Columns(ColCounter).ColumnName
                WorkSheet.Range(1, ColCounter + 1).Value = headline
                WorkSheet.Range(1, ColCounter + 1).Style.Font.IsBold = True
                WorkSheet.Range(1, ColCounter + 1).Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Medium
                WorkSheet.Range(1, ColCounter + 1).Borders(BordersLineType.EdgeBottom).Color = System.Drawing.Color.FromArgb(0, 0, 0)
            Next

            'Fehlerwert Rückgabe von FEHLER.TYP 
            '#NULL! 1 
            '#DIV/0! 2  --> NaN
            '#VALUE! 3 
            '#REF! 4 
            '#NAME? 5
            '#NUM! 6   --> Infinity (Positive/Negative)
            '#NA 7      
            'Sonstiges #NA 
            '{blank}    --> DBNull

            'Paste the data from the datatable
            For RowCounter As Integer = 0 To dataTable.Rows.Count - 1
                For ColCounter As Integer = 0 To dataTable.Columns.Count - 1
                    Dim value As Object = dataTable.Rows(RowCounter)(ColCounter)
                    If IsDBNull(value) Then
                        WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value = Nothing
                    ElseIf value.GetType Is GetType(String) Then
                        'Excel requires line-breaks to be an LF character only, not a windows typical CR+LF
                        If CType(value, String) <> Nothing Then
                            value = Replace(CType(value, String), ControlChars.CrLf, ControlChars.Lf) 'Windows line breaks
                            value = Replace(CType(value, String), ControlChars.Cr, ControlChars.Lf) 'Mac or Linux line break
                        End If
                        WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Formula = ""
                        WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value = CType(value, String)
                    ElseIf value.GetType Is GetType(DateTime) Then
                        Dim datevalue As DateTime = CType(value, DateTime)
                        Try
                            'Re-create datevalue to strip off any other additional properties
                            datevalue = New DateTime(datevalue.Year, datevalue.Month, datevalue.Day, datevalue.Hour, datevalue.Minute, datevalue.Second, datevalue.Millisecond)
                            'Write back the new cell value
                            If datevalue = New DateTime Then
                                WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = Nothing
                            Else
                                'WorkSheet.Workbook.DateTimeToNumber(datevalue)
                                WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = datevalue
                                If dataTable.Columns(ColCounter).ExtendedProperties.ContainsKey("Format") Then
                                    WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Style.NumberFormat = CType(dataTable.Columns(ColCounter).ExtendedProperties("Format"), String)
                                Else
                                    WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Style.NumberFormat = "yyyy-MM-dd HH:mm:ss"
                                End If
                            End If
                        Catch ex As Exception
                            Throw New Exception("Error writing a date/time value """ & datevalue.ToString & """ in row " & (RowCounter + 1), ex)
                        End Try
                    ElseIf value.GetType Is GetType(Decimal) Then
                        Dim decimalValue As Decimal = CType(value, Decimal)
                        WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = decimalValue
                    ElseIf value.GetType Is GetType(Double) Then
                        Dim doubleValue As Double = CType(value, Double)
                        If doubleValue = Double.PositiveInfinity Then
                            WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = value 'OfficeOpenXml.ExcelErrorValue.Values.Num
                        ElseIf doubleValue = Double.NegativeInfinity Then
                            WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = value 'OfficeOpenXml.ExcelErrorValue.Values.Num
                        ElseIf Double.IsNaN(doubleValue) Then
                            WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = value 'OfficeOpenXml.ExcelErrorValue.Values.Div0
                        ElseIf Double.Epsilon = doubleValue Then
                            'too small number would be rounded to just 0
                            WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = value 'OfficeOpenXml.ExcelErrorValue.Values.Num
                        Else
                            WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = CType(value, Double)
                        End If
                    ElseIf value.GetType Is GetType(Int16) OrElse value.GetType Is GetType(Int32) Then
                        WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = CType(value, Int32)
                    ElseIf value.GetType Is GetType(Int64) Then
                        WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = CType(value, Double)
                    ElseIf value.GetType Is GetType(Boolean) Then
                        WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value2 = CType(value, Boolean)
                    Else
                        WorkSheet.Range(RowCounter + 1 + 1, ColCounter + 1).Value = CType(value, Object).ToString
                    End If
                Next
            Next

            ' Auto size all worksheet columns which contain data
            For MyCounter As Integer = 0 To WorkSheet.LastColumn - 1
                WorkSheet.AutoFitColumn(MyCounter)
            Next

            WorkSheet.TabColor = System.Drawing.Color.Black
        End Sub

        Public Overrides Function LookupCellErrorValue(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If Me.Workbook.Worksheets(sheetName).Range(rowIndex + 1, columnIndex + 1).HasError Then
                Return Me.Workbook.Worksheets(sheetName).Range(rowIndex + 1, columnIndex + 1).ErrorValue
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Sub ClearSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Item(sheetName).Clear()
        End Sub

        Public Overrides Sub CopySheetContent(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)
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

        Public Overrides Sub SelectCell(cell As ExcelCell)
            If cell.SheetName = Nothing Then Throw New ArgumentException("Sheet name required", NameOf(cell))
            Dim WorkSheet As Worksheet = Me.Workbook.Worksheets.Item(cell.SheetName)
            WorkSheet.Range(cell.Address(False)).Activate(True)
        End Sub

        Public Overrides ReadOnly Property HasVbaProject As Boolean
            Get
                Return Me.Workbook.HasMacros
            End Get
        End Property

        Public Overrides Sub RemoveVbaProject()
            If Me.Workbook.HasMacros = False Then Return 'Shortcut and circumvent following workaround

            'NOTE: Manufacturer component doesn't provide a direct way to remove the VBA project (setting Me.Workbook.HasMacros = False has no effect)
            'NOTE: VBA project will be removed automatically when saving as non-xlsm-file            

            '0. Lookup required private field of Spire.Xls
            Dim XlsWorkbookMembers = CompuMaster.Reflection.NonPublicInstanceMembers.GetMembers(Of System.Reflection.FieldInfo)(Me.Workbook.GetType, GetType(Spire.Xls.Core.Spreadsheet.XlsWorkbook))
            If XlsWorkbookMembers.Count <> 1 Then
                Throw New NotSupportedException("Spire.Xls incompatibility, please open an issue at https://github.com/CompuMasterGmbH/CompuMaster.Excel")
            End If

            '0. Preserve required values for later reset
            Dim XlsWb = CompuMaster.Reflection.NonPublicInstanceMembers.InvokeFieldGet(Of Spire.Xls.Core.Spreadsheet.XlsWorkbook)(Me.Workbook, Me.Workbook.GetType, XlsWorkbookMembers(0).Name)
            Dim PreservedFileName As String = XlsWb.FullFileName
            Dim PreservedIsSavedState As Boolean = Me.Workbook.IsSaved

            '1. Save to temp file
            Dim TempFile As String = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx"
            Me.Workbook.SaveToFile(TempFile)

            '2. Reload
            Me.Workbook.LoadFromFile(TempFile)

            '3. Reset FileName property
            Me.SetWorkbookFilePath(PreservedFileName)

            '4. Reset IsSaved property
            Me.Workbook.IsSaved = PreservedIsSavedState
        End Sub

    End Class

End Namespace