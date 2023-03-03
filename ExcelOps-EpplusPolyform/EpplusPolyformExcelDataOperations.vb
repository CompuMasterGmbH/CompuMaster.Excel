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

        Public Overrides Sub Close()
            If Me.IsClosed = False Then Me._WorkbookPackage.Dispose()
            Me._WorkbookPackage = Nothing
        End Sub

        Public Overrides ReadOnly Property IsClosed As Boolean
            Get
                Return Me._WorkbookPackage Is Nothing
            End Get
        End Property

        Protected Overrides Sub SaveInternal()
            If Me.PasswordForOpening <> Nothing Then
                Me.WorkbookPackage.Save(Me.PasswordForOpening)
            Else
                Me.WorkbookPackage.Save()
            End If
        End Sub

        Protected Overrides Sub SaveInternal_ApplyCachedCalculationOption(cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            If cachedCalculationsOption = SaveOptionsForDisabledCalculationEngines.DefaultBehaviour Then
                cachedCalculationsOption = SaveOptionsForDisabledCalculationEngines.NoReset
            End If
            Select Case cachedCalculationsOption
                Case SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation
                    'Me.ResetCellValueFromFormulaCellInWholeWorkbook()
                Case SaveOptionsForDisabledCalculationEngines.ResetCalculatedValuesForForcedCellRecalculationIfRecalculationRequired
                    'If Me.RecalculationRequired Then Me.ResetCellValueFromFormulaCellInWholeWorkbook()
                Case SaveOptionsForDisabledCalculationEngines.NoReset
                    'do nothing
                Case Else
                    Throw New NotImplementedException("Invalid option: " & cachedCalculationsOption)
            End Select
        End Sub

        Protected Overrides Sub SaveAsInternal(fileName As String, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            Dim FullPath As New System.IO.FileInfo(fileName)
            If Me.PasswordForOpening <> Nothing Then
                Me.WorkbookPackage.SaveAs(FullPath, Me.PasswordForOpening)
            Else
                Me.WorkbookPackage.SaveAs(FullPath)
            End If
            Me._FilePath = FullPath.FullName
        End Sub

        Protected Overrides ReadOnly Property WorkbookFilePath As String
            Get
                If Me.IsClosed Then
                    Return Nothing
                Else
                    Return Me.WorkbookPackage.File?.FullName
                End If
            End Get
        End Property

        Public Overrides Function SheetNames() As List(Of String)
            Dim Result As New List(Of String)
            For Each ws As ExcelWorksheet In Me.Workbook.Worksheets
                Result.Add(ws.Name)
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Remove all named ranges in Excel Workbook since this feature is involved but not actively used in Master Template V15; but this feature might throw exceptions in EPPlus when removing rows
        ''' </summary>
        Public Overrides Sub CleanupRangeNames()
            Dim NamesToRemove As New List(Of String)
            For Each namedRange In Me.Workbook.Names
                NamesToRemove.Add(namedRange.Name)
            Next
            For Each Name As String In NamesToRemove
                Me.Workbook.Names.Remove(Name)
            Next

            For Each worksheet In Me.Workbook.Worksheets
                Dim NamesInWorksheetToRemove As New List(Of String)
                For Each namedRange In worksheet.Names
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
            Dim MyExcelCellAddress As ExcelCellAddress = New ExcelAddress(cell.Address).Start
            Try
                Select Case GetType(T)
                    Case GetType(Integer), GetType(Long), GetType(Byte), GetType(Double), GetType(Decimal), GetType(Boolean)
                        If Trim(Me.Workbook.Worksheets(cell.SheetName).GetValue(Of String)(MyExcelCellAddress.Row, MyExcelCellAddress.Column)) = Nothing Then
                            Return Nothing
                        Else
                            Return Me.Workbook.Worksheets(cell.SheetName).GetValue(Of T)(MyExcelCellAddress.Row, MyExcelCellAddress.Column)
                        End If
                    Case GetType(String)
                        Try
                            Return Me.Workbook.Worksheets(cell.SheetName).GetValue(Of T)(MyExcelCellAddress.Row, MyExcelCellAddress.Column)
                        Catch ex As InvalidCastException
                            Dim CellValue As Object = Me.Workbook.Worksheets(cell.SheetName).GetValue(MyExcelCellAddress.Row, MyExcelCellAddress.Column)
                            If CellValue.GetType Is GetType(OfficeOpenXml.ExcelErrorValue) Then
                                Dim ErrorValue As String = CType(CellValue, OfficeOpenXml.ExcelErrorValue).ToString
                                Return CType(CType(ErrorValue, Object), T)
                            Else
                                Throw
                            End If
                        End Try
                    Case GetType(Object)
                        Return CType(Me.Workbook.Worksheets(cell.SheetName).GetValue(MyExcelCellAddress.Row, MyExcelCellAddress.Column), T)
                    Case Else
                        Return Me.Workbook.Worksheets(cell.SheetName).GetValue(Of T)(MyExcelCellAddress.Row, MyExcelCellAddress.Column)
                End Select
            Catch ex As InvalidCastException
                Throw New System.FormatException("Value """ & Me.LookupCellFormattedText(cell) & """ in cell """ & cell.Address(True) & """ can't be converted to " & GetType(T).Name, ex)
            Catch ex As FormatException
                Throw New System.FormatException("Value """ & Me.LookupCellFormattedText(cell) & """ in cell """ & cell.Address(True) & """ can't be converted to " & GetType(T).Name, ex)
            End Try
        End Function

        '''' <summary>
        '''' Read a cell value
        '''' </summary>
        '''' <typeparam name="T"></typeparam>
        '''' <param name="cell"></param>
        '''' <returns></returns>
        '''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        'Private Overloads Function LookupCellValue(Of T)(cell As OfficeOpenXml.ExcelRange) As T
        '    Select Case GetType(T)
        '        Case GetType(Integer), GetType(Long), GetType(Byte), GetType(Double), GetType(Decimal), GetType(Boolean)
        '            If Trim(cell.GetValue(Of String)) = Nothing Then
        '                Return Nothing
        '            Else
        '                Return cell.GetValue(Of T)
        '            End If
        '        Case GetType(String)
        '            Try
        '                Return cell.GetValue(Of T)
        '            Catch ex As InvalidCastException
        '                Dim CellValue As Object = cell.Value
        '                If CellValue.GetType Is GetType(OfficeOpenXml.ExcelErrorValue) Then
        '                    Dim ErrorValue As String = CType(CellValue, OfficeOpenXml.ExcelErrorValue).ToString
        '                    Return CType(CType(ErrorValue, Object), T)
        '                Else
        '                    Throw
        '                End If
        '            End Try
        '        Case GetType(Object)
        '            Return CType(cell.Value, T)
        '        Case Else
        '            Return cell.GetValue(Of T)
        '    End Select
        'End Function

        ''' <summary>
        ''' Read the cell format string
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Public Overrides Function LookupCellFormat(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return CompuMaster.Data.Utils.StringNotEmptyOrNothing(Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Style.Numberformat.Format)
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
            Try
                Select Case GetType(T)
                    Case GetType(Integer), GetType(Long), GetType(Byte), GetType(Double), GetType(Decimal), GetType(Boolean)
                        If Trim(Me.Workbook.Worksheets(sheetName).GetValue(Of String)(rowIndex + 1, columnIndex + 1)) = Nothing Then
                            Return Nothing
                        Else
                            Return Me.Workbook.Worksheets(sheetName).GetValue(Of T)(rowIndex + 1, columnIndex + 1)
                        End If
                    Case GetType(String)
                        Try
                            Return Me.Workbook.Worksheets(sheetName).GetValue(Of T)(rowIndex + 1, columnIndex + 1)
                        Catch ex As InvalidCastException
                            Dim CellValue As Object = Me.Workbook.Worksheets(sheetName).GetValue(rowIndex + 1, columnIndex + 1)
                            If CellValue.GetType Is GetType(OfficeOpenXml.ExcelErrorValue) Then
                                Dim ErrorValue As String = CType(CellValue, OfficeOpenXml.ExcelErrorValue).ToString
                                Return CType(CType(ErrorValue, Object), T)
                            Else
                                Throw
                            End If
                        End Try
                    Case GetType(Object)
                        Return CType(Me.Workbook.Worksheets(sheetName).GetValue(rowIndex + 1, columnIndex + 1), T)
                    Case Else
                        Return Me.Workbook.Worksheets(sheetName).GetValue(Of T)(rowIndex + 1, columnIndex + 1)
                End Select
            Catch ex As InvalidCastException
                Throw New System.FormatException("Value """ & Me.LookupCellFormattedText(sheetName, rowIndex, columnIndex) & """ in cell """ & New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All).Address(True) & """ can't be converted to " & GetType(T).Name, ex)
            Catch ex As FormatException
                Throw New System.FormatException("Value """ & Me.LookupCellFormattedText(sheetName, rowIndex, columnIndex) & """ in cell """ & New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All).Address(True) & """ can't be converted to " & GetType(T).Name, ex)
            End Try
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        ''' <remarks>Cell values with spaces will be converted to null values in case of method call with types bool, byte, int32, int64, double, decimal</remarks>
        Public Overrides Function TryLookupCellValue(Of T As Structure)(cell As ExcelCell) As T?
            Dim MyExcelCellAddress As ExcelCellAddress = New ExcelAddress(cell.Address).Start
            Try
                Select Case GetType(T)
                    Case GetType(Integer), GetType(Long), GetType(Byte), GetType(Double), GetType(Decimal), GetType(Boolean)
                        If Trim(Me.Workbook.Worksheets(cell.SheetName).GetValue(Of String)(MyExcelCellAddress.Row, MyExcelCellAddress.Column)) = Nothing Then
                            Return Nothing
                        Else
                            Return Me.Workbook.Worksheets(cell.SheetName).GetValue(Of T)(MyExcelCellAddress.Row, MyExcelCellAddress.Column)
                        End If
                    Case Else
                        Return Me.Workbook.Worksheets(cell.SheetName).GetValue(Of T)(MyExcelCellAddress.Row, MyExcelCellAddress.Column)
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
            Try
                Select Case GetType(T)
                    Case GetType(Integer), GetType(Long), GetType(Byte), GetType(Double), GetType(Decimal), GetType(Boolean)
                        If Trim(Me.Workbook.Worksheets(sheetName).GetValue(Of String)(rowIndex + 1, columnIndex + 1)) = Nothing Then
                            Return Nothing
                        Else
                            Return Me.Workbook.Worksheets(sheetName).GetValue(Of T)(rowIndex + 1, columnIndex + 1)
                        End If
                    Case Else
                        Return Me.Workbook.Worksheets(sheetName).GetValue(Of T)(rowIndex + 1, columnIndex + 1)
                End Select
            Catch ex As Exception
                Return Nothing
            End Try
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
            If Me.Workbook.Worksheets(sheetName) Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & sheetName, NameOf(sheetName))
            Return Me.Workbook.Worksheets(sheetName).GetValue(rowIndex + 1, columnIndex + 1)
        End Function

        ''' <summary>
        ''' Read a cell value
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        Public Overrides Function LookupCellFormula(cell As ExcelCell) As String
            Dim MyExcelCellAddress As ExcelCellAddress = New ExcelAddress(cell.Address).Start
            If Me.Workbook.Worksheets(cell.SheetName) Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & cell.SheetName, NameOf(cell))
            Return CompuMaster.Data.Utils.StringNotEmptyOrNothing(Me.Workbook.Worksheets(cell.SheetName).Cells(MyExcelCellAddress.Row, MyExcelCellAddress.Column).Formula)
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
            If Me.Workbook.Worksheets(sheetName) Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & sheetName, NameOf(sheetName))
            If rowIndex < 0 Then Throw New ArgumentOutOfRangeException("RowIndex " & rowIndex & " must be equal or bigger than 0", NameOf(rowIndex))
            If columnIndex < 0 Then Throw New ArgumentOutOfRangeException("ColumnIndex " & columnIndex & " must be equal or bigger than 0", NameOf(columnIndex))
            Return CompuMaster.Data.Utils.StringNotEmptyOrNothing(Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Formula)
        End Function

        Public Overrides Function LookupCellIsLocked(cell As ExcelCell) As Boolean
            Dim MyExcelCellAddress As ExcelCellAddress = New ExcelAddress(cell.Address).Start
            Return Me.Workbook.Worksheets(cell.SheetName).Cells(MyExcelCellAddress.Row, MyExcelCellAddress.Column).Style.Locked
        End Function

        Public Overrides Function LookupCellIsLocked(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Style.Locked
        End Function

        ''' <summary>
        ''' Write a cell value
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="cell"></param>
        ''' <param name="value"></param>
        Public Overrides Sub WriteCellValue(Of T)(cell As ExcelCell, value As T)
            Dim MyExcelCellAddress As ExcelCellAddress = New ExcelAddress(cell.Address).Start
            Me.Workbook.Worksheets(cell.SheetName).Cells(MyExcelCellAddress.Row, MyExcelCellAddress.Column).Formula = Nothing
            Me.Workbook.Worksheets(cell.SheetName).Cells(MyExcelCellAddress.Row, MyExcelCellAddress.Column).Value = value
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
            Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Formula = Nothing
            Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Value = value
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
            Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Formula = formula
            If immediatelyCalculateCellValue Then
                Me.RecalculateCell(sheetName, rowIndex, columnIndex)
            Else
                Me.RecalculationRequired = True
            End If
        End Sub

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
                Me.FullLog.AppendLine("WARNING at " & context.Scopes.ToString & ": " & message)
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
                ' GC.SuppressFinalize(Me)
            End Sub
#End Region
        End Class

        Public Sub CalculationEngineAttachLogger()
            Me.CalculationEngineLog.FullLog.Clear()
            Me.CalculationEngineLog.ExceptionsLog.Clear()
            Me.Workbook.FormulaParserManager.AttachLogger(Me.CalculationEngineLog)
        End Sub

        Public Sub CalculationEngineDetachLogger()
            Me.Workbook.FormulaParserManager.DetachLogger()
        End Sub

        Public ReadOnly Property CalculationEngineLog As New FormulaParserLogger

        ''' <summary>
        ''' Recalculate a cell based on its formula
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">0-based row number</param>
        ''' <param name="columnIndex">0-based column number</param>
        Public Overrides Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer, throwExceptionOnCalculationError As Boolean)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Calculate
            If throwExceptionOnCalculationError AndAlso Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Value?.GetType Is GetType(OfficeOpenXml.ExcelErrorValue) Then
                Dim Cell As New ExcelCell(sheetName, rowIndex, columnIndex, ExcelCell.ValueTypes.All)
                Throw New NotSupportedException("Epplus calculation at " & Cell.Address(True) & " resulted in #" & UCase(CType(Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Value, OfficeOpenXml.ExcelErrorValue).Type.ToString) & "!" & " for formula =" & Me.LookupCellFormula(Cell))
            End If
        End Sub

        ''' <summary>
        ''' Recalculate all cells of a sheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Sub RecalculateSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If Me.CalculationModuleDisabled Then Throw New InvalidOperationException("Calculation engine is disabled, requested recalculation failed")
            Me.Workbook.Worksheets(sheetName).Calculate
        End Sub

        Protected Overrides Sub CreateWorkbook()
            Me._WorkbookPackage = New OfficeOpenXml.ExcelPackage()
            Me._WorkbookPackage.Compatibility.IsWorksheets1Based = False

            'set workbook FullCalcOnLoad always to False since it's already triggered using property of Me.AutoCalculationOnLoad
            Me.Workbook.FullCalcOnLoad = True 'unknown if executed after loading already completed or if it's a workbook setting with effect on opening as user in MS Excel, too
            Me.Workbook.Worksheets.Add("Sheet1")
        End Sub

        Protected Overrides Sub LoadWorkbook(file As System.IO.FileInfo)
            If Me.PasswordForOpening <> Nothing Then
                Me._WorkbookPackage = New OfficeOpenXml.ExcelPackage(file, Me.PasswordForOpening)
            Else
                Me._WorkbookPackage = New OfficeOpenXml.ExcelPackage(file)
            End If
            Me._WorkbookPackage.Compatibility.IsWorksheets1Based = False

            'set workbook FullCalcOnLoad always to False since it's already triggered using property of Me.AutoCalculationOnLoad
            Me.Workbook.FullCalcOnLoad = True 'unknown if executed after loading already completed or if it's a workbook setting with effect on opening as user in MS Excel, too
        End Sub

        ''' <summary>
        ''' Lookup the last content column index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastColumnIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = Me.Workbook.Worksheets(sheetName)
            If Sheet.Dimension Is Nothing Then Return 0
            Return Sheet.Dimension.End.Column - 1
        End Function

        ''' <summary>
        ''' Lookup the last content column index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastContentColumnIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = Me.Workbook.Worksheets(sheetName)
            If Sheet Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & sheetName, NameOf(sheetName))
            If Sheet.Dimension Is Nothing Then Return 0
            Dim autoSuggestionLastRowIndex As Integer = Sheet.Dimension.End.Row - 1
            Dim autoSuggestedResult As Integer = Sheet.Dimension.End.Column - 1
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
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = Me.Workbook.Worksheets(sheetName)
            If Sheet Is Nothing Then Throw New NullReferenceException("Sheet """ & sheetName & """ doesn't exist in workbook")
            If Sheet.Dimension Is Nothing Then Return 0
            Return Sheet.Dimension.End.Row - 1
        End Function

        ''' <summary>
        ''' Lookup the last content row index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastContentRowIndex(sheetName As String) As Integer
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = Me.Workbook.Worksheets(sheetName)
            If Sheet Is Nothing Then Throw New ArgumentException("Must be an existing sheet name: """ & sheetName & """", NameOf(sheetName))
            If Sheet.Dimension Is Nothing Then Return 0
            Dim autoSuggestionLastColumnIndex As Integer = Sheet.Dimension.End.Column - 1
            Dim autoSuggestedResult As Integer = Sheet.Dimension.End.Row - 1
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
            Return New ExcelOps.ExcelCell(sheetName, Me.Workbook.Worksheets(sheetName).Cells(CellRowIndex + 1, CellColIndex + 1).Address, Nothing)
        End Function

        ''' <summary>
        ''' Lookup the last content cell (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Function LookupLastContentCell(sheetName As String) As ExcelOps.ExcelCell
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim CellRowIndex As Integer = Me.LookupLastContentRowIndex(sheetName)
            Dim CellColIndex As Integer = Me.LookupLastContentColumnIndex(sheetName)
            Return New ExcelOps.ExcelCell(sheetName, Me.Workbook.Worksheets(sheetName).Cells(CellRowIndex + 1, CellColIndex + 1).Address, Nothing)
        End Function

        ''' <summary>
        ''' Lookup the row index (zero based index)
        ''' </summary>
        ''' <param name="cell"></param>
        Public Overrides Function LookupRowIndex(cell As ExcelOps.ExcelCell) As Integer
            Return New OfficeOpenXml.ExcelCellAddress(cell.Address).Row - 1
        End Function

        ''' <summary>
        ''' Lookup the column index (zero based index)
        ''' </summary>
        ''' <param name="cell"></param>
        Public Overrides Function LookupColumnIndex(cell As ExcelOps.ExcelCell) As Integer
            Return New OfficeOpenXml.ExcelCellAddress(cell.Address).Column - 1
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
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = Me.Workbook.Worksheets(sheetName)
            If Sheet.Dimension Is Nothing Then Throw New Exception("Specified worksheet is not a data worksheet")
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
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = Me.Workbook.Worksheets(sheetName)
            If Sheet.Dimension Is Nothing Then Throw New Exception("Specified worksheet is not a data worksheet")
            Dim ws As ExcelWorksheet = Me.WorkbookPackage.Workbook.Worksheets.Item(rangeFirstCell.SheetName)
            ws.Cells(rangeFirstCell.Address & ":" & rangeLastCell.Address).Clear()
        End Sub

        ''' <summary>
        ''' Determine if a cell contains empty content
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex">Zero-based index</param>
        ''' <param name="columnIndex">Zero-based index</param>
        Public Overrides Function IsEmptyCell(sheetName As String, ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = Me.Workbook.Worksheets(sheetName)
            Return IsEmptyCell(Sheet, rowIndex, columnIndex)
        End Function

        ''' <summary>
        ''' Determine if a cell contains empty content (cells with formulas are always considered as filled cells)
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="rowIndex">Zero-based index</param>
        ''' <param name="columnIndex">Zero-based index</param>
        Private Overloads Function IsEmptyCell(ByVal sheet As OfficeOpenXml.ExcelWorksheet, ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Boolean
            If sheet.Cells(rowIndex + 1, columnIndex + 1).Formula <> Nothing Then
                Return False
            End If
            Dim value As Object = sheet.Cells(rowIndex + 1, columnIndex + 1).Value
            If value Is Nothing Then
                Return True
            ElseIf value.GetType Is GetType(String) AndAlso CType(value, String) = Nothing Then
                Return True
            Else
                Return False
            End If
        End Function

#Region "Inactive code"
#If False Then
    Private Function ReadCellData(ByVal sheet As OfficeOpenXml.ExcelWorksheet, cellAddress As OfficeOpenXml.ExcelCellAddress) As Object
        Dim value As Object
        Select Case LookupDotNetType(sheet.Cells(cellAddress.Address))
            Case VariantType.Empty
                value = DBNull.Value
            Case VariantType.Boolean
                value = CType(sheet.Cells(cellAddress.Address).Value, Boolean)
            Case VariantType.Error
                value = Double.NaN
            Case VariantType.Double
                'Handle as normal double
                value = CType(sheet.Cells(cellAddress.Address).Value, Double)
            Case VariantType.String
                Dim cellValue As String
                cellValue = CType(sheet.Cells(cellAddress.Address).Value, String)
                If cellValue <> "" AndAlso System.Environment.NewLine <> ControlChars.Lf Then
                    cellValue = Replace(cellValue, ControlChars.Lf, System.Environment.NewLine, , , CompareMethod.Binary)
                End If
                value = cellValue
            Case VariantType.Date
                If sheet.Cells(cellAddress.Address).Value.GetType Is GetType(Double) Then
                    value = sheet.Cells(cellAddress.Address).GetValue(Of DateTime)
                Else
                    value = CType(sheet.Cells(cellAddress.Address).Value, DateTime)
                End If
            Case VariantType.Decimal
                value = CType(sheet.Cells(cellAddress.Address).Value, Decimal)
            Case VariantType.Char
                value = CType(sheet.Cells(cellAddress.Address).Value, Char)
            Case VariantType.Byte
                value = CType(sheet.Cells(cellAddress.Address).Value, Byte)
            Case VariantType.Currency
                value = CType(sheet.Cells(cellAddress.Address).Value, Decimal)
            Case VariantType.Integer
                value = CType(sheet.Cells(cellAddress.Address).Value, Integer)
            Case VariantType.Long
                value = CType(sheet.Cells(cellAddress.Address).Value, Long)
            Case VariantType.Short
                value = CType(sheet.Cells(cellAddress.Address).Value, Short)
            Case VariantType.Single
                value = CType(sheet.Cells(cellAddress.Address).Value, Single)
            Case Else
                'Case VariantType.DataObject
                'Case VariantType.Array
                'Case VariantType.Null
                'Case VariantType.UserDefinedType
                'Case VariantType.Variant
                value = New NotImplementedException("Error in sheet cell " & cellAddress.Address & "
            Unknown cell type")
        End Select
        If value.GetType Is GetType(String) AndAlso CType(value, String) = "" Then
            'Handle situation that a cell might contain a "" instead of a blank value because of some user-defined Excel formulas which shall return "blank" cell content by using "" - irrespective to the regular column data type
            'e.g. following formula: =IF($F6=I$1;1;"")
            value = DBNull.Value
        End If
        Return value
    End Function

    'Private Function NumberToDateTime(value As Double) As DateTime
    '    Return DateTime.FromOADate(value)
    'End Function

    Private Function LookupDotNetType(xlsCell As OfficeOpenXml.ExcelRange) As VariantType
        If xlsCell Is Nothing OrElse xlsCell.Value Is Nothing Then
            Return VariantType.Empty
        Else
            Select Case xlsCell.Value.GetType
                Case GetType(String)
                    Return VariantType.String
                Case GetType(Double)
                    If IsDateTimeFormat(xlsCell.Style.Numberformat.Format) Then
                        Return VariantType.Date
                    Else
                        Return VariantType.Double
                    End If
                Case GetType(Boolean)
                    Return VariantType.Boolean
                Case GetType(DateTime)
                    Return VariantType.Date
                Case GetType(OfficeOpenXml.ExcelErrorValue)
                    Return VariantType.Error
                Case Else
                    Return VariantType.Object
            End Select
        End If
    End Function

    ''' <summary>
    ''' Detect custom date/time format strings which haven't been detected by Epplus
    ''' </summary>
    ''' <param name="cellFormat"></param>
    ''' <returns></returns>
    Private Function IsDateTimeFormat(cellFormat As String) As Boolean
        If cellFormat = "" Then
            Return False
        ElseIf cellFormat.StartsWith("yyyy-MM-dd") OrElse cellFormat.StartsWith("HH
            mm
            ss") Then
            Return True
        Else
            Return False
        End If
    End Function

#Region "Inactive code"
    '    ''' <summary>
    '    '''     Lookup the (zero-based) index number of a work sheet
    '    ''' </summary>
    '    ''' <param name="workbook">The excel workbook</param>
    '    ''' <param name="worksheetName">A work sheet name</param>
    '    ''' <returns>-1 if the sheet name doesn't exist, otherwise its index value</returns>
    '    Private Function ResolveWorksheetIndex(ByVal workbook As OfficeOpenXml.ExcelPackage, ByVal worksheetName As String) As Integer
    '        Dim sheetIndex As Integer = -1
    '        For MyCounter As Integer = 0 To workbook.Workbook.Worksheets.Count - 1
    '            Dim sheet As OfficeOpenXml.ExcelWorksheet = workbook.Workbook.Worksheets(MyCounter + 1)
    '            If sheet.Name.ToLower = worksheetName.ToLower Then
    '                sheetIndex = MyCounter
    '            End If
    '        Next
    '        Return sheetIndex
    '    End Function

    '    ''' <param name="workbook">The excel workbook</param>
    '    ''' <param name="sheetName">A sheet name</param>
    '    ''' <returns>An excel sheet</returns>
    '    Private Function LookupWorksheet(ByVal workbook As OfficeOpenXml.ExcelPackage, ByVal sheetName As String) As OfficeOpenXml.ExcelWorksheet
    '        Dim resolvedIndex As Integer = ResolveWorksheetIndex(workbook, sheetName)
    '        If resolvedIndex = -1 Then
    '            Throw New Exception("Worksheet """ & sheetName & """ hasn't been found")
        End If
    '        Else
    '            Return workbook.Workbook.Worksheets(resolvedIndex + 1)
    '        End If
    '    End Function

    '    ''' <summary>
    '    ''' Lookup if the value is a DateTime value and not a normal number
    '    ''' </summary>
    '    ''' <param name="cell"></param>
    '    ''' <returns>True for DateTime, False for Number(Double)</returns>
    '    ''' <remarks></remarks>
    '    Private Function IsDateTimeInsteadOfNumber(ByVal cell As OfficeOpenXml.ExcelRange) As Boolean
    '        Dim numFormat As String = cell.Style.Numberformat.Format
    '        If numFormat.ToLower.IndexOf("y") > 0 OrElse numFormat.ToLower.IndexOf("m") > 0 OrElse numFormat.ToLower.IndexOf("d") > 0 OrElse numFormat.ToLower.IndexOf("h") > 0 Then
    '            Try
    '                DateTime.FromOADate(CType(cell.Value, Double))
    '                Return True
    '            Catch
    '                Return False
    '            End Try
    '        Else
    '            Return False
    '        End If

    '    End Function
#End If
#End Region

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
                Dim CellValue As Object = Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Value
                If CellValue IsNot Nothing AndAlso CellValue.GetType() Is GetType(Boolean) AndAlso Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Style.Numberformat.Format = "General" Then
                    Return Me.LookupCellValue(Of String)(sheetName, rowIndex, columnIndex)
                Else
                    Return Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Text
                End If
            Catch ex As InvalidCastException
                Dim CellValue As Object = Me.Workbook.Worksheets(sheetName).GetValue(rowIndex + 1, columnIndex + 1)
                If CellValue.GetType Is GetType(OfficeOpenXml.ExcelErrorValue) Then
                    Return CType(CellValue, OfficeOpenXml.ExcelErrorValue).ToString
                Else
                    Throw
                End If
            Catch ex As Exception
                Return "#ERROR: " & ex.Message
            End Try
        End Function

        '''' <summary>
        '''' Try to lookup the cell's value to a string anyhow
        '''' </summary>
        '''' <param name="cell"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Private Overloads Function LookupCellFormattedText(ByVal cell As OfficeOpenXml.ExcelRange) As String
        '    Try
        '        Dim CellValue As Object = cell.Value
        '        If CellValue IsNot Nothing AndAlso CellValue.GetType() Is GetType(Boolean) AndAlso cell.Style.Numberformat.Format = "General" Then
        '            Return Me.LookupCellValue(Of String)(cell)
        '        Else
        '            Return cell.Text
        '        End If
        '    Catch ex As InvalidCastException
        '        Dim CellValue As Object = cell.Value
        '        If CellValue.GetType Is GetType(OfficeOpenXml.ExcelErrorValue) Then
        '            Dim ErrorValue As String = CType(CellValue, OfficeOpenXml.ExcelErrorValue).ToString
        '            Return ErrorValue
        '        Else
        '            Throw
        '        End If
        '    Catch ex As Exception
        '        Return "#ERROR: " & ex.Message
        '    End Try
        'End Function

        Public Overrides Sub UnprotectSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Item(sheetName).Protection.IsProtected = False
        End Sub

        Public Overrides Sub ProtectSheet(sheetName As String, level As ProtectionLevel)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Select Case level
                Case ProtectionLevel.StandardWithInsertDeleteRows
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectUnlockedCells = True
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectLockedCells = True
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowDeleteRows = True
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowInsertRows = True
                Case ProtectionLevel.Standard
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectUnlockedCells = True
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectLockedCells = True
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowDeleteRows = False
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowInsertRows = False
                Case ProtectionLevel.SelectAndEditUnlockedCellsOnly
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectUnlockedCells = True
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectLockedCells = False
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowDeleteRows = False
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowInsertRows = False
                Case ProtectionLevel.SelectAndEditAllCellsButNoFurtherEditing
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectUnlockedCells = True
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectLockedCells = True
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowDeleteRows = False
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowInsertRows = False
                Case ProtectionLevel.SelectNoCellsAndNoEditing
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectUnlockedCells = False
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSelectLockedCells = False
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowDeleteRows = False
                    Me.Workbook.Worksheets.Item(sheetName).Protection.AllowInsertRows = False
                Case Else
                    Throw New NotImplementedException
            End Select
            'Me.Workbook.Worksheets.Item(sheetName).Protection.AllowInsertingHyperlinks = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowEditScenarios = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowEditObject = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowDeleteColumns = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowInsertColumns = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowFormatCells = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowFormatColumns = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowFormatRows = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowSort = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowPivotTables = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.AllowAutoFilter = False
            Me.Workbook.Worksheets.Item(sheetName).Protection.IsProtected = True
        End Sub

        Protected Overrides Sub RecalculateAllInternal()
            If Me.CalculationModuleDisabled Then Throw New FeatureDisabledException("Calculation engine")
            OfficeOpenXml.CalculationExtension.Calculate(Me.Workbook)
        End Sub

        ''' <summary>
        ''' Is the Excel engine allowed to automatically/continuously calculate on every change or does the user has to manually force a recalculation (typically by pressing F9 key in MS Excel)
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Property AutoCalculationEnabled As Boolean
            Get
                Return (Me.Workbook.CalcMode = ExcelCalcMode.Automatic)
            End Get
            Set(value As Boolean)
                If value Then
                    Me.Workbook.CalcMode = ExcelCalcMode.Automatic
                Else
                    Me.Workbook.CalcMode = ExcelCalcMode.Manual
                End If
            End Set
        End Property

        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "Epplus (Polyform license edition)"
            End Get
        End Property

        Public Overrides Function IsProtectedSheet(sheetName As String) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.Workbook.Worksheets.Item(sheetName).Protection.IsProtected
        End Function

        Public Overrides Sub RemoveSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets.Delete(sheetName)
        End Sub

        Public Overrides Sub AddSheet(sheetName As String, beforeSheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim ws As ExcelWorksheet = Me.Workbook.Worksheets.Add(sheetName)
            If beforeSheetName <> Nothing Then Me.Workbook.Worksheets.MoveBefore(sheetName, beforeSheetName)
        End Sub

        ''' <summary>
        ''' Select a worksheet
        ''' </summary>
        ''' <param name="sheetName"></param>
        Public Overrides Sub SelectSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets(sheetName).Select()
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
            Me.Workbook.Worksheets.Item(sheetName).Hidden = eWorkSheetHidden.Visible
        End Sub

        Public Overrides Sub HideSheet(sheetName As String, stronglyHide As Boolean)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If stronglyHide Then
                Me.Workbook.Worksheets.Item(sheetName).Hidden = eWorkSheetHidden.VeryHidden
            Else
                Me.Workbook.Worksheets.Item(sheetName).Hidden = eWorkSheetHidden.Hidden
            End If
        End Sub

        Public Overrides Function IsHiddenSheet(sheetName As String) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Return Me.Workbook.Worksheets.Item(sheetName).Hidden <> eWorkSheetHidden.Visible
        End Function

        Public Sub WriteTableToSheet(dataTable As DataTable, sheetName As String)

            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim WorkSheet As OfficeOpenXml.ExcelWorksheet = Me.Workbook.Worksheets(sheetName)
            WorkSheet.Cells.Clear()

            'Paste the column headers
            For ColCounter As Integer = 0 To dataTable.Columns.Count - 1
                Dim headline As String = dataTable.Columns(ColCounter).ColumnName
                WorkSheet.Cells(1, ColCounter + 1).Value = headline
                WorkSheet.Cells(1, ColCounter + 1).Style.Font.Bold = True
                WorkSheet.Cells(1, ColCounter + 1).Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium
                WorkSheet.Cells(1, ColCounter + 1).Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(0, 0, 0))
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
                        WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = Nothing
                    ElseIf value.GetType Is GetType(String) Then
                        'Excel requires line-breaks to be an LF character only, not a windows typical CR+LF
                        Dim cell As OfficeOpenXml.ExcelRange = WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1)
                        If CType(value, String) <> Nothing Then
                            value = Replace(CType(value, String), ControlChars.CrLf, ControlChars.Lf) 'Windows line breaks
                            value = Replace(CType(value, String), ControlChars.Cr, ControlChars.Lf) 'Mac or Linux line break
                        End If
                        cell.Formula = ""
                        cell.Value = value
                    ElseIf value.GetType Is GetType(DateTime) Then
                        Dim datevalue As DateTime = CType(value, DateTime)
                        Try
                            'Re-create datevalue to strip off any other additional properties
                            datevalue = New DateTime(datevalue.Year, datevalue.Month, datevalue.Day, datevalue.Hour, datevalue.Minute, datevalue.Second, datevalue.Millisecond)
                            'Write back the new cell value
                            If datevalue = New DateTime Then
                                WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = Nothing
                            Else
                                'WorkSheet.Workbook.DateTimeToNumber(datevalue)
                                WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = datevalue
                                If dataTable.Columns(ColCounter).ExtendedProperties.ContainsKey("Format") Then
                                    WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Style.Numberformat.Format = CType(dataTable.Columns(ColCounter).ExtendedProperties("Format"), String)
                                Else
                                    WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss"
                                End If
                            End If
                        Catch ex As Exception
                            Throw New Exception("Error writing a date/time value """ & datevalue.ToString & """ in row " & (RowCounter + 1), ex)
                        End Try
                    ElseIf value.GetType Is GetType(Decimal) Then
                        Dim decimalValue As Decimal = CType(value, Decimal)
                        WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = decimalValue
                    ElseIf value.GetType Is GetType(Double) Then
                        Dim doubleValue As Double = CType(value, Double)
                        If doubleValue = Double.PositiveInfinity Then
                            WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = OfficeOpenXml.ExcelErrorValue.Values.Num
                        ElseIf doubleValue = Double.NegativeInfinity Then
                            WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = OfficeOpenXml.ExcelErrorValue.Values.Num
                        ElseIf Double.IsNaN(doubleValue) Then
                            WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = OfficeOpenXml.ExcelErrorValue.Values.Div0
                        ElseIf Double.Epsilon = doubleValue Then
                            'too small number would be rounded to just 0
                            WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = OfficeOpenXml.ExcelErrorValue.Values.Num
                        Else
                            WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = CType(value, Double)
                        End If
                    ElseIf value.GetType Is GetType(Int16) OrElse value.GetType Is GetType(Int32) Then
                        WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = CType(value, Int32)
                    ElseIf value.GetType Is GetType(Int64) Then
                        WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = CType(value, Double)
                    ElseIf value.GetType Is GetType(Boolean) Then
                        WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = CType(value, Boolean)
                    Else
                        WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = CType(value, Object).ToString
                    End If
                Next
            Next

            ' Auto size all worksheet columns which contain data
            For MyCounter As Integer = 1 To WorkSheet.Dimension.End.Column
                WorkSheet.Column(MyCounter).AutoFit(0.5)
            Next

            WorkSheet.TabColor = System.Drawing.Color.Black
        End Sub

        Public Overrides Function LookupCellErrorValue(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim CellValue As Object = Me.LookupCellValueAsObject(sheetName, rowIndex, columnIndex)
            If CellValue IsNot Nothing AndAlso CellValue.GetType Is GetType(OfficeOpenXml.ExcelErrorValue) Then
                Return CType(CellValue, OfficeOpenXml.ExcelErrorValue).ToString
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Sub ClearSheet(sheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            Me.Workbook.Worksheets(sheetName).Cells.Clear()
        End Sub

        Public Overrides Sub CopySheetContentInternal(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If targetWorkbook.GetType IsNot GetType(EpplusPolyformExcelDataOperations) Then Throw New NotSupportedException("Excel engines must be the same for source and target workbook for copying worksheets")
            'Me.Workbook.Worksheets.Copy(sheetName, targetSheetName)
            Throw New NotSupportedException("Epplus doesn't support copying of sheets with data + formats + locks")
            Dim LastCell As ExcelCell = Me.LookupLastCell(sheetName)
            targetWorkbook.ClearSheet(targetSheetName)
            Me.Workbook.Worksheets(sheetName).Cells(1, 1, LastCell.RowIndex + 1, LastCell.ColumnIndex + 1).Copy(CType(targetWorkbook, EpplusPolyformExcelDataOperations).Workbook.Worksheets(targetSheetName).Cells)
            'Me.Workbook.Worksheets(sheetName).Cells.Copy(CType(targetWorkbook, EpplusExcelDataOperations).Workbook.Worksheets(targetSheetName).Cells)
        End Sub

        Public Overrides Sub SelectCell(cell As ExcelCell)
            If cell.SheetName = Nothing Then Throw New ArgumentException("Sheet name required", NameOf(cell))
            Dim WorkSheet As OfficeOpenXml.ExcelWorksheet = Me.Workbook.Worksheets(cell.SheetName)
            WorkSheet.Select(cell.Address(False), False)
        End Sub

        Public Overrides Sub RemoveVbaProject()
            If Me.Workbook.VbaProject IsNot Nothing Then
                Me.Workbook.VbaProject.Remove()
            End If
            Me.Workbook.RemoveVBAProject()
        End Sub

        Public Overrides Function IsMergedCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
            If Me.Workbook.Worksheets(sheetName) Is Nothing Then Throw New ArgumentOutOfRangeException("Sheet not found: " & sheetName, NameOf(sheetName))
            Return Me.Workbook.Worksheets(sheetName).Cells(rowIndex + 1, columnIndex + 1).Merge
        End Function

        Public Overrides Sub UnMergeCells(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            Me.Workbook.Worksheets(sheetName).Cells(IsMergedCellOfRange(sheetName, rowIndex, columnIndex)).Merge = False
        End Sub

        Public Overrides Sub MergeCells(sheetName As String, fromRowIndex As Integer, fromColumnIndex As Integer, toRowIndex As Integer, toColumnIndex As Integer)
            Me.Workbook.Worksheets(sheetName).Cells(fromRowIndex + 1, fromColumnIndex + 1, toRowIndex + 1, toColumnIndex + 1).Merge = True
        End Sub

        Public Function IsMergedCellOfRange(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Return Me.Workbook.Worksheets(sheetName).MergedCells(Me.Workbook.Worksheets(sheetName).GetMergeCellId(rowIndex + 1, columnIndex + 1) - 1)
        End Function

        Public Overrides Sub AutoFitColumns(sheetName As String)
            Me.Workbook.Worksheets(sheetName).Cells.AutoFitColumns()
        End Sub

        Public Overrides Sub AutoFitColumns(sheetName As String, minimumWidth As Double)
            Me.Workbook.Worksheets(sheetName).Cells.AutoFitColumns(minimumWidth)
        End Sub

        Public Overrides Sub AutoFitColumns(sheetName As String, columnIndex As Integer)
            Me.Workbook.Worksheets(sheetName).Column(columnIndex + 1).AutoFit()
        End Sub

        Public Overrides Sub AutoFitColumns(sheetName As String, columnIndex As Integer, minimumWidth As Double)
            Me.Workbook.Worksheets(sheetName).Column(columnIndex + 1).AutoFit(minimumWidth)
        End Sub

        Public Overrides ReadOnly Property HasVbaProject As Boolean
            Get
                Return Me.Workbook.VbaProject IsNot Nothing
            End Get
        End Property

        'Public ReadOnly Property DrawingsCount As Integer
        '    Get
        '        Return Me.Workbook.Worksheets
        '        OfficeOpenXml.Drawing.ExcelPicture
        '    End Get
        'End Property
        '
        'Public ReadOnly Property Drawings As OfficeOpenXml.Drawing.ExcelPicture

    End Class

End Namespace