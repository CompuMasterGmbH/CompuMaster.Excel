Option Explicit On
Option Strict On

'NOTE:    THIS FILE IS UPDATED IN DIRECTORY CM.Data.EpplusFixCalcsEdition FIRST AND COPIED TO CM.Data.EpplusPolyformEdition AFTERWARDS
'SEE:     clone-build-files.cmd/.sh/.ps1
'WARNING: PLEASE CHANGE THIS FILE ONLY AT REQUIRED LOCATION, OR CHANGES WILL BE LOST!

Imports System.Data
Imports System.Linq

Namespace CompuMaster.Data

    ''' <summary>
    ''' Provides simplified write access to XLS files.
    ''' </summary>
    ''' <remarks>
    '''     Please pay attention to following circumstances
    '''     - Written null values (Nothing in VisualBasic) will be re-read as DBNull.Value
    '''     - Written zero-DateTime value will be re-read as DBNull.Value
    '''     - Excel supports DateTime values starting from 01.01.1900, only. Lower date values will throw an exception when assigning.
    '''     - Excel DateTime values are limited to year, month, day, hour, minute, second. Milliseconds and ticks will be dropped.
    '''     - Lines with only DBNull.Value or null (Nothing in VisualBasic) will be considered as not-existing if they are the last lines
    ''' </remarks>
    Public NotInheritable Class XlsEpplusPolyformEdition

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Defines immutable defaults for reading worksheet data.
        ''' </summary>
        Public NotInheritable Class ReadOptions

            ''' <summary>
            ''' Initializes read options.
            ''' </summary>
            ''' <param name="firstRowContainsColumnNames">Indicates whether the first imported row contains column names.</param>
            ''' <param name="startReadingAtRowIndex">The zero-based row index at which reading starts.</param>
            ''' <exception cref="ArgumentOutOfRangeException"><paramref name="startReadingAtRowIndex"/> is negative.</exception>
            Public Sub New(Optional firstRowContainsColumnNames As Boolean = True, Optional startReadingAtRowIndex As Integer = 0)
                If startReadingAtRowIndex < 0 Then
                    Throw New ArgumentOutOfRangeException(NameOf(startReadingAtRowIndex), "The start row index must not be negative")
                End If

                Me.FirstRowContainsColumnNames = firstRowContainsColumnNames
                Me.StartReadingAtRowIndex = startReadingAtRowIndex
#If CM_FIXCALCS Then
                Me.PackageLoadLimits = OfficeOpenXml.ExcelPackageLoadLimits.Default
#End If
            End Sub

#If CM_FIXCALCS Then
            ''' <summary>
            ''' Initializes read options with explicit XLSX resource limits.
            ''' </summary>
            ''' <param name="firstRowContainsColumnNames">Indicates whether the first imported row contains column names.</param>
            ''' <param name="startReadingAtRowIndex">The zero-based row index at which reading starts.</param>
            ''' <param name="packageLoadLimits">The resource limits to apply while opening the XLSX file.</param>
            ''' <exception cref="ArgumentNullException"><paramref name="packageLoadLimits"/> is <see langword="Nothing"/>.</exception>
            ''' <exception cref="ArgumentOutOfRangeException"><paramref name="startReadingAtRowIndex"/> is negative.</exception>
            Public Sub New(firstRowContainsColumnNames As Boolean, startReadingAtRowIndex As Integer, packageLoadLimits As OfficeOpenXml.ExcelPackageLoadLimits)
                Me.New(firstRowContainsColumnNames, startReadingAtRowIndex)
                If packageLoadLimits Is Nothing Then
                    Throw New ArgumentNullException(NameOf(packageLoadLimits))
                End If
                Me.PackageLoadLimits = packageLoadLimits
            End Sub

            ''' <summary>
            ''' Gets the resource limits applied while opening an XLSX file.
            ''' </summary>
            Public ReadOnly Property PackageLoadLimits As OfficeOpenXml.ExcelPackageLoadLimits
#End If

            ''' <summary>
            ''' Gets whether the first imported row contains column names.
            ''' </summary>
            ''' <value>Whether the first imported row contains column names.</value>
            Public ReadOnly Property FirstRowContainsColumnNames As Boolean

            ''' <summary>
            ''' Gets the zero-based row index at which reading starts.
            ''' </summary>
            ''' <value>The zero-based row index at which importing starts.</value>
            Public ReadOnly Property StartReadingAtRowIndex As Integer

        End Class

        ''' <summary>
        ''' Defines immutable defaults for writing worksheet data.
        ''' </summary>
        Public NotInheritable Class WriteOptions

            ''' <summary>
            ''' Initializes write options.
            ''' </summary>
            ''' <param name="errorLevel">The compatibility error level. Zero writes fallback error values; any other value throws for invalid values.</param>
            Public Sub New(Optional errorLevel As Byte = 0)
                Me.ErrorLevel = errorLevel
            End Sub

            ''' <summary>
            ''' Gets the compatibility error level used while writing values.
            ''' </summary>
            ''' <value>The configured error reporting level.</value>
            Public ReadOnly Property ErrorLevel As Byte

        End Class

        Private Enum VariantType
            Empty
            [Object]
            [Error]
            [Boolean]
            [Byte]
            [Short]
            [Integer]
            [Long]
            [Single]
            [Double]
            [Decimal]
            [Currency]
            [Date]
            [String]
            [Char]
        End Enum

        Private Shared ReadOnly CarriageReturn As String = Char.ConvertFromUtf32(13)
        Private Shared ReadOnly LineFeed As String = Char.ConvertFromUtf32(10)

        Private Shared _ErrorLevel As Byte = 0
        ''' <summary>
        ''' Error level 0 doesn't throw exception when writing e.g. invalid date/time values (invalid for excel); Error level 1 throws them.
        ''' </summary>
        ''' <value>The configured error reporting level.</value>
        Public Shared Property ErrorLevel() As Byte
            Get
                Return _ErrorLevel
            End Get
            Set(ByVal Value As Byte)
                _ErrorLevel = Value
            End Set
        End Property

        Private Shared Function LegacyWriteOptions() As WriteOptions
            Return New WriteOptions(ErrorLevel)
        End Function

        Private Shared Function RequiredReadOptions(ByVal options As ReadOptions) As ReadOptions
            If options Is Nothing Then
                Throw New ArgumentNullException(NameOf(options))
            End If
            Return options
        End Function

        Private Shared Function RequiredWriteOptions(ByVal options As WriteOptions) As WriteOptions
            If options Is Nothing Then
                Throw New ArgumentNullException(NameOf(options))
            End If
            Return options
        End Function

        ''' <summary>
        ''' Creates a new Excel file with data.
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataSet">A dataset to write into the workbook</param>
        Public Shared Sub WriteDataSetToXlsFile(ByVal outputPath As String, ByVal dataSet As System.Data.DataSet)
            WriteDataSetToXlsFile(Nothing, outputPath, dataSet)
        End Sub

        ''' <summary>
        ''' Creates a workbook from a data set using explicit write options.
        ''' </summary>
        ''' <param name="outputPath">The output file name.</param>
        ''' <param name="dataSet">The data set to write.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataSetToXlsFileWithOptions(ByVal outputPath As String, ByVal dataSet As System.Data.DataSet, ByVal options As WriteOptions)
            WriteDataSetToXlsFileWithOptions(Nothing, outputPath, dataSet, options)
        End Sub

        ''' <summary>
        ''' Loads an Excel file, writes data into it, and saves the file again.
        ''' </summary>
        ''' <param name="inputPath">A file which shall be loaded</param>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataSet">A dataset to write into the workbook</param>
        Public Shared Sub WriteDataSetToXlsFile(ByVal inputPath As String, ByVal outputPath As String, ByVal dataSet As System.Data.DataSet)
            WriteDataSetToXlsFileWithOptions(inputPath, outputPath, dataSet, LegacyWriteOptions())
        End Sub

        ''' <summary>
        ''' Updates or creates a workbook from a data set using explicit write options.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template.</param>
        ''' <param name="outputPath">The output file name.</param>
        ''' <param name="dataSet">The data set to write.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataSetToXlsFileWithOptions(ByVal inputPath As String, ByVal outputPath As String, ByVal dataSet As System.Data.DataSet, ByVal options As WriteOptions)
            options = RequiredWriteOptions(options)
            Dim tables As New ArrayList
            Dim tableNames As New ArrayList
            If Not dataSet Is Nothing AndAlso dataSet.Tables.Count > 0 Then
                For MyCounter As Integer = 0 To dataSet.Tables.Count - 1
                    tables.Add(dataSet.Tables(MyCounter))
                    tableNames.Add(dataSet.Tables(MyCounter).TableName)
                Next
            End If
            WriteDataTableToXlsFileWithOptions(inputPath, outputPath, CType(tables.ToArray(GetType(DataTable)), DataTable()), CType(tableNames.ToArray(GetType(String)), String()), options)
        End Sub

        ''' <summary>
        ''' Creates a new Excel file with data.
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <remarks>
        ''' The data will be written to the sheet with the name as the datatable's name
        ''' </remarks>
        Public Shared Sub WriteDataTableToXlsFile(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable)
            WriteDataTableToXlsFile(Nothing, outputPath, dataTable, CType(Nothing, String))
        End Sub

        ''' <summary>
        ''' Creates a workbook from a data table using explicit write options.
        ''' </summary>
        ''' <param name="outputPath">The output file name.</param>
        ''' <param name="dataTable">The data table to write.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataTableToXlsFileWithOptions(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable, ByVal options As WriteOptions)
            WriteDataTableToXlsFileWithOptions(Nothing, outputPath, dataTable, CType(Nothing, String), options)
        End Sub

        ''' <summary>
        ''' Creates a new Excel file with data.
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <remarks>
        ''' The data will be written to the sheet with the name as the datatable's name
        ''' </remarks>
        Public Shared Sub WriteDataTableToXlsFileAndFirstSheet(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable)
            WriteDataTableToXlsFileAndFirstSheetWithOptions(outputPath, dataTable, LegacyWriteOptions())
        End Sub

        ''' <summary>
        ''' Writes a data table to the first worksheet using explicit write options.
        ''' </summary>
        ''' <param name="outputPath">The output file name.</param>
        ''' <param name="dataTable">The data table to write.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataTableToXlsFileAndFirstSheetWithOptions(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable, ByVal options As WriteOptions)
            options = RequiredWriteOptions(options)
            If outputPath = Nothing OrElse (New System.IO.FileInfo(outputPath)).FullName = Nothing Then
                Throw New ArgumentNullException(NameOf(outputPath), "The output filename is required")
            End If

            Dim exportWorkbook As OfficeOpenXml.ExcelPackage
            exportWorkbook = OpenAndWriteDataTableToXlsFile(Nothing, New DataTable() {dataTable}, Array.Empty(Of String)(), SpecialSheet.FirstSheet, options)
            If exportWorkbook Is Nothing Then
                Return
            End If
            SaveWorkbook(exportWorkbook, outputPath)
        End Sub

        ''' <summary>
        ''' Creates a new Excel file with data.
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <remarks>
        ''' The data will be written to the sheet with the name as the datatable's name
        ''' </remarks>
        Public Shared Sub WriteDataTableToXlsFileAndCurrentSheet(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable)
            WriteDataTableToXlsFileAndCurrentSheetWithOptions(outputPath, dataTable, LegacyWriteOptions())
        End Sub

        ''' <summary>
        ''' Writes a data table to the current worksheet using explicit write options.
        ''' </summary>
        ''' <param name="outputPath">The output file name.</param>
        ''' <param name="dataTable">The data table to write.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataTableToXlsFileAndCurrentSheetWithOptions(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable, ByVal options As WriteOptions)
            options = RequiredWriteOptions(options)
            If outputPath = Nothing OrElse (New System.IO.FileInfo(outputPath)).FullName = Nothing Then
                Throw New ArgumentNullException(NameOf(outputPath), "The output filename is required")
            End If

            Dim exportWorkbook As OfficeOpenXml.ExcelPackage
            exportWorkbook = OpenAndWriteDataTableToXlsFile(Nothing, New DataTable() {dataTable}, Array.Empty(Of String)(), SpecialSheet.CurrentSheet, options)
            If exportWorkbook Is Nothing Then
                Return
            End If
            SaveWorkbook(exportWorkbook, outputPath)
        End Sub

        ''' <summary>
        ''' Creates a new Excel file with data.
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <param name="sheetName">The name the sheet which shall be updated/added</param>
        Public Shared Sub WriteDataTableToXlsFile(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable, ByVal sheetName As String)
            WriteDataTableToXlsFile(Nothing, outputPath, dataTable, sheetName)
        End Sub

        ''' <summary>
        ''' Creates a workbook with a named worksheet using explicit write options.
        ''' </summary>
        ''' <param name="outputPath">The output file name.</param>
        ''' <param name="dataTable">The data table to write.</param>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataTableToXlsFileWithOptions(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable, ByVal sheetName As String, ByVal options As WriteOptions)
            WriteDataTableToXlsFileWithOptions(Nothing, outputPath, dataTable, sheetName, options)
        End Sub

        ''' <summary>
        ''' Loads an Excel file, writes data into it, and saves the file again.
        ''' </summary>
        ''' <param name="inputPath">A file which shall be loaded</param>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <param name="sheetName">The name the sheet which shall be updated/added</param>
        Public Shared Sub WriteDataTableToXlsFile(ByVal inputPath As String, ByVal outputPath As String, ByVal dataTable As System.Data.DataTable, ByVal sheetName As String)
            WriteDataTableToXlsFile(inputPath, outputPath, New DataTable() {dataTable}, New String() {sheetName})
        End Sub

        ''' <summary>
        ''' Updates or creates a named worksheet using explicit write options.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template.</param>
        ''' <param name="outputPath">The output file name.</param>
        ''' <param name="dataTable">The data table to write.</param>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataTableToXlsFileWithOptions(ByVal inputPath As String, ByVal outputPath As String, ByVal dataTable As System.Data.DataTable, ByVal sheetName As String, ByVal options As WriteOptions)
            WriteDataTableToXlsFileWithOptions(inputPath, outputPath, New DataTable() {dataTable}, New String() {sheetName}, options)
        End Sub

        ''' <summary>
        ''' Updates or creates an Excel file, writes data into it, and saves the file again.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template</param>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTables">Data tables to write to the workbook.</param>
        ''' <param name="sheetNames">Worksheet names to update or add in the same order as <paramref name="dataTables"/>.</param>
        Public Shared Sub WriteDataTableToXlsFile(ByVal inputPath As String, ByVal outputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String())
            WriteDataTableToXlsFileWithOptions(inputPath, outputPath, dataTables, sheetNames, LegacyWriteOptions())
        End Sub

        ''' <summary>
        ''' Updates or creates worksheets using explicit write options.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template.</param>
        ''' <param name="outputPath">The output file name.</param>
        ''' <param name="dataTables">The data tables to write.</param>
        ''' <param name="sheetNames">The corresponding worksheet names.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataTableToXlsFileWithOptions(ByVal inputPath As String, ByVal outputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String(), ByVal options As WriteOptions)
            options = RequiredWriteOptions(options)
            If outputPath = Nothing OrElse (New System.IO.FileInfo(outputPath)).FullName = Nothing Then
                Throw New ArgumentNullException(NameOf(outputPath), "The output filename is required")
            End If

            Dim exportWorkbook As OfficeOpenXml.ExcelPackage
            exportWorkbook = OpenAndWriteDataTableToXlsFile(inputPath, dataTables, sheetNames, SpecialSheet.AsDefinedInSheetNamesCollection, options)
            If exportWorkbook Is Nothing Then
                Return
            End If
            SaveWorkbook(exportWorkbook, outputPath)
        End Sub

        ''' <summary>
        ''' Saves the changed worksheet.
        ''' </summary>
        ''' <param name="exportWorkbook"></param>
        ''' <param name="outputPath"></param>
        ''' <remarks></remarks>
        Private Shared Sub SaveWorkbook(ByVal exportWorkbook As OfficeOpenXml.ExcelPackage, ByVal outputPath As String)
            If outputPath <> Nothing AndAlso outputPath.ToLower.EndsWith(".xlsb") Then
                'Excel 2007 binary format
                Throw New NotSupportedException("Excel2007 binary file format not supported yet")
            ElseIf outputPath <> Nothing AndAlso outputPath.ToLower.EndsWith(".xlsm") Then
                'Excel 2007 macro format
                exportWorkbook.SaveAs(New IO.FileInfo(outputPath))
            Else
                'Excel 2007 standard format
                exportWorkbook.SaveAs(New IO.FileInfo(outputPath))
            End If
        End Sub

        ''' <summary>
        ''' The sheet which is subject of operations.
        ''' </summary>
        ''' <remarks></remarks>
        Private Enum SpecialSheet As Byte
            AsDefinedInSheetNamesCollection = 0
            FirstSheet = 1
            CurrentSheet = 2
        End Enum

        ''' <summary>
        ''' Updates or creates an Excel file, writes data into it, and saves the file again.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template</param>
        ''' <param name="dataTables">Data tables to write to the workbook.</param>
        ''' <param name="sheetNames">Worksheet names to update or add in the same order as <paramref name="dataTables"/>.</param>
        ''' <param name="specialSheet">A special sheet</param>
        ''' <returns>A Workbook object</returns>
        ''' <remarks></remarks>
        Private Shared Function OpenAndWriteDataTableToXlsFile(ByVal inputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetnames As String(), ByVal specialSheet As SpecialSheet, ByVal options As WriteOptions) As OfficeOpenXml.ExcelPackage

            'Some parameter validation, first
            If dataTables Is Nothing Then
                Return Nothing
            ElseIf specialSheet = SpecialSheet.AsDefinedInSheetNamesCollection AndAlso (sheetnames Is Nothing OrElse dataTables.Length <> sheetnames.Length) Then
                Throw New ArgumentException("Arrays must have the same length", NameOf(sheetnames))
            Else
                Select Case specialSheet
                    Case SpecialSheet.CurrentSheet, SpecialSheet.FirstSheet
                        If dataTables.Length <> 1 Then
                            Throw New ArgumentException("Tables array must contain exactly 1 item for sheet mode FirstSheet or CurrentSheet", NameOf(dataTables))
                        End If
                End Select
            End If

            Dim exportWorkbook As OfficeOpenXml.ExcelPackage

            'Read existing file
            If inputPath <> Nothing Then
                exportWorkbook = LoadWorkbookFile(inputPath)
            Else
                exportWorkbook = New OfficeOpenXml.ExcelPackage
            End If

            For MyDataTableCounter As Integer = 0 To dataTables.Length - 1
                Dim dataTable As DataTable = dataTables(MyDataTableCounter)
                Dim sheetName As String
                Select Case specialSheet
                    Case SpecialSheet.CurrentSheet
                        If exportWorkbook.Workbook.Worksheets.Count = 0 Then
                            exportWorkbook.Workbook.Worksheets.Add(dataTable.TableName)
                            sheetName = dataTable.TableName
                        Else
                            sheetName = exportWorkbook.Workbook.Worksheets(exportWorkbook.Workbook.View.ActiveTab).Name
                        End If
                    Case SpecialSheet.FirstSheet
                        If exportWorkbook.Workbook.Worksheets.Count > 0 Then
                            sheetName = exportWorkbook.Workbook.Worksheets(0).Name
                        Else
                            sheetName = dataTable.TableName
                        End If
                    Case Else 'XlsEpplus.SpecialSheet.AsDefinedInSheetNamesCollection
                        sheetName = sheetnames(MyDataTableCounter)
                        If sheetName = Nothing Then sheetName = dataTable.TableName
                End Select

                'Find existing work sheet or add new one
                Dim SheetIndex As Integer = ResolveWorksheetIndex(exportWorkbook, sheetName)
                If SheetIndex = -1 Then
                    Dim sheet As OfficeOpenXml.ExcelWorksheet
                    sheet = exportWorkbook.Workbook.Worksheets.Add(sheetName)
                    SheetIndex = sheet.Index
                End If

                Dim WorkSheet As OfficeOpenXml.ExcelWorksheet = exportWorkbook.Workbook.Worksheets(SheetIndex) 'CType(exportWorkbook.Workbook.Worksheets(SheetIndex), EpplusFreeOfficeOpenXml.ExcelWorksheet)
                'WorkSheet.Cells(1, 1).LoadFromDataTable(dataTable, True)

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
                        If Convert.IsDBNull(value) Then
                            WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = Nothing
                        ElseIf value.GetType Is GetType(String) Then
                            'Excel requires line-breaks to be an LF character only, not a windows typical CR+LF
                            Dim cell As OfficeOpenXml.ExcelRange = WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1)
                            If CType(value, String) <> "" Then
                                value = CType(value, String).Replace(CarriageReturn & LineFeed, LineFeed) 'Windows line breaks
                                value = CType(value, String).Replace(CarriageReturn, LineFeed) 'Mac or Linux line break
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
                                If options.ErrorLevel = 0 Then
                                    WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = Double.NaN
                                Else
                                    Throw New InvalidOperationException("Error writing a date/time value """ & datevalue.ToString(System.Globalization.CultureInfo.InvariantCulture) & """ in row " & (RowCounter + 1), ex)
                                End If
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
                Try
                    For MyCounter As Integer = 1 To WorkSheet.Dimension.End.Column
                        WorkSheet.Column(MyCounter).AutoFit(0.5)
                    Next
                    'For MyCounter As Integer = 1 To WorkSheet.Dimension.End.Row
                    '    WorkSheet.Row(MyCounter).AutoFit()
                    'Next
                Catch ex As PlatformNotSupportedException
                    'System.Drawing.Common is not supported on platform
                    'just ignore AutoFit feature
                Catch ex As System.TypeInitializationException
                    'The type initializer for 'Gdip' threw an exception.
                    '---> System.PlatformNotSupportedException System.Drawing.Common Is Not supported on non-Windows platforms. See https://aka.ms/systemdrawingnonwindows for more information.
                    'just ignore AutoFit feature
                End Try
            Next

            Return exportWorkbook

        End Function

        ''' <summary>
        ''' Defines Excel file formats.
        ''' </summary>
        Public Enum FileFormat As Byte
            ''' <summary>
            ''' Excel 2007 or newer workbook format without macros.
            ''' </summary>
            Excel2007 = 1
            ''' <summary>
            ''' Excel 2007 or newer macro-enabled workbook format.
            ''' </summary>
            Excel2007Macro = 2
        End Enum

        ''' <summary>
        ''' Updates or creates an Excel file, writes data into it, and saves the file to the output stream.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template</param>
        ''' <param name="outputStream">An opened output stream</param>
        ''' <param name="dataTables">Data tables to write to the workbook.</param>
        ''' <param name="sheetNames">Worksheet names to update or add in the same order as <paramref name="dataTables"/>.</param>
        ''' <param name="fileFormat">Workbook file format.</param>
        Public Shared Sub WriteDataTableToXlsStream(ByVal inputPath As String, ByVal outputStream As System.IO.Stream, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String(), ByVal fileFormat As FileFormat)
            WriteDataTableToXlsStreamWithOptions(inputPath, outputStream, dataTables, sheetNames, fileFormat, LegacyWriteOptions())
        End Sub

        ''' <summary>
        ''' Updates or creates a workbook and writes it to a stream using explicit write options.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template.</param>
        ''' <param name="outputStream">The opened output stream.</param>
        ''' <param name="dataTables">The data tables to write.</param>
        ''' <param name="sheetNames">The corresponding worksheet names.</param>
        ''' <param name="fileFormat">The workbook file format.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataTableToXlsStreamWithOptions(ByVal inputPath As String, ByVal outputStream As System.IO.Stream, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String(), ByVal fileFormat As FileFormat, ByVal options As WriteOptions)
            options = RequiredWriteOptions(options)
            Dim exportWorkbook As OfficeOpenXml.ExcelPackage
            exportWorkbook = OpenAndWriteDataTableToXlsFile(inputPath, dataTables, sheetNames, SpecialSheet.AsDefinedInSheetNamesCollection, options)
            If exportWorkbook Is Nothing Then
                Return
            Else
                If fileFormat = FileFormat.Excel2007 Then
                    exportWorkbook.SaveAs(outputStream)
                ElseIf fileFormat = FileFormat.Excel2007Macro Then
                    exportWorkbook.SaveAs(outputStream)
                Else
                    Throw New NotSupportedException("value for fileformat is invalid")
                End If
            End If
        End Sub

        ''' <summary>
        ''' Writes a data table to an HTTP response.
        ''' </summary>
        ''' <param name="dataTable">The data table to write.</param>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="httpContext">The HTTP listener context receiving the response.</param>
        ''' <param name="fileFormat">The workbook file format.</param>
        ''' <param name="suggestedFileNameToBrowser">The file name suggested to the browser.</param>
        Public Shared Sub WriteDataTableToXlsHttpResponse(ByVal dataTable As System.Data.DataTable, ByVal sheetName As String, ByVal httpContext As System.Net.HttpListenerContext, ByVal fileFormat As FileFormat, ByVal suggestedFileNameToBrowser As String)
            WriteDataTableToXlsHttpResponse(String.Empty, New DataTable() {dataTable}, New String() {sheetName}, httpContext, fileFormat, suggestedFileNameToBrowser)
        End Sub

        ''' <summary>
        ''' Writes a data table to an HTTP response using explicit write options.
        ''' </summary>
        ''' <param name="dataTable">The data table to write.</param>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="httpContext">The HTTP listener context receiving the response.</param>
        ''' <param name="fileFormat">The workbook file format.</param>
        ''' <param name="suggestedFileNameToBrowser">The file name suggested to the browser.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataTableToXlsHttpResponseWithOptions(ByVal dataTable As System.Data.DataTable, ByVal sheetName As String, ByVal httpContext As System.Net.HttpListenerContext, ByVal fileFormat As FileFormat, ByVal suggestedFileNameToBrowser As String, ByVal options As WriteOptions)
            WriteDataTableToXlsHttpResponseWithOptions(String.Empty, New DataTable() {dataTable}, New String() {sheetName}, httpContext, fileFormat, suggestedFileNameToBrowser, options)
        End Sub

        ''' <summary>
        ''' Updates or creates a workbook and writes it to an HTTP response.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template.</param>
        ''' <param name="dataTables">The data tables to write.</param>
        ''' <param name="sheetNames">The worksheet names corresponding to <paramref name="dataTables"/>.</param>
        ''' <param name="httpContext">The HTTP listener context receiving the response.</param>
        ''' <param name="fileFormat">The workbook file format.</param>
        Public Shared Sub WriteDataTableToXlsHttpResponse(ByVal inputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String(), ByVal httpContext As System.Net.HttpListenerContext, ByVal fileFormat As FileFormat)
            WriteDataTableToXlsHttpResponse(inputPath, dataTables, sheetNames, httpContext, fileFormat, String.Empty)
        End Sub

        ''' <summary>
        ''' Updates or creates a workbook and writes it to an HTTP response using explicit write options.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template.</param>
        ''' <param name="dataTables">The data tables to write.</param>
        ''' <param name="sheetNames">The corresponding worksheet names.</param>
        ''' <param name="httpContext">The HTTP listener context receiving the response.</param>
        ''' <param name="fileFormat">The workbook file format.</param>
        ''' <param name="options">The immutable write options.</param>
        Public Shared Sub WriteDataTableToXlsHttpResponseWithOptions(ByVal inputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String(), ByVal httpContext As System.Net.HttpListenerContext, ByVal fileFormat As FileFormat, ByVal options As WriteOptions)
            WriteDataTableToXlsHttpResponseWithOptions(inputPath, dataTables, sheetNames, httpContext, fileFormat, String.Empty, options)
        End Sub

        ''' <summary>
        ''' Updates or creates a workbook and writes it to an HTTP response.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template.</param>
        ''' <param name="dataTables">The data tables to write.</param>
        ''' <param name="sheetNames">The worksheet names corresponding to <paramref name="dataTables"/>.</param>
        ''' <param name="httpContext">The HTTP listener context receiving the response.</param>
        ''' <param name="fileFormat">The workbook file format.</param>
        ''' <param name="suggestedFileNameToBrowser">The file name suggested to the browser. The default is <c>report.xlsx</c> or <c>report.xlsm</c>.</param>
        ''' <exception cref="ArgumentNullException"><paramref name="dataTables"/> or <paramref name="httpContext"/> is <see langword="Nothing"/>.</exception>
        ''' <exception cref="InvalidOperationException">The workbook could not be created.</exception>
        ''' <exception cref="NotSupportedException"><paramref name="fileFormat"/> is not supported.</exception>
        Public Shared Sub WriteDataTableToXlsHttpResponse(ByVal inputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String(), ByVal httpContext As System.Net.HttpListenerContext, ByVal fileFormat As FileFormat, ByVal suggestedFileNameToBrowser As String)
            WriteDataTableToXlsHttpResponseWithOptions(inputPath, dataTables, sheetNames, httpContext, fileFormat, suggestedFileNameToBrowser, LegacyWriteOptions())
        End Sub

        ''' <summary>
        ''' Updates or creates a workbook and writes it to an HTTP response using explicit write options.
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template.</param>
        ''' <param name="dataTables">The data tables to write.</param>
        ''' <param name="sheetNames">The corresponding worksheet names.</param>
        ''' <param name="httpContext">The HTTP listener context receiving the response.</param>
        ''' <param name="fileFormat">The workbook file format.</param>
        ''' <param name="suggestedFileNameToBrowser">The file name suggested to the browser.</param>
        ''' <param name="options">The immutable write options.</param>
        ''' <exception cref="ArgumentNullException"><paramref name="dataTables"/>, <paramref name="httpContext"/>, or <paramref name="options"/> is <see langword="Nothing"/>.</exception>
        ''' <exception cref="InvalidOperationException">The workbook could not be created.</exception>
        ''' <exception cref="NotSupportedException"><paramref name="fileFormat"/> is not supported.</exception>
        Public Shared Sub WriteDataTableToXlsHttpResponseWithOptions(ByVal inputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String(), ByVal httpContext As System.Net.HttpListenerContext, ByVal fileFormat As FileFormat, ByVal suggestedFileNameToBrowser As String, ByVal options As WriteOptions)
            options = RequiredWriteOptions(options)
            If dataTables Is Nothing Then
                Throw New ArgumentNullException(NameOf(dataTables))
            End If
            If httpContext Is Nothing Then
                Throw New ArgumentNullException(NameOf(httpContext))
            End If

            Dim exportWorkbook = OpenAndWriteDataTableToXlsFile(inputPath, dataTables, sheetNames, SpecialSheet.AsDefinedInSheetNamesCollection, options)
            If exportWorkbook Is Nothing Then
                Throw New InvalidOperationException("Workbook creation failed - missing workbook")
            End If

            httpContext.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            If String.IsNullOrEmpty(suggestedFileNameToBrowser) Then
                If fileFormat = FileFormat.Excel2007Macro Then
                    suggestedFileNameToBrowser = "report.xlsm"
                Else
                    suggestedFileNameToBrowser = "report.xlsx"
                End If
            End If
            httpContext.Response.AddHeader("Content-Disposition", "attachment; filename=" & System.Net.WebUtility.UrlEncode(suggestedFileNameToBrowser))

            If fileFormat = FileFormat.Excel2007 OrElse fileFormat = FileFormat.Excel2007Macro Then
                exportWorkbook.SaveAs(httpContext.Response.OutputStream)
            Else
                Throw New NotSupportedException("value for fileformat is invalid")
            End If
        End Sub

        ''' <summary>
        ''' Reads all sheets from an excel sheet into a dataset.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <returns>A dataset with one or more, independent tables.</returns>
        ''' <remarks>
        '''     The table names are as the sheet names.
        ''' 
        '''     Conversion errors will not be ignored!
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' </remarks>
        Public Shared Function ReadDataSetFromXlsFile(ByVal inputPath As String, ByVal firstRowContainsColumnNames As Boolean) As DataSet

            Return ReadDataSetFromXlsFileWithOptions(inputPath, New ReadOptions(firstRowContainsColumnNames))

        End Function

        ''' <summary>
        ''' Reads all worksheets into a data set using explicit read options.
        ''' </summary>
        ''' <param name="inputPath">The workbook file name.</param>
        ''' <param name="options">The immutable read options.</param>
        ''' <returns>A data set containing one table for each worksheet.</returns>
        ''' <exception cref="ArgumentNullException"><paramref name="inputPath"/> or <paramref name="options"/> is <see langword="Nothing"/>.</exception>
        Public Shared Function ReadDataSetFromXlsFileWithOptions(ByVal inputPath As String, ByVal options As ReadOptions) As DataSet

            options = RequiredReadOptions(options)

            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException(NameOf(inputPath), "The input filename is required")
            End If

            Dim importWorkbook As OfficeOpenXml.ExcelPackage

            'Load the worksheet
#If CM_FIXCALCS Then
            importWorkbook = LoadWorkbookFile(inputPath, options.PackageLoadLimits)
#Else
            importWorkbook = LoadWorkbookFile(inputPath)
#End If

            Dim Result As New DataSet

            For sheetCounter As Integer = 0 To importWorkbook.Workbook.Worksheets.Count - 1
                Dim Sheet As OfficeOpenXml.ExcelWorksheet = importWorkbook.Workbook.Worksheets(sheetCounter)

                'Detect the column types which must be used
                Dim sheetData As DataTable = ReadDataTableFromXlsFileCreateDataTableSuggestion(Sheet, Sheet.Name, options.StartReadingAtRowIndex, options.FirstRowContainsColumnNames)

                'Read all data and put it into the datatable
                ReadDataTableFromXlsFile(Sheet, options.StartReadingAtRowIndex, options.FirstRowContainsColumnNames, sheetData)

                Result.Tables.Add(sheetData)
            Next

            'Return the result
            Return Result

        End Function

        Private Shared Function LoadWorkbookFile(inputPath As String) As OfficeOpenXml.ExcelPackage
            'Load the changed worksheet
            Dim file As New System.IO.FileInfo(inputPath)
            If file.Exists = False Then
                Throw New System.IO.FileNotFoundException("Missing file: " & file.ToString, file.ToString)
            End If
            Dim importWorkbook As New OfficeOpenXml.ExcelPackage(file)
            Return importWorkbook
        End Function

#If CM_FIXCALCS Then
        Private Shared Function LoadWorkbookFile(inputPath As String, loadLimits As OfficeOpenXml.ExcelPackageLoadLimits) As OfficeOpenXml.ExcelPackage
            Dim file As New System.IO.FileInfo(inputPath)
            If Not file.Exists Then
                Throw New System.IO.FileNotFoundException("Missing file: " & file.ToString(), file.ToString())
            End If
            Return OfficeOpenXml.ExcelPackage.OpenWithLoadLimits(file, loadLimits)
        End Function
#End If

        ''' <summary>
        ''' Reads the data from an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <returns>The worksheet data.</returns>
        ''' <remarks>
        '''     The first sheet will be used for reading data.
        '''     Values in first row will be assigned as column names.
        ''' 
        '''     Conversion errors will not be ignored!
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' </remarks>
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String) As DataTable
            Return ReadDataTableFromXlsFile(inputPath, True)
        End Function

        ''' <summary>
        ''' Reads the first worksheet into a data table using explicit read options.
        ''' </summary>
        ''' <param name="inputPath">The workbook file name.</param>
        ''' <param name="options">The immutable read options.</param>
        ''' <returns>A data table containing the worksheet data.</returns>
        ''' <exception cref="ArgumentNullException"><paramref name="inputPath"/> or <paramref name="options"/> is <see langword="Nothing"/>.</exception>
        Public Shared Function ReadDataTableFromXlsFileWithOptions(ByVal inputPath As String, ByVal options As ReadOptions) As DataTable
            options = RequiredReadOptions(options)
            Return ReadDataTableFromXlsFileCore(inputPath, options.StartReadingAtRowIndex, options.FirstRowContainsColumnNames, options)
        End Function

        ''' <summary>
        ''' Reads the data from an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <returns>The worksheet data.</returns>
        ''' <remarks>
        '''     The first sheet will be used for reading data.
        ''' 
        '''     Conversion errors will not be ignored!
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' </remarks>
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal firstRowContainsColumnNames As Boolean) As DataTable
            Return ReadDataTableFromXlsFile(inputPath, 0, firstRowContainsColumnNames)
        End Function

        ''' <summary>
        ''' Reads the data from an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <returns>The worksheet data.</returns>
        ''' <remarks>
        '''     The first sheet will be used for reading data.
        ''' 
        '''     Conversion errors will not be ignored!
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' </remarks>
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean) As DataTable
            Return ReadDataTableFromXlsFileCore(inputPath, startReadingAtRowIndex, firstRowContainsColumnNames, Nothing)
        End Function

        Private Shared Function ReadDataTableFromXlsFileCore(ByVal inputPath As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean, ByVal options As ReadOptions) As DataTable

            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException(NameOf(inputPath), "The input filename is required")
            End If

            Dim importWorkbook As OfficeOpenXml.ExcelPackage

            'Save the changed worksheet
#If CM_FIXCALCS Then
            importWorkbook = If(options Is Nothing, LoadWorkbookFile(inputPath), LoadWorkbookFile(inputPath, options.PackageLoadLimits))
#Else
            importWorkbook = LoadWorkbookFile(inputPath)
#End If
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = importWorkbook.Workbook.Worksheets(0)

            'Detect the column types which must be used
            Dim Result As DataTable = ReadDataTableFromXlsFileCreateDataTableSuggestion(Sheet, Sheet.Name, startReadingAtRowIndex, firstRowContainsColumnNames)

            'Read all data and put it into the datatable
            ReadDataTableFromXlsFile(Sheet, startReadingAtRowIndex, firstRowContainsColumnNames, Result)

            'Return the result
            Return Result

        End Function

        ''' <summary>
        ''' Reads the data from an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <returns>The worksheet data.</returns>
        ''' <remarks>
        '''     Values in first row will be assigned as column names.
        '''     Conversion errors will not be ignored!
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' </remarks>
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String) As DataTable
            Return ReadDataTableFromXlsFile(inputPath, sheetName, True)
        End Function

        ''' <summary>
        ''' Reads a named worksheet into a data table using explicit read options.
        ''' </summary>
        ''' <param name="inputPath">The workbook file name.</param>
        ''' <param name="sheetName">The worksheet name, or <see langword="Nothing"/> for the first worksheet.</param>
        ''' <param name="options">The immutable read options.</param>
        ''' <returns>A data table containing the worksheet data.</returns>
        ''' <exception cref="ArgumentNullException"><paramref name="inputPath"/> or <paramref name="options"/> is <see langword="Nothing"/>.</exception>
        Public Shared Function ReadDataTableFromXlsFileWithOptions(ByVal inputPath As String, ByVal sheetName As String, ByVal options As ReadOptions) As DataTable
            options = RequiredReadOptions(options)
            Return ReadDataTableFromXlsFileCore(inputPath, sheetName, options.StartReadingAtRowIndex, options.FirstRowContainsColumnNames, options)
        End Function

        ''' <summary>
        ''' Reads the data from an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <returns>The worksheet data.</returns>
        ''' <remarks>
        '''     Conversion errors will not be ignored!
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' </remarks>
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String, ByVal firstRowContainsColumnNames As Boolean) As DataTable
            Return ReadDataTableFromXlsFile(inputPath, sheetName, 0, firstRowContainsColumnNames)
        End Function

        ''' <summary>
        ''' Reads the data from an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <returns>The worksheet data.</returns>
        ''' <remarks>
        '''     Conversion errors will not be ignored!
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' </remarks>
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean) As DataTable
            Return ReadDataTableFromXlsFileCore(inputPath, sheetName, startReadingAtRowIndex, firstRowContainsColumnNames, Nothing)
        End Function

        Private Shared Function ReadDataTableFromXlsFileCore(ByVal inputPath As String, ByVal sheetName As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean, ByVal options As ReadOptions) As DataTable
            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException(NameOf(inputPath), "The input filename is required")
            End If

            Dim importWorkbook As OfficeOpenXml.ExcelPackage

            'Save the changed worksheet
#If CM_FIXCALCS Then
            importWorkbook = If(options Is Nothing, LoadWorkbookFile(inputPath), LoadWorkbookFile(inputPath, options.PackageLoadLimits))
#Else
            importWorkbook = LoadWorkbookFile(inputPath)
#End If
            If sheetName = Nothing Then
                sheetName = importWorkbook.Workbook.Worksheets.First.Name
            End If
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = LookupWorksheet(importWorkbook, sheetName)

            'Detect the column types which must be used
            Dim Result As DataTable = ReadDataTableFromXlsFileCreateDataTableSuggestion(Sheet, sheetName, startReadingAtRowIndex, firstRowContainsColumnNames)

            'Read all data and put it into the datatable
            ReadDataTableFromXlsFile(Sheet, startReadingAtRowIndex, firstRowContainsColumnNames, Result)

            'Return the result
            Return Result

        End Function

        ''' <summary>
        ''' Reads the data from an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <param name="data">The datatable which shall be filled; only columns which exist in this target table will be imported</param>
        ''' <remarks>
        '''     Conversion errors will not be ignored!
        ''' 
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' 
        '''     Dependent on the firstRowContainsColumnNames parameter, the datatable parameter must contain a table with column names as they're defined in the first row of the excel sheet or the table's columnn must have the name of the column index in excel ("1", "2", "3", ...)
        ''' </remarks>
        Public Shared Sub ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String, ByVal firstRowContainsColumnNames As Boolean, ByVal data As DataTable)
            ReadDataTableFromXlsFile(inputPath, sheetName, 0, firstRowContainsColumnNames, data)
        End Sub

        ''' <summary>
        ''' Reads a named worksheet into an existing data table using explicit read options.
        ''' </summary>
        ''' <param name="inputPath">The workbook file name.</param>
        ''' <param name="sheetName">The worksheet name, or <see langword="Nothing"/> for the first worksheet.</param>
        ''' <param name="options">The immutable read options.</param>
        ''' <param name="data">The data table to fill.</param>
        ''' <exception cref="ArgumentNullException"><paramref name="inputPath"/>, <paramref name="options"/>, or <paramref name="data"/> is <see langword="Nothing"/>.</exception>
        Public Shared Sub ReadDataTableFromXlsFileWithOptions(ByVal inputPath As String, ByVal sheetName As String, ByVal options As ReadOptions, ByVal data As DataTable)
            options = RequiredReadOptions(options)
            ReadDataTableFromXlsFileCore(inputPath, sheetName, options.StartReadingAtRowIndex, options.FirstRowContainsColumnNames, data, options)
        End Sub

        ''' <summary>
        ''' Reads the data from first sheet of an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <param name="data">The datatable which shall be filled; only columns which exist in this target table will be imported</param>
        ''' <remarks>
        '''     Conversion errors will not be ignored!
        ''' 
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' 
        '''     Dependent on the firstRowContainsColumnNames parameter, the datatable parameter must contain a table with column names as they're defined in the first row of the excel sheet or the table's columnn must have the name of the column index in excel ("1", "2", "3", ...)
        ''' </remarks>
        Public Shared Sub ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean, ByVal data As DataTable)
            ReadDataTableFromXlsFile(inputPath, Nothing, startReadingAtRowIndex, firstRowContainsColumnNames, data)
        End Sub

        ''' <summary>
        ''' Reads the data from an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <param name="data">The datatable which shall be filled; only columns which exist in this target table will be imported</param>
        ''' <remarks>
        '''     Conversion errors will not be ignored!
        ''' 
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3   --> Cell value of type System.Exception with error details
        '''     #REF! 4  --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6   --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' 
        '''     Dependent on the firstRowContainsColumnNames parameter, the datatable parameter must contain a table with column names as they're defined in the first row of the excel sheet or the table's columnn must have the name of the column index in excel ("1", "2", "3", ...)
        ''' </remarks>
        Public Shared Sub ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean, ByVal data As DataTable)
            ReadDataTableFromXlsFileCore(inputPath, sheetName, startReadingAtRowIndex, firstRowContainsColumnNames, data, Nothing)
        End Sub

        Private Shared Sub ReadDataTableFromXlsFileCore(ByVal inputPath As String, ByVal sheetName As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean, ByVal data As DataTable, ByVal options As ReadOptions)

            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException(NameOf(inputPath), "The input filename is required")
            ElseIf data Is Nothing Then
                Throw New ArgumentNullException(NameOf(data), "A datatable must be predefined which shall hold all the data")
            End If

            Dim importWorkbook As OfficeOpenXml.ExcelPackage

            'Save the changed worksheet
#If CM_FIXCALCS Then
            importWorkbook = If(options Is Nothing, LoadWorkbookFile(inputPath), LoadWorkbookFile(inputPath, options.PackageLoadLimits))
#Else
            importWorkbook = LoadWorkbookFile(inputPath)
#End If
            If sheetName = Nothing Then
                sheetName = importWorkbook.Workbook.Worksheets.First.Name
            End If
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = LookupWorksheet(importWorkbook, sheetName)

            'Extend table's column set as long as columns count matches
            ReadDataTableFromXlsFileExtendDataTableColumns(data, Sheet, startReadingAtRowIndex, firstRowContainsColumnNames)

            'Read all data and put it into the datatable
            ReadDataTableFromXlsFile(Sheet, startReadingAtRowIndex, firstRowContainsColumnNames, data)

        End Sub

        ''' <summary>
        ''' Reads the available sheet names from an XLS file.
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <returns>The worksheet names.</returns>
        Public Shared Function ReadSheetNamesFromXlsFile(ByVal inputPath As String) As String()

            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException(NameOf(inputPath), "The input filename is required")
            End If

            Dim importWorkbook As OfficeOpenXml.ExcelPackage

            'Save the changed worksheet
            importWorkbook = LoadWorkbookFile(inputPath)

            Dim Result As New ArrayList

            For sheetCounter As Integer = 0 To importWorkbook.Workbook.Worksheets.Count - 1
                Dim Sheet As OfficeOpenXml.ExcelWorksheet = importWorkbook.Workbook.Worksheets(sheetCounter)
                Result.Add(Sheet.Name)
            Next

            'Return the result
            Return CType(Result.ToArray(GetType(String)), String())

        End Function

#Region "Internal tools"

        ''' <summary>
        ''' Looks up the last content column index (zero-based index) (the last content cell might differ from Excel's special cell xlLastCell).
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function LookupLastContentColumnIndex(ByVal sheet As OfficeOpenXml.ExcelWorksheet) As Integer
            If sheet.Dimension Is Nothing Then Return 0
            Dim autoSuggestionLastRowIndex As Integer = sheet.Dimension.End.Row - 1
            Dim autoSuggestedResult As Integer = sheet.Dimension.End.Column - 1
            For colCounter As Integer = autoSuggestedResult To 0 Step -1
                For rowCounter As Integer = 0 To autoSuggestionLastRowIndex
                    If IsEmptyCell(sheet, rowCounter, colCounter) = False Then
                        Return colCounter
                    End If
                Next
            Next
            Return 0
        End Function

        ''' <summary>
        ''' Looks up the last content row index (zero-based index) (the last content cell might differ from Excel's special cell xlLastCell).
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function LookupLastContentRowIndex(ByVal sheet As OfficeOpenXml.ExcelWorksheet) As Integer
            If sheet.Dimension Is Nothing Then Return 0
            Dim autoSuggestionLastColumnIndex As Integer = sheet.Dimension.End.Column - 1
            Dim autoSuggestedResult As Integer = sheet.Dimension.End.Row - 1
            For rowCounter As Integer = autoSuggestedResult To 0 Step -1
                For colCounter As Integer = 0 To autoSuggestionLastColumnIndex
                    If IsEmptyCell(sheet, rowCounter, colCounter) = False Then
                        Return rowCounter
                    End If
                Next
            Next
            Return 0
        End Function

        ''' <summary>
        ''' Determines whether a cell contains empty content.
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="rowIndex">Zero-based index</param>
        ''' <param name="columnIndex">Zero-based index</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function IsEmptyCell(ByVal sheet As OfficeOpenXml.ExcelWorksheet, ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Boolean
            Dim value As Object = sheet.Cells(rowIndex + 1, columnIndex + 1).Value
            If value Is Nothing Then
                Return True
            ElseIf value.GetType Is GetType(String) AndAlso CType(value, String) = Nothing Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Reads the data from an excel sheet into a datatable.
        ''' </summary>
        ''' <param name="sheet">An excel sheet containing the required data</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <param name="data">The datatable which shall be filled; only columns which exist in this target table will be imported</param>
        ''' <remarks>
        '''     Excel error values
        '''     #NULL! 1   --> Cell value of type System.Exception with error details
        '''     #DIV/0! 2  --> Double.NaN
        '''     #VALUE! 3  --> Cell value of type System.Exception with error details
        '''     #REF! 4    --> Cell value of type System.Exception with error details
        '''     #NAME? 5   --> Cell value of type System.Exception with error details
        '''     #NUM! 6    --> Cell value of type System.Exception with error details
        '''     #NA 7      --> Cell value of type System.Exception with error details
        '''     {blank}    --> DBNull
        ''' 
        '''     Dependent on the firstRowContainsColumnNames parameter, the datatable parameter must contain a table with column names as they're defined in the first row of the excel sheet or the table's columnn must have the name of the column index in excel ("1", "2", "3", ...)
        ''' </remarks>
        Private Shared Sub ReadDataTableFromXlsFile(ByVal sheet As OfficeOpenXml.ExcelWorksheet, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean, ByVal data As DataTable)
            'Read all data and put it into the datatable (pay attention to field with blank content, #DIV/0 and all the other error types
            Dim firstRowIndexWithContent As Integer
            If firstRowContainsColumnNames Then
                firstRowIndexWithContent = 1
            Else
                firstRowIndexWithContent = 0
            End If
            firstRowIndexWithContent += startReadingAtRowIndex
            'sheet.CalcDimensions() 'Calculate the sheet end positions (to prevent bug that this information is 0, e. g. after saving and reloading with this component)
            For rowCounter As Integer = firstRowIndexWithContent To LookupLastContentRowIndex(sheet)
                Dim row As DataRow = data.NewRow
                For colCounter As Integer = 0 To LookupLastContentColumnIndex(sheet)
                    Dim value As Object
                    Select Case LookupDotNetType(sheet.Cells(rowCounter + 1, colCounter + 1))
                        Case VariantType.Empty
                            value = DBNull.Value
                        Case VariantType.Boolean
                            value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Boolean)
                        Case VariantType.Error
                            If data.Columns(colCounter).DataType Is GetType(Double) Then
                                If CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, OfficeOpenXml.ExcelErrorValue).ToString = OfficeOpenXml.ExcelErrorValue.Values.Div0 Then
                                    value = Double.NaN
                                ElseIf CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, OfficeOpenXml.ExcelErrorValue).ToString = OfficeOpenXml.ExcelErrorValue.Values.Num Then
                                    value = Double.PositiveInfinity
                                Else
                                    value = DBNull.Value
                                End If
                            ElseIf data.Columns(colCounter).DataType Is GetType(String) Then
                                value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, OfficeOpenXml.ExcelErrorValue).ToString
                            Else
                                value = DBNull.Value
                            End If
                        Case VariantType.Double
                            If data.Columns(colCounter).DataType Is GetType(DateTime) Then
                                'Handle as date value
                                Dim datevalue As DateTime
                                datevalue = NumberToDateTime(CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Double))
                                'watch out for the milliseconds: the returned value may differ in milliseconds, 20.000 sec might be returned as 19.999 sec! --> round it!
                                Dim RoundedSeconds As Double = System.Math.Round((datevalue.Second * 1000 + datevalue.Millisecond) / 1000)
                                datevalue = New DateTime(datevalue.Year, datevalue.Month, datevalue.Day, datevalue.Hour, datevalue.Minute, CType(RoundedSeconds, Integer))
                                value = datevalue
                            Else
                                'Handle as normal double
                                value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Double)
                            End If
                        Case VariantType.String
                            Dim cellValue As String
                            cellValue = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, String)
                            If cellValue <> "" AndAlso System.Environment.NewLine <> LineFeed Then
                                cellValue = cellValue.Replace(LineFeed, System.Environment.NewLine)
                            End If
                            value = cellValue
                        Case VariantType.Date
                            If sheet.Cells(rowCounter + 1, colCounter + 1).Value.GetType Is GetType(Double) Then
                                value = sheet.Cells(rowCounter + 1, colCounter + 1).GetValue(Of DateTime)
                            Else
                                value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, DateTime)
                            End If
                        Case VariantType.Decimal
                            value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Decimal)
                        Case VariantType.Char
                            value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Char)
                        Case VariantType.Byte
                            value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Byte)
                        Case VariantType.Currency
                            value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Decimal)
                        Case VariantType.Integer
                            value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Integer)
                        Case VariantType.Long
                            value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Long)
                        Case VariantType.Short
                            value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Short)
                        Case VariantType.Single
                            value = CType(sheet.Cells(rowCounter + 1, colCounter + 1).Value, Single)
                        Case Else
                            'Case VariantType.DataObject
                            'Case VariantType.Array
                            'Case VariantType.Null
                            'Case VariantType.UserDefinedType
                            'Case VariantType.Variant
                            If data.Columns(colCounter).DataType Is GetType(String) Then
                                value = New NotImplementedException("Error in sheet row " & (rowCounter + 1) & ", column " & (colCounter + 1) & ": Unknown cell type")
                            Else
                                value = DBNull.Value
                            End If
                    End Select
                    If value.GetType Is GetType(String) AndAlso CType(value, String) = "" AndAlso Not data.Columns(colCounter).DataType Is GetType(String) Then
                        'Handle situation that a cell might contain a "" instead of a blank value because of some user-defined Excel formulas which shall return "blank" cell content by using "" - irrespective to the regular column data type
                        'e.g. following formula: =IF($F6=I$1;1;"")
                        value = DBNull.Value
                    End If
                    row(colCounter) = value
                Next
                data.Rows.Add(row)
            Next

        End Sub

        Private Shared Function NumberToDateTime(value As Double) As DateTime
            Return DateTime.FromOADate(value)
        End Function

        Private Shared Function LookupDotNetType(xlsCell As OfficeOpenXml.ExcelRange) As VariantType
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
        ''' Detect custom date/time format strings which haven't been detected by Epplus.
        ''' </summary>
        ''' <param name="cellFormat"></param>
        ''' <returns></returns>
        Private Shared Function IsDateTimeFormat(cellFormat As String) As Boolean
            If cellFormat = "" Then
                Return False
            ElseIf cellFormat.StartsWith("yyyy-MM-dd") OrElse cellFormat.StartsWith("HH:mm:ss") Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Analyze the values in the complete sheet for their data type and create a data table with those corresponding column data types to hold all the data of the sheet.
        ''' </summary>
        ''' <param name="sheet">A sheet</param>
        ''' <param name="tableName">A table name for the new table</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <returns>A data table with the suggested structure to be able to hold all the data of the sheet</returns>
        ''' <remarks>
        ''' </remarks>
        Private Shared Function ReadDataTableFromXlsFileCreateDataTableSuggestion(ByVal sheet As OfficeOpenXml.ExcelWorksheet, ByVal tableName As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean) As DataTable
            'Create a datatable which can hold all the data available in that sheet (pay attention to the automatic column type detection)
            Dim Result As New DataTable(tableName)
            ReadDataTableFromXlsFileExtendDataTableColumns(Result, sheet, startReadingAtRowIndex, firstRowContainsColumnNames)
            Return Result
        End Function

        ''' <summary>
        ''' Analyze the values in the complete sheet for their data type and create a data table with those corresponding column data types to hold all the data of the sheet.
        ''' </summary>
        ''' <param name="inputTable">The target table</param>
        ''' <param name="sheet">A sheet</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Whether the first row contains column names instead of values.</param>
        ''' <remarks>
        ''' </remarks>
        Private Shared Sub ReadDataTableFromXlsFileExtendDataTableColumns(inputTable As System.Data.DataTable, ByVal sheet As OfficeOpenXml.ExcelWorksheet, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean)
            'Add required amount of columns
            'sheet.CalcDimensions() 'Calculate the sheet end positions (to prevent bug that this information is 0, e. g. after saving and reloading with this component)
            Dim LastSheetContentRowIndex As Integer = LookupLastContentRowIndex(sheet)
            Dim LastSheetContentColumnIndex As Integer = LookupLastContentColumnIndex(sheet)
            For colCounter As Integer = inputTable.Columns.Count To LastSheetContentColumnIndex
                'step through all rows and determine if there is a common data type, e. g. Date, String, Double
                Dim fieldType As System.Type = Nothing
                Dim firstContentRowIndex As Integer
                If firstRowContainsColumnNames Then
                    firstContentRowIndex = 1
                Else
                    firstContentRowIndex = 0
                End If
                firstContentRowIndex += startReadingAtRowIndex
                For RowCounter As Integer = firstContentRowIndex To LastSheetContentRowIndex
                    Select Case LookupDotNetType(sheet.Cells(RowCounter + 1, colCounter + 1))
                        Case VariantType.Empty
                            'no decision here
                        Case VariantType.Error
                            'value forces string-type and breaks for loop
                            Select Case CType(sheet.Cells(RowCounter + 1, colCounter + 1).Value, OfficeOpenXml.ExcelErrorValue).ToString
                                Case OfficeOpenXml.ExcelErrorValue.Values.Div0, OfficeOpenXml.ExcelErrorValue.Values.Num
                                    fieldType = GetType(Double)
                                Case Else
                                    fieldType = Nothing
                                    Exit For
                            End Select
                        Case VariantType.Boolean
                            If fieldType Is Nothing Then
                                fieldType = GetType(Boolean)
                            ElseIf fieldType Is GetType(Boolean) Then
                                'keep it
                            Else
                                'another value forces string-type and breaks for loop
                                fieldType = Nothing
                                Exit For
                            End If
                        Case VariantType.Double
                            If fieldType Is Nothing Then
                                fieldType = GetType(Double)
                            ElseIf fieldType Is GetType(Double) Then
                                'keep it
                            Else
                                'another value forces string-type and breaks for loop
                                fieldType = Nothing
                                Exit For
                            End If
                        Case VariantType.String
                            If String.IsNullOrEmpty(sheet.Cells(RowCounter + 1, colCounter + 1).Value.ToString) Then
                                'keep it
                            ElseIf fieldType Is Nothing Then
                                fieldType = GetType(String)
                            ElseIf fieldType Is GetType(String) Then
                                'keep it
                            Else
                                'another value forces string-type and breaks for loop
                                fieldType = Nothing
                                Exit For
                            End If
                        Case VariantType.Date
                            If fieldType Is Nothing Then
                                fieldType = GetType(DateTime)
                            ElseIf fieldType Is GetType(DateTime) Then
                                'keep it
                            Else
                                'another value forces string-type and breaks for loop
                                fieldType = Nothing
                                Exit For
                            End If
                    End Select
                Next
                If fieldType Is Nothing Then
                    fieldType = GetType(String)
                End If
                'Add the column
                Dim newCol As DataColumn
                If firstRowContainsColumnNames Then
                    'hint: also detect e.g. column header with date formats, e.g. "May 2005"
                    Dim ColName As String = CellValueAsString(sheet.Cells(startReadingAtRowIndex + 1, colCounter + 1))
                    ColName = Utils.LookupUniqueColumnName(inputTable, ColName)
                    newCol = New DataColumn(ColName, fieldType) 'column gets column name of 1st row
                Else
                    newCol = New DataColumn(Nothing, fieldType)
                End If
                inputTable.Columns.Add(newCol)
            Next
        End Sub

        ''' <summary>
        ''' Try to lookup the cell's value to a string anyhow.
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function CellValueAsString(ByVal cell As OfficeOpenXml.ExcelRange) As String
            Try
                Return cell.Text
            Catch ex As Exception
                Return "#ERROR: " & ex.Message
            End Try
        End Function
        ''' <summary>
        ''' Looks up the (zero-based) index number of a worksheet.
        ''' </summary>
        ''' <param name="workbook">The excel workbook</param>
        ''' <param name="worksheetName">A worksheet name</param>
        ''' <returns>-1 if the sheet name doesn't exist, otherwise its index value</returns>
        ''' <remarks>
        ''' </remarks>
        Private Shared Function ResolveWorksheetIndex(ByVal workbook As OfficeOpenXml.ExcelPackage, ByVal worksheetName As String) As Integer
            Dim sheetIndex As Integer = -1
            For MyCounter As Integer = 0 To workbook.Workbook.Worksheets.Count - 1
                Dim sheet As OfficeOpenXml.ExcelWorksheet = workbook.Workbook.Worksheets(MyCounter)
                If String.Equals(sheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase) Then
                    sheetIndex = MyCounter
                End If
            Next
            Return sheetIndex
        End Function

        ''' <summary>
        ''' Looks for a sheet with the specified name.
        ''' </summary>
        ''' <param name="workbook">The excel workbook</param>
        ''' <param name="sheetName">A sheet name</param>
        ''' <returns>An excel sheet</returns>
        ''' <remarks>
        ''' </remarks>
        Private Shared Function LookupWorksheet(ByVal workbook As OfficeOpenXml.ExcelPackage, ByVal sheetName As String) As OfficeOpenXml.ExcelWorksheet
            Dim resolvedIndex As Integer = ResolveWorksheetIndex(workbook, sheetName)
            If resolvedIndex = -1 Then
                Throw New ArgumentException("Worksheet """ & sheetName & """ hasn't been found")
            Else
                Return workbook.Workbook.Worksheets(resolvedIndex)
            End If
        End Function
#End Region

        ''' <summary>
        ''' Determines whether the value is a DateTime value and not a regular number.
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns>True for DateTime, False for Number(Double)</returns>
        ''' <remarks></remarks>
        Private Shared Function IsDateTimeInsteadOfNumber(ByVal cell As OfficeOpenXml.ExcelRange) As Boolean
            Dim numFormat As String = cell.Style.Numberformat.Format.ToLowerInvariant
            If numFormat.Contains("y"c) OrElse numFormat.Contains("m"c) OrElse numFormat.Contains("d"c) OrElse numFormat.Contains("h"c) Then
                Try
                    DateTime.FromOADate(CType(cell.Value, Double))
                    Return True
                Catch
                    Return False
                End Try
            Else
                Return False
            End If

        End Function

    End Class

End Namespace
