Option Explicit On
Option Strict On

'NOTE:    THIS FILE IS UPDATED IN DIRECTORY CM.Data.EpplusFixCalcsEdition FIRST AND COPIED TO CM.Data.EpplusPolyformEdition AFTERWARDS
'SEE:     clone-build-files.cmd/.sh/.ps1
'WARNING: PLEASE CHANGE THIS FILE ONLY AT REQUIRED LOCATION, OR CHANGES WILL BE LOST!

Imports System.Data
Imports System.Linq

Namespace CompuMaster.Data

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Provides simplified write access to XLS files
    ''' </summary>
    ''' <remarks>
    '''     Please pay attention to following circumstances
    '''     - Written null values (Nothing in VisualBasic) will be re-read as DBNull.Value
    '''     - Written zero-DateTime value will be re-read as DBNull.Value
    '''     - Excel supports DateTime values starting from 01.01.1900, only. Lower date values will throw an exception when assigning.
    '''     - Excel DateTime values are limited to year, month, day, hour, minute, second. Milliseconds and ticks will be dropped.
    '''     - Lines with only DBNull.Value or null (Nothing in VisualBasic) will be considered as not-existing if they are the last lines
    ''' </remarks>
    ''' <history>
    ''' 	[adminwezel]	30.05.2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class XlsEpplusPolyformEdition

        Private Shared _ErrorLevel As Byte = 0
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Error level 0 doesn't throw exception when writing e.g. invalid date/time values (invalid for excel); Error level 1 throws them
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	31.05.2010	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Property ErrorLevel() As Byte
            Get
                Return _ErrorLevel
            End Get
            Set(ByVal Value As Byte)
                _ErrorLevel = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Create an new excel file with some data
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataSet">A dataset to write into the workbook</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	05.07.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub WriteDataSetToXlsFile(ByVal outputPath As String, ByVal dataSet As System.Data.DataSet)
            WriteDataSetToXlsFile(Nothing, outputPath, dataSet)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Load an excel file, put some data into it and save the file again
        ''' </summary>
        ''' <param name="inputPath">A file which shall be loaded</param>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataSet">A dataset to write into the workbook</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	05.07.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub WriteDataSetToXlsFile(ByVal inputPath As String, ByVal outputPath As String, ByVal dataSet As System.Data.DataSet)
            Dim tables As New ArrayList
            Dim tableNames As New ArrayList
            If Not dataSet Is Nothing AndAlso dataSet.Tables.Count > 0 Then
                For MyCounter As Integer = 0 To dataSet.Tables.Count - 1
                    tables.Add(dataSet.Tables(MyCounter))
                    tableNames.Add(dataSet.Tables(MyCounter).TableName)
                Next
            End If
            WriteDataTableToXlsFile(inputPath, outputPath, CType(tables.ToArray(GetType(DataTable)), DataTable()), CType(tableNames.ToArray(GetType(String)), String()))
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Create an new excel file with some data
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <remarks>
        ''' The data will be written to the sheet with the name as the datatable's name
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	05.07.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub WriteDataTableToXlsFile(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable)
            WriteDataTableToXlsFile(Nothing, outputPath, dataTable, CType(Nothing, String))
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Create an new excel file with some data
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <remarks>
        ''' The data will be written to the sheet with the name as the datatable's name
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	05.07.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub WriteDataTableToXlsFileAndFirstSheet(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable)
            If outputPath = Nothing OrElse (New System.IO.FileInfo(outputPath)).FullName = Nothing Then
                Throw New ArgumentNullException("outputPath", "The output filename is required")
            End If

            Dim exportWorkbook As OfficeOpenXml.ExcelPackage
            exportWorkbook = OpenAndWriteDataTableToXlsFile(Nothing, New DataTable() {dataTable}, New String() {}, SpecialSheet.FirstSheet)
            If exportWorkbook Is Nothing Then
                Return
            End If
            SaveWorkbook(exportWorkbook, outputPath)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Create an new excel file with some data
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <remarks>
        ''' The data will be written to the sheet with the name as the datatable's name
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	05.07.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub WriteDataTableToXlsFileAndCurrentSheet(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable)
            If outputPath = Nothing OrElse (New System.IO.FileInfo(outputPath)).FullName = Nothing Then
                Throw New ArgumentNullException("outputPath", "The output filename is required")
            End If

            Dim exportWorkbook As OfficeOpenXml.ExcelPackage
            exportWorkbook = OpenAndWriteDataTableToXlsFile(Nothing, New DataTable() {dataTable}, New String() {}, SpecialSheet.CurrentSheet)
            If exportWorkbook Is Nothing Then
                Return
            End If
            SaveWorkbook(exportWorkbook, outputPath)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Create an new excel file with some data
        ''' </summary>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <param name="sheetName">The name the sheet which shall be updated/added</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	05.07.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub WriteDataTableToXlsFile(ByVal outputPath As String, ByVal dataTable As System.Data.DataTable, ByVal sheetName As String)
            WriteDataTableToXlsFile(Nothing, outputPath, dataTable, sheetName)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Load an excel file, put some data into it and save the file again
        ''' </summary>
        ''' <param name="inputPath">A file which shall be loaded</param>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTable">A datatable to write into one of the sheets</param>
        ''' <param name="sheetName">The name the sheet which shall be updated/added</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	05.07.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub WriteDataTableToXlsFile(ByVal inputPath As String, ByVal outputPath As String, ByVal dataTable As System.Data.DataTable, ByVal sheetName As String)
            WriteDataTableToXlsFile(inputPath, outputPath, New DataTable() {dataTable}, New String() {sheetName})
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Update/create an excel file, put some data into it and save the file again
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template</param>
        ''' <param name="outputPath">The output file</param>
        ''' <param name="dataTables">Some datatables to write into the workbook</param>
        ''' <param name="sheetNames">The name the sheets which shall be updated/added in the order as defined by parameter dataTables</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	05.07.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub WriteDataTableToXlsFile(ByVal inputPath As String, ByVal outputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String())
            If outputPath = Nothing OrElse (New System.IO.FileInfo(outputPath)).FullName = Nothing Then
                Throw New ArgumentNullException("outputPath", "The output filename is required")
            End If

            Dim exportWorkbook As OfficeOpenXml.ExcelPackage
            exportWorkbook = OpenAndWriteDataTableToXlsFile(inputPath, dataTables, sheetNames, SpecialSheet.AsDefinedInSheetNamesCollection)
            If exportWorkbook Is Nothing Then
                Return
            End If
            SaveWorkbook(exportWorkbook, outputPath)
        End Sub

        ''' <summary>
        ''' Save the changed worksheet
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
        ''' The sheet which is subject of operations
        ''' </summary>
        ''' <remarks></remarks>
        Private Enum SpecialSheet As Byte
            AsDefinedInSheetNamesCollection = 0
            FirstSheet = 1
            CurrentSheet = 2
        End Enum

        ''' <summary>
        ''' Update/create an excel file, put some data into it and save the file again
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template</param>
        ''' <param name="dataTables">Some datatables to write into the workbook</param>
        ''' <param name="sheetNames">The name the sheets which shall be updated/added in the order as defined by parameter dataTables</param>
        ''' <param name="specialSheet">A special sheet</param>
        ''' <returns>A Workbook object</returns>
        ''' <remarks></remarks>
        Private Shared Function OpenAndWriteDataTableToXlsFile(ByVal inputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetnames As String(), ByVal specialSheet As SpecialSheet) As OfficeOpenXml.ExcelPackage

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
                        If IsDBNull(value) Then
                            WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = Nothing
                        ElseIf value.GetType Is GetType(String) Then
                            'Excel requires line-breaks to be an LF character only, not a windows typical CR+LF
                            Dim cell As OfficeOpenXml.ExcelRange = WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1)
                            If CType(value, String) <> "" Then
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
                                If ErrorLevel = 0 Then
                                    WorkSheet.Cells(RowCounter + 1 + 1, ColCounter + 1).Value = Double.NaN
                                Else
                                    Throw New Exception("Error writing a date/time value """ & datevalue.ToString & """ in row " & (RowCounter + 1), ex)
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
        ''' Excel file formats
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum FileFormat As Byte
            Excel2007 = 1
            Excel2007Macro = 2
        End Enum

        ''' <summary>
        ''' Update/create an excel file, put some data into it and save the file to the output stream
        ''' </summary>
        ''' <param name="inputPath">An optional path to a template</param>
        ''' <param name="outputStream">An opened output stream</param>
        ''' <param name="dataTables">Some datatables to write into the workbook</param>
        ''' <param name="sheetNames">The name the sheets which shall be updated/added in the order as defined by parameter dataTables</param>
        ''' <remarks></remarks>
        Public Sub WriteDataTableToXlsStream(ByVal inputPath As String, ByVal outputStream As System.IO.Stream, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String(), ByVal fileFormat As FileFormat)
            Dim exportWorkbook As OfficeOpenXml.ExcelPackage
            exportWorkbook = OpenAndWriteDataTableToXlsFile(inputPath, dataTables, sheetNames, SpecialSheet.AsDefinedInSheetNamesCollection)
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

        '''' <summary>
        '''' Directly send the new workbook file to the browser
        '''' </summary>
        '''' <param name="inputPath">An optional path to a template</param>
        '''' <param name="dataTables">Some datatables to write into the workbook</param>
        '''' <param name="sheetNames">The name the sheets which shall be updated/added in the order as defined by parameter dataTables</param>
        '''' <param name="httpContext">The current HTTP context</param>
        '''' <remarks></remarks>
        'Public Sub WriteDataTableToXlsHttpResponse(ByVal inputPath As String, ByVal dataTables As System.Data.DataTable(), ByVal sheetNames As String(), ByVal httpContext As System.Web.HttpContext, ByVal fileFormat As FileFormat)
        '    If dataTables Is Nothing Then
        '        Throw New ArgumentNullException("dataTables")
        '    End If

        '    Dim exportWorkbook As EpplusFreeOfficeOpenXml.ExcelPackage
        '    exportWorkbook = OpenAndWriteDataTableToXlsFile(inputPath, dataTables, sheetNames, SpecialSheet.AsDefinedInSheetNamesCollection)
        '    If exportWorkbook Is Nothing Then
        '        Throw New Exception("Workbook creation failed - missing workbook")
        '    End If

        '    ' compatible with Excel 97/2000/XP/2003/2007.
        '    httpContext.Response.Clear()
        '    httpContext.Response.ContentType = "application/vnd.ms-excel"
        '    httpContext.Response.AddHeader("Content-Disposition", "attachment; filename=report.xls")
        '    If fileFormat = FileFormat.Excel2007 Then
        '        'Excel 2007 format
        '        exportWorkbook.SaveAs(httpContext.Response.OutputStream)
        '    ElseIf fileFormat = FileFormat.Excel2007Macro Then
        '        'Excel 2007 format
        '        exportWorkbook.SaveAs(httpContext.Response.OutputStream)
        '    Else
        '        Throw New NotSupportedException("value for fileformat is invalid")
        '    End If
        'End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read all sheets from an excel sheet into a dataset
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ReadDataSetFromXlsFile(ByVal inputPath As String, ByVal firstRowContainsColumnNames As Boolean) As DataSet

            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException("inputPath", "The input filename is required")
            End If

            Dim importWorkbook As OfficeOpenXml.ExcelPackage

            'Load the worksheet
            importWorkbook = LoadWorkbookFile(inputPath)

            Dim Result As New DataSet

            For sheetCounter As Integer = 0 To importWorkbook.Workbook.Worksheets.Count - 1
                Dim Sheet As OfficeOpenXml.ExcelWorksheet = importWorkbook.Workbook.Worksheets(sheetCounter)

                'Detect the column types which must be used
                Dim sheetData As DataTable = ReadDataTableFromXlsFileCreateDataTableSuggestion(Sheet, Sheet.Name, 0, firstRowContainsColumnNames)

                'Read all data and put it into the datatable
                ReadDataTableFromXlsFile(Sheet, 0, firstRowContainsColumnNames, sheetData)

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

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read the data from an excel sheet into a datatable
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <returns></returns>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String) As DataTable
            Return ReadDataTableFromXlsFile(inputPath, True)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read the data from an excel sheet into a datatable
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
        ''' <returns></returns>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal firstRowContainsColumnNames As Boolean) As DataTable
            Return ReadDataTableFromXlsFile(inputPath, 0, firstRowContainsColumnNames)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read the data from an excel sheet into a datatable
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
        ''' <returns></returns>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean) As DataTable

            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException("inputPath", "The input filename is required")
            End If

            Dim importWorkbook As OfficeOpenXml.ExcelPackage

            'Save the changed worksheet
            importWorkbook = LoadWorkbookFile(inputPath)
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = importWorkbook.Workbook.Worksheets(1)

            'Detect the column types which must be used
            Dim Result As DataTable = ReadDataTableFromXlsFileCreateDataTableSuggestion(Sheet, Sheet.Name, startReadingAtRowIndex, firstRowContainsColumnNames)

            'Read all data and put it into the datatable
            ReadDataTableFromXlsFile(Sheet, startReadingAtRowIndex, firstRowContainsColumnNames, Result)

            'Return the result
            Return Result

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read the data from an excel sheet into a datatable
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <returns></returns>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String) As DataTable
            Return ReadDataTableFromXlsFile(inputPath, sheetName, True)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read the data from an excel sheet into a datatable
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
        ''' <returns></returns>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String, ByVal firstRowContainsColumnNames As Boolean) As DataTable
            Return ReadDataTableFromXlsFile(inputPath, sheetName, 0, firstRowContainsColumnNames)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read the data from an excel sheet into a datatable
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
        ''' <returns></returns>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean) As DataTable
            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException("inputPath", "The input filename is required")
            End If

            Dim importWorkbook As OfficeOpenXml.ExcelPackage

            'Save the changed worksheet
            importWorkbook = LoadWorkbookFile(inputPath)
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

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read the data from an excel sheet into a datatable
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String, ByVal firstRowContainsColumnNames As Boolean, ByVal data As DataTable)
            ReadDataTableFromXlsFile(inputPath, sheetName, 0, firstRowContainsColumnNames, data)
        End Sub

        ''' <summary>
        '''     Read the data from first sheet of an excel sheet into a datatable
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
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

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read the data from an excel sheet into a datatable
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <param name="sheetName">The sheet which contains the import data</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Sub ReadDataTableFromXlsFile(ByVal inputPath As String, ByVal sheetName As String, ByVal startReadingAtRowIndex As Integer, ByVal firstRowContainsColumnNames As Boolean, ByVal data As DataTable)

            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException("inputPath", "The input filename is required")
            ElseIf data Is Nothing Then
                Throw New ArgumentNullException("data", "A datatable must be predefined which shall hold all the data")
            End If

            Dim importWorkbook As OfficeOpenXml.ExcelPackage

            'Save the changed worksheet
            importWorkbook = LoadWorkbookFile(inputPath)
            If sheetName = Nothing Then
                sheetName = importWorkbook.Workbook.Worksheets.First.Name
            End If
            Dim Sheet As OfficeOpenXml.ExcelWorksheet = LookupWorksheet(importWorkbook, sheetName)

            'Extend table's column set as long as columns count matches
            ReadDataTableFromXlsFileExtendDataTableColumns(data, Sheet, 0, firstRowContainsColumnNames)

            'Read all data and put it into the datatable
            ReadDataTableFromXlsFile(Sheet, startReadingAtRowIndex, firstRowContainsColumnNames, data)

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Read the available sheet names from an XLS file
        ''' </summary>
        ''' <param name="inputPath">The filename of the excel document</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	21.04.2010	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ReadSheetNamesFromXlsFile(ByVal inputPath As String) As String()

            If inputPath = Nothing OrElse (New System.IO.FileInfo(inputPath)).FullName = Nothing Then
                Throw New ArgumentNullException("inputPath", "The input filename is required")
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
        ''' Lookup the last content column index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
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
        ''' Lookup the last content row index (zero based index) (the last content cell might differ from Excel's special cell xlLastCell)
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
        ''' Determine if a cell contains empty content
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

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Read the data from an excel sheet into a datatable
        ''' </summary>
        ''' <param name="sheet">An excel sheet containing the required data</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
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
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
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
                            If cellValue <> "" AndAlso System.Environment.NewLine <> ControlChars.Lf Then
                                cellValue = Replace(cellValue, ControlChars.Lf, System.Environment.NewLine, , , CompareMethod.Binary)
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
        ''' Detect custom date/time format strings which haven't been detected by Epplus
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
        '''     Analyze the values in the complete sheet for their data type and create a data table with those corresponding column data types to hold all the data of the sheet
        ''' </summary>
        ''' <param name="sheet">A sheet</param>
        ''' <param name="tableName">A table name for the new table</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
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
        '''     Analyze the values in the complete sheet for their data type and create a data table with those corresponding column data types to hold all the data of the sheet
        ''' </summary>
        ''' <param name="inputTable">The target table</param>
        ''' <param name="sheet">A sheet</param>
        ''' <param name="startReadingAtRowIndex">Sometimes, excel sheets start with an introductional/explaining header instead of just column names, e.g. a table may start at row index 2 (in excel line 3)</param>
        ''' <param name="firstRowContainsColumnNames">Indicate wether the first row contains column names (true) or values (false)</param>
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
        ''' Try to lookup the cell's value to a string anyhow
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
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Lookup the (zero-based) index number of a work sheet
        ''' </summary>
        ''' <param name="workbook">The excel workbook</param>
        ''' <param name="worksheetName">A work sheet name</param>
        ''' <returns>-1 if the sheet name doesn't exist, otherwise its index value</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[AdminSupport]	29.09.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Shared Function ResolveWorksheetIndex(ByVal workbook As OfficeOpenXml.ExcelPackage, ByVal worksheetName As String) As Integer
            Dim sheetIndex As Integer = -1
            For MyCounter As Integer = 0 To workbook.Workbook.Worksheets.Count - 1
                Dim sheet As OfficeOpenXml.ExcelWorksheet = workbook.Workbook.Worksheets(MyCounter)
                If sheet.Name.ToLower = worksheetName.ToLower Then
                    sheetIndex = MyCounter
                End If
            Next
            Return sheetIndex
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Lookup for a sheet with the specified name
        ''' </summary>
        ''' <param name="workbook">The excel workbook</param>
        ''' <param name="sheetName">A sheet name</param>
        ''' <returns>An excel sheet</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	02.02.2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Shared Function LookupWorksheet(ByVal workbook As OfficeOpenXml.ExcelPackage, ByVal sheetName As String) As OfficeOpenXml.ExcelWorksheet
            Dim resolvedIndex As Integer = ResolveWorksheetIndex(workbook, sheetName)
            If resolvedIndex = -1 Then
                Throw New Exception("Worksheet """ & sheetName & """ hasn't been found")
            Else
                Return workbook.Workbook.Worksheets(resolvedIndex)
            End If
        End Function
#End Region

        ''' <summary>
        ''' Lookup if the value is a DateTime value and not a normal number
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns>True for DateTime, False for Number(Double)</returns>
        ''' <remarks></remarks>
        Private Shared Function IsDateTimeInsteadOfNumber(ByVal cell As OfficeOpenXml.ExcelRange) As Boolean
            Dim numFormat As String = cell.Style.Numberformat.Format
            If numFormat.ToLower.IndexOf("y") > 0 OrElse numFormat.ToLower.IndexOf("m") > 0 OrElse numFormat.ToLower.IndexOf("d") > 0 OrElse numFormat.ToLower.IndexOf("h") > 0 Then
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
