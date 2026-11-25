Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    ''' <summary>Edits XLSX workbooks using the Open XML SDK.</summary>
    ''' <remarks>The engine edits a private package snapshot. It preserves untouched package parts and supports structural reference updates on compatible workbooks. Formula evaluation is unavailable.</remarks>
    Public NotInheritable Partial Class OpenXmlExcelDataOperations
        Inherits ExcelDataOperationsBase
        Implements IDisposable

        Private _source As MemoryStream
        Private _document As SpreadsheetDocument
        Private ReadOnly _changedSheets As New HashSet(Of WorksheetPart)
        Private _workbookChanged As Boolean
        Private _calculationChainInvalid As Boolean
        Private _recalculateOnOpen As Boolean
        Private _initialAutomaticCalculation As Boolean?
        Private _automaticCalculation As Boolean?
        Private _openedFromFile As Boolean

        ''' <summary>Opens or creates an XLSX workbook.</summary>
        ''' <param name="file">Source path or intended path for a new workbook.</param>
        ''' <param name="mode">Mode used to open or create the workbook.</param>
        ''' <param name="options">File and engine options.</param>
        Public Sub New(file As String, mode As OpenMode, options As ExcelDataOperationsOptions)
            Me.New(options)
            ExecuteOpenModeActions(If(String.IsNullOrEmpty(file), Nothing, Path.GetFullPath(file)), mode, LoadOptions)
        End Sub

        ''' <summary>Opens an XLSX workbook from bytes.</summary>
        ''' <param name="data">XLSX package bytes to copy.</param>
        ''' <param name="options">File and engine options.</param>
        Public Sub New(data As Byte(), options As ExcelDataOperationsOptions)
            Me.New(options)
            LoadAndInitializeWorkbookFile(data, LoadOptions)
        End Sub

        ''' <summary>Opens an XLSX workbook from a stream.</summary>
        ''' <param name="data">Stream positioned at the start of the package. The caller retains ownership.</param>
        ''' <param name="options">File and engine options.</param>
        Public Sub New(data As Stream, options As ExcelDataOperationsOptions)
            Me.New(options)
            LoadAndInitializeWorkbookFile(data, LoadOptions)
        End Sub

        Private Sub New(options As ExcelDataOperationsOptions)
            MyBase.New(options)
            If options.DisableAutoCalculationInWorkbook.HasValue Then
                _initialAutomaticCalculation = Not options.DisableAutoCalculationInWorkbook.Value
                ' Apply the option to the package after it has been opened.
                LoadOptions.DisableAutoCalculationInWorkbook = Nothing
            End If
            AutoCalculationResetToEnabledForAllSavedWorkbooks = False
        End Sub

        ''' <inheritdoc/>
        Public Overrides ReadOnly Property EngineName As String
            Get
                Return "Open XML SDK"
            End Get
        End Property

        ''' <inheritdoc/>
        Protected Overrides ReadOnly Property DefaultCalculationOptions As ExcelEngineDefaultOptions
            Get
                Return New ExcelEngineDefaultOptions(disableInitialCalculation:=True, disableCalculationEngine:=True)
            End Get
        End Property

        ''' <inheritdoc/>
        Protected Overrides ReadOnly Property AutomaticallyUpdatesFormulasAndReferencesForStructuralChanges As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <inheritdoc/>
        Protected Overrides Sub ValidateLoadOptions(options As ExcelDataOperationsOptions)
            If Not options.DisableCalculationEngine.Value OrElse Not options.DisableInitialCalculation.Value Then
                Throw CalculationFeatureDisabled()
            End If
            If Not String.IsNullOrEmpty(options.PasswordForOpening) Then
                Throw New NotSupportedException("The Open XML SDK engine does not open password-encrypted workbooks.")
            End If
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub LoadWorkbook(file As FileInfo)
            If String.Equals(file.Extension, ".xls", StringComparison.OrdinalIgnoreCase) Then Throw New BinaryXlsFileNotSupportedException(file)
            If String.Equals(file.Extension, ".xlsb", StringComparison.OrdinalIgnoreCase) Then Throw New BinaryXlsbFileNotSupportedException(file)
            Using input = file.OpenRead()
                LoadWorkbook(input)
            End Using
            _openedFromFile = True
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub LoadWorkbook(data As Byte())
            If data Is Nothing Then Throw New ArgumentNullException(NameOf(data))
            Using input As New MemoryStream(data, writable:=False)
                LoadWorkbook(input)
            End Using
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub LoadWorkbook(data As Stream)
            If data Is Nothing Then Throw New ArgumentNullException(NameOf(data))
            ValidateLoadOptions(LoadOptions)
            Dim snapshot As New MemoryStream()
            Try
                data.CopyTo(snapshot)
                OpenSnapshot(snapshot)
            Catch
                snapshot.Dispose()
                Throw
            End Try
        End Sub

        Private Sub OpenSnapshot(snapshot As MemoryStream)
            snapshot.Position = 0
            Dim document As SpreadsheetDocument = Nothing
            Try
                document = SpreadsheetDocument.Open(snapshot, False, New OpenSettings With {.AutoSave = False})
                If document.DocumentType <> SpreadsheetDocumentType.Workbook Then
                    Throw New NotSupportedException("The Open XML SDK engine supports XLSX workbooks.")
                End If
                If document.WorkbookPart?.Workbook?.Sheets Is Nothing Then
                    Throw New InvalidDataException("The package does not contain a workbook with sheets.")
                End If
                _source = snapshot
                _document = document
                _automaticCalculation = _initialAutomaticCalculation
                RecalculationRequired = False
            Catch ex As Exception
                document?.Dispose()
                If TypeOf ex Is OpenXmlPackageException OrElse TypeOf ex Is FileFormatException OrElse
                   TypeOf ex Is InvalidDataException OrElse TypeOf ex Is XmlException Then
                    Throw New FileCorruptedOrInvalidFileFormatException(FilePath, ex)
                End If
                Throw
            End Try
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub CreateWorkbook()
            Dim snapshot As New MemoryStream()
            Try
                Using document = SpreadsheetDocument.Create(snapshot, SpreadsheetDocumentType.Workbook)
                    Dim workbookPart = document.AddWorkbookPart()
                    Dim worksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
                    worksheetPart.Worksheet = New Worksheet(New SheetData())
                    workbookPart.Workbook = New Workbook(New Sheets(New Sheet With {
                        .Name = "Sheet1", .SheetId = 1UI, .Id = workbookPart.GetIdOfPart(worksheetPart)
                    }))
                End Using
                OpenSnapshot(snapshot)
            Catch
                snapshot.Dispose()
                Throw
            End Try
        End Sub

        Private ReadOnly Property WorkbookPart As WorkbookPart
            Get
                If IsClosed Then Throw New InvalidOperationException("No XLSX workbook is open.")
                Return _document.WorkbookPart
            End Get
        End Property

        ''' <inheritdoc/>
        Protected Overrides ReadOnly Property WorkbookFilePath As String
            Get
                ' The engine stores the package in memory, so the base class owns file identity.
                Return If(_openedFromFile OrElse (Not IsClosed AndAlso Not Me.ReadOnly AndAlso File.Exists(FilePath)), FilePath, Nothing)
            End Get
        End Property

        ''' <inheritdoc/>
        Public Overrides ReadOnly Property IsClosed As Boolean
            Get
                Return _document Is Nothing
            End Get
        End Property

        ''' <inheritdoc/>
        Public Overrides Sub Close()
            Try
                _document?.Dispose()
            Finally
                _document = Nothing
                _source?.Dispose()
                _source = Nothing
                _changedSheets.Clear()
                _changedTables.Clear()
                _formulaRanges.Clear()
                _workbookChanged = False
                _calculationChainInvalid = False
                _recalculateOnOpen = False
                _automaticCalculation = Nothing
                _openedFromFile = False
            End Try
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub CloseExcelAppInstance()
            ' There is no external application to close.
        End Sub

        ''' <inheritdoc/>
        Public Sub Dispose() Implements IDisposable.Dispose
            Close()
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Controls the saved workbook's calculation mode. The engine itself cannot evaluate formulas.</remarks>
        Public Overrides Property AutoCalculationEnabledWorkbookSetting As Boolean
            Get
                Return If(_automaticCalculation, Not IsClosed AndAlso If(WorkbookPart.Workbook.CalculationProperties?.CalculationMode?.Value, CalculateModeValues.Auto) <> CalculateModeValues.Manual)
            End Get
            Set(value As Boolean)
                ' The base constructor calls this setter before a package is open.
                If IsClosed Then Return
                Dim originalMode = If(WorkbookPart.Workbook.CalculationProperties?.CalculationMode?.Value, CalculateModeValues.Auto)
                _automaticCalculation = If(value = (originalMode <> CalculateModeValues.Manual), CType(Nothing, Boolean?), value)
            End Set
        End Property

        ''' <summary>Requests a full recalculation when the saved workbook is opened in Excel.</summary>
        ''' <remarks>The request sets automatic workbook calculation and the full-calculation-on-load flag in subsequent saves. Use a save cache option to clear existing cached formula results.</remarks>
        Public Sub RequestRecalculationOnOpen()
            If IsClosed Then Throw New InvalidOperationException("No XLSX workbook is open.")
            _recalculateOnOpen = True
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub SaveInternal(cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            SaveAsInternal(FilePath, cachedCalculationsOption)
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub SaveAsInternal(fileName As String, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            Dim sourceWorkbook = WorkbookPart
            If Not String.Equals(Path.GetExtension(fileName), ".xlsx", StringComparison.OrdinalIgnoreCase) Then
                Throw New NotSupportedException("The Open XML SDK engine saves XLSX files.")
            End If
            Dim destination = Path.GetFullPath(fileName)
            Dim pathComparison = If(Path.DirectorySeparatorChar = "\"c, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
            If Me.ReadOnly AndAlso _openedFromFile AndAlso String.Equals(destination, FilePath, pathComparison) Then
                Throw New FileReadOnlyException(destination)
            End If
            Dim clearCache = cachedCalculationsOption = SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation OrElse
                (cachedCalculationsOption = SaveOptionsForDisabledCalculationEngines.ResetCalculatedValuesForForcedCellRecalculationIfRecalculationRequired AndAlso RecalculationRequired)

            Dim temporaryPath = Path.Combine(Path.GetDirectoryName(destination), Path.GetRandomFileName())
            Dim temporaryFile As New FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)
            Try
                Using output = temporaryFile
                    _source.WriteTo(output)
                    output.Position = 0
                    If _workbookChanged OrElse _changedSheets.Count > 0 OrElse _automaticCalculation.HasValue OrElse _recalculateOnOpen OrElse clearCache Then
                        Using target = SpreadsheetDocument.Open(output, True, New OpenSettings With {.AutoSave = False})
                            SaveChangedParts(sourceWorkbook, target.WorkbookPart, clearCache)
                        End Using
                    End If
                End Using
                If File.Exists(destination) Then
                    File.Replace(temporaryPath, destination, Nothing)
                Else
                    File.Move(temporaryPath, destination)
                End If
            Finally
                File.Delete(temporaryPath)
            End Try
        End Sub

        Private Sub SaveChangedParts(source As WorkbookPart, target As WorkbookPart, clearCache As Boolean)
            For Each sourceSheet In _changedSheets
                Dim targetSheet = DirectCast(target.GetPartById(source.GetIdOfPart(sourceSheet)), WorksheetPart)
                Using output = targetSheet.GetStream(FileMode.Create, FileAccess.Write)
                    sourceSheet.Worksheet.Save(output)
                End Using
            Next
            SaveChangedTables(source, target)
            If _workbookChanged Then
                Using output = target.GetStream(FileMode.Create, FileAccess.Write)
                    source.Workbook.Save(output)
                End Using
            End If
            If clearCache Then
                For Each sheetPart In target.WorksheetParts
                    If ClearFormulaResults(GetSheetData(sheetPart)) Then sheetPart.Worksheet.Save()
                Next
            End If
            If (_calculationChainInvalid OrElse clearCache) AndAlso target.CalculationChainPart IsNot Nothing Then
                target.DeletePart(target.CalculationChainPart)
            End If
            If _automaticCalculation.HasValue OrElse _recalculateOnOpen OrElse clearCache Then
                Dim workbook = target.Workbook
                Dim calculation = workbook.CalculationProperties
                If calculation Is Nothing Then
                    calculation = New CalculationProperties()
                    workbook.CalculationProperties = calculation
                End If
                If _recalculateOnOpen OrElse clearCache Then
                    calculation.FullCalculationOnLoad = True
                End If
                If _recalculateOnOpen Then
                    calculation.CalculationMode = CalculateModeValues.Auto
                ElseIf _automaticCalculation.HasValue Then
                    calculation.CalculationMode = If(_automaticCalculation.Value, CalculateModeValues.Auto, CalculateModeValues.Manual)
                End If
                workbook.Save()
            End If
        End Sub

        Private Const CalculationUnavailableMessage As String =
            "The Open XML SDK engine cannot evaluate formulas. Use RequestRecalculationOnOpen to have Excel recalculate the saved workbook."

        Private Shared Function CalculationFeatureDisabled() As FeatureDisabledException
            Return New FeatureDisabledException("Calculation engine", CalculationUnavailableMessage)
        End Function

        ''' <inheritdoc/>
        Protected Overrides Sub CalculateMissingFormulaResult(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            Throw New NotSupportedException(CalculationUnavailableMessage)
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub RecalculateAllInternal()
            Throw CalculationFeatureDisabled()
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub RecalculateSheet(sheetName As String)
            Throw CalculationFeatureDisabled()
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub RecalculateCell(sheetName As String, rowIndex As Integer, columnIndex As Integer, throwExceptionOnCalculationError As Boolean)
            Throw CalculationFeatureDisabled()
        End Sub

        ''' <inheritdoc/>
        Public Overrides ReadOnly Property HasVbaProject As Boolean
            Get
                Return WorkbookPart.VbaProjectPart IsNot Nothing
            End Get
        End Property

        ''' <inheritdoc/>
        Public Overrides Sub RemoveVbaProject()
            If HasVbaProject Then Throw UnsupportedOperation("Removing VBA projects")
        End Sub

    End Class

End Namespace
