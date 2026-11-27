Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    ''' <summary>Edits Open XML workbooks and templates using the Open XML SDK.</summary>
    ''' <remarks>The engine edits a private package snapshot. It preserves untouched package parts and supports opt-in formula evaluation and structural reference updates on compatible workbooks.</remarks>
    Public NotInheritable Partial Class OpenXmlExcelDataOperations
        Inherits ExcelDataOperationsBase
        Implements IDisposable, ICloneable

        Private _source As MemoryStream
        Private _document As SpreadsheetDocument
        Private _currentDocumentType As SpreadsheetDocumentType
        Private ReadOnly _changedSheets As New HashSet(Of WorksheetPart)
        Private ReadOnly _changedChartSheets As New HashSet(Of ChartsheetPart)
        Private _workbookChanged As Boolean
        Private _calculationChainInvalid As Boolean
        Private _recalculateOnOpen As Boolean
        Private _initialAutomaticCalculation As Boolean?
        Private _automaticCalculation As Boolean?
        Private _openedFromFile As Boolean
        Private _vbaProjectRemoved As Boolean
        Private ReadOnly _formulaSession As ExcelFormulaSession
        Private ReadOnly _configuredFormulaName As String
        Private _formulaWorkbookName As String

        ''' <summary>Opens or creates an Open XML workbook or template.</summary>
        ''' <param name="file">Source path or intended path for a new workbook.</param>
        ''' <param name="mode">Mode used to open or create the workbook.</param>
        ''' <param name="options">File and engine options.</param>
        Public Sub New(file As String, mode As OpenMode, options As ExcelDataOperationsOptions)
            Me.New(options)
            Try
                ExecuteOpenModeActions(If(String.IsNullOrEmpty(file), Nothing, Path.GetFullPath(file)), mode, LoadOptions)
            Catch
                Close()
                Throw
            End Try
        End Sub

        ''' <summary>Opens an Open XML workbook or template from bytes.</summary>
        ''' <param name="data">Open XML package bytes to copy.</param>
        ''' <param name="options">File and engine options.</param>
        Public Sub New(data As Byte(), options As ExcelDataOperationsOptions)
            Me.New(options)
            Try
                LoadAndInitializeWorkbookFile(data, LoadOptions)
            Catch
                Close()
                Throw
            End Try
        End Sub

        ''' <summary>Opens an Open XML workbook or template from a stream.</summary>
        ''' <param name="data">Stream positioned at the start of the package. The caller retains ownership.</param>
        ''' <param name="options">File and engine options.</param>
        Public Sub New(data As Stream, options As ExcelDataOperationsOptions)
            Me.New(options)
            Try
                LoadAndInitializeWorkbookFile(data, LoadOptions)
            Catch
                Close()
                Throw
            End Try
        End Sub

        Private Sub New(options As ExcelDataOperationsOptions)
            MyBase.New(options)
            _formulaSession = LoadOptions.FormulaSession
            _configuredFormulaName = LoadOptions.FormulaWorkbookName
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
            If options.DisableCalculationEngine.Value AndAlso Not options.DisableInitialCalculation.Value Then
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
                Select Case document.DocumentType
                    Case SpreadsheetDocumentType.Workbook, SpreadsheetDocumentType.MacroEnabledWorkbook,
                         SpreadsheetDocumentType.Template, SpreadsheetDocumentType.MacroEnabledTemplate
                    Case Else
                        Throw New NotSupportedException("The Open XML SDK engine supports XLSX, XLSM, XLTX and XLTM files.")
                End Select
                If document.WorkbookPart?.Workbook?.Sheets Is Nothing Then
                    Throw New InvalidDataException("The package does not contain a workbook with sheets.")
                End If
                SyncLock WorkbookSaveLock
                    BeginWorkbookIdentity()
                    If _formulaSession IsNot Nothing Then
                        Dim name = If(_configuredFormulaName, Path.GetFileName(FilePath))
                        If String.IsNullOrWhiteSpace(name) Then Throw New ArgumentException("A workbook name is required to attach to a formula session.", NameOf(ExcelDataOperationsOptions.FormulaWorkbookName))
                        _formulaSession.AttachWorkbook(name, Me)
                        _formulaWorkbookName = name
                    End If
                    _source = snapshot
                    _document = document
                    _currentDocumentType = document.DocumentType
                    _automaticCalculation = _initialAutomaticCalculation
                    RecalculationRequired = False
                End SyncLock
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
                Using document = SpreadsheetDocument.Create(snapshot, If(String.IsNullOrEmpty(FilePath), SpreadsheetDocumentType.Workbook, GetDocumentType(FilePath)))
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
                If IsClosed Then Throw New InvalidOperationException("No workbook is open.")
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
            SyncLock WorkbookSaveLock
                Try
                    _document?.Dispose()
                Finally
                    If _formulaWorkbookName IsNot Nothing Then _formulaSession.DetachWorkbook(_formulaWorkbookName, Me)
                    _formulaWorkbookName = Nothing
                    _document = Nothing
                    CloseNativeRtd()
                    _source?.Dispose()
                    _source = Nothing
                    _changedSheets.Clear()
                    _changedChartSheets.Clear()
                    _changedTables.Clear()
                    _changedCharts.Clear()
                    _changedExtendedCharts.Clear()
                    _changedDrawings.Clear()
                    _formulaDependencies.Clear()
                    _pendingCellEdits.Clear()
                    _completeDependencyRevision = Nothing
                    _formulaRanges.Clear()
                    _authoredDynamicMetadata = Nothing
                    _workbookChanged = False
                    _calculationChainInvalid = False
                    _recalculateOnOpen = False
                    _automaticCalculation = Nothing
                    _openedFromFile = False
                    _vbaProjectRemoved = False
                End Try
            End SyncLock
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
        ''' <remarks>Controls the workbook calculation mode. Set CalculateOnCellEdit to opt in to in-process automatic calculation after value writes.</remarks>
        Public Overrides Property AutoCalculationEnabledWorkbookSetting As Boolean
            Get
                Return If(_automaticCalculation, Not IsClosed AndAlso If(WorkbookPart.Workbook.CalculationProperties?.CalculationMode?.Value, CalculateModeValues.Auto) <> CalculateModeValues.Manual)
            End Get
            Set(value As Boolean)
                SyncLock WorkbookSaveLock
                    ' The base constructor calls this setter before a package is open.
                    If IsClosed Then Return
                    If AutoCalculationEnabledWorkbookSetting <> value Then _completeDependencyRevision = Nothing
                    Dim originalMode = If(WorkbookPart.Workbook.CalculationProperties?.CalculationMode?.Value, CalculateModeValues.Auto)
                    _automaticCalculation = If(value = (originalMode <> CalculateModeValues.Manual), CType(Nothing, Boolean?), value)
                End SyncLock
            End Set
        End Property

        ''' <summary>Requests a full recalculation when the saved workbook is opened in Excel.</summary>
        ''' <remarks>The request sets automatic workbook calculation and the full-calculation-on-load flag in subsequent saves. Use a save cache option to clear existing cached formula results.</remarks>
        Public Sub RequestRecalculationOnOpen()
            SyncLock WorkbookSaveLock
                If IsClosed Then Throw New InvalidOperationException("No workbook is open.")
                _recalculateOnOpen = True
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub SaveInternal(cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            SaveAsInternal(FilePath, cachedCalculationsOption)
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub SaveAsInternal(fileName As String, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            Dim destination = ValidateSaveDestination(fileName)
            Dim documentType = GetDocumentType(destination)
            Dim clearCache = cachedCalculationsOption = SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation OrElse
                (cachedCalculationsOption = SaveOptionsForDisabledCalculationEngines.ResetCalculatedValuesForForcedCellRecalculationIfRecalculationRequired AndAlso RecalculationRequired)

            EnsureVisibleSheetForSave()
            Dim temporaryPath = Path.Combine(Path.GetDirectoryName(destination), Path.GetRandomFileName())
            Dim temporaryFile As New FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)
            Try
                Using output = temporaryFile
                    WritePackage(output, clearCache, documentType, _vbaProjectRemoved OrElse IsMacroFree(documentType))
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

        ''' <inheritdoc/>
        Protected Overrides Sub SaveCopyAsCore(filePath As String, recalculateWorkbook As Boolean?, cachedCalculationsOption As SaveOptionsForDisabledCalculationEngines)
            Dim destination = ValidateSaveDestination(filePath)
            EnsureVisibleSheetForSave()
            If Not recalculateWorkbook.GetValueOrDefault(RecalculationRequired AndAlso Not CalculationModuleDisabled) Then
                SaveInternal_ApplyCachedCalculationOption(cachedCalculationsOption)
                SaveAsInternal(destination, cachedCalculationsOption)
                Return
            End If

            Using copy = Clone()
                If _formulaSession Is Nothing Then
                    copy.RecalculateAll()
                Else
                    Dim session = _formulaSession.Snapshot()
                    Dim workbooks = session.Workbooks.ToDictionary(Function(item) item.Key,
                        Function(item) If(item.Value Is Me, copy, item.Value), StringComparer.OrdinalIgnoreCase)
                    Dim copySession As New ExcelFormulaSessionSnapshot(workbooks, session.Functions, session.RtdProviders, session.ActiveWorkbook,
                        session.ActiveCell, session.Revision)
                    Dim calculation As New FormulaCalculation(copy, copySession, Me.FilePath)
                    calculation.RecalculateWorkbook()
                    calculation.Commit()
                    copy.RecalculationRequired = False
                End If
                copy.SaveAs(destination, recalculateWorkbook:=False, cachedCalculationsOption)
            End Using
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub PrepareVbaForSave(destination As String)
            ' Apply format-specific VBA removal to the output package.
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub OnSavedAs(destination As String)
            _currentDocumentType = GetDocumentType(destination)
            If IsMacroFree(_currentDocumentType) AndAlso HasVbaProject Then _vbaProjectRemoved = True
            WorkbookSession?.BindTemporaryLink(TemporaryWorkbookName, destination)
        End Sub

        Private Function ValidateSaveDestination(fileName As String) As String
            Dim destination = Path.GetFullPath(fileName)
            Dim pathComparison = If(Path.DirectorySeparatorChar = ChrW(92), StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
            If Me.ReadOnly AndAlso String.Equals(destination, FilePath, pathComparison) Then
                Throw New FileReadOnlyException(destination)
            End If
            Return destination
        End Function

        Private Shared Function GetDocumentType(fileName As String) As SpreadsheetDocumentType
            Select Case Path.GetExtension(fileName).ToLowerInvariant()
                Case ".xlsx" : Return SpreadsheetDocumentType.Workbook
                Case ".xlsm" : Return SpreadsheetDocumentType.MacroEnabledWorkbook
                Case ".xltx" : Return SpreadsheetDocumentType.Template
                Case ".xltm" : Return SpreadsheetDocumentType.MacroEnabledTemplate
                Case Else : Throw New NotSupportedException("The Open XML SDK engine saves XLSX, XLSM, XLTX and XLTM files.")
            End Select
        End Function

        Private Shared Function IsMacroFree(documentType As SpreadsheetDocumentType) As Boolean
            Return documentType = SpreadsheetDocumentType.Workbook OrElse documentType = SpreadsheetDocumentType.Template
        End Function

        ''' <summary>Creates an independent copy of the current workbook.</summary>
        ''' <returns>A writable workbook with no associated file path.</returns>
        ''' <remarks>Includes unsaved changes, VBA and cached formula results. Use SaveAs to give the clone a file path.</remarks>
        ''' <exception cref="InvalidOperationException">The workbook is closed.</exception>
        Public Function Clone() As OpenXmlExcelDataOperations
            SyncLock WorkbookSaveLock
                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite,
                    Nothing, disableInitialCalculation:=True, disableAutoCalculation:=Nothing, disableCalculationEngine:=CalculationModuleDisabled) With {
                        .FormulaCulture = LoadOptions.FormulaCulture, .FormulaLocalClock = LoadOptions.FormulaLocalClock,
                        .CalculateOnCellEdit = LoadOptions.CalculateOnCellEdit,
                        .AllowedNativeRtdProgIds = LoadOptions.AllowedNativeRtdProgIds?.ToArray(),
                        .AllowKnownStructuralReferenceDifferences = LoadOptions.AllowKnownStructuralReferenceDifferences
                    }
                Dim copy As New OpenXmlExcelDataOperations(options)
                Dim snapshot As New MemoryStream()
                Try
                    WritePackage(snapshot, clearCache:=False, documentType:=_currentDocumentType, removeVba:=_vbaProjectRemoved)
                    copy.OpenSnapshot(snapshot)
                    copy.RecalculationRequired = RecalculationRequired
                    copy._recalculateOnOpen = _recalculateOnOpen
                    copy.AutoCalculationResetToEnabledForAllSavedWorkbooks = AutoCalculationResetToEnabledForAllSavedWorkbooks
                    Return copy
                Catch
                    snapshot.Dispose()
                    Throw
                End Try
            End SyncLock
        End Function

        Private Function CloneObject() As Object Implements ICloneable.Clone
            Return Clone()
        End Function

        Private Sub WritePackage(output As Stream, clearCache As Boolean, documentType As SpreadsheetDocumentType, removeVba As Boolean)
            Dim sourceWorkbook = WorkbookPart
            Dim linkedPaths = ResolvedSessionLinks(sourceWorkbook)
            _source.WriteTo(output)
            If Not (_workbookChanged OrElse _changedSheets.Count > 0 OrElse _changedChartSheets.Count > 0 OrElse _calculationChainInvalid OrElse
                    _automaticCalculation.HasValue OrElse _recalculateOnOpen OrElse clearCache OrElse
                    linkedPaths.Count > 0 OrElse
                    (removeVba AndAlso sourceWorkbook.VbaProjectPart IsNot Nothing) OrElse
                    documentType <> _document.DocumentType OrElse (IsMacroFree(documentType) AndAlso HasMacroSheets(sourceWorkbook))) Then Return

            output.Position = 0
            Using target = SpreadsheetDocument.Open(output, True, New OpenSettings With {.AutoSave = False})
                Dim targetWorkbook = target.WorkbookPart
                If IsMacroFree(documentType) AndAlso HasMacroSheets(targetWorkbook) Then
                    Throw UnsupportedOperation("Saving Excel 4.0 macro sheets to a macro-free format")
                End If
                SaveChangedParts(sourceWorkbook, targetWorkbook, clearCache)
                For Each change In linkedPaths
                    Dim linkPart = DirectCast(targetWorkbook.GetPartById(change.Id), ExternalWorkbookPart)
                    linkPart.DeleteExternalRelationship(change.Link.Id)
                    linkPart.AddExternalRelationship(change.Link.RelationshipType, change.Path, change.Link.Id)
                Next
                If removeVba AndAlso targetWorkbook.VbaProjectPart IsNot Nothing Then
                    targetWorkbook.DeletePart(targetWorkbook.VbaProjectPart)
                End If
                If documentType <> target.DocumentType Then target.ChangeDocumentType(documentType)
            End Using
        End Sub

        Private Shared Function HasMacroSheets(workbook As WorkbookPart) As Boolean
            Return workbook.Parts.Any(Function(part) part.OpenXmlPart.ContentType.EndsWith("macrosheet+xml", StringComparison.OrdinalIgnoreCase))
        End Function

        Private Sub SaveChangedParts(source As WorkbookPart, target As WorkbookPart, clearCache As Boolean)
            If _workbookChanged Then
                Using output = target.GetStream(FileMode.Create, FileAccess.Write)
                    source.Workbook.Save(output)
                End Using
            End If
            If _authoredDynamicMetadata IsNot Nothing Then
                Dim metadataPart = If(target.CellMetadataPart, target.AddNewPart(Of CellMetadataPart)())
                metadataPart.Metadata = DirectCast(_authoredDynamicMetadata.CloneNode(True), Metadata)
                metadataPart.Metadata.Save()
                EnsureWorkbookCalculationFeature(target.Workbook, "microsoft.com:Single")
            End If
            Dim richErrorMetadataIndices As New Dictionary(Of String, UInteger)(StringComparer.Ordinal)
            For Each sourceSheet In _changedSheets
                Dim targetSheet = DirectCast(target.GetPartById(source.GetIdOfPart(sourceSheet)), WorksheetPart)
                SaveWorksheet(sourceSheet.Worksheet, targetSheet)
                If Not clearCache AndAlso GetSheetData(sourceSheet).Descendants(Of Cell)().Any(AddressOf IsUnencodedRichError) Then
                    SaveRichFormulaErrors(target, targetSheet, richErrorMetadataIndices)
                End If
            Next
            For Each sourceChartSheet In _changedChartSheets
                Dim targetChartSheet = DirectCast(target.GetPartById(source.GetIdOfPart(sourceChartSheet)), ChartsheetPart)
                Using output = targetChartSheet.GetStream(FileMode.Create, FileAccess.Write)
                    sourceChartSheet.Chartsheet.Save(output)
                End Using
            Next
            SaveChangedTables(source, target)
            SaveChangedCharts(source, target)
            If clearCache Then
                For Each sheetPart In target.WorksheetParts
                    If ClearFormulaResults(GetSheetData(sheetPart), target.CellMetadataPart?.Metadata) Then SaveWorksheet(sheetPart.Worksheet, sheetPart)
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

        Private Shared Sub SaveWorksheet(worksheet As Worksheet, target As WorksheetPart)
            ' EPPlus 4 replaces worksheet content with unprefixed elements. Keep that
            ' content in SpreadsheetML even when the SDK writes an explicit root prefix.
            Using packaged As New MemoryStream()
                Using bufferWriter = XmlWriter.Create(packaged)
                    worksheet.WriteTo(bufferWriter)
                End Using
                packaged.Position = 0
                Using reader = XmlReader.Create(packaged), output = target.GetStream(FileMode.Create, FileAccess.Write)
                    reader.MoveToContent()
                    If reader.LookupNamespace(String.Empty) = reader.NamespaceURI Then
                        packaged.Position = 0
                        packaged.CopyTo(output)
                        Return
                    End If

                    Dim spreadsheetNamespace = reader.NamespaceURI
                    Using writer = XmlWriter.Create(output)
                        writer.WriteStartElement(reader.Prefix, reader.LocalName, spreadsheetNamespace)
                        While reader.MoveToNextAttribute()
                            If reader.Name <> "xmlns" Then writer.WriteAttributeString(reader.Prefix, reader.LocalName, reader.NamespaceURI, reader.Value)
                        End While
                        writer.WriteAttributeString("xmlns", spreadsheetNamespace)
                        reader.MoveToElement()
                        If Not reader.IsEmptyElement Then
                            reader.Read()
                            While reader.Depth > 0
                                writer.WriteNode(reader, defattr:=True)
                            End While
                        End If
                        writer.WriteEndElement()
                    End Using
                End Using
            End Using
        End Sub

        Private Shared Function CalculationFeatureDisabled() As FeatureDisabledException
            Return New FeatureDisabledException("Calculation engine", "Initial formula evaluation requires the Open XML calculation module to be enabled.")
        End Function

        ''' <inheritdoc/>
        Public Overrides ReadOnly Property HasVbaProject As Boolean
            Get
                SyncLock WorkbookSaveLock
                    Return WorkbookPart.VbaProjectPart IsNot Nothing AndAlso Not _vbaProjectRemoved
                End SyncLock
            End Get
        End Property

        ''' <inheritdoc/>
        ''' <remarks>The project is removed from the package on the next save.</remarks>
        Public Overrides Sub RemoveVbaProject()
            SyncLock WorkbookSaveLock
                If HasVbaProject Then _vbaProjectRemoved = True
            End SyncLock
        End Sub

    End Class

End Namespace
