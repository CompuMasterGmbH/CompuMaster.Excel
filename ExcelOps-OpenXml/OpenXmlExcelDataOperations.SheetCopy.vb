Imports System.Globalization
Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports ChartXml = DocumentFormat.OpenXml.Drawing.Charts
Imports ExtendedChartXml = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <inheritdoc/>
        Protected Overrides Function TransferSheetsCore(sheetNames As IReadOnlyList(Of String), targetWorkbook As ExcelDataOperationsBase,
                                                        targetSheetIndex As Integer, move As Boolean,
                                                        formulaMode As CopyWorksheetFormulaMode) As IReadOnlyList(Of String)
            Dim destination = TryCast(targetWorkbook, OpenXmlExcelDataOperations)
            If destination Is Nothing Then Throw UnsupportedOperation("Transferring sheets between different engines")
            SyncLock WorkbookSessionLock
                Dim books = If(move AndAlso WorkbookSession IsNot Nothing AndAlso WorkbookSession Is destination.WorkbookSession,
                               WorkbookSession.Snapshot(), {Me, destination})
                Return WithLockedWorkbooks(books, Function() TransferSheetsCore(sheetNames, destination, targetSheetIndex, move, formulaMode, books))
            End SyncLock
        End Function

        Private Overloads Function TransferSheetsCore(sheetNames As IReadOnlyList(Of String), destination As OpenXmlExcelDataOperations,
                                                      targetSheetIndex As Integer, move As Boolean,
                                                      formulaMode As CopyWorksheetFormulaMode,
                                                      sessionBooks As OpenXmlExcelDataOperations()) As IReadOnlyList(Of String)
            If If(destination.WorkbookPart.Workbook.WorkbookProtection?.LockStructure?.Value, False) Then
                Throw UnsupportedOperation("Adding sheets to a structure-protected workbook")
            End If
            ValidateCopyWorkbookStyles(destination)

            Dim assigned As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim taken As New HashSet(Of String)(destination.SheetNames(), StringComparer.OrdinalIgnoreCase)
            For Each name In sheetNames
                assigned.Add(name, NextCopiedSheetName(name, taken))
            Next

            Dim workbook = WorkbookPart.Workbook
            Dim sheetOrder = WorkbookSheets()
            Dim names As New TransferNamePlan(workbook, destination.WorkbookPart.Workbook, sheetOrder, assigned)
            Dim sourceWorksheets As New HashSet(Of String)(WorkSheetNames(), StringComparer.OrdinalIgnoreCase)
            Dim remaining = sheetOrder.Where(Function(sheet) Not assigned.ContainsKey(sheet.Name.Value)).ToArray()
            Dim sourceLinkNames = If(move, remaining, sheetOrder.ToArray()).Select(Function(sheet) sheet.Name.Value).ToArray()
            Dim destinationOrder = destination.SheetNames()
            destinationOrder.InsertRange(targetSheetIndex, sheetNames.Select(Function(name) assigned(name)))
            Dim targetLinkNames = destinationOrder.ToArray()
            Dim sameSession = WorkbookSession IsNot Nothing AndAlso WorkbookSession Is destination.WorkbookSession
            Dim targetLinks As New SheetTransferLinks(destination)
            Dim sourceLinks As New SheetTransferLinks(Me)
            Dim chartNames As New TransferredChartNames(destination.WorkbookPart.Workbook, assigned, sheetNames)

            Dim retainedTargetLinks As New HashSet(Of Integer)
            Dim oldSourceLinks As Integer() = Nothing
            If move Then
                ValidateMoveWorkbook()
                destination.ValidateMoveWorkbook()
                oldSourceLinks = destination.FindLinksTo(Me)
                If If(workbook.WorkbookProtection?.LockStructure?.Value, False) Then
                    Throw UnsupportedOperation("Removing sheets from a structure-protected workbook")
                End If
                If Not remaining.Any(AddressOf IsVisible) Then
                    Throw New NotSupportedException("The source workbook must retain at least one visible sheet after the move.")
                End If
            End If

            Dim resolveQualifier As Func(Of String, String, String) =
                Function(referencedName, originalQualifier)
                    If referencedName.StartsWith("[", StringComparison.Ordinal) Then
                        If Not move Then Throw UnsupportedReferenceSyntax()
                        Dim external = ReadTransferExternalReference(referencedName)
                        If destination.MatchesTransferPath(external.Path) Then Return FormatSheetName(external.Sheet)
                        Return targetLinks.ReferenceTo(external.Path, external.SheetNames, external.Sheet)
                    End If
                    If Not sourceWorksheets.Contains(referencedName) Then
                        Throw UnsupportedOperation("Transferring a formula with an unknown worksheet qualifier")
                    End If
                    Dim newName As String = Nothing
                    If assigned.TryGetValue(referencedName, newName) Then Return FormatSheetName(newName)
                    Return targetLinks.ReferenceTo(TransferWorkbookPath(sameSession), sourceLinkNames, referencedName)
                End Function

            Dim copiedSheets As New List(Of (Name As String, State As SheetStateValues?, Worksheet As Worksheet,
                                             SourceWorksheet As WorksheetPart, ChartSheet As ChartsheetPart, Drawing As DrawingsPart,
                                             Charts As Dictionary(Of String, OpenXmlPartRootElement)))
            For Each name In sheetNames
                Dim source = FindSheet(name)
                Dim part = WorkbookPart.GetPartById(source.Id.Value)
                Dim worksheet = TryCast(part, WorksheetPart)
                Dim chartSheet = TryCast(part, ChartsheetPart)
                If worksheet IsNot Nothing Then
                    Dim prepared = PrepareWorksheetCopy(worksheet, name, formulaMode, resolveQualifier, move,
                                                        Sub(identifier, qualifier) names.Observe(name, identifier, qualifier))
                    Dim drawing = worksheet.DrawingsPart
                    Dim placement As DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing = Nothing
                    If drawing IsNot Nothing AndAlso _changedDrawings.TryGetValue(drawing, placement) AndAlso Not placement.HasChildren Then
                        prepared.GetFirstChild(Of Drawing)().Remove()
                        drawing = Nothing
                    End If
                    Dim charts = If(drawing Is Nothing, Nothing, PrepareDrawingChartCopy(drawing, name, formulaMode, resolveQualifier, move, chartNames:=chartNames))
                    copiedSheets.Add((assigned(name), source.State?.Value, prepared, worksheet, Nothing, drawing, charts))
                ElseIf chartSheet IsNot Nothing Then
                    Dim charts = PrepareChartSheetCopy(chartSheet, name, formulaMode, resolveQualifier, move, chartNames:=chartNames)
                    copiedSheets.Add((assigned(name), source.State?.Value, Nothing, Nothing, chartSheet, chartSheet.DrawingsPart, charts))
                Else
                    Throw UnsupportedOperation("Copying this sheet type")
                End If
            Next
            If WorksheetThemesDiffer(destination.WorkbookPart) AndAlso copiedSheets.Any(Function(sheet) sheet.Drawing IsNot Nothing OrElse sheet.ChartSheet IsNot Nothing) Then
                Throw UnsupportedOperation("Transferring chart or drawing parts between workbooks with different themes")
            End If
            Dim copiedNames = PrepareTransferredNames(names, sheetNames, targetSheetIndex, resolveQualifier, formulaMode, move)
            If copiedNames.Any(Function(item) chartNames.Imported.Any(
                Function(chartName) String.Equals(item.Name.Value, chartName.Name.Value, StringComparison.OrdinalIgnoreCase))) Then
                Throw UnsupportedOperation("Transferring a worksheet formula that uses an internal ChartEx source name")
            End If

            Using changes As New StagedWorkbookChanges()
                changes.Add(destination,
                    Sub(bookPart)
                        TransferWorksheetStyles(bookPart, copiedSheets.Where(Function(sheet) sheet.Worksheet IsNot Nothing).Select(Function(sheet) sheet.Worksheet))
                        If WorkbookPart.ThemePart IsNot Nothing AndAlso bookPart.ThemePart Is Nothing Then
                            bookPart.AddPart(WorkbookPart.ThemePart)
                        End If
                        If move AndAlso oldSourceLinks.Length > 0 Then destination.RewriteMoveDependents(bookPart, Nothing,
                            Function(referencedName, originalQualifier)
                                If Not referencedName.StartsWith("[", StringComparison.Ordinal) Then Return originalQualifier
                                Dim external = destination.ReadTransferExternalReference(referencedName)
                                Dim assignedName As String = Nothing
                                If MatchesTransferPath(external.Path) AndAlso assigned.TryGetValue(external.Sheet, assignedName) Then
                                    Return FormatSheetName(assignedName)
                                End If
                                retainedTargetLinks.Add(external.Index)
                                Return originalQualifier
                            End Function)
                        targetLinks.ApplyTo(bookPart)
                        Dim destinationSheets = bookPart.Workbook.Sheets
                        Dim anchor = destinationSheets.Elements(Of Sheet)().ElementAtOrDefault(targetSheetIndex)
                        ShiftDestinationSheetIndexes(bookPart.Workbook, CUInt(targetSheetIndex), CUInt(copiedSheets.Count))
                        Dim nextSheetId = destinationSheets.Elements(Of Sheet)().Select(Function(sheet) sheet.SheetId.Value).DefaultIfEmpty(0UI).Max()
                        For Each copied In copiedSheets
                            Dim part As OpenXmlPart
                            If copied.Worksheet IsNot Nothing Then
                                Dim worksheetPart = bookPart.AddNewPart(Of WorksheetPart)()
                                worksheetPart.Worksheet = copied.Worksheet
                                For Each settings In copied.SourceWorksheet.Parts.Where(Function(item) TypeOf item.OpenXmlPart Is SpreadsheetPrinterSettingsPart)
                                    worksheetPart.AddPart(settings.OpenXmlPart, settings.RelationshipId)
                                Next
                                If copied.Drawing IsNot Nothing Then
                                    Dim drawingId = copied.Worksheet.GetFirstChild(Of Drawing)().Id.Value
                                    Dim drawing = worksheetPart.AddPart(copied.Drawing, drawingId)
                                    Dim placement As DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing = Nothing
                                    If _changedDrawings.TryGetValue(copied.Drawing, placement) Then
                                        drawing.WorksheetDrawing = DirectCast(placement.CloneNode(True), DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing)
                                        drawing.WorksheetDrawing.Save()
                                    End If
                                    SaveTransferredCharts(drawing, copied.Charts)
                                    If placement IsNot Nothing Then RemoveDetachedChartParts(drawing, placement)
                                End If
                                SaveWorksheet(worksheetPart.Worksheet, worksheetPart)
                                part = worksheetPart
                            Else
                                Dim chartSheet = bookPart.AddPart(copied.ChartSheet)
                                For Each view In chartSheet.Chartsheet.Descendants(Of ChartSheetView)()
                                    view.TabSelected = Nothing
                                Next
                                chartSheet.Chartsheet.Save()
                                SaveTransferredCharts(chartSheet.DrawingsPart, copied.Charts)
                                part = chartSheet
                            End If

                            nextSheetId += 1UI
                            Dim sheet As New Sheet With {.Name = copied.Name, .SheetId = nextSheetId, .Id = bookPart.GetIdOfPart(part)}
                            If copied.State.HasValue Then sheet.State = copied.State.Value
                            If anchor Is Nothing Then
                                destinationSheets.Append(sheet)
                            Else
                                destinationSheets.InsertBefore(sheet, anchor)
                            End If
                        Next
                        InstallTransferredNames(bookPart.Workbook, copiedNames, formulaMode)
                        InstallTransferredNames(bookPart.Workbook, chartNames.Imported, formulaMode)
                        If move Then
                            Dim obsolete = oldSourceLinks.Where(Function(index) Not retainedTargetLinks.Contains(index) AndAlso Not targetLinks.UsesIndex(index)).ToArray()
                            destination.RemoveUnusedMoveLinks(bookPart, obsolete)
                        End If
                    End Sub)
                If move Then
                    changes.Add(Me,
                        Sub(bookPart)
                            Dim removedChartNames = RemoveMovedChartSourceNames(bookPart, assigned, chartNames.SourceNames)
                            RewriteMoveDependents(bookPart, assigned,
                                Function(referencedName, originalQualifier)
                                    Dim assignedName As String = Nothing
                                    If assigned.TryGetValue(referencedName, assignedName) Then
                                        Return sourceLinks.ReferenceTo(destination.TransferWorkbookPath(sameSession), targetLinkNames, assignedName)
                                    End If
                                    Return originalQualifier
                                End Function, observeName:=Sub(identifier, qualifier)
                                    If removedChartNames.Contains(identifier) Then
                                        Throw UnsupportedOperation("Moving a chart whose source name is also used in a retained formula")
                                    End If
                                    names.ValidateRetainedReference(identifier, qualifier)
                                End Sub)
                            sourceLinks.ApplyTo(bookPart)
                            RemoveTransferredSheets(bookPart, assigned)
                        End Sub)
                    StageSessionMoveDependents(changes, sessionBooks, destination, assigned, targetLinkNames)
                End If
                changes.Apply()
            End Using
            Return copiedSheets.Select(Function(sheet) sheet.Name).ToArray()
        End Function

        Private Function PrepareWorksheetCopy(part As WorksheetPart, name As String, formulaMode As CopyWorksheetFormulaMode,
                                              resolveQualifier As Func(Of String, String, String), move As Boolean,
                                              observeName As Action(Of String, String)) As Worksheet
            ValidateCopySourceWorksheet(part)
            Dim worksheet = DirectCast(part.Worksheet.CloneNode(True), Worksheet)
            Dim data = worksheet.GetFirstChild(Of SheetData)()
            If data.Descendants(Of CellFormula)().Any() Then
                ValidateCopiedFormulaContext(move)
                If formulaMode = CopyWorksheetFormulaMode.ExcelCompatible Then
                    Dim expanded = ExpandSharedFormulas(data, name, Nothing)
                    If expanded IsNot data Then
                        worksheet.ReplaceChild(expanded, data)
                        data = expanded
                    End If
                    For Each formula In data.Descendants(Of CellFormula)()
                        formula.Text = RewriteFormula(formula.Text, name, name, Nothing, Nothing, False,
                                                      rewriteSheetQualifier:=resolveQualifier, observeName:=observeName)
                    Next
                End If
                ClearFormulaResults(data)
            End If
            ConvertCopiedSharedStrings(data)
            For Each view In worksheet.Descendants(Of SheetView)()
                view.TabSelected = Nothing
            Next
            Return worksheet
        End Function

        Private Sub ValidateCopiedFormulaContext(move As Boolean)
            Dim workbook = WorkbookPart.Workbook
            If Not move AndAlso If(workbook.ExternalReferences?.HasChildren, False) Then
                Throw UnsupportedOperation("Copying formulas from a workbook with external links")
            End If
        End Sub

        Private Function PrepareChartSheetCopy(part As ChartsheetPart, name As String, formulaMode As CopyWorksheetFormulaMode,
                                               resolveQualifier As Func(Of String, String, String), move As Boolean,
                                               Optional retained As Boolean = False, Optional formulaWorkbook As WorkbookPart = Nothing,
                                               Optional chartNames As TransferredChartNames = Nothing) As Dictionary(Of String, OpenXmlPartRootElement)
            Dim drawings = part.DrawingsPart
            If drawings?.WorksheetDrawing Is Nothing OrElse
               Not drawings.Parts.Any(Function(item) TypeOf item.OpenXmlPart Is ChartPart OrElse TypeOf item.OpenXmlPart Is ExtendedChartPart) OrElse
               part.Parts.Any(Function(item) item.OpenXmlPart IsNot drawings AndAlso Not TypeOf item.OpenXmlPart Is SpreadsheetPrinterSettingsPart) Then
                Throw UnsupportedOperation("Copying chart sheets with unsupported drawings or chart types")
            End If
            ValidateChartCopyRelationships(part)
            Return PrepareDrawingChartCopy(drawings, name, formulaMode, resolveQualifier, move, retained, formulaWorkbook, chartNames)
        End Function

        Private Function PrepareDrawingChartCopy(drawings As DrawingsPart, name As String, formulaMode As CopyWorksheetFormulaMode,
                                                  resolveQualifier As Func(Of String, String, String), move As Boolean,
                                                  Optional retained As Boolean = False,
                                                  Optional formulaWorkbook As WorkbookPart = Nothing,
                                                  Optional chartNames As TransferredChartNames = Nothing) As Dictionary(Of String, OpenXmlPartRootElement)
            If drawings.WorksheetDrawing Is Nothing OrElse drawings.Parts.Any(
                Function(item) Not TypeOf item.OpenXmlPart Is ChartPart AndAlso
                    Not TypeOf item.OpenXmlPart Is ExtendedChartPart AndAlso Not TypeOf item.OpenXmlPart Is ImagePart) Then
                Throw UnsupportedOperation("Transferring a worksheet with unsupported drawings or chart types")
            End If
            ValidateChartCopyRelationships(drawings)
            Dim placement As DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing = Nothing
            If Not _changedDrawings.TryGetValue(drawings, placement) Then placement = drawings.WorksheetDrawing
            Dim chartIds = ReadDrawingRelationshipIds(placement)
            Dim extended = drawings.Parts.Where(Function(item) chartIds.Contains(item.RelationshipId)).
                Select(Function(item) item.OpenXmlPart).OfType(Of ExtendedChartPart)().ToArray()
            StageChartAnchors(placement, Nothing, allowExtendedCharts:=extended.Length > 0)

            Dim result As New Dictionary(Of String, OpenXmlPartRootElement)(StringComparer.Ordinal)
            If extended.Length > 0 Then
                StageTransferredExtendedCharts(drawings, extended, name, formulaMode, resolveQualifier, move, retained, formulaWorkbook, chartNames, result)
            End If
            For Each chart In drawings.ChartParts
                Dim chartId = drawings.GetIdOfPart(chart)
                If Not chartIds.Contains(chartId) Then Continue For
                If extended.Length > 0 AndAlso IsExtendedChartFallback(placement, chartId) Then
                    ValidateExtendedChartFallback(chart)
                    Continue For
                End If
                If chart.ChartSpace Is Nothing OrElse chart.ChartSpace.Descendants(Of ChartXml.ExternalData)().Any() OrElse
                   chart.Parts.Any(Function(item) Not TypeOf item.OpenXmlPart Is ChartStylePart AndAlso
                                                 Not TypeOf item.OpenXmlPart Is ChartColorStylePart AndAlso
                                                 Not TypeOf item.OpenXmlPart Is ThemeOverridePart AndAlso
                                                 Not TypeOf item.OpenXmlPart Is ImagePart) Then
                    Throw UnsupportedOperation("Copying charts with unsupported related parts or external data")
                End If
                ValidateChartCopyRelationships(chart)

                Dim sourceChart As ChartXml.ChartSpace = Nothing
                If Not _changedCharts.TryGetValue(chart, sourceChart) Then sourceChart = chart.ChartSpace
                Dim chartSpace = DirectCast(sourceChart.CloneNode(True), ChartXml.ChartSpace)
                Dim formulas = chartSpace.Descendants(Of ChartXml.Formula)().ToArray()
                If formulas.Length > 0 AndAlso Not retained Then ValidateCopiedFormulaContext(move)
                For Each formula In formulas
                    If formulaMode = CopyWorksheetFormulaMode.ExcelCompatible Then
                        Dim qualifier = FormulaSheet.Match(formula.Text, 0)
                        Dim endpointEnd As Integer
                        If Not qualifier.Success OrElse ReadEndpoint(formula.Text, qualifier.Length, endpointEnd) Is Nothing Then
                            Throw UnsupportedOperation("Adapting a chart formula without an explicit worksheet reference")
                        End If
                        formula.Text = RewriteFormula(formula.Text, name, name, Nothing, Nothing, False,
                                                      rewriteSheetQualifier:=resolveQualifier, formulaWorkbook:=formulaWorkbook)
                    Else
                        ' The same reference text may now resolve to different data in the destination.
                        formula.Parent.RemoveAllChildren(Of ChartXml.NumberingCache)()
                        formula.Parent.RemoveAllChildren(Of ChartXml.StringCache)()
                        formula.Parent.RemoveAllChildren(Of ChartXml.MultiLevelStringCache)()
                    End If
                Next
                result.Add(drawings.GetIdOfPart(chart), chartSpace)
            Next
            Return result
        End Function

        Private Shared Sub SaveTransferredCharts(drawing As DrawingsPart, charts As Dictionary(Of String, OpenXmlPartRootElement))
            For Each chart In charts
                Dim target = drawing.GetPartById(chart.Key)
                If TypeOf target Is ChartPart Then
                    DirectCast(target, ChartPart).ChartSpace = DirectCast(chart.Value, ChartXml.ChartSpace)
                Else
                    DirectCast(target, ExtendedChartPart).ChartSpace = DirectCast(chart.Value, ExtendedChartXml.ChartSpace)
                End If
                chart.Value.Save()
            Next
        End Sub

        Private Sub ValidateChartCopyRelationships(part As OpenXmlPart)
            If part.ExternalRelationships.Any() OrElse part.HyperlinkRelationships.Any() OrElse part.DataPartReferenceRelationships.Any() Then
                Throw UnsupportedOperation("Copying a chart sheet with external relationships")
            End If
        End Sub

        Private Shared Sub ShiftDestinationSheetIndexes(workbook As Workbook, firstIndex As UInteger, count As UInteger)
            If workbook.DefinedNames IsNot Nothing Then
                For Each name In workbook.DefinedNames.Elements(Of DefinedName)()
                    If name.LocalSheetId IsNot Nothing AndAlso name.LocalSheetId.Value >= firstIndex Then
                        name.LocalSheetId = name.LocalSheetId.Value + count
                    End If
                Next
            End If
            If workbook.BookViews IsNot Nothing Then
                For Each view In workbook.BookViews.Elements(Of WorkbookView)()
                    Dim activeTab = If(view.ActiveTab?.Value, 0UI)
                    If activeTab >= firstIndex Then view.ActiveTab = activeTab + count
                    If view.FirstSheet IsNot Nothing AndAlso view.FirstSheet.Value >= firstIndex Then view.FirstSheet = view.FirstSheet.Value + count
                Next
            End If
        End Sub

        Private Sub ValidateCopyWorkbookStyles(destination As OpenXmlExcelDataOperations)
            Dim sourceProperties = WorkbookPart.Workbook.WorkbookProperties
            Dim targetProperties = destination.WorkbookPart.Workbook.WorkbookProperties
            If If(sourceProperties?.Date1904?.Value, False) <> If(targetProperties?.Date1904?.Value, False) Then
                Throw UnsupportedOperation("Copying between workbooks with different date systems")
            End If
        End Sub

        Private Sub ValidateCopySourceWorksheet(part As WorksheetPart)
            If part.Parts.Any(Function(item) Not (TypeOf item.OpenXmlPart Is DrawingsPart OrElse TypeOf item.OpenXmlPart Is SpreadsheetPrinterSettingsPart)) OrElse
               part.ExternalRelationships.Any() OrElse part.HyperlinkRelationships.Any() OrElse
               part.Worksheet.ChildElements.Any(Function(item) Not StandardWorksheetElements.Contains(item.LocalName)) Then
                Throw UnsupportedOperation("Copying a worksheet with related parts or unsupported worksheet elements")
            End If
            ValidateWorksheetStructure(part, Nothing)
            If GetSheetData(part).Descendants(Of CellFormula)().Any(Function(formula)
                Dim kind = If(formula.FormulaType?.Value, CellFormulaValues.Normal)
                Return kind <> CellFormulaValues.Normal AndAlso kind <> CellFormulaValues.Shared
            End Function) Then
                Throw UnsupportedOperation("Copying worksheets with array or data-table formulas")
            End If
        End Sub

        Private Sub ConvertCopiedSharedStrings(data As SheetData)
            Dim strings = WorkbookPart.SharedStringTablePart?.SharedStringTable?.Elements(Of SharedStringItem)().ToArray()
            For Each cell In data.Descendants(Of Cell)()
                If If(cell.DataType?.Value, CellValues.Number) <> CellValues.SharedString Then Continue For
                Dim index = Integer.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture)
                If strings Is Nothing OrElse index < 0 OrElse index >= strings.Length Then
                    Throw New InvalidDataException("A copied worksheet refers to a missing shared string.")
                End If
                cell.CellValue = Nothing
                cell.DataType = CellValues.InlineString
                cell.InlineString = New InlineString(strings(index).ChildElements.Select(Function(item) item.CloneNode(True)))
            Next
        End Sub

        Private Shared Function NextCopiedSheetName(name As String, taken As HashSet(Of String)) As String
            If taken.Add(name) Then Return name
            Dim index = 2
            Dim candidate As String
            Do
                Dim suffix = " (" & index.ToString(CultureInfo.InvariantCulture) & ")"
                Dim baseName = name.Substring(0, System.Math.Min(name.Length, 31 - suffix.Length))
                If Char.IsHighSurrogate(baseName(baseName.Length - 1)) Then baseName = baseName.Substring(0, baseName.Length - 1)
                candidate = baseName & suffix
                index += 1
            Loop Until taken.Add(candidate)
            Return candidate
        End Function

        Private Shared Function AddTransferLink(target As WorkbookPart, source As Uri, sheets As IEnumerable(Of String)) As ExternalBook
            Dim part = target.AddNewPart(Of ExternalWorkbookPart)()
            Dim relationship = part.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath", source)
            Dim sheetNames As New SheetNames(sheets.Select(Function(name) New SheetName With {.Val = name}))
            Dim book As New ExternalBook(sheetNames) With {.Id = relationship.Id}
            part.ExternalLink = New ExternalLink(book)
            part.ExternalLink.Save()
            If target.Workbook.ExternalReferences Is Nothing Then target.Workbook.ExternalReferences = New ExternalReferences()
            target.Workbook.ExternalReferences.Append(New ExternalReference With {.Id = target.GetIdOfPart(part)})
            Return book
        End Function

        Private Sub InstallTransferSnapshot(snapshot As MemoryStream, document As SpreadsheetDocument)
            _document.Dispose()
            _source.Dispose()
            _source = snapshot
            _document = document
            _changedSheets.Clear()
            _changedChartSheets.Clear()
            _changedTables.Clear()
            _changedCharts.Clear()
            _changedExtendedCharts.Clear()
            _changedDrawings.Clear()
            _formulaRanges.Clear()
            ResetCalculationDependencies()
            _workbookChanged = False
            _calculationChainInvalid = False
            RecalculationRequired = True
        End Sub

    End Class

End Namespace
