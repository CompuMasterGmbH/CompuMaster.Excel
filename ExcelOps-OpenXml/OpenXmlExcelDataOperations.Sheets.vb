Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <inheritdoc/>
        ''' <remarks>Removes definitions with broken references or invalid worksheet scope indexes.</remarks>
        Public Overrides Sub CleanupRangeNames()
            SyncLock WorkbookSaveLock
                Dim workbook = WorkbookPart.Workbook
                Dim names = workbook.DefinedNames
                If names Is Nothing Then Return
                Dim sheetCount = workbook.Sheets.Elements(Of Sheet)().Count()
                Dim invalid = names.Elements(Of DefinedName)().Where(
                    Function(name) (name.LocalSheetId IsNot Nothing AndAlso name.LocalSheetId.Value >= sheetCount) OrElse
                        DefinedNameTokens.Matches(If(name.Text, String.Empty)).Cast(Of Match)().Any(Function(token) token.Value(0) = "#"c)).ToList()
                If invalid.Count = 0 Then Return
                For Each name In invalid
                    name.Remove()
                Next
                If Not names.HasChildren Then names.Remove()
                _workbookChanged = True
                _calculationChainInvalid = True
                RecalculationRequired = True
                ResetCalculationDependencies()
            End SyncLock
        End Sub

        Private Shared ReadOnly DefinedNameTokens As New Regex(
            """(?:[^""]|"""")*""|'(?:[^']|'')*'|#REF!", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

        Private Function WorkbookSheets() As List(Of Sheet)
            Return WorkbookPart.Workbook.Sheets.Elements(Of Sheet)().ToList()
        End Function

        Private Function FindSheet(sheetName As String) As Sheet
            If String.IsNullOrEmpty(sheetName) Then Throw New ArgumentNullException(NameOf(sheetName))
            Dim sheet = WorkbookSheets().Find(Function(item) String.Equals(item.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
            If sheet Is Nothing Then Throw New ArgumentOutOfRangeException(NameOf(sheetName), "Sheet not found: " & sheetName)
            Return sheet
        End Function

        Private Function GetWorksheetPart(sheetName As String) As WorksheetPart
            Dim sheet = FindSheet(sheetName)
            Dim part = TryCast(WorkbookPart.GetPartById(sheet.Id.Value), WorksheetPart)
            If part Is Nothing Then Throw New ArgumentException("The selected sheet is not a worksheet.", NameOf(sheetName))
            Return part
        End Function

        ''' <inheritdoc/>
        Public Overrides Function SheetNames() As List(Of String)
            Return WorkbookSheets().Select(Function(sheet) sheet.Name.Value).ToList()
        End Function

        ''' <inheritdoc/>
        Public Overrides Function WorkSheetNames() As List(Of String)
            Return WorkbookSheets().Where(Function(sheet) TypeOf WorkbookPart.GetPartById(sheet.Id.Value) Is WorksheetPart).
                Select(Function(sheet) sheet.Name.Value).ToList()
        End Function

        ''' <inheritdoc/>
        Public Overrides Function ChartSheetNames() As List(Of String)
            Return WorkbookSheets().Where(Function(sheet) TypeOf WorkbookPart.GetPartById(sheet.Id.Value) Is ChartsheetPart).
                Select(Function(sheet) sheet.Name.Value).ToList()
        End Function

        Private Shared Function IsVisible(sheet As Sheet) As Boolean
            Return sheet.State Is Nothing OrElse sheet.State.Value = SheetStateValues.Visible
        End Function

        ''' <inheritdoc/>
        Public Overrides Function IsHiddenSheet(sheetName As String) As Boolean
            Return Not IsVisible(FindSheet(sheetName))
        End Function

        ''' <inheritdoc/>
        Public Overrides Sub UnhideSheet(sheetName As String)
            SyncLock WorkbookSaveLock
                    Dim sheet = FindSheet(sheetName)
                    If IsVisible(sheet) Then Return
                    sheet.State = SheetStateValues.Visible
                    _workbookChanged = True
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub HideSheet(sheetName As String, stronglyHide As Boolean)
            SyncLock WorkbookSaveLock
                    Dim sheet = FindSheet(sheetName)
                    Dim state = If(stronglyHide, SheetStateValues.VeryHidden, SheetStateValues.Hidden)
                    If sheet.State?.Value = state Then Return
                    Dim sheets = WorkbookSheets()
                    Dim replacementIndex = sheets.FindIndex(Function(item) item IsNot sheet AndAlso IsVisible(item))
                    If replacementIndex < 0 Then Throw New InvalidOperationException("At least one sheet must remain visible.")
                    If IsVisible(sheet) Then SwitchFromHiddenSheet(sheets, sheets.IndexOf(sheet), replacementIndex)
                    sheet.State = state
                    _workbookChanged = True
            End SyncLock
        End Sub

        Private Sub SwitchFromHiddenSheet(sheets As List(Of Sheet), oldIndex As Integer, newIndex As Integer)
            Dim workbook = WorkbookPart.Workbook
            Dim views = workbook.BookViews
            If views Is Nothing AndAlso oldIndex = 0 Then
                views = New BookViews(New WorkbookView())
                workbook.BookViews = views
            End If
            If views Is Nothing Then Return
            Dim viewIndex As UInteger = 0
            For Each view In views.Elements(Of WorkbookView)()
                Dim wasActive = If(view.ActiveTab?.Value, 0UI) = CUInt(oldIndex)
                If wasActive Then view.ActiveTab = CUInt(newIndex)
                If If(view.FirstSheet?.Value, 0UI) = CUInt(oldIndex) Then view.FirstSheet = CUInt(newIndex)
                SetTabSelected(sheets(oldIndex), viewIndex, False)
                If wasActive Then SetTabSelected(sheets(newIndex), viewIndex, True)
                viewIndex += 1UI
            Next
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>The name must identify a visible worksheet.</remarks>
        Public Overrides Sub SelectSheet(sheetName As String)
            SyncLock WorkbookSaveLock
                Dim sheet = FindSheet(sheetName)
                If Not TypeOf WorkbookPart.GetPartById(sheet.Id.Value) Is WorksheetPart Then
                    Throw New ArgumentException("The selected sheet is not a worksheet.", NameOf(sheetName))
                End If
                SelectVisibleWorksheet(WorkbookSheets(), sheet)
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Hidden worksheets count toward the index. Chart sheets are skipped.</remarks>
        Public Overrides Sub SelectSheet(sheetIndex As Integer)
            SyncLock WorkbookSaveLock
                Dim sheets = WorkbookSheets()
                Dim sheet = If(sheetIndex < 0, Nothing, sheets.Where(
                    Function(item) TypeOf WorkbookPart.GetPartById(item.Id.Value) Is WorksheetPart).ElementAtOrDefault(sheetIndex))
                If sheet Is Nothing Then Throw New ArgumentOutOfRangeException(NameOf(sheetIndex))
                SelectVisibleWorksheet(sheets, sheet)
            End SyncLock
        End Sub

        Private Sub SelectVisibleWorksheet(sheets As List(Of Sheet), selected As Sheet)
            If Not IsVisible(selected) Then Throw New InvalidOperationException("A hidden worksheet cannot be selected.")
            Dim view = GetOrCreateFirstWorkbookView()
            Dim index = CUInt(sheets.IndexOf(selected))
            If If(view.ActiveTab?.Value, 0UI) <> index Then
                view.ActiveTab = index
                _workbookChanged = True
            End If
            If If(view.FirstSheet?.Value, 0UI) > index Then
                view.FirstSheet = index
                _workbookChanged = True
            End If
            For Each sheet In sheets
                SetTabSelected(sheet, 0UI, sheet Is selected)
            Next
        End Sub

        Private Function GetOrCreateFirstWorkbookView() As WorkbookView
            Dim workbook = WorkbookPart.Workbook
            Dim workbookViews = workbook.BookViews
            If workbookViews Is Nothing Then
                workbookViews = New BookViews()
                workbook.AddChild(workbookViews)
            End If
            Dim view = workbookViews.GetFirstChild(Of WorkbookView)()
            If view Is Nothing Then
                view = workbookViews.AppendChild(New WorkbookView())
                _workbookChanged = True
            End If
            Return view
        End Function

        ''' <inheritdoc/>
        ''' <remarks>Updates the target worksheet's selection in the first workbook view and preserves the active sheet.</remarks>
        Public Overrides Sub SelectCell(cell As ExcelCell)
            ValidateAddress(cell)
            ValidateCoordinates(cell.RowIndex, cell.ColumnIndex)
            SyncLock WorkbookSaveLock
                Dim part = GetWorksheetPart(cell.SheetName)
                Dim merged = MergedCells(cell.SheetName).FirstOrDefault(Function(range) CellIsInRange(cell.RowIndex, cell.ColumnIndex, range))
                Dim address = If(merged?.AddressStart.Address, ExcelCell.LocalCellAddress(cell.RowIndex, cell.ColumnIndex))
                Dim reference = If(merged Is Nothing, address, address & ":" & merged.AddressEnd.Address)
                GetOrCreateFirstWorkbookView()
                Dim view = GetWorksheetView(part, 0UI, create:=True)
                Dim pane = view.GetFirstChild(Of Pane)()
                Dim selectedPane = If(pane?.ActivePane?.Value, PaneValues.TopLeft)
                Dim paneState = pane?.State?.Value
                If paneState = PaneStateValues.Frozen OrElse paneState = PaneStateValues.FrozenSplit Then
                    Dim row = If(merged?.AddressStart.RowIndex, cell.RowIndex)
                    Dim column = If(merged?.AddressStart.ColumnIndex, cell.ColumnIndex)
                    Dim rowSplit = If(pane.VerticalSplit?.Value, 0.0)
                    Dim columnSplit = If(pane.HorizontalSplit?.Value, 0.0)
                    Dim bottom = rowSplit > 0 AndAlso row >= rowSplit
                    Dim right = columnSplit > 0 AndAlso column >= columnSplit
                    selectedPane = If(bottom, If(right, PaneValues.BottomRight, PaneValues.BottomLeft), If(right, PaneValues.TopRight, PaneValues.TopLeft))
                End If

                Dim selection = view.Elements(Of Selection)().FirstOrDefault(Function(item) If(item.Pane?.Value, PaneValues.TopLeft) = selectedPane)
                If selection Is Nothing Then
                    selection = New Selection()
                    If pane IsNot Nothing Then selection.Pane = selectedPane
                    Dim preceding As OpenXmlElement = If(DirectCast(view.Elements(Of Selection)().LastOrDefault(), OpenXmlElement), pane)
                    If preceding Is Nothing Then
                        view.PrependChild(selection)
                    Else
                        view.InsertAfter(selection, preceding)
                    End If
                End If
                If pane IsNot Nothing AndAlso selectedPane <> If(pane.ActivePane?.Value, PaneValues.TopLeft) Then pane.ActivePane = selectedPane
                selection.ActiveCell = address
                selection.ActiveCellId = Nothing
                selection.SequenceOfReferences = New ListValue(Of StringValue) With {.InnerText = reference}
                _changedSheets.Add(part)
            End SyncLock
        End Sub

        Private Shared Function CellIsInRange(rowIndex As Integer, columnIndex As Integer, range As ExcelRange) As Boolean
            Return rowIndex >= range.AddressStart.RowIndex AndAlso rowIndex <= range.AddressEnd.RowIndex AndAlso
                columnIndex >= range.AddressStart.ColumnIndex AndAlso columnIndex <= range.AddressEnd.ColumnIndex
        End Function

        Private Shared Function GetWorksheetView(part As WorksheetPart, viewIndex As UInteger, create As Boolean) As SheetView
            Dim views = part.Worksheet.SheetViews
            Dim view = views?.Elements(Of SheetView)().FirstOrDefault(Function(item) item.WorkbookViewId IsNot Nothing AndAlso item.WorkbookViewId.Value = viewIndex)
            If view IsNot Nothing OrElse Not create Then Return view
            If views Is Nothing Then
                views = New SheetViews()
                part.Worksheet.AddChild(views)
            End If
            Return views.AppendChild(New SheetView With {.WorkbookViewId = viewIndex})
        End Function

        Private Sub SetTabSelected(sheet As Sheet, viewIndex As UInteger, selected As Boolean)
            Dim part = WorkbookPart.GetPartById(sheet.Id.Value)
            Dim worksheet = TryCast(part, WorksheetPart)
            If worksheet IsNot Nothing Then
                Dim view = GetWorksheetView(worksheet, viewIndex, create:=selected)
                If view IsNot Nothing AndAlso If(view.TabSelected?.Value, False) <> selected Then
                    view.TabSelected = selected
                    _changedSheets.Add(worksheet)
                End If
                Return
            End If

            Dim chart = TryCast(part, ChartsheetPart)
            Dim chartView = chart?.Chartsheet?.GetFirstChild(Of ChartSheetViews)()?.Elements(Of ChartSheetView)().
                FirstOrDefault(Function(item) item.WorkbookViewId IsNot Nothing AndAlso item.WorkbookViewId.Value = viewIndex)
            If chartView IsNot Nothing AndAlso If(chartView.TabSelected?.Value, False) <> selected Then
                chartView.TabSelected = selected
                _changedChartSheets.Add(chart)
            End If
        End Sub

        ''' <inheritdoc/>
        Public Overrides Function SelectedSheetName() As String
            Dim sheets = WorkbookSheets()
            Dim index = WorkbookPart.Workbook.BookViews?.GetFirstChild(Of WorkbookView)()?.ActiveTab?.Value
            Dim active = sheets.ElementAtOrDefault(CInt(If(index, 0UI)))
            Return If(active IsNot Nothing AndAlso IsVisible(active), active.Name?.Value, Nothing)
        End Function

        ''' <inheritdoc/>
        Public Overrides Function IsProtectedSheet(sheetName As String) As Boolean
            Return If(GetWorksheetPart(sheetName).Worksheet.GetFirstChild(Of SheetProtection)()?.Sheet?.Value, False)
        End Function

        ''' <inheritdoc/>
        Public Overrides Sub UnprotectSheet(sheetName As String)
            SyncLock WorkbookSaveLock
                    Dim part = GetWorksheetPart(sheetName)
                    Dim protection = part.Worksheet.GetFirstChild(Of SheetProtection)()
                    If protection Is Nothing Then Return
                    If protection.Password IsNot Nothing OrElse protection.HashValue IsNot Nothing Then
                        Throw UnsupportedOperation("Removing password-protected worksheet protection")
                    End If
                    protection.Remove()
                    _changedSheets.Add(part)
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>The engine supports the protection levels that leave existing cell-lock styles in place.</remarks>
        Public Overrides Sub ProtectSheet(sheetName As String, level As ProtectionLevel)
            SyncLock WorkbookSaveLock
                    Select Case level
                        Case ProtectionLevel.Standard, ProtectionLevel.StandardWithInsertDeleteRows,
                             ProtectionLevel.SelectAndEditUnlockedCellsOnly, ProtectionLevel.SelectNoCellsAndNoEditing
                        Case ProtectionLevel.SelectAndEditAllCellsButNoFurtherEditing
                            Throw UnsupportedOperation("Protecting a worksheet while allowing edits to locked cells")
                        Case Else
                            Throw New ArgumentOutOfRangeException(NameOf(level))
                    End Select
                    Dim part = GetWorksheetPart(sheetName)
                    Dim protection = part.Worksheet.GetFirstChild(Of SheetProtection)()
                    If protection Is Nothing Then
                        protection = New SheetProtection()
                        part.Worksheet.AddChild(protection)
                    End If
                    Dim preventRowChanges = level <> ProtectionLevel.StandardWithInsertDeleteRows
                    protection.Sheet = True
                    protection.SelectLockedCells = level = ProtectionLevel.SelectAndEditUnlockedCellsOnly OrElse level = ProtectionLevel.SelectNoCellsAndNoEditing
                    protection.SelectUnlockedCells = level = ProtectionLevel.SelectNoCellsAndNoEditing
                    protection.InsertRows = preventRowChanges
                    protection.DeleteRows = preventRowChanges
                    protection.Objects = True
                    protection.Scenarios = True
                    protection.InsertColumns = True
                    protection.DeleteColumns = True
                    protection.FormatCells = True
                    protection.FormatColumns = True
                    protection.FormatRows = True
                    protection.InsertHyperlinks = True
                    protection.Sort = True
                    protection.AutoFilter = True
                    protection.PivotTables = True
                    _changedSheets.Add(part)
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        Public Overrides Function LookupCellIsLocked(cell As ExcelCell) As Boolean
            ValidateAddress(cell)
            Return LookupCellIsLocked(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <inheritdoc/>
        Public Overrides Function LookupCellIsLocked(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            ValidateCoordinates(rowIndex, columnIndex)
            Dim part = GetWorksheetPart(sheetName)
            Dim row = FindRow(GetSheetData(part), rowIndex, False)
            Dim cell = FindCell(row, rowIndex, columnIndex, False)
            Dim stylesheet = WorkbookPart.WorkbookStylesPart?.Stylesheet
            Dim format = GetCellFormat(stylesheet, WorksheetCellStyleIndex(part.Worksheet, row, cell, columnIndex))
            If format Is Nothing Then Return True
            If format.Protection IsNot Nothing AndAlso (format.ApplyProtection Is Nothing OrElse format.ApplyProtection.Value) Then
                Return If(format.Protection.Locked?.Value, True)
            End If
            Dim baseFormat = stylesheet.CellStyleFormats?.Elements(Of CellFormat)().ElementAtOrDefault(CInt(If(format.FormatId?.Value, 0UI)))
            Return If(baseFormat?.Protection?.Locked?.Value, True)
        End Function

    End Class

End Namespace
