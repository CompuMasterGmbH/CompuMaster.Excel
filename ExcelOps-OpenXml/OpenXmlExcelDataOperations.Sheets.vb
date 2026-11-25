Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

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
            Dim sheet = FindSheet(sheetName)
            If IsVisible(sheet) Then Return
            sheet.State = SheetStateValues.Visible
            _workbookChanged = True
        End Sub

        ''' <inheritdoc/>
        Public Overrides Sub HideSheet(sheetName As String, stronglyHide As Boolean)
            Dim sheet = FindSheet(sheetName)
            Dim state = If(stronglyHide, SheetStateValues.VeryHidden, SheetStateValues.Hidden)
            If sheet.State?.Value = state Then Return
            Dim sheets = WorkbookSheets()
            Dim replacementIndex = sheets.FindIndex(Function(item) item IsNot sheet AndAlso IsVisible(item))
            If replacementIndex < 0 Then Throw New InvalidOperationException("At least one sheet must remain visible.")
            If IsVisible(sheet) Then SwitchFromHiddenSheet(sheets, sheets.IndexOf(sheet), replacementIndex)
            sheet.State = state
            _workbookChanged = True
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

        Private Sub SetTabSelected(sheet As Sheet, viewIndex As UInteger, selected As Boolean)
            Dim part = TryCast(WorkbookPart.GetPartById(sheet.Id.Value), WorksheetPart)
            Dim view = part?.Worksheet?.SheetViews?.Elements(Of SheetView)().FirstOrDefault(
                Function(item) item.WorkbookViewId IsNot Nothing AndAlso item.WorkbookViewId.Value = viewIndex)
            If view IsNot Nothing AndAlso If(view.TabSelected?.Value, False) <> selected Then
                view.TabSelected = selected
                _changedSheets.Add(part)
            End If
        End Sub

        ''' <inheritdoc/>
        Public Overrides Function SelectedSheetName() As String
            Dim sheets = WorkbookSheets()
            Dim index = WorkbookPart.Workbook.BookViews?.GetFirstChild(Of WorkbookView)()?.ActiveTab?.Value
            Return sheets.ElementAtOrDefault(CInt(If(index, 0UI)))?.Name?.Value
        End Function

        ''' <inheritdoc/>
        Public Overrides Function IsProtectedSheet(sheetName As String) As Boolean
            Return If(GetWorksheetPart(sheetName).Worksheet.GetFirstChild(Of SheetProtection)()?.Sheet?.Value, False)
        End Function

        ''' <inheritdoc/>
        Public Overrides Sub UnprotectSheet(sheetName As String)
            Dim part = GetWorksheetPart(sheetName)
            Dim protection = part.Worksheet.GetFirstChild(Of SheetProtection)()
            If protection Is Nothing Then Return
            If protection.Password IsNot Nothing OrElse protection.HashValue IsNot Nothing Then
                Throw UnsupportedOperation("Removing password-protected worksheet protection")
            End If
            protection.Remove()
            _changedSheets.Add(part)
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>The engine supports the protection levels that leave existing cell-lock styles in place.</remarks>
        Public Overrides Sub ProtectSheet(sheetName As String, level As ProtectionLevel)
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
