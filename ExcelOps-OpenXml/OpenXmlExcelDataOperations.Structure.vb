Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Class StructuralShift
            Friend ReadOnly Vertical As Boolean
            Friend ReadOnly Start As Integer
            Friend ReadOnly Count As Integer
            Friend ReadOnly Inserting As Boolean
            Private ReadOnly _otherStart As Integer
            Private ReadOnly _otherCount As Integer
            Private ReadOnly _expandedTables As New List(Of StructuralTable)
            Private ReadOnly _tablesWithDeletedColumns As New List(Of StructuralTable)
            Private ReadOnly _emptiedTable As StructuralTable

            Friend Sub New(vertical As Boolean, start As Integer, count As Integer, otherStart As Integer, otherCount As Integer, inserting As Boolean)
                Me.Vertical = vertical
                Me.Start = start
                Me.Count = count
                Me.Inserting = inserting
                _otherStart = otherStart
                _otherCount = otherCount
                If start < 0 OrElse start >= AxisSize Then Throw New ArgumentOutOfRangeException(NameOf(start))
                If count < 1 OrElse count > AxisSize - start Then Throw New ArgumentOutOfRangeException(NameOf(count))
                If otherStart < 0 OrElse otherStart >= OtherAxisSize Then Throw New ArgumentOutOfRangeException(NameOf(otherStart))
                If otherCount < 1 OrElse otherCount > OtherAxisSize - otherStart Then Throw New ArgumentOutOfRangeException(NameOf(otherCount))
            End Sub

            Private Sub New(original As StructuralShift, emptiedTable As StructuralTable)
                Vertical = True
                Start = original.Start + 1
                Count = original.Count - 1
                _otherCount = ColumnCount
                _emptiedTable = emptiedTable
            End Sub

            Friend Function RetainEmptyTableRow(worksheet As WorksheetPart, tables As IEnumerable(Of StructuralTable)) As StructuralShift
                If Not Vertical OrElse Inserting OrElse Not WholeAxis Then Return Me
                Dim affected = tables.Where(Function(table) table.Worksheet Is worksheet AndAlso
                    Start <= table.Last.Row AndAlso Start + Count > table.First.Row).ToArray()
                Dim emptied = affected.FirstOrDefault(Function(table) table.HeaderCount = 1 AndAlso
                    Start = table.FirstDataRow AndAlso Start + Count = table.LastDataRow + 1)
                If emptied Is Nothing Then Return Me
                If affected.Length <> 1 OrElse If(emptied.Source.GetFirstChild(Of AutoFilter)()?.Elements(Of FilterColumn)().Any(), False) Then
                    Throw New NotSupportedException("Deleting a table's entire data body with another intersecting table or an active filter is unsupported.")
                End If
                ' Excel retains the first row across the worksheet and clears only the cells inside the table.
                Return New StructuralShift(Me, emptied)
            End Function

            Friend ReadOnly Property ClearsTableDataBody As Boolean
                Get
                    Return _emptiedTable IsNot Nothing
                End Get
            End Property

            Friend Function ClearsTableCell(row As Integer, column As Integer) As Boolean
                Return _emptiedTable IsNot Nothing AndAlso row = _emptiedTable.FirstDataRow AndAlso
                    column >= _emptiedTable.First.Column AndAlso column <= _emptiedTable.Last.Column
            End Function

            Friend Function ClearsTableRange(first As ReferenceEndpoint, last As ReferenceEndpoint) As Boolean
                If _emptiedTable Is Nothing Then Return False
                Dim finalCell = If(last, first)
                Return first.Row <= _emptiedTable.FirstDataRow AndAlso finalCell.Row >= _emptiedTable.FirstDataRow AndAlso
                    first.Column <= _emptiedTable.Last.Column AndAlso finalCell.Column >= _emptiedTable.First.Column
            End Function

            Friend ReadOnly Property RetainsClearedCellStyles As Boolean
                Get
                    Return _emptiedTable.TotalsCount = 0
                End Get
            End Property

            Private ReadOnly Property AxisSize As Integer
                Get
                    Return If(Vertical, RowCount, ColumnCount)
                End Get
            End Property

            Private ReadOnly Property OtherAxisSize As Integer
                Get
                    Return If(Vertical, ColumnCount, RowCount)
                End Get
            End Property

            Friend ReadOnly Property WholeAxis As Boolean
                Get
                    Return _otherCount = OtherAxisSize
                End Get
            End Property

            Friend Function Includes(otherIndex As Integer) As Boolean
                Return otherIndex >= _otherStart AndAlso otherIndex < _otherStart + _otherCount
            End Function

            Friend Function MoveIndex(index As Integer) As Integer
                If index < Start Then Return index
                If Inserting Then
                    If index + Count >= AxisSize Then Throw New InvalidOperationException("The insertion would move existing cells or rows outside the worksheet.")
                    Return index + Count
                End If
                Return If(index < Start + Count, -1, index - Count)
            End Function

            Friend Function RewriteReference(first As ReferenceEndpoint, second As ReferenceEndpoint, separator As String) As String
                Dim last = If(second, first)
                Dim firstIndex = If(Vertical, first.Row, first.Column)
                Dim lastIndex = If(Vertical, last.Row, last.Column)
                Dim firstOtherIndex = If(Vertical, first.Column, first.Row)
                Dim lastOtherIndex = If(Vertical, last.Column, last.Row)

                ' Excel shifts a reference after a partial cell edit only if the strip contains its full width or height.
                If firstIndex >= 0 AndAlso (WholeAxis OrElse (Includes(firstOtherIndex) AndAlso Includes(lastOtherIndex))) Then
                    Dim low = System.Math.Min(firstIndex, lastIndex)
                    Dim high = System.Math.Max(firstIndex, lastIndex)
                    If (low <> 0 OrElse high <> AxisSize - 1) Then
                        If second IsNot Nothing AndAlso ExtendsTableReference(low, high, firstOtherIndex, lastOtherIndex) Then
                            high += Count
                        ElseIf high >= Start Then
                            ShiftBounds(low, high)
                        End If
                        If low > high Then Return "#REF!"
                        If firstIndex <= lastIndex Then
                            firstIndex = low
                            lastIndex = high
                        Else
                            firstIndex = high
                            lastIndex = low
                        End If
                    End If
                End If

                Dim text = first.Format(Vertical, firstIndex)
                If second IsNot Nothing Then text &= separator & second.Format(Vertical, lastIndex)
                Return text
            End Function

            Friend Function RewriteDefinedNameReference(first As ReferenceEndpoint, second As ReferenceEndpoint,
                                                         separator As String, unchanged As String) As String
                ' A whole-axis edit can move absolute coordinates on that axis even if the other coordinate is relative.
                ' A partial edit needs both coordinates fixed to identify the shifted cell.
                Dim firstFixed = If(WholeAxis, first.IsAbsoluteOnAxis(Vertical), first.IsAbsolute)
                Dim last = If(second, first)
                Dim lastFixed = If(WholeAxis, last.IsAbsoluteOnAxis(Vertical), last.IsAbsolute)
                If firstFixed AndAlso lastFixed Then Return RewriteReference(first, second, separator)
                If Not firstFixed AndAlso Not lastFixed Then Return unchanged

                ' Excel applies a shift to the fixed endpoint of a mixed range as if it were a single reference.
                Dim fixedEndpoint = If(firstFixed, first, last)
                Dim rewritten = RewriteReference(fixedEndpoint, Nothing, Nothing)
                If rewritten = "#REF!" Then Return rewritten
                If firstFixed Then Return rewritten & separator & last.ColumnText & last.RowText
                Return first.ColumnText & first.RowText & separator & rewritten
            End Function

            Friend Sub PlanTableShift(table As StructuralTable)
                If table Is _emptiedTable Then
                    table.TargetLastRow -= Count
                    table.ClearsDataBody = True
                    Return
                End If
                If Count = 0 Then Return
                Dim low = If(Vertical, table.First.Row, table.First.Column)
                Dim high = If(Vertical, table.Last.Row, table.Last.Column)
                Dim otherLow = If(Vertical, table.First.Column, table.First.Row)
                Dim otherHigh = If(Vertical, table.Last.Column, table.Last.Row)
                If high < Start OrElse otherLow >= _otherStart + _otherCount OrElse otherHigh < _otherStart Then Return
                If Not Includes(otherLow) OrElse Not Includes(otherHigh) Then
                    Throw New NotSupportedException("This structural edit would change part of an Excel table in an unsupported way.")
                End If
                If Not Vertical Then
                    If Not WholeAxis Then
                        Throw New NotSupportedException("This structural edit would change the columns of an Excel table in an unsupported way.")
                    End If
                    If (Inserting AndAlso Start <= low) OrElse (Not Inserting AndAlso Start + Count <= low) Then
                        table.TargetFirstColumn = MoveIndex(low)
                        table.TargetLastColumn = MoveIndex(high)
                        Return
                    End If
                    If Not Inserting AndAlso (Start < low OrElse Start + Count > high + 1 OrElse Count >= high - low + 1) Then
                        Throw New NotSupportedException("Removing Excel table columns together with adjoining columns or removing an entire table is unsupported.")
                    End If
                    table.ColumnEditOffset = Start - low
                    If Inserting Then
                        table.TargetLastColumn = MoveIndex(high)
                    Else
                        table.TargetLastColumn = high - Count
                        Dim columns = table.Source.GetFirstChild(Of TableColumns)()?.Elements(Of TableColumn)().ToArray()
                        If columns Is Nothing OrElse columns.Length <> high - low + 1 Then
                            Throw New IO.InvalidDataException("The table's columns do not match its range.")
                        End If
                        table.DeletedColumnNames.UnionWith(columns.Skip(table.ColumnEditOffset).Take(Count).Select(Function(column) column.Name.Value))
                        _tablesWithDeletedColumns.Add(table)
                    End If
                    Return
                End If
                If If(table.Source.InsertRow?.Value, False) AndAlso Start > low Then
                    Throw New NotSupportedException("Shifting through an empty Excel table's insertion row is unsupported.")
                End If
                If Not Inserting AndAlso Start + Count > low Then
                    If WholeAxis AndAlso table.TotalsCount = 1 AndAlso Start = high AndAlso Count = 1 AndAlso
                       table.LastDataRow >= table.FirstDataRow Then
                        table.RemovesTotalsRow = True
                        table.TargetLastRow = high - 1
                        Return
                    End If
                    If Not WholeAxis OrElse Start <= low OrElse Start < table.FirstDataRow OrElse
                       Start + Count > table.LastDataRow + 1 OrElse Count >= table.LastDataRow - table.FirstDataRow + 1 Then
                        Throw New NotSupportedException("This deletion would remove rows of an Excel table in an unsupported way.")
                    End If
                    table.RemovesFirstDataRow = Start = table.FirstDataRow
                End If
                table.TargetLastRow = MoveIndex(high)
                If Start <= low Then
                    table.TargetFirstRow = MoveIndex(low)
                ElseIf Inserting Then
                    table.AddsDataRows = True
                    _expandedTables.Add(table)
                End If
            End Sub

            Friend Function TableForCell(sheetName As String, address As String) As StructuralTable
                If _tablesWithDeletedColumns.Count = 0 OrElse address Is Nothing Then Return Nothing
                Dim cell = ReadStructuralRange(address).First
                Return _tablesWithDeletedColumns.FirstOrDefault(Function(table)
                    Return String.Equals(table.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) AndAlso
                        cell.Row >= table.TargetFirstRow AndAlso cell.Row <= table.TargetLastRow AndAlso
                        cell.Column >= table.TargetFirstColumn AndAlso cell.Column <= table.TargetLastColumn
                End Function)
            End Function

            Friend Function ReferencesDeletedTableColumn(value As String, selectorStart As Integer, selectorEnd As Integer,
                                                         tableName As String, contextualTable As StructuralTable, isDefinedName As Boolean) As Boolean
                If _tablesWithDeletedColumns.Count = 0 Then Return False
                If tableName IsNot Nothing Then
                    contextualTable = _tablesWithDeletedColumns.FirstOrDefault(Function(table)
                        Return String.Equals(table.Source.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase) OrElse
                            String.Equals(table.Source.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase)
                    End Function)
                End If
                If contextualTable IsNot Nothing Then
                    Return TableSelectorUsesColumn(value, selectorStart, selectorEnd, contextualTable.DeletedColumnNames)
                End If
                If tableName Is Nothing AndAlso isDefinedName AndAlso _tablesWithDeletedColumns.Any(
                    Function(table) TableSelectorUsesColumn(value, selectorStart, selectorEnd, table.DeletedColumnNames)) Then
                    Throw New NotSupportedException("A defined name contains a contextual reference to a removed table column.")
                End If
                Return False
            End Function

            Private Function ExtendsTableReference(low As Integer, high As Integer, otherStart As Integer, otherEnd As Integer) As Boolean
                For Each table In _expandedTables
                    If System.Math.Min(otherStart, otherEnd) < table.First.Column OrElse
                       System.Math.Max(otherStart, otherEnd) > table.Last.Column Then Continue For
                    If Start = table.FirstDataRow AndAlso low = Start AndAlso
                       (high = table.LastDataRow OrElse high = table.Last.Row) Then Return True
                    If Start = table.LastDataRow + 1 AndAlso high = table.LastDataRow AndAlso
                       (low = table.FirstDataRow OrElse low = table.First.Row) Then Return True
                Next
                Return False
            End Function

            Friend Function RewriteMergedArea(first As ReferenceEndpoint, last As ReferenceEndpoint) As String
                Dim low = If(Vertical, first.Row, first.Column)
                Dim high = If(Vertical, last.Row, last.Column)
                Dim otherLow = If(Vertical, first.Column, first.Row)
                Dim otherHigh = If(Vertical, last.Column, last.Row)

                If high >= Start AndAlso otherLow < _otherStart + _otherCount AndAlso otherHigh >= _otherStart Then
                    ' Excel removes the merge when only part of its width or height moves.
                    If Not Includes(otherLow) OrElse Not Includes(otherHigh) Then Return Nothing
                    If Inserting Then MoveIndex(high)
                    ShiftBounds(low, high)
                    If low > high OrElse (low = high AndAlso otherLow = otherHigh) Then Return Nothing
                End If

                Return first.Format(Vertical, low) & ":" & last.Format(Vertical, high)
            End Function

            Private Sub ShiftBounds(ByRef low As Integer, ByRef high As Integer)
                If Inserting Then
                    If low >= Start Then low += Count
                    high = System.Math.Min(high + Count, AxisSize - 1)
                Else
                    Dim after = Start + Count
                    low = If(low >= after, low - Count, System.Math.Min(low, Start))
                    high = If(high >= after, high - Count, Start - 1)
                End If
            End Sub
        End Class

        ''' <inheritdoc/>
        ''' <remarks>Supports automatic reference updates for the workbook structures documented by this engine.</remarks>
        Public Overrides Sub AddColumn(sheetName As String, columnIndex As Integer, columns As Integer, updateFormulasAndReferences As Boolean)
            ValidateFormulaReferenceUpdateRequest(updateFormulasAndReferences)
            ApplyStructuralEdit(sheetName, New StructuralShift(False, columnIndex, columns, 0, RowCount, True), Nothing)
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Supports automatic reference updates for the workbook structures documented by this engine.</remarks>
        Public Overrides Sub AddRow(sheetName As String, rowIndex As Integer, rows As Integer, updateFormulasAndReferences As Boolean)
            ValidateFormulaReferenceUpdateRequest(updateFormulasAndReferences)
            ApplyStructuralEdit(sheetName, New StructuralShift(True, rowIndex, rows, 0, ColumnCount, True), Nothing)
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Supports automatic reference updates for the workbook structures documented by this engine.</remarks>
        Public Overrides Sub AddCells(sheetName As String, rowIndex As Integer, columnIndex As Integer, rows As Integer, columns As Integer, shiftDirection As CellInsertShiftDirection, updateFormulasAndReferences As Boolean)
            ValidateFormulaReferenceUpdateRequest(updateFormulasAndReferences)
            Select Case shiftDirection
                Case CellInsertShiftDirection.ShiftCellsDown
                    ApplyStructuralEdit(sheetName, New StructuralShift(True, rowIndex, rows, columnIndex, columns, True), Nothing)
                Case CellInsertShiftDirection.ShiftCellsRight
                    ApplyStructuralEdit(sheetName, New StructuralShift(False, columnIndex, columns, rowIndex, rows, True), Nothing)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(shiftDirection))
            End Select
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Supports automatic reference updates for the workbook structures documented by this engine.</remarks>
        Public Overrides Sub RemoveColumns(sheetName As String, startColumnIndex As Integer, columns As Integer, updateFormulasAndReferences As Boolean)
            ValidateFormulaReferenceUpdateRequest(updateFormulasAndReferences)
            ApplyStructuralEdit(sheetName, New StructuralShift(False, startColumnIndex, columns, 0, RowCount, False), Nothing)
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Supports automatic reference updates for the workbook structures documented by this engine.</remarks>
        Public Overrides Sub RemoveRows(sheetName As String, startRowIndex As Integer, rows As Integer, updateFormulasAndReferences As Boolean)
            ValidateFormulaReferenceUpdateRequest(updateFormulasAndReferences)
            ApplyStructuralEdit(sheetName, New StructuralShift(True, startRowIndex, rows, 0, ColumnCount, False), Nothing)
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Supports automatic reference updates for the workbook structures documented by this engine.</remarks>
        Public Overrides Sub RemoveCells(sheetName As String, rowIndex As Integer, columnIndex As Integer, rows As Integer, columns As Integer, shiftDirection As CellRemoveShiftDirection, updateFormulasAndReferences As Boolean)
            ValidateFormulaReferenceUpdateRequest(updateFormulasAndReferences)
            Select Case shiftDirection
                Case CellRemoveShiftDirection.ShiftCellsUp
                    ApplyStructuralEdit(sheetName, New StructuralShift(True, rowIndex, rows, columnIndex, columns, False), Nothing)
                Case CellRemoveShiftDirection.ShiftCellsLeft
                    ApplyStructuralEdit(sheetName, New StructuralShift(False, columnIndex, columns, rowIndex, rows, False), Nothing)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(shiftDirection))
            End Select
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Supports automatic updates of sheet qualifiers for the workbook structures documented by this engine.</remarks>
        Public Overrides Sub RenameWorksheet(sheetName As String, newSheetName As String, updateFormulasAndReferences As Boolean)
            ValidateFormulaReferenceUpdateRequest(updateFormulasAndReferences)
            If newSheetName Is Nothing Then Throw New ArgumentNullException(NameOf(newSheetName))
            If newSheetName.Length = 0 OrElse newSheetName.Length > 31 OrElse
               newSheetName.IndexOfAny(InvalidWorksheetNameCharacters) >= 0 OrElse
               newSheetName(0) = "'"c OrElse newSheetName(newSheetName.Length - 1) = "'"c OrElse
               newSheetName.Any(Function(character) Char.IsControl(character)) Then
                Throw New ArgumentException("The requested name is not a valid Excel worksheet name.", NameOf(newSheetName))
            End If
            Try
                System.Xml.XmlConvert.VerifyXmlChars(newSheetName)
            Catch ex As System.Xml.XmlException
                Throw New ArgumentException("The requested name contains characters that cannot be stored in an XLSX workbook.", NameOf(newSheetName), ex)
            End Try
            ApplyStructuralEdit(sheetName, Nothing, newSheetName)
        End Sub

        Private Shared ReadOnly InvalidWorksheetNameCharacters As Char() = {":"c, "\"c, "/"c, "?"c, "*"c, "["c, "]"c}

        Private Shared ReadOnly StandardWorkbookElements As New HashSet(Of String)(StringComparer.Ordinal) From {
            "fileVersion", "workbookPr", "workbookProtection", "bookViews", "sheets", "definedNames", "calcPr", "fileRecoveryPr"
        }

        Private Shared ReadOnly StandardWorksheetElements As New HashSet(Of String)(StringComparer.Ordinal) From {
            "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection",
            "printOptions", "pageMargins", "pageSetup", "headerFooter", "drawing", "mergeCells", "tableParts"
        }

        Private Shared ReadOnly LocalWorksheetReferences As New HashSet(Of String)(StringComparer.Ordinal) From {
            "autoFilter", "sortState", "rowBreaks", "colBreaks", "protectedRanges"
        }

        ' Modern Excel adds these workbook metadata elements even to a new, empty workbook.
        Private Shared Function IsReferenceFreeWorkbookMetadata(element As OpenXmlElement) As Boolean
            Select Case element.LocalName
                Case "revisionPtr"
                    Return element.NamespaceUri = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision" AndAlso Not element.HasChildren
                Case "AlternateContent"
                    Const compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006"
                    Return element.NamespaceUri = compatibility AndAlso element.ChildElements.Count = 1 AndAlso
                        element.FirstChild.LocalName = "Choice" AndAlso element.FirstChild.NamespaceUri = compatibility AndAlso
                        element.FirstChild.ChildElements.Count = 1 AndAlso element.FirstChild.FirstChild.LocalName = "absPath" AndAlso
                        element.FirstChild.FirstChild.NamespaceUri = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"
                Case "extLst"
                    If element.NamespaceUri <> "http://schemas.openxmlformats.org/spreadsheetml/2006/main" Then Return False
                    Return element.ChildElements.All(AddressOf IsReferenceFreeWorkbookExtension)
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsReferenceFreeWorkbookExtension(extension As OpenXmlElement) As Boolean
            If extension.LocalName <> "ext" OrElse extension.NamespaceUri <> "http://schemas.openxmlformats.org/spreadsheetml/2006/main" OrElse
               extension.ChildElements.Count <> 1 Then Return False
            Dim content = extension.FirstChild
            Select Case content.NamespaceUri
                Case "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                    Return content.LocalName = "workbookPr" AndAlso Not content.HasChildren
                Case "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures"
                    Return content.LocalName = "calcFeatures" AndAlso content.ChildElements.All(
                        Function(feature) feature.LocalName = "feature" AndAlso feature.NamespaceUri = content.NamespaceUri AndAlso Not feature.HasChildren)
                Case "http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion"
                    Return content.LocalName = "version" AndAlso Not content.HasChildren
                Case Else
                    Return False
            End Select
        End Function

        Private Sub ValidateWorkbookStructure(Optional allowChartSheets As Boolean = False, Optional allowSpillMetadata As Boolean = False)
            For Each child In WorkbookPart.Workbook.ChildElements
                If TypeOf child Is ExternalReferences Then Continue For
                If Not IsReferenceFreeWorkbookMetadata(child) AndAlso
                   (TypeOf child Is OpenXmlUnknownElement OrElse Not StandardWorkbookElements.Contains(child.LocalName)) Then
                    Throw UnsupportedOperation("Updating references in a workbook with " & child.LocalName)
                End If
            Next
            For Each relationship In WorkbookPart.Parts
                Dim part = relationship.OpenXmlPart
                If allowChartSheets AndAlso TypeOf part Is ChartsheetPart Then Continue For
                If allowSpillMetadata AndAlso TypeOf part Is CellMetadataPart Then Continue For
                Dim externalPart = TryCast(part, ExternalWorkbookPart)
                Dim externalBookId = externalPart?.ExternalLink?.GetFirstChild(Of ExternalBook)()?.Id?.Value
                If externalBookId IsNot Nothing AndAlso
                   externalPart.ExternalRelationships.Any(Function(link) link.Id = externalBookId AndAlso
                                                         link.RelationshipType.EndsWith("/externalLinkPath", StringComparison.Ordinal)) Then Continue For
                If Not (TypeOf part Is WorksheetPart OrElse TypeOf part Is WorkbookStylesPart OrElse
                        TypeOf part Is SharedStringTablePart OrElse TypeOf part Is ThemePart OrElse TypeOf part Is CalculationChainPart) Then
                    Throw UnsupportedOperation("Updating references in a workbook with " & part.GetType().Name)
                End If
            Next
        End Sub

        Private Sub ValidateWorksheetStructure(part As WorksheetPart, shift As StructuralShift, Optional allowSpills As Boolean = False)
            Dim worksheet = part.Worksheet
            For Each child In worksheet.ChildElements
                If TypeOf child Is OpenXmlUnknownElement OrElse
                   Not (StandardWorksheetElements.Contains(child.LocalName) OrElse
                        (shift Is Nothing AndAlso LocalWorksheetReferences.Contains(child.LocalName))) Then
                    Throw UnsupportedOperation("Updating references in a worksheet with " & child.LocalName)
                End If
            Next
            If shift IsNot Nothing Then
                If If(worksheet.GetFirstChild(Of SheetProtection)()?.Sheet?.Value, False) Then
                    Throw UnsupportedOperation("Shifting cells in a protected worksheet")
                End If
            End If

            For Each relationship In part.Parts
                If TypeOf relationship.OpenXmlPart Is SpreadsheetPrinterSettingsPart OrElse
                   TypeOf relationship.OpenXmlPart Is TableDefinitionPart Then Continue For
                If TypeOf relationship.OpenXmlPart Is DrawingsPart Then Continue For
                Throw UnsupportedOperation("Updating references in a worksheet with " & relationship.OpenXmlPart.GetType().Name)
            Next
            For Each rowElement In GetSheetData(part).ChildElements
                Dim row = TryCast(rowElement, Row)
                If row Is Nothing Then Throw UnsupportedOperation("Updating references in a worksheet with " & rowElement.LocalName)
                For Each cellElement In row.ChildElements
                    Dim cell = TryCast(cellElement, Cell)
                    If cell Is Nothing OrElse cell.ChildElements.Any(
                        Function(child) Not (TypeOf child Is CellFormula OrElse TypeOf child Is CellValue OrElse TypeOf child Is InlineString)) Then
                        Throw UnsupportedOperation("Updating references in a worksheet with row or cell extensions")
                    End If
                    Dim formula = cell.CellFormula
                    Dim formulaType = If(formula?.FormulaType?.Value, CellFormulaValues.Normal)
                    If allowSpills AndAlso If(cell.CellMetaIndex?.Value, 0UI) = 1UI AndAlso cell.ValueMetaIndex Is Nothing AndAlso
                       formulaType = CellFormulaValues.Array AndAlso Not String.IsNullOrEmpty(formula.Reference?.Value) Then Continue For
                    If cell.CellMetaIndex IsNot Nothing OrElse cell.ValueMetaIndex IsNot Nothing OrElse
                       (formulaType <> CellFormulaValues.Shared AndAlso
                        (formulaType <> CellFormulaValues.Normal OrElse formula?.Reference IsNot Nothing)) Then
                        Throw UnsupportedOperation("Updating references in a workbook with array, data-table or metadata-backed cells")
                    End If
                Next
            Next
        End Sub

        Private Sub ApplyStructuralEdit(sheetName As String, shift As StructuralShift, newSheetName As String)
            ApplyWorkbookSessionEdit(sheetName, shift, newSheetName)
        End Sub

        Private Sub ApplyStructuralEditCore(sheetName As String, shift As StructuralShift, newSheetName As String,
                                            stageDependents As Action(Of String, StructuralShift))
            Dim editedSheet = FindSheet(sheetName)
            Dim editedPart = TryCast(WorkbookPart.GetPartById(editedSheet.Id.Value), WorksheetPart)
            If editedPart Is Nothing Then Throw New ArgumentException("The selected sheet is not a worksheet.", NameOf(sheetName))
            Dim sheets = WorkbookSheets()
            Dim workbook = WorkbookPart.Workbook
            Dim originalName = editedSheet.Name.Value
            If newSheetName IsNot Nothing Then
                If String.Equals(originalName, newSheetName, StringComparison.Ordinal) Then Return
                If sheets.Any(Function(item) item IsNot editedSheet AndAlso String.Equals(item.Name.Value, newSheetName, StringComparison.OrdinalIgnoreCase)) Then
                    Throw New ArgumentException("A sheet with the requested name already exists.", NameOf(newSheetName))
                End If
                If If(workbook.WorkbookProtection?.LockStructure?.Value, False) Then Throw UnsupportedOperation("Renaming a worksheet in a structure-protected workbook")
            End If

            ' Verify and stage all changes before modifying the open workbook.
            Dim allowSpills = HasSupportedSpillMetadata(FormulaCellMetadata)
            ValidateWorkbookStructure(allowChartSheets:=True, allowSpillMetadata:=allowSpills)
            Dim tables = GetStructuralTables(sheets)
            If shift IsNot Nothing Then shift = shift.RetainEmptyTableRow(editedPart, tables)
            Dim originalColumns As Columns = Nothing
            Dim shiftedColumns As Columns = Nothing
            If shift IsNot Nothing AndAlso Not shift.Vertical AndAlso shift.WholeAxis Then
                Dim layouts = editedPart.Worksheet.Elements(Of Columns)().ToArray()
                If layouts.Length > 1 Then Throw UnsupportedOperation("Shifting whole columns with several column layout groups")
                originalColumns = layouts.SingleOrDefault()
                If originalColumns IsNot Nothing Then
                    Dim hasFilter = tables.Any(Function(table) table.Worksheet Is editedPart AndAlso
                        If(table.Source.GetFirstChild(Of AutoFilter)()?.Elements(Of FilterColumn)().Any(Function(field) field.HasChildren), False))
                    shiftedColumns = ShiftColumnLayout(originalColumns, shift, hasFilter)
                End If
            End If
            Dim tableNames = GetStructuralTableNames(tables)
            Dim tableFills As New List(Of (Row As Integer, Column As Integer, Formula As String, Header As String))
            Dim tableChanges = StageTableChanges(tables, editedPart, originalName, shift, newSheetName, tableNames, tableFills,
                                                 allowSpillReferences:=allowSpills)
            Dim filledCells As New HashSet(Of Cell)
            Dim formulaChanges As New List(Of (Part As WorksheetPart, Cell As Cell, Text As String))
            Dim revisedSheetData As New Dictionary(Of WorksheetPart, SheetData)
            Dim originalMerges As MergeCells = Nothing
            Dim shiftedMerges As MergeCells = Nothing
            Dim shiftedData As SheetData = Nothing
            Dim dimension As String = Nothing
            For Each sheet In sheets
                Dim part = TryCast(WorkbookPart.GetPartById(sheet.Id.Value), WorksheetPart)
                If part Is Nothing Then Continue For
                Dim sheetShift = If(part Is editedPart, shift, Nothing)
                ValidateWorksheetStructure(part, sheetShift, allowSpills)
                Dim originalData = GetSheetData(part)
                Dim data = ExpandSharedFormulas(originalData, sheet.Name.Value, tableNames, allowSpills)
                If allowSpills Then data = PrepareStructuralSpills(data, sheetShift)
                If sheetShift IsNot Nothing Then
                    originalMerges = part.Worksheet.GetFirstChild(Of MergeCells)()
                    shiftedMerges = ShiftMergedCells(originalMerges, sheetShift)
                    shiftedData = ShiftSheetData(data, sheetShift, shiftedMerges, tableFills, filledCells, dimension)
                    ApplyTableFilterVisibility(shiftedData, tables)
                    data = shiftedData
                End If
                If data IsNot originalData Then revisedSheetData.Add(part, data)
                For Each cell In data.Descendants(Of Cell)()
                    Dim formula = cell.CellFormula
                    If formula Is Nothing OrElse filledCells.Contains(cell) Then Continue For
                    Dim text = RewriteFormula(formula.Text, sheet.Name.Value, originalName, shift, newSheetName, isDefinedName:=False,
                                              tableNames:=tableNames, contextualTable:=shift?.TableForCell(sheet.Name.Value, cell.CellReference?.Value),
                                              allowSpillReferences:=allowSpills)
                    If text <> formula.Text OrElse cell.CellValue IsNot Nothing Then
                        formulaChanges.Add((part, cell, text))
                    End If
                Next
            Next

            Dim nameChanges As New List(Of (Definition As DefinedName, Text As String))
            If workbook.DefinedNames IsNot Nothing Then
                For Each definition In workbook.DefinedNames.Elements(Of DefinedName)()
                    Dim scopeSheet As String = Nothing
                    If definition.LocalSheetId IsNot Nothing Then
                        If definition.LocalSheetId.Value >= sheets.Count Then
                            Throw New IO.InvalidDataException("A defined name has an invalid worksheet scope.")
                        End If
                        scopeSheet = sheets(CInt(definition.LocalSheetId.Value)).Name.Value
                    End If
                    Dim text = RewriteFormula(definition.Text, scopeSheet, originalName, shift, newSheetName, isDefinedName:=True,
                                              tableNames:=tableNames, allowSpillReferences:=allowSpills)
                    If text <> definition.Text Then nameChanges.Add((definition, text))
                Next
            End If

            Dim chartChanges = StageChartChanges(sheets, editedPart, originalName, shift, newSheetName,
                                                 editedData:=shiftedData, editedColumns:=shiftedColumns)
            stageDependents(originalName, shift)
            For Each revision In revisedSheetData
                revision.Key.Worksheet.ReplaceChild(revision.Value, GetSheetData(revision.Key))
                InvalidateGroupedFormulaRanges(revision.Key)
                _changedSheets.Add(revision.Key)
            Next
            If shiftedData IsNot Nothing Then
                If originalColumns IsNot Nothing Then
                    If shiftedColumns.HasChildren Then
                        editedPart.Worksheet.ReplaceChild(shiftedColumns, originalColumns)
                    Else
                        originalColumns.Remove()
                    End If
                End If
                If originalMerges IsNot Nothing Then
                    If shiftedMerges Is Nothing Then
                        originalMerges.Remove()
                    Else
                        editedPart.Worksheet.ReplaceChild(shiftedMerges, originalMerges)
                    End If
                End If
                If editedPart.Worksheet.SheetDimension IsNot Nothing Then editedPart.Worksheet.SheetDimension.Reference = dimension
            End If
            For Each change In formulaChanges
                change.Cell.CellFormula.Text = change.Text
                ClearFormulaResult(change.Cell)
                _changedSheets.Add(change.Part)
            Next
            For Each change In nameChanges
                change.Definition.Text = change.Text
            Next
            For Each change In tableChanges
                _changedTables(change.Key) = change.Value
            Next
            For Each change In chartChanges.Charts
                _changedCharts(change.Key) = change.Value
            Next
            For Each change In chartChanges.ExtendedCharts
                _changedExtendedCharts(change.Key) = change.Value
            Next
            For Each change In chartChanges.Drawings
                _changedDrawings(change.Key) = change.Value
            Next
            If newSheetName IsNot Nothing Then editedSheet.Name = newSheetName

            InvalidateWorkbookFormulaResults()
        End Sub

        Private Function ShiftColumnLayout(source As Columns, shift As StructuralShift, hasFilter As Boolean) As Columns
            Dim widths(ColumnCount - 1) As Column
            Dim supportedAttributes = {"min", "max", "width", "customWidth", "hidden"}
            Dim lastColumn As UInteger = 0
            For Each item In source.ChildElements
                Dim column = TryCast(item, Column)
                If column Is Nothing OrElse column.GetAttributes().Any(Function(attribute) attribute.NamespaceUri <> "" OrElse
                    Not supportedAttributes.Contains(attribute.LocalName)) Then
                    Throw UnsupportedOperation("Shifting whole columns with unsupported width, style or outline information")
                End If
                Dim first = column.Min?.Value
                Dim last = column.Max?.Value
                If Not first.HasValue OrElse Not last.HasValue OrElse first.Value <= lastColumn OrElse
                   last.Value < first.Value OrElse last.Value > ColumnCount Then
                    Throw New IO.InvalidDataException("The worksheet contains invalid or overlapping column layout ranges.")
                End If
                For index As Integer = CInt(first.Value) - 1 To CInt(last.Value) - 1
                    widths(index) = column
                Next
                lastColumn = last.Value
            Next

            If Not shift.Inserting AndAlso hasFilter AndAlso Not LoadOptions.AllowKnownStructuralReferenceDifferences AndAlso
               widths.Skip(shift.Start).Take(shift.Count).All(Function(column) If(column?.Hidden?.Value, False)) Then
                Throw New StructuralReferenceCompatibilityException(EngineName,
                    "Excel refuses to remove columns when every selected column is hidden and a table filter is active on the worksheet")
            End If

            Dim inserted = If(shift.Inserting AndAlso shift.Start > 0, widths(shift.Start - 1), Nothing)
            If If(inserted?.Width?.Value, 0.0) > 0.0 Then
                ' Excel inherits the width from the left while the new columns remain visible.
                inserted = DirectCast(inserted.CloneNode(False), Column)
                inserted.Hidden = Nothing
            Else
                inserted = Nothing
            End If

            Dim result = DirectCast(source.CloneNode(False), Columns)
            Dim previous As Column = Nothing
            Dim offset = If(shift.Inserting, -shift.Count, shift.Count)
            For index = 0 To ColumnCount - 1
                Dim layout As Column = Nothing
                If index < shift.Start Then
                    layout = widths(index)
                ElseIf shift.Inserting AndAlso index < shift.Start + shift.Count Then
                    layout = inserted
                ElseIf index + offset < ColumnCount Then
                    layout = widths(index + offset)
                End If
                If layout Is Nothing Then
                    previous = Nothing
                ElseIf previous IsNot Nothing AndAlso Nullable.Equals(previous.Width?.Value, layout.Width?.Value) AndAlso
                    Nullable.Equals(previous.CustomWidth?.Value, layout.CustomWidth?.Value) AndAlso
                    Nullable.Equals(previous.Hidden?.Value, layout.Hidden?.Value) Then
                    previous.Max = CUInt(index + 1)
                Else
                    previous = DirectCast(layout.CloneNode(False), Column)
                    previous.Min = CUInt(index + 1)
                    previous.Max = previous.Min.Value
                    result.Append(previous)
                End If
            Next
            Return result
        End Function

        Private Shared Function ReadStructuralRange(reference As String) As (First As ReferenceEndpoint, Last As ReferenceEndpoint)
            If String.IsNullOrEmpty(reference) OrElse reference.IndexOf("$"c) >= 0 Then
                Throw New IO.InvalidDataException("The worksheet contains an invalid rectangular range.")
            End If

            Dim firstEnd As Integer
            Dim first = ReadEndpoint(reference, 0, firstEnd)
            If first Is Nothing OrElse Not first.IsCell Then
                Throw New IO.InvalidDataException("Invalid rectangular range: " & reference)
            End If

            If firstEnd = reference.Length Then Return (first, first)
            If reference(firstEnd) <> ":"c Then Throw New IO.InvalidDataException("Invalid rectangular range: " & reference)

            Dim lastEnd As Integer
            Dim last = ReadEndpoint(reference, firstEnd + 1, lastEnd)
            If last Is Nothing OrElse Not last.IsCell OrElse lastEnd <> reference.Length OrElse
               first.Row > last.Row OrElse first.Column > last.Column Then
                Throw New IO.InvalidDataException("Invalid rectangular range: " & reference)
            End If
            Return (first, last)
        End Function

        Private Function ShiftMergedCells(source As MergeCells, shift As StructuralShift) As MergeCells
            If source Is Nothing Then Return Nothing

            Dim result = DirectCast(source.CloneNode(False), MergeCells)
            For Each element In source.ChildElements
                Dim merge = TryCast(element, MergeCell)
                If merge Is Nothing OrElse merge.HasChildren Then
                    Throw UnsupportedOperation("Updating a worksheet with extensions in merged cell ranges")
                End If
                Dim area = ReadStructuralRange(merge.Reference?.Value)
                Dim reference = shift.RewriteMergedArea(area.First, area.Last)
                If reference IsNot Nothing Then
                    Dim shifted = DirectCast(merge.CloneNode(False), MergeCell)
                    shifted.Reference = reference
                    result.Append(shifted)
                End If
            Next
            If Not result.HasChildren Then Return Nothing
            If source.Count IsNot Nothing Then result.Count = CUInt(result.ChildElements.Count)
            Return result
        End Function

        Private Shared Function ShiftSheetData(source As SheetData, shift As StructuralShift, merges As MergeCells,
                                               tableFills As List(Of (Row As Integer, Column As Integer, Formula As String, Header As String)),
                                               filledCells As HashSet(Of Cell), ByRef dimension As String) As SheetData
            Dim rows As New SortedDictionary(Of Integer, Row)
            Dim cells As New Dictionary(Of Integer, List(Of (Column As Integer, Cell As Cell)))
            Dim rowIndex = -1
            Dim minRow = RowCount
            Dim minColumn = ColumnCount
            Dim maxRow = -1
            Dim maxColumn = -1

            For Each sourceRow In source.Elements(Of Row)()
                Dim nextRowIndex = If(sourceRow.RowIndex?.Value, CUInt(rowIndex + 2)) - 1L
                If nextRowIndex <= rowIndex OrElse nextRowIndex >= RowCount Then Throw New IO.InvalidDataException("Invalid row order or index in the worksheet.")
                rowIndex = CInt(nextRowIndex)

                Dim newRowIndex = If(shift.Vertical AndAlso shift.WholeAxis, shift.MoveIndex(rowIndex), rowIndex)
                If newRowIndex < 0 Then Continue For
                rows.Add(newRowIndex, DirectCast(sourceRow.CloneNode(False), Row))

                Dim columnIndex = -1
                For Each sourceCell In sourceRow.Elements(Of Cell)()
                    Dim nextColumnIndex = columnIndex + 1
                    If sourceCell.CellReference IsNot Nothing Then
                        Dim address = sourceCell.CellReference.Value
                        Dim tokenEnd As Integer
                        Dim endpoint = ReadEndpoint(address, 0, tokenEnd)
                        If endpoint Is Nothing OrElse Not endpoint.IsCell OrElse tokenEnd <> address.Length OrElse endpoint.Row <> rowIndex Then
                            Throw New IO.InvalidDataException("Invalid cell address in the worksheet: " & address)
                        End If
                        nextColumnIndex = endpoint.Column
                    End If
                    If nextColumnIndex <= columnIndex OrElse nextColumnIndex >= ColumnCount Then
                        Throw New IO.InvalidDataException("Invalid cell order or column index in the worksheet.")
                    End If
                    columnIndex = nextColumnIndex

                    Dim destinationRow = rowIndex
                    Dim destinationColumn = columnIndex
                    If shift.Vertical AndAlso shift.Includes(columnIndex) Then destinationRow = shift.MoveIndex(rowIndex)
                    If Not shift.Vertical AndAlso shift.Includes(rowIndex) Then destinationColumn = shift.MoveIndex(columnIndex)
                    If destinationRow < 0 OrElse destinationColumn < 0 Then Continue For

                    Dim cell As Cell
                    If shift.ClearsTableCell(rowIndex, columnIndex) Then
                        If Not shift.RetainsClearedCellStyles OrElse sourceCell.StyleIndex Is Nothing Then Continue For
                        cell = New Cell With {.StyleIndex = sourceCell.StyleIndex.Value}
                    Else
                        cell = DirectCast(sourceCell.CloneNode(True), Cell)
                    End If
                    cell.CellReference = ExcelCell.LocalCellAddress(destinationRow, destinationColumn)
                    Dim rowCells As List(Of (Column As Integer, Cell As Cell)) = Nothing
                    If Not cells.TryGetValue(destinationRow, rowCells) Then
                        rowCells = New List(Of (Column As Integer, Cell As Cell))
                        cells.Add(destinationRow, rowCells)
                    End If
                    rowCells.Add((destinationColumn, cell))
                    minRow = System.Math.Min(minRow, destinationRow)
                    maxRow = System.Math.Max(maxRow, destinationRow)
                    minColumn = System.Math.Min(minColumn, destinationColumn)
                    maxColumn = System.Math.Max(maxColumn, destinationColumn)
                Next
            Next

            For Each fill In tableFills
                Dim cell As New Cell With {.CellReference = ExcelCell.LocalCellAddress(fill.Row, fill.Column)}
                Dim rowCells As List(Of (Column As Integer, Cell As Cell)) = Nothing
                If Not cells.TryGetValue(fill.Row, rowCells) Then
                    rowCells = New List(Of (Column As Integer, Cell As Cell))
                    cells.Add(fill.Row, rowCells)
                End If
                If fill.Header IsNot Nothing Then
                    cell.DataType = CellValues.InlineString
                    cell.InlineString = New InlineString(New Text(EncodeText(fill.Header)) With {.Space = SpaceProcessingModeValues.Preserve})
                    Dim previous = rowCells.Find(Function(item) item.Column = fill.Column - 1).Cell
                    If previous?.StyleIndex IsNot Nothing Then cell.StyleIndex = previous.StyleIndex.Value
                Else
                    cell.CellFormula = New CellFormula(fill.Formula)
                End If
                rowCells.Add((fill.Column, cell))
                filledCells.Add(cell)
                minRow = System.Math.Min(minRow, fill.Row)
                maxRow = System.Math.Max(maxRow, fill.Row)
                minColumn = System.Math.Min(minColumn, fill.Column)
                maxColumn = System.Math.Max(maxColumn, fill.Column)
            Next

            For Each index In cells.Keys
                If Not rows.ContainsKey(index) Then rows.Add(index, New Row())
            Next
            Dim precedingRow As Row = Nothing
            If shift.Vertical AndAlso shift.WholeAxis AndAlso shift.Inserting AndAlso
               rows.TryGetValue(shift.Start - 1, precedingRow) AndAlso If(precedingRow.Height?.Value, 0.0) > 0 Then
                For index = shift.Start To shift.Start + shift.Count - 1
                    Dim inserted As Row = Nothing
                    If Not rows.TryGetValue(index, inserted) Then
                        inserted = New Row()
                        rows.Add(index, inserted)
                    End If
                    inserted.Height = precedingRow.Height.Value
                    inserted.CustomHeight = True
                Next
            End If
            Dim result = DirectCast(source.CloneNode(False), SheetData)
            For Each entry In rows
                Dim row = entry.Value
                row.RowIndex = CUInt(entry.Key + 1)
                row.Spans = Nothing
                Dim rowCells As List(Of (Column As Integer, Cell As Cell)) = Nothing
                If cells.TryGetValue(entry.Key, rowCells) Then
                    row.Append(rowCells.OrderBy(Function(item) item.Column).Select(Function(item) item.Cell))
                End If
                result.Append(row)
            Next
            If merges IsNot Nothing Then
                For Each merge In merges.Elements(Of MergeCell)()
                    Dim area = ReadStructuralRange(merge.Reference.Value)
                    minRow = System.Math.Min(minRow, area.First.Row)
                    maxRow = System.Math.Max(maxRow, area.Last.Row)
                    minColumn = System.Math.Min(minColumn, area.First.Column)
                    maxColumn = System.Math.Max(maxColumn, area.Last.Column)
                Next
            End If
            dimension = "A1"
            If maxRow >= 0 Then
                dimension = ExcelCell.LocalCellAddress(minRow, minColumn)
                If maxRow <> minRow OrElse maxColumn <> minColumn Then dimension &= ":" & ExcelCell.LocalCellAddress(maxRow, maxColumn)
            End If
            Return result
        End Function

    End Class

End Namespace
