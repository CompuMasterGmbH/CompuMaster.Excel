Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports SpreadsheetDrawing = DocumentFormat.OpenXml.Drawing.Spreadsheet

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

            Friend Sub PlanTableRows(table As StructuralTable)
                Dim low = If(Vertical, table.First.Row, table.First.Column)
                Dim high = If(Vertical, table.Last.Row, table.Last.Column)
                Dim otherLow = If(Vertical, table.First.Column, table.First.Row)
                Dim otherHigh = If(Vertical, table.Last.Column, table.Last.Row)
                If high < Start OrElse otherLow >= _otherStart + _otherCount OrElse otherHigh < _otherStart Then Return
                If Not Vertical OrElse Not Includes(otherLow) OrElse Not Includes(otherHigh) OrElse
                   (Not Inserting AndAlso Start + Count > low) Then
                    Throw New NotSupportedException("This structural edit would change part of an Excel table in an unsupported way.")
                End If
                table.TargetLastRow = MoveIndex(high)
                If Start <= low Then
                    table.TargetFirstRow = MoveIndex(low)
                Else
                    table.AddsDataRows = True
                    _expandedTables.Add(table)
                End If
            End Sub

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
            "fileVersion", "workbookPr", "workbookProtection", "bookViews", "sheets", "definedNames", "calcPr"
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

        Private Sub ValidateWorkbookStructure()
            For Each child In WorkbookPart.Workbook.ChildElements
                If Not IsReferenceFreeWorkbookMetadata(child) AndAlso
                   (TypeOf child Is OpenXmlUnknownElement OrElse Not StandardWorkbookElements.Contains(child.LocalName)) Then
                    Throw UnsupportedOperation("Updating references in a workbook with " & child.LocalName)
                End If
            Next
            For Each relationship In WorkbookPart.Parts
                Dim part = relationship.OpenXmlPart
                If Not (TypeOf part Is WorksheetPart OrElse TypeOf part Is WorkbookStylesPart OrElse
                        TypeOf part Is SharedStringTablePart OrElse TypeOf part Is ThemePart OrElse TypeOf part Is CalculationChainPart) Then
                    Throw UnsupportedOperation("Updating references in a workbook with " & part.GetType().Name)
                End If
            Next
        End Sub

        Private Sub ValidateWorksheetStructure(part As WorksheetPart, shift As StructuralShift)
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
                If Not shift.Vertical AndAlso shift.WholeAxis AndAlso worksheet.GetFirstChild(Of Columns)() IsNot Nothing Then
                    Throw UnsupportedOperation("Shifting whole columns with explicit column formatting")
                End If
            End If

            For Each relationship In part.Parts
                If TypeOf relationship.OpenXmlPart Is SpreadsheetPrinterSettingsPart OrElse
                   TypeOf relationship.OpenXmlPart Is TableDefinitionPart Then Continue For
                Dim drawings = TryCast(relationship.OpenXmlPart, DrawingsPart)
                If drawings?.WorksheetDrawing Is Nothing OrElse Not drawings.Parts.All(Function(item) TypeOf item.OpenXmlPart Is ImagePart) OrElse
                   Not drawings.WorksheetDrawing.ChildElements.All(
                       Function(anchor) anchor.GetFirstChild(Of SpreadsheetDrawing.Picture)() IsNot Nothing AndAlso
                           (shift Is Nothing OrElse TypeOf anchor Is SpreadsheetDrawing.AbsoluteAnchor)) Then
                    Throw UnsupportedOperation("Updating references in a worksheet with " & relationship.OpenXmlPart.GetType().Name)
                End If
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
                    If cell.CellMetaIndex IsNot Nothing OrElse cell.ValueMetaIndex IsNot Nothing OrElse
                       (formulaType <> CellFormulaValues.Shared AndAlso
                        (formulaType <> CellFormulaValues.Normal OrElse formula?.Reference IsNot Nothing)) Then
                        Throw UnsupportedOperation("Updating references in a workbook with array, data-table or metadata-backed cells")
                    End If
                Next
            Next
        End Sub

        Private Sub ApplyStructuralEdit(sheetName As String, shift As StructuralShift, newSheetName As String)
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
            ValidateWorkbookStructure()
            Dim tables = GetStructuralTables(sheets)
            Dim tableNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each table In tables
                If table.Source.Name IsNot Nothing Then tableNames.Add(table.Source.Name.Value)
                If table.Source.DisplayName IsNot Nothing Then tableNames.Add(table.Source.DisplayName.Value)
            Next
            Dim tableFills As New List(Of (Row As Integer, Column As Integer, Formula As String))
            Dim tableChanges = StageTableChanges(tables, editedPart, originalName, shift, newSheetName, tableNames, tableFills)
            Dim filledCells As New HashSet(Of Cell)
            Dim formulaChanges As New List(Of (Part As WorksheetPart, Cell As Cell, Text As String))
            Dim revisedSheetData As New Dictionary(Of WorksheetPart, SheetData)
            Dim originalMerges As MergeCells = Nothing
            Dim shiftedMerges As MergeCells = Nothing
            Dim shiftedData As SheetData = Nothing
            Dim dimension As String = Nothing
            For Each sheet In sheets
                Dim part = DirectCast(WorkbookPart.GetPartById(sheet.Id.Value), WorksheetPart)
                Dim sheetShift = If(part Is editedPart, shift, Nothing)
                ValidateWorksheetStructure(part, sheetShift)
                Dim originalData = GetSheetData(part)
                Dim data = ExpandSharedFormulas(originalData, sheet.Name.Value, tableNames)
                If sheetShift IsNot Nothing Then
                    originalMerges = part.Worksheet.GetFirstChild(Of MergeCells)()
                    shiftedMerges = ShiftMergedCells(originalMerges, sheetShift)
                    shiftedData = ShiftSheetData(data, sheetShift, shiftedMerges, tableFills, filledCells, dimension)
                    data = shiftedData
                End If
                If data IsNot originalData Then revisedSheetData.Add(part, data)
                For Each cell In data.Descendants(Of Cell)()
                    Dim formula = cell.CellFormula
                    If formula Is Nothing OrElse filledCells.Contains(cell) Then Continue For
                    Dim text = RewriteFormula(formula.Text, sheet.Name.Value, originalName, shift, newSheetName, isDefinedName:=False, tableNames:=tableNames)
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
                    Dim text = RewriteFormula(definition.Text, scopeSheet, originalName, shift, newSheetName, isDefinedName:=True, tableNames:=tableNames)
                    If text <> definition.Text Then nameChanges.Add((definition, text))
                Next
            End If

            For Each revision In revisedSheetData
                revision.Key.Worksheet.ReplaceChild(revision.Value, GetSheetData(revision.Key))
                _formulaRanges.Remove(revision.Key)
                _changedSheets.Add(revision.Key)
            Next
            If shiftedData IsNot Nothing Then
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
            If newSheetName IsNot Nothing Then editedSheet.Name = newSheetName

            InvalidateWorkbookFormulaResults()
        End Sub

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
                                               tableFills As List(Of (Row As Integer, Column As Integer, Formula As String)),
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

                    Dim cell = DirectCast(sourceCell.CloneNode(True), Cell)
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
                Dim cell As New Cell With {
                    .CellReference = ExcelCell.LocalCellAddress(fill.Row, fill.Column),
                    .CellFormula = New CellFormula(fill.Formula)
                }
                Dim rowCells As List(Of (Column As Integer, Cell As Cell)) = Nothing
                If Not cells.TryGetValue(fill.Row, rowCells) Then
                    rowCells = New List(Of (Column As Integer, Cell As Cell))
                    cells.Add(fill.Row, rowCells)
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
