Imports System.Globalization
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Shared ReadOnly EscapedText As New Regex("_x([0-9a-f]{4})_", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
        Private Shared ReadOnly TextToEscape As New Regex("_(?=[xX][0-9A-Fa-f]{4}_)|[\u0000-\u0008\u000B-\u001F\uFFFE\uFFFF]")
        Private ReadOnly _formulaRanges As New Dictionary(Of WorksheetPart, List(Of ExcelRange))

        Private Shared Sub ValidateCoordinates(rowIndex As Integer, columnIndex As Integer)
            If rowIndex < 0 OrElse rowIndex >= 1048576 Then Throw New ArgumentOutOfRangeException(NameOf(rowIndex))
            If columnIndex < 0 OrElse columnIndex >= 16384 Then Throw New ArgumentOutOfRangeException(NameOf(columnIndex))
        End Sub

        Private Shared Sub ValidateAddress(cell As ExcelCell)
            If cell Is Nothing Then Throw New ArgumentNullException(NameOf(cell))
            If Not cell.ValidateFullCellAddressInclSheetName() Then Throw New InvalidCellAddressException(cell)
        End Sub

        Private Shared Function FindRow(data As SheetData, rowIndex As Integer, create As Boolean) As Row
            Dim rowNumber = CUInt(rowIndex + 1)
            Dim currentNumber As UInteger
            Dim nextRow As Row = Nothing
            For Each row In data.Elements(Of Row)()
                currentNumber = If(row.RowIndex?.Value, currentNumber + 1UI)
                If currentNumber = rowNumber Then Return row
                If currentNumber > rowNumber Then
                    nextRow = row
                    Exit For
                End If
            Next
            If Not create Then Return Nothing
            Return data.InsertBefore(New Row With {.RowIndex = rowNumber}, nextRow)
        End Function

        Private Shared Function FindCell(row As Row, rowIndex As Integer, columnIndex As Integer, create As Boolean) As Cell
            If row Is Nothing Then Return Nothing
            Dim currentIndex = -1
            Dim nextCell As Cell = Nothing
            For Each cell In row.Elements(Of Cell)()
                currentIndex = If(cell.CellReference Is Nothing, currentIndex + 1,
                                  New ExcelCell(Nothing, cell.CellReference.Value, ExcelCell.ValueTypes.All).ColumnIndex)
                If currentIndex = columnIndex Then Return cell
                If currentIndex > columnIndex Then
                    nextCell = cell
                    Exit For
                End If
            Next
            If Not create Then Return Nothing
            Dim result As New Cell With {.CellReference = ExcelCell.LocalCellAddress(rowIndex, columnIndex)}
            If nextCell Is Nothing Then
                Return row.InsertBefore(result, row.GetFirstChild(Of ExtensionList)())
            End If
            Return row.InsertBefore(result, nextCell)
        End Function

        Private Function GetCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Cell
            ValidateCoordinates(rowIndex, columnIndex)
            Dim data = GetSheetData(GetWorksheetPart(sheetName))
            Return FindCell(FindRow(data, rowIndex, False), rowIndex, columnIndex, False)
        End Function

        Private Shared Function GetSheetData(part As WorksheetPart) As SheetData
            Dim data = part.Worksheet.GetFirstChild(Of SheetData)()
            If data Is Nothing Then Throw New IO.InvalidDataException("The worksheet is missing its sheet data.")
            Return data
        End Function

        Private Sub EditCell(sheetName As String, rowIndex As Integer, columnIndex As Integer, payload As Cell)
            ValidateCoordinates(rowIndex, columnIndex)
            Dim part = GetWorksheetPart(sheetName)
            Dim data = GetSheetData(part)
            Dim row = FindRow(data, rowIndex, False)
            Dim existing = FindCell(row, rowIndex, columnIndex, False)
            If existing Is Nothing AndAlso Not payload.HasChildren Then Return
            If IsGroupedFormulaCell(part, sheetName, rowIndex, columnIndex, existing?.CellFormula) Then
                Throw UnsupportedOperation("Editing cells in shared, array or data-table formula ranges")
            End If
            Dim target = existing
            If target Is Nothing Then
                row = If(row, FindRow(data, rowIndex, True))
                target = FindCell(row, rowIndex, columnIndex, True)
                UpdateDimension(part.Worksheet, sheetName, rowIndex, columnIndex)
            End If
            _calculationChainInvalid = _calculationChainInvalid OrElse target.CellFormula IsNot Nothing OrElse payload.CellFormula IsNot Nothing
            target.CellFormula = Nothing
            target.CellValue = Nothing
            target.InlineString = Nothing
            For Each child In payload.ChildElements.ToList()
                child.Remove()
                target.AddChild(child)
            Next
            target.DataType = payload.DataType
            _changedSheets.Add(part)
            RecalculationRequired = True
        End Sub

        Private Function IsGroupedFormulaCell(part As WorksheetPart, sheetName As String, rowIndex As Integer, columnIndex As Integer, formula As CellFormula) As Boolean
            If formula?.FormulaType IsNot Nothing AndAlso formula.FormulaType.Value <> CellFormulaValues.Normal Then Return True
            Dim ranges As List(Of ExcelRange) = Nothing
            If Not _formulaRanges.TryGetValue(part, ranges) Then
                ranges = GetSheetData(part).Descendants(Of CellFormula)().
                    Where(Function(item) item.Reference IsNot Nothing AndAlso item.FormulaType IsNot Nothing AndAlso item.FormulaType.Value <> CellFormulaValues.Normal).
                    Select(Function(item) New ExcelRange(sheetName, item.Reference.Value)).ToList()
                _formulaRanges.Add(part, ranges)
            End If
            Return ranges.Any(Function(range) rowIndex >= range.AddressStart.RowIndex AndAlso rowIndex <= range.AddressEnd.RowIndex AndAlso
                                               columnIndex >= range.AddressStart.ColumnIndex AndAlso columnIndex <= range.AddressEnd.ColumnIndex)
        End Function

        Private Shared Sub UpdateDimension(worksheet As Worksheet, sheetName As String, rowIndex As Integer, columnIndex As Integer)
            Dim dimension = worksheet.SheetDimension
            If dimension?.Reference Is Nothing Then Return
            Dim oldBounds As New ExcelRange(sheetName, dimension.Reference.Value)
            Dim first = ExcelCell.LocalCellAddress(System.Math.Min(rowIndex, oldBounds.AddressStart.RowIndex), System.Math.Min(columnIndex, oldBounds.AddressStart.ColumnIndex))
            Dim last = ExcelCell.LocalCellAddress(System.Math.Max(rowIndex, oldBounds.AddressEnd.RowIndex), System.Math.Max(columnIndex, oldBounds.AddressEnd.ColumnIndex))
            If first <> oldBounds.AddressStart.Address OrElse last <> oldBounds.AddressEnd.Address Then
                dimension.Reference = first & ":" & last
            End If
        End Sub

        ''' <inheritdoc/>
        Public Overrides Function HasMissingCachedFormulaResult(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            Return IsMissingFormulaResult(GetCell(sheetName, rowIndex, columnIndex))
        End Function

        Private Shared Function IsMissingFormulaResult(cell As Cell) As Boolean
            If cell?.CellFormula Is Nothing Then Return False
            Dim value = cell.CellValue?.Text
            Return value Is Nothing OrElse (value.Length = 0 AndAlso If(cell.DataType?.Value, CellValues.Number) <> CellValues.String)
        End Function

        ''' <inheritdoc/>
        <Obsolete("Use the overload with an explicit FormulaCalculationMode argument.", True)>
        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Public Overrides Function LookupCellValueAsObject(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Object
            Return ReadCellValue(sheetName, rowIndex, columnIndex, GetCell(sheetName, rowIndex, columnIndex))
        End Function

        Private Function ReadCellValue(sheetName As String, rowIndex As Integer, columnIndex As Integer, cell As Cell) As Object
            If cell Is Nothing Then Return Nothing
            Dim value = cell.CellValue?.Text
            Dim dataType = If(cell.DataType?.Value, CellValues.Number)
            If IsMissingFormulaResult(cell) Then
                Throw New MissingCalculatedCellValueException(FilePath, sheetName, rowIndex, columnIndex, cell.CellFormula.Text)
            End If
            If dataType = CellValues.InlineString Then Return DecodeText(cell.InlineString)
            If value Is Nothing Then Return Nothing
            If dataType = CellValues.SharedString Then
                Dim stringIndex = Integer.Parse(value, CultureInfo.InvariantCulture)
                Dim item = WorkbookPart.SharedStringTablePart?.SharedStringTable?.Elements(Of SharedStringItem)().ElementAtOrDefault(stringIndex)
                If item Is Nothing Then Throw New FormatException("The cell refers to a missing shared string.")
                Return DecodeText(item)
            End If
            If dataType = CellValues.Boolean Then Return Xml.XmlConvert.ToBoolean(value)
            If dataType = CellValues.Number Then Return If(value.Length = 0, Nothing, CType(Double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture), Object))
            If dataType = CellValues.Date Then Return Xml.XmlConvert.ToDateTime(value, Xml.XmlDateTimeSerializationMode.RoundtripKind)
            Return DecodeText(value)
        End Function

        Private Shared Function DecodeText(container As OpenXmlCompositeElement) As String
            If container Is Nothing Then Return Nothing
            ' Phonetic annotations are separate from the displayed value.
            Return DecodeText(String.Concat(container.ChildElements.Select(
                Function(child) If(TypeOf child Is Text OrElse TypeOf child Is Run, child.InnerText, Nothing))))
        End Function

        Private Shared Function EncodeText(value As String) As String
            Return TextToEscape.Replace(value, Function(m) "_x" & (AscW(m.Value(0)) And &HFFFF).ToString("X4", CultureInfo.InvariantCulture) & "_")
        End Function

        Private Shared Function DecodeText(value As String) As String
            Return EscapedText.Replace(value, Function(m) ChrW(Integer.Parse(m.Groups(1).Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString())
        End Function

        ''' <inheritdoc/>
        <Obsolete("Use the overload with an explicit FormulaCalculationMode argument.", True)>
        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Public Overrides Function LookupCellValue(Of T)(cell As ExcelCell) As T
            ValidateAddress(cell)
            Return ReadLegacyCellValue(Of T)(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <inheritdoc/>
        <Obsolete("Use the overload with an explicit FormulaCalculationMode argument.", True)>
        <System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)>
        Public Overrides Function LookupCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T
            Dim value = ReadLegacyCellValueAsObject(sheetName, rowIndex, columnIndex)
            Return ConvertValue(Of T)(value)
        End Function

        Private Shared Function ConvertValue(Of T)(value As Object) As T
            If value Is Nothing Then Return Nothing
            If TypeOf value Is T Then Return DirectCast(value, T)
            Dim targetType = If(Nullable.GetUnderlyingType(GetType(T)), GetType(T))
            Try
                Return CType(Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture), T)
            Catch ex As Exception When TypeOf ex Is InvalidCastException OrElse TypeOf ex Is OverflowException OrElse TypeOf ex Is FormatException
                Throw New FormatException("The cell value cannot be converted to " & targetType.Name & ".", ex)
            End Try
        End Function

        ''' <inheritdoc/>
        Public Overrides Function TryLookupCellValue(Of T As Structure)(cell As ExcelCell) As T?
            ValidateAddress(cell)
            Return TryLookupCellValue(Of T)(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <inheritdoc/>
        Public Overrides Function TryLookupCellValue(Of T As Structure)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T?
            Dim value = ReadLegacyCellValueAsObject(sheetName, rowIndex, columnIndex)
            If value Is Nothing OrElse (TypeOf value Is String AndAlso String.IsNullOrWhiteSpace(DirectCast(value, String))) Then Return Nothing
            Try
                Return ConvertValue(Of T)(value)
            Catch ex As FormatException
                Return Nothing
            End Try
        End Function

        ''' <inheritdoc/>
        Public Overrides Function LookupCellFormula(cell As ExcelCell) As String
            ValidateAddress(cell)
            Return LookupCellFormula(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <inheritdoc/>
        ''' <remarks>Returns the stored formula text. A shared-formula follower may have an empty formula element.</remarks>
        Public Overrides Function LookupCellFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Return GetCell(sheetName, rowIndex, columnIndex)?.CellFormula?.Text
        End Function

        ''' <inheritdoc/>
        Public Overrides Function LookupCellErrorValue(sheetName As String, rowIndex As Integer, columnIndex As Integer) As String
            Dim cell = GetCell(sheetName, rowIndex, columnIndex)
            If cell?.DataType?.Value = CellValues.Error Then Return cell.CellValue?.Text
            Return Nothing
        End Function

        ''' <inheritdoc/>
        Public Overrides Function IsEmptyCell(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            Dim cell = GetCell(sheetName, rowIndex, columnIndex)
            Return cell Is Nothing OrElse (cell.CellFormula Is Nothing AndAlso cell.CellValue Is Nothing AndAlso cell.InlineString Is Nothing)
        End Function

        ''' <inheritdoc/>
        Public Overrides Sub WriteCellValue(Of T)(cell As ExcelCell, value As T)
            ValidateAddress(cell)
            WriteCellValue(cell.SheetName, cell.RowIndex, cell.ColumnIndex, value)
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Supports strings, booleans, finite numbers and blank values. Existing cell formatting is retained.</remarks>
        Public Overrides Sub WriteCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer, value As T)
            Dim boxed As Object = value
            Dim payload As New Cell()
            If TypeOf boxed Is String Then
                payload.DataType = CellValues.InlineString
                Dim encoded = EncodeText(DirectCast(boxed, String))
                payload.InlineString = New InlineString(New Text(encoded) With {.Space = SpaceProcessingModeValues.Preserve})
            ElseIf TypeOf boxed Is Boolean Then
                payload.DataType = CellValues.Boolean
                payload.CellValue = New CellValue(If(CBool(boxed), "1", "0"))
            ElseIf boxed IsNot Nothing Then
                Select Case Convert.GetTypeCode(boxed)
                    Case TypeCode.Single, TypeCode.Double
                        Dim number = Convert.ToDouble(boxed, CultureInfo.InvariantCulture)
                        If Double.IsNaN(number) OrElse Double.IsInfinity(number) Then Throw New ArgumentOutOfRangeException(NameOf(value), "An XLSX number must be finite.")
                        payload.CellValue = New CellValue(number.ToString("R", CultureInfo.InvariantCulture))
                    Case TypeCode.SByte, TypeCode.Byte, TypeCode.Int16, TypeCode.UInt16, TypeCode.Int32, TypeCode.UInt32, TypeCode.Int64, TypeCode.UInt64, TypeCode.Decimal
                        payload.CellValue = New CellValue(Convert.ToString(boxed, CultureInfo.InvariantCulture))
                    Case Else
                        Throw UnsupportedOperation("Writing values of type " & boxed.GetType().Name)
                End Select
                If boxed.GetType().IsEnum Then Throw UnsupportedOperation("Writing enum values")
                payload.DataType = CellValues.Number
            End If
            EditCell(sheetName, rowIndex, columnIndex, payload)
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Evaluates the formula before editing the cell when immediate calculation is requested.</remarks>
        Public Overrides Sub WriteCellFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer, formula As String, immediatelyCalculateCellValue As Boolean)
            If immediatelyCalculateCellValue Then EnsureCalculationEnabled()
            If String.IsNullOrWhiteSpace(formula) Then Throw New ArgumentException("A formula is required.", NameOf(formula))
            Dim expression = If(formula(0) = "="c, formula.Substring(1), formula)
            If String.IsNullOrWhiteSpace(expression) Then Throw New ArgumentException("A formula is required.", NameOf(formula))
            Dim replacement As New Cell With {.CellFormula = New CellFormula(expression)}
            If Not immediatelyCalculateCellValue Then
                EditCell(sheetName, rowIndex, columnIndex, replacement)
                Return
            End If

            Dim existing = GetCell(sheetName, rowIndex, columnIndex)
            If IsGroupedFormulaCell(GetWorksheetPart(sheetName), sheetName, rowIndex, columnIndex, existing?.CellFormula) Then
                Throw UnsupportedOperation("Editing cells in shared, array or data-table formula ranges")
            End If
            Dim calculation As New FormulaCalculation(Me)
            Dim result = calculation.EvaluateReplacement(sheetName, rowIndex, columnIndex, replacement)
            EditCell(sheetName, rowIndex, columnIndex, replacement)
            calculation.Commit()
            ThrowIfFormulaError(sheetName, rowIndex, columnIndex, result)
        End Sub

        ''' <inheritdoc/>
        Public Overrides Function LookupRowIndex(cell As ExcelCell) As Integer
            ValidateAddress(cell)
            Return cell.RowIndex
        End Function

        ''' <inheritdoc/>
        Public Overrides Function LookupColumnIndex(cell As ExcelCell) As Integer
            ValidateAddress(cell)
            Return cell.ColumnIndex
        End Function

    End Class

End Namespace
