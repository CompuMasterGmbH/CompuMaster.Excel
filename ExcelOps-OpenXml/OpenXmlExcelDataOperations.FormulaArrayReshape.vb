Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateReshape(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                             formulaCell As ExcelCell) As Object
                Dim minimum = If(name = "EXPAND" OrElse name.StartsWith("WRAP", StringComparison.Ordinal), 2, 1)
                If arguments.Count < minimum OrElse arguments.Count > If(name = "EXPAND", 4, 3) Then Return New FormulaError("#VALUE!")
                Dim input = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                If name = "TRIMRANGE" Then Return EvaluateTrimRange(input, arguments, formulaCell)
                Dim matrix = AsMatrix(input)
                Select Case name
                    Case "TOCOL", "TOROW"
                        Dim mode = ReadReshapeNumber(arguments, 1, formulaCell, 0.0)
                        If TypeOf mode Is FormulaError Then Return mode
                        Dim ignore = Math.Truncate(DirectCast(mode, Double))
                        If ignore < 0.0 OrElse ignore > 3.0 Then Return New FormulaError("#VALUE!")
                        Dim rawOrder = If(arguments.Count < 3, Nothing, Scalar(EvaluateExpression(arguments(2), formulaCell)))
                        Dim order = LogicalOptionValue(rawOrder)
                        If TypeOf order Is FormulaError Then Return order
                        Return FlattenArray(matrix, name = "TOCOL", CInt(ignore), DirectCast(order, Boolean))
                    Case "WRAPCOLS", "WRAPROWS"
                        If matrix.Rows > 1 AndAlso matrix.Columns > 1 Then Return New FormulaError("#VALUE!")
                        Dim count = ReadReshapeNumber(arguments, 1, formulaCell, 0.0)
                        If TypeOf count Is FormulaError Then Return count
                        Dim width = Math.Truncate(DirectCast(count, Double))
                        If width < 1.0 Then Return New FormulaError("#NUM!")
                        Dim padding = ReadReshapePadding(arguments, 2, formulaCell)
                        If IsMultiValue(padding) Then Return New FormulaError("#VALUE!")
                        Return WrapArray(matrix, name = "WRAPCOLS", CInt(Math.Min(width, matrix.Values.Length)), padding)
                    Case Else
                        Return ExpandArray(matrix, arguments, formulaCell)
                End Select
            End Function

            Private Function ReadReshapeNumber(arguments As IReadOnlyList(Of FormulaExpression), index As Integer,
                                               formulaCell As ExcelCell, defaultValue As Double) As Object
                If index >= arguments.Count Then Return defaultValue
                Dim input = EvaluateExpression(arguments(index), formulaCell)
                If IsMultiValue(input) Then Return New FormulaError("#VALUE!")
                Return ReadFinancialNumber(If(Scalar(input), defaultValue))
            End Function

            Private Function ReadReshapePadding(arguments As IReadOnlyList(Of FormulaExpression), index As Integer, formulaCell As ExcelCell) As Object
                If index >= arguments.Count Then Return New FormulaError("#N/A")
                Dim padding = EvaluateExpression(arguments(index), formulaCell)
                If IsMultiValue(padding) Then Return padding
                Return If(Scalar(padding), New FormulaError("#N/A"))
            End Function

            Private Shared Function FlattenArray(matrix As FormulaMatrix, vertical As Boolean, ignore As Integer, scanColumns As Boolean) As Object
                Dim values As New List(Of Object)
                Dim major = If(scanColumns, matrix.Columns, matrix.Rows)
                Dim minor = If(scanColumns, matrix.Rows, matrix.Columns)
                For outer = 0 To major - 1
                    For inner = 0 To minor - 1
                        Dim value = If(scanColumns, matrix.At(inner, outer), matrix.At(outer, inner))
                        If (ignore And 1) <> 0 AndAlso value Is Nothing Then Continue For
                        If (ignore And 2) <> 0 AndAlso TypeOf value Is FormulaError Then Continue For
                        values.Add(If(value, 0.0))
                    Next
                Next
                If values.Count = 0 Then Return New FormulaError("#CALC!")
                If values.Count = 1 Then Return values(0)
                Dim rows = If(vertical, values.Count, 1)
                Dim columns = If(vertical, 1, values.Count)
                If rows > RowCount OrElse columns > ColumnCount Then Return New FormulaError("#NUM!")
                Return New FormulaMatrix(rows, columns, values.ToArray())
            End Function

            Private Shared Function WrapArray(matrix As FormulaMatrix, columnwise As Boolean, count As Integer, padding As Object) As Object
                Dim other = (matrix.Values.Length + count - 1) \ count
                Dim rows = If(columnwise, count, other)
                Dim columns = If(columnwise, other, count)
                If rows > RowCount OrElse columns > ColumnCount Then Return New FormulaError("#NUM!")
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim index = If(columnwise, column * rows + row, row * columns + column)
                        values(row * columns + column) = If(index < matrix.Values.Length, If(matrix.Values(index), 0.0), padding)
                    Next
                Next
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Function ExpandArray(matrix As FormulaMatrix, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Dim rawRows = ReadReshapeNumber(arguments, 1, formulaCell, matrix.Rows)
                If TypeOf rawRows Is FormulaError Then Return rawRows
                Dim rawColumns = ReadReshapeNumber(arguments, 2, formulaCell, matrix.Columns)
                If TypeOf rawColumns Is FormulaError Then Return rawColumns
                Dim rows = Math.Truncate(DirectCast(rawRows, Double))
                Dim columns = Math.Truncate(DirectCast(rawColumns, Double))
                If rows < matrix.Rows OrElse columns < matrix.Columns Then Return New FormulaError("#VALUE!")
                If rows > RowCount OrElse columns > ColumnCount Then Return New FormulaError("#NUM!")
                Dim padding = ReadReshapePadding(arguments, 3, formulaCell)
                If IsMultiValue(padding) Then Return New FormulaError("#VALUE!")
                Dim values = AllocateMatrixValues(CInt(rows), CInt(columns))
                For row = 0 To CInt(rows) - 1
                    For column = 0 To CInt(columns) - 1
                        values(row * CInt(columns) + column) = If(row < matrix.Rows AndAlso column < matrix.Columns,
                            If(matrix.At(row, column), 0.0), padding)
                    Next
                Next
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(CInt(rows), CInt(columns), values)
            End Function

            Private Function EvaluateTrimRange(input As Object, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Dim vertical = ReadReshapeNumber(arguments, 1, formulaCell, 3.0)
                If TypeOf vertical Is FormulaError Then Return vertical
                Dim horizontal = ReadReshapeNumber(arguments, 2, formulaCell, 3.0)
                If TypeOf horizontal Is FormulaError Then Return horizontal
                Dim rows = DirectCast(vertical, Double)
                Dim columns = DirectCast(horizontal, Double)
                If rows < 0.0 OrElse rows > 3.0 OrElse rows <> Math.Truncate(rows) OrElse
                    columns < 0.0 OrElse columns > 3.0 OrElse columns <> Math.Truncate(columns) Then Return New FormulaError("#VALUE!")
                If rows = 0.0 AndAlso columns = 0.0 Then Return input
                Dim area = TryCast(input, FormulaArea)
                If area IsNot Nothing Then Return TrimArea(area, CInt(rows), CInt(columns))
                Return input
            End Function

            Private Function TrimArea(area As FormulaArea, vertical As Integer, horizontal As Integer) As Object
                Dim source = ResolveReference(area)
                Dim worksheet As WorksheetSnapshot = Nothing
                If Not source._worksheets.TryGetValue(area.Sheet, worksheet) Then Return New FormulaError("#REF!")
                RecordPrecedent(New FormulaPrecedent(New FormulaLocation(source._owner, area.Sheet, area.FirstRow, area.FirstColumn),
                    area.LastRow, area.LastColumn))

                RecalculateReferencedSpills(source, worksheet, area)
                Dim stagedSpills = worksheet.SpillResults.Keys.ToArray()
                Dim firstRow = area.LastRow + 1
                Dim firstColumn = area.LastColumn + 1
                Dim lastRow = area.FirstRow - 1
                Dim lastColumn = area.FirstColumn - 1
                For Each item In worksheet.Cells
                    Dim row = item.Key.RowIndex
                    Dim column = item.Key.ColumnIndex
                    If row < area.FirstRow OrElse row > area.LastRow OrElse column < area.FirstColumn OrElse column > area.LastColumn Then Continue For
                    Dim cell = item.Value
                    If cell.CellFormula Is Nothing AndAlso cell.CellValue Is Nothing AndAlso cell.InlineString Is Nothing Then Continue For
                    If stagedSpills.Any(Function(spill) spill.Contains(row, column) AndAlso
                        (spill.First.Row <> row OrElse spill.First.Column <> column) AndAlso
                        Not worksheet.SpillChildren.ContainsKey(item.Key)) Then Continue For
                    firstRow = Math.Min(firstRow, row)
                    firstColumn = Math.Min(firstColumn, column)
                    lastRow = Math.Max(lastRow, row)
                    lastColumn = Math.Max(lastColumn, column)
                Next
                For Each item In worksheet.SpillChildren.Keys
                    Dim row = item.RowIndex
                    Dim column = item.ColumnIndex
                    If row < area.FirstRow OrElse row > area.LastRow OrElse column < area.FirstColumn OrElse column > area.LastColumn Then Continue For
                    firstRow = Math.Min(firstRow, row)
                    firstColumn = Math.Min(firstColumn, column)
                    lastRow = Math.Max(lastRow, row)
                    lastColumn = Math.Max(lastColumn, column)
                Next
                If firstRow > lastRow Then
                    Dim multipleRows = area.FirstRow <> area.LastRow
                    Dim multipleColumns = area.FirstColumn <> area.LastColumn
                    If (vertical <> 0 AndAlso multipleRows) OrElse (horizontal <> 0 AndAlso multipleColumns) OrElse
                        (Not multipleRows AndAlso Not multipleColumns) Then Return New FormulaError("#REF!")
                    Return area
                End If
                Return SliceArea(area, If((vertical And 1) = 0, 0, firstRow - area.FirstRow),
                    If((horizontal And 1) = 0, 0, firstColumn - area.FirstColumn),
                    If((vertical And 2) = 0, area.LastRow, lastRow) - area.FirstRow,
                    If((horizontal And 2) = 0, area.LastColumn, lastColumn) - area.FirstColumn)
            End Function

            Private Function EvaluatePercentOf(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                Dim first = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf first Is FormulaError Then Return first
                Dim second = EvaluateExpression(arguments(1), formulaCell)
                If TypeOf second Is FormulaError Then Return second
                Dim part = PercentOfSum(first)
                If TypeOf part Is FormulaError Then Return part
                Dim whole = PercentOfSum(second)
                If TypeOf whole Is FormulaError Then Return whole
                Dim divisor = DirectCast(whole, Double)
                If divisor = 0.0 Then Return New FormulaError("#DIV/0!")
                Return FiniteNumber(DirectCast(part, Double) / divisor)
            End Function

            Private Function PercentOfSum(value As Object) As Object
                If Not TypeOf value Is FormulaArea AndAlso Not TypeOf value Is FormulaMatrix AndAlso Not TypeOf value Is FormulaReferenceUnion Then
                    Return ReadFinancialNumber(value)
                End If
                Dim sum = 0.0
                For Each item In ReferenceArgumentValues(value)
                    If TypeOf item Is FormulaError Then Return New FormulaError("#NUM!")
                    If TypeOf item Is Double Then sum += DirectCast(item, Double)
                Next
                Return FiniteNumber(sum)
            End Function

        End Class

    End Class

End Namespace
