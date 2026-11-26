Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateClassicLookup(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim needle = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf needle Is FormulaError Then Return needle
                Dim input = EvaluateExpression(arguments(1), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim table = AsMatrix(input)
                Dim vertical = table.Rows >= table.Columns
                Dim count = If(vertical, table.Rows, table.Columns)
                Dim keys = Enumerable.Range(0, count).Select(Function(index) If(vertical, table.At(index, 0), table.At(0, index))).ToArray()

                Dim selectValue As Func(Of Integer, Object)
                If arguments.Count = 2 Then
                    selectValue = Function(index) If(vertical, table.At(index, table.Columns - 1), table.At(table.Rows - 1, index))
                Else
                    Dim result = EvaluateExpression(arguments(2), formulaCell)
                    If TypeOf result Is FormulaError Then Return result
                    Dim area = TryCast(result, FormulaArea)
                    If area IsNot Nothing Then
                        selectValue = Function(index) ClassicLookupAreaValue(area, index)
                    Else
                        Dim values = AsMatrix(result)
                        Dim byRow = values.Rows >= values.Columns
                        Dim length = If(byRow, values.Rows, values.Columns)
                        selectValue = Function(index) If(index >= length, CObj(New FormulaError("#N/A")),
                            If(byRow, values.At(index, 0), values.At(0, index)))
                    End If
                End If

                Dim positions = Enumerable.Range(0, keys.Length).Where(Function(index) keys(index) IsNot Nothing AndAlso
                    Not TypeOf keys(index) Is FormulaError).GroupBy(Function(index) keys(index).GetType()).
                    ToDictionary(Function(group) group.Key, Function(group) group.ToArray())
                If Not IsMultiValue(needle) Then Return ClassicLookupMatch(Scalar(needle), keys, positions, selectValue)
                Dim queries = AsMatrix(needle)
                Dim results = queries.Values.Select(Function(value) ClassicLookupMatch(value, keys, positions, selectValue)).ToArray()
                Return New FormulaMatrix(queries.Rows, queries.Columns, results)
            End Function

            Private Function ClassicLookupAreaValue(area As FormulaArea, index As Integer) As Object
                Dim rows = area.LastRow - area.FirstRow + 1
                Dim columns = area.LastColumn - area.FirstColumn + 1
                If rows = 1 AndAlso columns = 1 AndAlso index > 0 Then Return 0.0
                Dim vertical = rows >= columns
                Dim row = If(vertical, index, 0)
                Dim column = If(vertical, 0, index)
                If area.FirstRow + row >= RowCount OrElse area.FirstColumn + column >= ColumnCount Then Return New FormulaError("#REF!")
                Return If(Scalar(SliceArea(area, row, column, row, column)), 0.0)
            End Function

            Private Shared Function ClassicLookupMatch(needle As Object, keys As Object(), positions As Dictionary(Of Type, Integer()),
                                                       selectValue As Func(Of Integer, Object)) As Object
                If TypeOf needle Is FormulaError Then Return needle
                needle = If(needle, 0.0)
                Dim candidates As Integer() = Nothing
                If Not positions.TryGetValue(needle.GetType(), candidates) Then Return New FormulaError("#N/A")
                Dim first = 0
                Dim last = candidates.Length
                While first < last
                    Dim middle = first + (last - first) \ 2
                    If EvaluateComparison("<=", keys(candidates(middle)), needle) Then
                        first = middle + 1
                    Else
                        last = middle
                    End If
                End While
                If first = 0 Then Return New FormulaError("#N/A")
                Return If(selectValue(candidates(first - 1)), 0.0)
            End Function

            Private Function EvaluateLookupFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                Select Case name
                    Case "INDEX"
                        Return EvaluateIndex(arguments, formulaCell)
                    Case "MATCH"
                        Return EvaluateMatch(arguments, formulaCell)
                    Case "VLOOKUP"
                        Return EvaluateLegacyLookup(arguments, formulaCell, horizontal:=False)
                    Case "HLOOKUP"
                        Return EvaluateLegacyLookup(arguments, formulaCell, horizontal:=True)
                    Case "XMATCH", "_XLFN.XMATCH"
                        Return EvaluateXMatch(arguments, formulaCell)
                    Case Else
                        Return EvaluateXLookup(arguments, formulaCell)
                End Select
            End Function

            Private Function EvaluateIndex(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim input = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim area = TryCast(input, FormulaArea)
                Dim matrix = TryCast(input, FormulaMatrix)
                If area Is Nothing AndAlso matrix Is Nothing Then Return New FormulaError("#VALUE!")
                Dim rows = If(area Is Nothing, matrix.Rows, area.LastRow - area.FirstRow + 1)
                Dim columns = If(area Is Nothing, matrix.Columns, area.LastColumn - area.FirstColumn + 1)

                Dim row = ReadIndex(arguments(1), formulaCell)
                If TypeOf row Is FormulaError Then Return row
                Dim rowNumber = DirectCast(row, Double)
                Dim columnNumber = 1.0
                If arguments.Count = 3 Then
                    Dim column = ReadIndex(arguments(2), formulaCell)
                    If TypeOf column Is FormulaError Then Return column
                    columnNumber = DirectCast(column, Double)
                ElseIf rows = 1 Then
                    columnNumber = rowNumber
                    rowNumber = 1.0
                ElseIf columns <> 1 Then
                    columnNumber = 0.0
                End If

                If rowNumber < 0 OrElse rowNumber > rows OrElse columnNumber < 0 OrElse columnNumber > columns Then Return New FormulaError("#REF!")
                Dim firstRow = CInt(Math.Max(0.0, rowNumber - 1.0))
                Dim firstColumn = CInt(Math.Max(0.0, columnNumber - 1.0))
                Dim lastRow = If(rowNumber = 0, rows - 1, firstRow)
                Dim lastColumn = If(columnNumber = 0, columns - 1, firstColumn)
                If area IsNot Nothing Then Return SliceArea(area, firstRow, firstColumn, lastRow, lastColumn)
                Return SliceMatrix(matrix, firstRow, firstColumn, lastRow, lastColumn)
            End Function

            Private Function EvaluateMatch(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim needle = Scalar(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf needle Is FormulaError Then Return needle
                Dim input = EvaluateExpression(arguments(1), formulaCell)
                If TypeOf input Is FormulaError Then Return input

                Dim mode = 1.0
                If arguments.Count = 3 Then
                    Dim optionValue = ReadIndex(arguments(2), formulaCell)
                    If TypeOf optionValue Is FormulaError Then Return optionValue
                    mode = DirectCast(optionValue, Double)
                End If
                If Math.Abs(mode) > 1.0 Then Return New FormulaError("#N/A")
                Dim matchMode = If(mode = 0, If(TypeOf needle Is String, 2, 0), -CInt(mode))
                Dim area = TryCast(input, FormulaArea)
                If area IsNot Nothing AndAlso area.FirstRow <> area.LastRow AndAlso area.FirstColumn <> area.LastColumn Then Return New FormulaError("#N/A")
                Dim index As Integer
                If area IsNot Nothing AndAlso Math.Max(area.LastRow - area.FirstRow + 1,
                                                       area.LastColumn - area.FirstColumn + 1) > MaxMaterializedReferenceCells Then
                    index = MatchSparseVector(needle, area, matchMode, preferLastEqual:=mode <> 0)
                Else
                    Dim keys = AsMatrix(input)
                    If keys.Rows > 1 AndAlso keys.Columns > 1 Then Return New FormulaError("#N/A")
                    index = FindLookupIndex(needle, keys.Values, matchMode, reverse:=False, preferLastEqual:=mode <> 0)
                End If
                Return If(index < 0, CType(New FormulaError("#N/A"), Object), CDbl(index + 1))
            End Function

            Private Function MatchSparseVector(needle As Object, area As FormulaArea, mode As Integer, preferLastEqual As Boolean) As Integer
                Dim reader = BindSparseRange(area)
                Dim columns = area.LastColumn - area.FirstColumn + 1
                Dim offsets As New HashSet(Of Long)
                reader.AddStoredPositions(offsets, columns)
                Dim sorted = offsets.OrderBy(Function(offset) offset).ToArray()
                Dim values = sorted.Select(Function(offset) reader.Read(CInt(offset \ columns), CInt(offset Mod columns))).ToArray()
                Dim selected = FindLookupIndex(needle, values, mode, reverse:=False, preferLastEqual:=preferLastEqual)
                Return If(selected < 0, -1, CInt(sorted(selected)))
            End Function

            Private Function EvaluateLegacyLookup(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell,
                                                  horizontal As Boolean) As Object
                If arguments.Count < 3 OrElse arguments.Count > 4 Then Return New FormulaError("#VALUE!")
                Dim input = EvaluateExpression(arguments(1), formulaCell)
                Dim area = TryCast(input, FormulaArea)
                Dim table = TryCast(input, FormulaMatrix)
                If TypeOf input Is FormulaError Then Return input
                If area Is Nothing AndAlso table Is Nothing Then Return New FormulaError("#N/A")
                Dim rowCount = If(area Is Nothing, table.Rows, area.LastRow - area.FirstRow + 1)
                Dim columnCount = If(area Is Nothing, table.Columns, area.LastColumn - area.FirstColumn + 1)
                Dim needles = AsMatrix(EvaluateExpression(arguments(0), formulaCell))
                Dim modes = AsMatrix(If(arguments.Count = 4, EvaluateExpression(arguments(3), formulaCell), True))
                Dim indexInput = EvaluateExpression(arguments(2), formulaCell)
                ' Excel uses only the first result index when the lookup value or match option is an array.
                If needles.Values.Length > 1 OrElse modes.Values.Length > 1 Then
                    Dim indexArea = TryCast(indexInput, FormulaArea)
                    Dim indexMatrix = TryCast(indexInput, FormulaMatrix)
                    If indexArea IsNot Nothing Then
                        indexInput = SliceArea(indexArea, 0, 0, 0, 0)
                    ElseIf indexMatrix IsNot Nothing Then
                        indexInput = indexMatrix.At(0, 0)
                    End If
                End If
                Dim indices = AsMatrix(indexInput)
                Dim keys As Object()
                If area IsNot Nothing Then
                    keys = MaterializeRange(SliceArea(area, 0, 0, If(horizontal, 0, rowCount - 1), If(horizontal, columnCount - 1, 0)))
                ElseIf horizontal Then
                    keys = table.Values.Take(columnCount).ToArray()
                Else
                    keys = Enumerable.Range(0, rowCount).Select(Function(row) table.At(row, 0)).ToArray()
                End If
                Dim resultCount = If(horizontal, rowCount, columnCount)
                Dim selectValue As Func(Of Integer, Integer, Object) =
                    Function(key, index)
                        Dim row = If(horizontal, index, key)
                        Dim column = If(horizontal, key, index)
                        If area Is Nothing Then Return table.At(row, column)
                        Return Scalar(SliceArea(area, row, column, row, column))
                    End Function

                Dim rows = Math.Max(Math.Max(needles.Rows, modes.Rows), indices.Rows)
                Dim columns = Math.Max(Math.Max(needles.Columns, modes.Columns), indices.Columns)
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        results(row * columns + column) = LegacyLookupValue(needles.BroadcastAt(row, column),
                            indices.BroadcastAt(row, column), modes.BroadcastAt(row, column), keys, resultCount, selectValue)
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function LegacyLookupValue(needle As Object, resultIndex As Object, lookupMode As Object,
                                               keys As Object(), resultCount As Integer,
                                               selectValue As Func(Of Integer, Integer, Object)) As Object
                If TypeOf needle Is FormulaError Then Return needle
                Dim optionValue = LogicalOptionValue(lookupMode)
                If TypeOf optionValue Is FormulaError Then Return optionValue
                Dim approximate = DirectCast(optionValue, Boolean)
                Dim mode = If(approximate, -1, If(TypeOf needle Is String, 2, 0))
                Dim key = FindLookupIndex(needle, keys, mode, reverse:=False, preferLastEqual:=approximate)
                If key < 0 Then Return New FormulaError("#N/A")

                Dim indexValue = NumberValue(resultIndex)
                If TypeOf indexValue Is FormulaError Then Return indexValue
                Dim index = Math.Truncate(DirectCast(indexValue, Double))
                If index < 1 Then Return New FormulaError("#VALUE!")
                If index > resultCount Then Return New FormulaError("#REF!")
                Return If(selectValue(key, CInt(index) - 1), 0.0)
            End Function

            Private Function EvaluateXMatch(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 4 Then Return New FormulaError("#VALUE!")
                Dim needle = EvaluateExpression(arguments(0), formulaCell)
                Dim input = EvaluateExpression(arguments(1), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim keys = AsMatrix(input)
                If keys.Rows > 1 AndAlso keys.Columns > 1 Then Return New FormulaError("#VALUE!")

                Dim mode As Integer, search As Integer
                Dim modeError = ReadModernLookupModes(arguments, 2, formulaCell, mode, search)
                If modeError IsNot Nothing Then Return modeError
                Return EvaluateLookupNeedles(needle, keys, mode, search,
                    Function(index) If(index < 0, CType(New FormulaError("#N/A"), Object), CDbl(index + 1)))
            End Function

            Private Function EvaluateXLookup(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 3 OrElse arguments.Count > 6 Then Return New FormulaError("#VALUE!")
                Dim needle = EvaluateExpression(arguments(0), formulaCell)
                Dim keyInput = EvaluateExpression(arguments(1), formulaCell)
                Dim valueInput = EvaluateExpression(arguments(2), formulaCell)
                If TypeOf keyInput Is FormulaError Then Return keyInput
                If TypeOf valueInput Is FormulaError Then Return valueInput
                Dim keys = AsMatrix(keyInput)
                If keys.Rows > 1 AndAlso keys.Columns > 1 Then Return New FormulaError("#VALUE!")

                Dim area = TryCast(valueInput, FormulaArea)
                Dim matrix = If(area Is Nothing, AsMatrix(valueInput), Nothing)
                Dim rows = If(area Is Nothing, matrix.Rows, area.LastRow - area.FirstRow + 1)
                Dim columns = If(area Is Nothing, matrix.Columns, area.LastColumn - area.FirstColumn + 1)
                Dim byRow = keys.Columns = 1 AndAlso rows = keys.Rows
                If Not byRow AndAlso (keys.Rows <> 1 OrElse columns <> keys.Columns) Then Return New FormulaError("#VALUE!")

                Dim mode As Integer, search As Integer
                Dim modeError = ReadModernLookupModes(arguments, 4, formulaCell, mode, search)
                If modeError IsNot Nothing Then Return modeError
                Dim fallback As New Lazy(Of Object)(
                    Function() If(If(arguments.Count >= 4, EvaluateExpression(arguments(3), formulaCell), Nothing), New FormulaError("#N/A")))
                Return EvaluateLookupNeedles(needle, keys, mode, search,
                    Function(index)
                        If index < 0 Then Return fallback.Value
                        Dim firstRow = If(byRow, index, 0)
                        Dim firstColumn = If(byRow, 0, index)
                        Dim lastRow = If(byRow, index, rows - 1)
                        Dim lastColumn = If(byRow, columns - 1, index)
                        If area IsNot Nothing Then Return SliceArea(area, firstRow, firstColumn, lastRow, lastColumn)
                        Return SliceMatrix(matrix, firstRow, firstColumn, lastRow, lastColumn)
                    End Function)
            End Function

            Private Function ReadModernLookupModes(arguments As IReadOnlyList(Of FormulaExpression), firstOption As Integer,
                                                   formulaCell As ExcelCell, ByRef mode As Integer, ByRef search As Integer) As FormulaError
                Dim rawMode = If(arguments.Count > firstOption, ReadIndex(arguments(firstOption), formulaCell), 0.0)
                If TypeOf rawMode Is FormulaError Then Return DirectCast(rawMode, FormulaError)
                Dim matchNumber = DirectCast(rawMode, Double)
                If matchNumber < -1.0 OrElse matchNumber > 3.0 Then Return New FormulaError("#VALUE!")
                mode = CInt(matchNumber)

                Dim searchNumber = 1.0
                If arguments.Count > firstOption + 1 Then
                    Dim rawSearch = Scalar(EvaluateExpression(arguments(firstOption + 1), formulaCell))
                    If rawSearch IsNot Nothing Then
                        Dim number = NumberValue(rawSearch)
                        If TypeOf number Is FormulaError Then Return DirectCast(number, FormulaError)
                        searchNumber = Math.Truncate(DirectCast(number, Double))
                    End If
                End If
                If Math.Abs(searchNumber) < 1.0 OrElse Math.Abs(searchNumber) > 2.0 Then Return New FormulaError("#VALUE!")
                If mode >= 2 AndAlso Math.Abs(searchNumber) = 2.0 Then Return New FormulaError("#VALUE!")
                search = CInt(searchNumber)
                Return Nothing
            End Function

            Private Function EvaluateLookupNeedles(needle As Object, keys As FormulaMatrix, mode As Integer, search As Integer,
                                                     selectValue As Func(Of Integer, Object)) As Object
                Dim binaryPositions As Dictionary(Of Type, Integer()) = Nothing
                If Math.Abs(search) = 2 Then
                    Dim positions = Enumerable.Range(0, keys.Values.Length).Where(Function(index) Not TypeOf keys.Values(index) Is FormulaError)
                    If search < 0 Then positions = positions.Reverse()
                    binaryPositions = positions.GroupBy(Function(index) LookupValueType(keys.Values(index))).
                        ToDictionary(Function(group) group.Key, Function(group) group.ToArray())
                End If
                If Not IsMultiValue(needle) Then
                    Return SelectLookupValue(Scalar(needle), keys.Values, mode, search, selectValue, binaryPositions)
                End If
                Dim matrix = AsMatrix(needle)
                Dim results = matrix.Values.Select(
                    Function(value) If(Scalar(SelectLookupValue(value, keys.Values, mode, search, selectValue, binaryPositions)), 0.0)).ToArray()
                Return New FormulaMatrix(matrix.Rows, matrix.Columns, results)
            End Function

            Private Function SelectLookupValue(needle As Object, keys As Object(), mode As Integer, search As Integer,
                                                       selectValue As Func(Of Integer, Object),
                                                       binaryPositions As Dictionary(Of Type, Integer())) As Object
                If TypeOf needle Is FormulaError Then Return needle
                If mode = 3 Then Return SelectRegexLookupValue(needle, keys, reverse:=search < 0, selectValue)
                If binaryPositions Is Nothing Then
                    Return selectValue(FindLookupIndex(needle, keys, mode, reverse:=search < 0, preferLastEqual:=False, matchBlank:=True))
                End If
                Dim candidates As Integer() = Nothing
                If Not binaryPositions.TryGetValue(LookupValueType(needle), candidates) Then Return selectValue(-1)
                Return selectValue(FindBinaryLookupIndex(needle, keys, candidates, mode))
            End Function

            Private Function ReadIndex(expression As FormulaExpression, formulaCell As ExcelCell) As Object
                Dim value = NumberValue(Scalar(EvaluateExpression(expression, formulaCell)))
                If TypeOf value Is FormulaError Then Return value
                Return Math.Truncate(DirectCast(value, Double))
            End Function

            Private Shared Function IsLookupCandidate(needle As Object, item As Object) As Boolean
                If needle Is Nothing Then Return item Is Nothing
                Return item IsNot Nothing AndAlso Not TypeOf item Is FormulaError AndAlso item.GetType() Is needle.GetType()
            End Function

            Private Shared Function LookupValueType(value As Object) As Type
                Return If(value?.GetType(), GetType(DBNull))
            End Function

            Private Shared Function FindBinaryLookupIndex(needle As Object, items As Object(), candidates As Integer(), mode As Integer) As Integer
                ' Compare candidates in ascending value order while retaining their original positions.
                Dim first = 0
                Dim last = candidates.Length
                While first < last
                    Dim middle = first + (last - first) \ 2
                    If EvaluateComparison("<", items(candidates(middle)), needle) Then
                        first = middle + 1
                    Else
                        last = middle
                    End If
                End While

                If first < candidates.Length AndAlso EvaluateComparison("=", items(candidates(first)), needle) Then Return candidates(first)
                If mode = -1 AndAlso first > 0 Then Return candidates(first - 1)
                If mode = 1 AndAlso first < candidates.Length Then Return candidates(first)
                Return -1
            End Function

            Private Shared Function FindLookupIndex(needle As Object, items As Object(), mode As Integer,
                                                    reverse As Boolean, preferLastEqual As Boolean, Optional matchBlank As Boolean = False) As Integer
                If needle Is Nothing AndAlso Not matchBlank Then needle = 0.0
                Dim pattern = If(mode = 2 AndAlso TypeOf needle Is String, BuildWildcard(DirectCast(needle, String)), Nothing)
                Dim best = -1
                Dim equalMatch = -1
                For offset = 0 To items.Length - 1
                    Dim index = If(reverse, items.Length - 1 - offset, offset)
                    Dim item = items(index)
                    If Not IsLookupCandidate(needle, item) Then Continue For
                    Dim equal = If(pattern Is Nothing, EvaluateComparison("=", item, needle), MatchesWildcard(pattern, DirectCast(item, String)))
                    If equal Then
                        If Not preferLastEqual Then Return index
                        equalMatch = index
                    ElseIf mode = -1 AndAlso EvaluateComparison("<", item, needle) Then
                        If best < 0 OrElse EvaluateComparison(If(preferLastEqual, ">=", ">"), item, items(best)) Then best = index
                    ElseIf mode = 1 AndAlso EvaluateComparison(">", item, needle) Then
                        If best < 0 OrElse EvaluateComparison(If(preferLastEqual, "<=", "<"), item, items(best)) Then best = index
                    End If
                Next
                Return If(equalMatch >= 0, equalMatch, best)
            End Function

        End Class

    End Class

End Namespace
