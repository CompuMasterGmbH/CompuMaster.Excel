Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateLookupFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                Select Case name
                    Case "INDEX"
                        Return EvaluateIndex(arguments, formulaCell)
                    Case "MATCH"
                        Return EvaluateMatch(arguments, formulaCell)
                    Case "VLOOKUP"
                        Return EvaluateVLookup(arguments, formulaCell)
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
                Return SliceLookupMatrix(matrix, firstRow, firstColumn, lastRow, lastColumn)
            End Function

            Private Function EvaluateMatch(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim needle = Scalar(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf needle Is FormulaError Then Return needle
                Dim input = EvaluateExpression(arguments(1), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim keys = AsMatrix(input)
                If keys.Rows > 1 AndAlso keys.Columns > 1 Then Return New FormulaError("#N/A")

                Dim mode = 1.0
                If arguments.Count = 3 Then
                    Dim optionValue = ReadIndex(arguments(2), formulaCell)
                    If TypeOf optionValue Is FormulaError Then Return optionValue
                    mode = DirectCast(optionValue, Double)
                End If
                If Math.Abs(mode) > 1.0 Then Return New FormulaError("#N/A")
                Dim matchMode = If(mode = 0, If(TypeOf needle Is String, 2, 0), -CInt(mode))
                Dim index = FindLookupIndex(needle, keys.Values, matchMode, reverse:=False, preferLastEqual:=mode <> 0)
                Return If(index < 0, CType(New FormulaError("#N/A"), Object), CDbl(index + 1))
            End Function

            Private Function EvaluateVLookup(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 3 OrElse arguments.Count > 4 Then Return New FormulaError("#VALUE!")
                Dim needle = Scalar(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf needle Is FormulaError Then Return needle
                Dim input = EvaluateExpression(arguments(1), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim table = TryCast(input, FormulaArea)
                If table Is Nothing Then Return New FormulaError("#VALUE!")
                Dim column = ReadIndex(arguments(2), formulaCell)
                If TypeOf column Is FormulaError Then Return column
                Dim columnNumber = DirectCast(column, Double)
                If columnNumber < 1 Then Return New FormulaError("#VALUE!")
                If columnNumber > table.LastColumn - table.FirstColumn + 1 Then Return New FormulaError("#REF!")

                Dim approximate = True
                If arguments.Count = 4 Then
                    Dim optionValue = NumberValue(Scalar(EvaluateExpression(arguments(3), formulaCell)))
                    If TypeOf optionValue Is FormulaError Then Return optionValue
                    approximate = DirectCast(optionValue, Double) <> 0
                End If
                Dim keys = SliceArea(table, 0, 0, table.LastRow - table.FirstRow, 0)
                Dim mode = If(approximate, -1, If(TypeOf needle Is String, 2, 0))
                Dim index = FindLookupIndex(needle, MaterializeRange(keys), mode, reverse:=False, preferLastEqual:=approximate)
                If index < 0 Then Return New FormulaError("#N/A")
                Return Scalar(SliceArea(table, index, CInt(columnNumber) - 1, index, CInt(columnNumber) - 1))
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
                        Return SliceLookupMatrix(matrix, firstRow, firstColumn, lastRow, lastColumn)
                    End Function)
            End Function

            Private Function ReadModernLookupModes(arguments As IReadOnlyList(Of FormulaExpression), firstOption As Integer,
                                                   formulaCell As ExcelCell, ByRef mode As Integer, ByRef search As Integer) As FormulaError
                Dim rawMode = If(arguments.Count > firstOption, ReadIndex(arguments(firstOption), formulaCell), 0.0)
                If TypeOf rawMode Is FormulaError Then Return DirectCast(rawMode, FormulaError)
                Dim matchNumber = DirectCast(rawMode, Double)
                If matchNumber = 3.0 Then Throw New NotSupportedException("Regular expression matching is outside the supported calculation scope.")
                If matchNumber < -1.0 OrElse matchNumber > 2.0 Then Return New FormulaError("#VALUE!")
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
                If mode = 2 AndAlso Math.Abs(searchNumber) = 2.0 Then Return New FormulaError("#VALUE!")
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

            Private Shared Function SelectLookupValue(needle As Object, keys As Object(), mode As Integer, search As Integer,
                                                       selectValue As Func(Of Integer, Object),
                                                       binaryPositions As Dictionary(Of Type, Integer())) As Object
                If TypeOf needle Is FormulaError Then Return needle
                If binaryPositions Is Nothing Then
                    Return selectValue(FindLookupIndex(needle, keys, mode, reverse:=search < 0, preferLastEqual:=False, matchBlank:=True))
                End If
                Dim candidates As Integer() = Nothing
                If Not binaryPositions.TryGetValue(LookupValueType(needle), candidates) Then Return selectValue(-1)
                Return selectValue(FindBinaryLookupIndex(needle, keys, candidates, mode))
            End Function

            Private Shared Function SliceLookupMatrix(matrix As FormulaMatrix, firstRow As Integer, firstColumn As Integer,
                                                       lastRow As Integer, lastColumn As Integer) As Object
                If firstRow = lastRow AndAlso firstColumn = lastColumn Then Return matrix.At(firstRow, firstColumn)
                Dim rows = lastRow - firstRow + 1
                Dim columns = lastColumn - firstColumn + 1
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        values(row * columns + column) = matrix.At(firstRow + row, firstColumn + column)
                    Next
                Next
                Return New FormulaMatrix(rows, columns, values)
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
