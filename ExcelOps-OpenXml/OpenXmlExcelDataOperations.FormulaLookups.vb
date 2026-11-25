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
                If firstRow = lastRow AndAlso firstColumn = lastColumn Then Return matrix.At(firstRow, firstColumn)

                Dim selectedRows = lastRow - firstRow + 1
                Dim selectedColumns = lastColumn - firstColumn + 1
                Dim values = AllocateMatrixValues(selectedRows, selectedColumns)
                For rowIndex = 0 To selectedRows - 1
                    For columnIndex = 0 To selectedColumns - 1
                        values(rowIndex * selectedColumns + columnIndex) = matrix.At(firstRow + rowIndex, firstColumn + columnIndex)
                    Next
                Next
                Return New FormulaMatrix(selectedRows, selectedColumns, values)
            End Function

            Private Function EvaluateMatch(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim needle = Scalar(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf needle Is FormulaError Then Return needle
                Dim input = EvaluateExpression(arguments(1), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim area = TryCast(input, FormulaArea)
                If Not IsVector(area) Then Return New FormulaError("#N/A")

                Dim mode = 1.0
                If arguments.Count = 3 Then
                    Dim optionValue = ReadIndex(arguments(2), formulaCell)
                    If TypeOf optionValue Is FormulaError Then Return optionValue
                    mode = DirectCast(optionValue, Double)
                End If
                If Math.Abs(mode) > 1.0 Then Return New FormulaError("#N/A")
                Dim matchMode = If(mode = 0, If(TypeOf needle Is String, 2, 0), -CInt(mode))
                Dim index = FindLookupIndex(needle, MaterializeRange(area), matchMode, reverse:=False, preferLastEqual:=mode <> 0)
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

            Private Function EvaluateXLookup(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 3 OrElse arguments.Count > 6 Then Return New FormulaError("#VALUE!")
                Dim needle = Scalar(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf needle Is FormulaError Then Return needle
                Dim first = EvaluateExpression(arguments(1), formulaCell)
                Dim second = EvaluateExpression(arguments(2), formulaCell)
                If TypeOf first Is FormulaError Then Return first
                If TypeOf second Is FormulaError Then Return second
                Dim keys = TryCast(first, FormulaArea)
                Dim values = TryCast(second, FormulaArea)
                If Not IsVector(keys) OrElse Not IsVector(values) OrElse Not SameSize(keys, values) Then Return New FormulaError("#VALUE!")

                Dim mode = 0.0
                Dim search = 1.0
                If arguments.Count >= 5 Then
                    Dim optionValue = ReadIndex(arguments(4), formulaCell)
                    If TypeOf optionValue Is FormulaError Then Return optionValue
                    mode = DirectCast(optionValue, Double)
                    If mode < -1.0 OrElse mode > 2.0 Then Return New FormulaError("#VALUE!")
                End If
                If arguments.Count = 6 Then
                    Dim optionValue = Scalar(EvaluateExpression(arguments(5), formulaCell))
                    If optionValue IsNot Nothing Then
                        optionValue = NumberValue(optionValue)
                        If TypeOf optionValue Is FormulaError Then Return optionValue
                        search = Math.Truncate(DirectCast(optionValue, Double))
                    End If
                    If Math.Abs(search) = 2.0 Then Throw New NotSupportedException("Binary search in XLOOKUP is outside the supported calculation scope.")
                    If Math.Abs(search) <> 1.0 Then Return New FormulaError("#VALUE!")
                End If
                Dim index = FindLookupIndex(needle, MaterializeRange(keys), CInt(mode), reverse:=search < 0, preferLastEqual:=False)
                If index < 0 Then
                    Dim fallback = If(arguments.Count >= 4, EvaluateExpression(arguments(3), formulaCell), Nothing)
                    Return If(fallback, New FormulaError("#N/A"))
                End If
                Dim row = If(values.FirstRow = values.LastRow, 0, index)
                Dim column = If(values.FirstColumn = values.LastColumn, 0, index)
                Return Scalar(SliceArea(values, row, column, row, column))
            End Function

            Private Function ReadIndex(expression As FormulaExpression, formulaCell As ExcelCell) As Object
                Dim value = NumberValue(Scalar(EvaluateExpression(expression, formulaCell)))
                If TypeOf value Is FormulaError Then Return value
                Return Math.Truncate(DirectCast(value, Double))
            End Function

            Private Shared Function IsVector(area As FormulaArea) As Boolean
                Return area IsNot Nothing AndAlso (area.FirstRow = area.LastRow OrElse area.FirstColumn = area.LastColumn)
            End Function

            Private Shared Function FindLookupIndex(needle As Object, items As Object(), mode As Integer,
                                                    reverse As Boolean, preferLastEqual As Boolean) As Integer
                If needle Is Nothing Then needle = 0.0
                Dim pattern = If(mode = 2 AndAlso TypeOf needle Is String, BuildWildcard(DirectCast(needle, String)), Nothing)
                Dim best = -1
                Dim equalMatch = -1
                For offset = 0 To items.Length - 1
                    Dim index = If(reverse, items.Length - 1 - offset, offset)
                    Dim item = items(index)
                    If item Is Nothing OrElse TypeOf item Is FormulaError OrElse item.GetType() IsNot needle.GetType() Then Continue For
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
