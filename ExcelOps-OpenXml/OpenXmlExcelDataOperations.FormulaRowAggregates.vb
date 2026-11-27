Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly NestedRowAggregate As New Regex(
                """(?:[^""]|"""")*""|'(?:[^']|'')*'|\[[^\]]*\]|(?<![\p{L}\p{Nd}_.])(?:_xlfn\.)?(?<aggregate>SUBTOTAL|AGGREGATE)\s*\(",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

            Private Function EvaluateRowAggregate(subtotal As Boolean, arguments As IReadOnlyList(Of FormulaExpression),
                                                  formulaCell As ExcelCell) As Object
                Dim controlCount = If(subtotal, 1, 2)
                If arguments.Count <= controlCount Then Return New FormulaError("#VALUE!")
                Dim inputs = arguments.Skip(controlCount).Select(
                    Function(argument) New Lazy(Of Object)(Function() EvaluateExpression(argument, formulaCell))).ToArray()

                Dim selected = BroadcastReferenceArguments(arguments.Take(controlCount).ToArray(), formulaCell,
                    Function(controls) EvaluateRowAggregateChoice(subtotal, controls, inputs))
                Dim selections = TryCast(selected, FormulaMatrix)
                If selections Is Nothing Then Return selected
                Dim rankResults = selections.Values.OfType(Of FormulaMatrix)().ToArray()
                If rankResults.Length = 0 Then Return selections

                Dim rows = Math.Max(selections.Rows, rankResults.Max(Function(result) result.Rows))
                Dim columns = Math.Max(selections.Columns, rankResults.Max(Function(result) result.Columns))
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim value = selections.BroadcastAt(row, column)
                        Dim ranks = TryCast(value, FormulaMatrix)
                        values(row * columns + column) = If(ranks Is Nothing, value, ranks.BroadcastAt(row, column))
                    Next
                Next
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Function EvaluateRowAggregateChoice(subtotal As Boolean, controls As Object(), inputs As Lazy(Of Object)()) As Object
                Dim codeValue = NumberValue(controls(0))
                If TypeOf codeValue Is FormulaError Then Return codeValue
                Dim code = Math.Truncate(DirectCast(codeValue, Double))
                Dim functionNumber As Integer
                Dim skipHidden, skipErrors, skipNested As Boolean
                If subtotal Then
                    skipHidden = code >= 101 AndAlso code <= 111
                    If skipHidden Then code -= 100
                    If code < 1 OrElse code > 11 Then Return New FormulaError("#VALUE!")
                    skipNested = True
                Else
                    If code < 1 OrElse code > 19 Then Return New FormulaError("#VALUE!")
                    Dim optionValue = NumberValue(controls(1))
                    If TypeOf optionValue Is FormulaError Then Return optionValue
                    Dim optionNumber = DirectCast(optionValue, Double)
                    If optionNumber < 0 OrElse optionNumber >= 8 Then Return New FormulaError("#VALUE!")
                    Dim flags = CInt(Math.Truncate(optionNumber))
                    skipHidden = (flags And 1) <> 0
                    skipErrors = (flags And 2) <> 0
                    skipNested = flags < 4
                End If
                functionNumber = CInt(code)
                Dim arrayForm = functionNumber >= 14
                If arrayForm AndAlso inputs.Length <> 2 Then Return New FormulaError("#VALUE!")

                Dim includeCell As Func(Of WorksheetSnapshot, Integer, Integer, Boolean) =
                    Function(sheet, row, column)
                        If (skipHidden OrElse (subtotal AndAlso sheet.HasActiveRowFilter.Value)) AndAlso
                            sheet.HiddenRows.Value.Contains(row) Then Return False
                        If Not skipNested Then Return True
                        Return Not IsNestedRowAggregate(sheet, row, column)
                    End Function

                Dim numbers As New List(Of Double)
                Dim occupied = 0
                For index = 0 To If(arrayForm, 0, inputs.Length - 1)
                    Dim input = inputs(index).Value
                    Dim area = TryCast(input, FormulaArea)
                    Dim union = TryCast(input, FormulaReferenceUnion)
                    Dim items As IEnumerable(Of Object)
                    If area IsNot Nothing Then
                        items = RangeValues(area, includeCell)
                    ElseIf union IsNot Nothing Then
                        items = union.Areas.SelectMany(Function(member) RangeValues(member, includeCell))
                    ElseIf arrayForm Then
                        items = ReferenceArgumentValues(input)
                    ElseIf TypeOf input Is FormulaError Then
                        Return input
                    Else
                        Return New FormulaError("#VALUE!")
                    End If

                    For Each item In items
                        If item Is Nothing Then Continue For
                        If TypeOf item Is FormulaError Then
                            If skipErrors Then Continue For
                            If functionNumber <> 2 AndAlso functionNumber <> 3 Then Return item
                        End If
                        occupied += 1
                        If TypeOf item Is Double Then numbers.Add(DirectCast(item, Double))
                    Next
                Next
                If Not arrayForm Then Return SummarizeRowAggregate(functionNumber, numbers, occupied)

                Dim rankInput = inputs(1).Value
                If TypeOf rankInput Is FormulaReferenceUnion OrElse (TypeOf rankInput Is FormulaArea AndAlso IsMultiValue(rankInput)) Then
                    Return New FormulaError("#VALUE!")
                End If
                Dim statisticNames = {"LARGE", "SMALL", "PERCENTILE.INC", "QUARTILE.INC", "PERCENTILE.EXC", "QUARTILE.EXC"}
                Dim statistic = statisticNames(functionNumber - 14)
                Dim ranks = AsMatrix(rankInput)
                numbers.Sort()
                Dim results = ranks.Values.Select(
                    Function(rank)
                        Dim converted = NumberValue(rank)
                        If TypeOf converted Is FormulaError Then Return converted
                        Return SelectOrderStatistic(statistic, numbers, DirectCast(converted, Double))
                    End Function).ToArray()
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(ranks.Rows, ranks.Columns, results)
            End Function

            Private Shared Function IsNestedRowAggregate(sheet As WorksheetSnapshot, row As Integer, column As Integer) As Boolean
                Dim cell As DocumentFormat.OpenXml.Spreadsheet.Cell = Nothing
                If Not sheet.Cells.TryGetValue((row, column), cell) Then Return False
                Dim formula = cell.CellFormula?.InnerText
                Return Not String.IsNullOrEmpty(formula) AndAlso
                    NestedRowAggregate.Matches(formula).Cast(Of Match)().Any(Function(token) token.Groups("aggregate").Success)
            End Function

            Private Shared Function SummarizeRowAggregate(functionNumber As Integer, numbers As List(Of Double), occupied As Integer) As Object
                Select Case functionNumber
                    Case 1 : Return If(numbers.Count = 0, CObj(New FormulaError("#DIV/0!")), FiniteNumber(numbers.Average()))
                    Case 2 : Return CDbl(numbers.Count)
                    Case 3 : Return CDbl(occupied)
                    Case 4 : Return If(numbers.Count = 0, 0.0, numbers.Max())
                    Case 5 : Return If(numbers.Count = 0, 0.0, numbers.Min())
                    Case 6
                        If numbers.Count = 0 Then Return 0.0
                        Return FiniteNumber(numbers.Aggregate(1.0, Function(product, value) product * value))
                    Case 7 : Return EvaluateDeviation("STDEV.S", numbers)
                    Case 8 : Return EvaluateDeviation("STDEV.P", numbers)
                    Case 9 : Return FiniteNumber(numbers.Sum())
                    Case 10 : Return EvaluateDeviation("VAR.S", numbers)
                    Case 11 : Return EvaluateDeviation("VAR.P", numbers)
                    Case 12
                        If numbers.Count = 0 Then Return New FormulaError("#NUM!")
                        numbers.Sort()
                        Return InterpolateOrderedValues(numbers, (numbers.Count - 1) * 0.5)
                    Case Else
                        Return Modes(numbers, multiple:=False)
                End Select
            End Function

        End Class

    End Class

End Namespace
