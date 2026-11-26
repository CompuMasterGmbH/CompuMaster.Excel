Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Structure PivotAxisMember
                Friend Group As WorksheetGroup
                Friend Parent As WorksheetGroup
                Friend Labels As Object()
                Friend Level As Integer
            End Structure

            Private Function EvaluatePivotBy(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 4 OrElse arguments.Count > 11 Then Return New FormulaError("#VALUE!")
                Dim inputs = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                Dim failure = inputs.OfType(Of FormulaError)().FirstOrDefault()
                If failure IsNot Nothing Then Return failure
                Dim values = AsMatrix(inputs(2))
                Dim rowFields = If(inputs(0) Is Nothing, New FormulaMatrix(values.Rows, 0, Array.Empty(Of Object)()), AsMatrix(inputs(0)))
                Dim columnFields = If(inputs(1) Is Nothing, New FormulaMatrix(values.Rows, 0, Array.Empty(Of Object)()), AsMatrix(inputs(1)))
                If rowFields.Rows <> values.Rows OrElse columnFields.Rows <> values.Rows OrElse
                    rowFields.Columns + columnFields.Columns = 0 Then Return New FormulaError("#VALUE!")

                Dim functionMatrix = AsMatrix(inputs(3))
                Dim functions = functionMatrix.Values.OfType(Of FormulaLambda)().ToArray()
                If functions.Length <> functionMatrix.Values.Length OrElse (functionMatrix.Rows > 1 AndAlso functionMatrix.Columns > 1) OrElse
                    (values.Columns > 1 AndAlso functions.Length > 1 AndAlso values.Columns <> functions.Length) OrElse
                    functions.Any(Function(item) Not item.Accepts(1) AndAlso Not item.Accepts(2)) Then Return New FormulaError("#VALUE!")
                Dim vertical = functionMatrix.Rows > 1
                Dim measures = Math.Max(values.Columns, functions.Length)
                Dim headers = GroupOption(inputs.ElementAtOrDefault(4), -1, 0, 3)
                Dim rowDepth = GroupOption(inputs.ElementAtOrDefault(5), If(rowFields.Columns = 0, 0, 1), -rowFields.Columns, rowFields.Columns)
                Dim columnDepth = GroupOption(inputs.ElementAtOrDefault(7), If(columnFields.Columns = 0, 0, 1), -columnFields.Columns, columnFields.Columns)
                Dim relative = GroupOption(inputs.ElementAtOrDefault(10), 0, 0, 4)
                If Not headers.HasValue OrElse Not rowDepth.HasValue OrElse Not columnDepth.HasValue OrElse Not relative.HasValue Then Return New FormulaError("#VALUE!")
                If values.Columns > 1 AndAlso functions.Length = 1 AndAlso Math.Abs(columnDepth.Value) > 1 Then
                    Throw New NotSupportedException("Excel's column subtotals with multiple values and a single aggregation require additional PIVOTBY support.")
                End If

                Dim rowSort = ReadGroupingSort(inputs.ElementAtOrDefault(6), rowFields.Columns, measures)
                Dim columnSort = ReadGroupingSort(inputs.ElementAtOrDefault(8), columnFields.Columns, measures)
                If rowSort Is Nothing OrElse columnSort Is Nothing Then Return New FormulaError("#VALUE!")
                If columnDepth.Value = 0 AndAlso rowSort.Any(Function(index) Math.Abs(index) > rowFields.Columns) Then rowSort = Enumerable.Range(1, rowFields.Columns).ToArray()
                Dim hasHeaderRow = headers.Value = 1 OrElse headers.Value = 3 OrElse (headers.Value = -1 AndAlso
                    values.Rows > 1 AndAlso TypeOf values.At(0, 0) Is String AndAlso TypeOf values.At(1, 0) Is Double)
                Dim filter = inputs.ElementAtOrDefault(9)
                Dim mask = If(filter Is Nothing, Nothing, AsMatrix(filter))
                If mask IsNot Nothing AndAlso (mask.Columns <> 1 OrElse (mask.Rows <> values.Rows AndAlso
                    (mask.Rows <> 1 OrElse rowFields.Columns > 1 OrElse columnFields.Columns > 1)) OrElse
                    mask.Values.Any(Function(value) value IsNot Nothing AndAlso Not TypeOf value Is Boolean AndAlso Not TypeOf value Is Double)) Then
                    Return New FormulaError("#VALUE!")
                End If
                Dim rows = Enumerable.Range(If(hasHeaderRow, 1, 0), values.Rows - If(hasHeaderRow, 1, 0)).
                    Where(Function(index) mask Is Nothing OrElse GroupingFilterIncludes(mask.Values(If(mask.Rows = 1, 0, index)))).ToArray()
                If rows.Length = 0 Then Return New FormulaError("#VALUE!")

                Dim allValues = Enumerable.Range(0, values.Columns).Select(Function(column) GroupValueArray(values, rows, column)).ToArray()
                Dim needsDenominator = functions.Any(AddressOf GroupLambdaNeedsTotal)
                Dim aggregate As Func(Of Integer(), Object()) =
                    Function(members) CalculateGroupValues(values, functions, members, allValues, formulaCell)
                Dim rowRoot As New WorksheetGroup(rows, aggregate)
                Dim columnRoot As New WorksheetGroup(rows, aggregate)
                Dim rowGroups = MakePivotAxis(rowRoot, rowFields, rowSort, rowDepth.Value, aggregate)
                Dim columnGroups = MakePivotAxis(columnRoot, columnFields, columnSort, columnDepth.Value, aggregate, valueSortFromLevel:=If(rowDepth.Value = 0, 1, 0))

                Dim showHeaders = headers.Value = 2 OrElse headers.Value = 3 OrElse
                    (headers.Value = -1 AndAlso hasHeaderRow AndAlso Not vertical AndAlso functions.Length > 1)
                Dim showColumnFields = showHeaders AndAlso columnFields.Columns > 0
                Dim showValueLabels = showHeaders AndAlso Not vertical
                Dim showLambdaLabels = Not vertical AndAlso functions.Length > 1
                Dim showVerticalMeasure = showHeaders AndAlso vertical AndAlso values.Columns > 1
                Dim labelColumns = rowFields.Columns + If(vertical, 1, 0) + If(showVerticalMeasure, 1, 0)
                Dim repeatColumns = If(vertical, 1, measures)
                Dim wide = labelColumns + CLng(columnGroups.Count) * repeatColumns
                Dim tall = columnFields.Columns + If(showColumnFields, 1, 0) + If(showValueLabels, 1, 0) +
                    If(showLambdaLabels, 1, 0) + CLng(rowGroups.Count) * If(vertical, functions.Length, 1)
                If wide > ColumnCount OrElse tall > RowCount Then Return New FormulaError("#NUM!")
                Dim width = CInt(wide)
                Dim height = CInt(tall)
                Dim output = AllocateMatrixValues(height, width)
                For index = 0 To output.Length - 1
                    output(index) = String.Empty
                Next

                Dim german = ConversionCulture.TwoLetterISOLanguageName = "de"
                Dim rowLabels = GroupingFieldLabels(rowFields, hasHeaderRow, If(german, "Zeilenfeld ", "Row Field "))
                Dim columnLabels = GroupingFieldLabels(columnFields, hasHeaderRow, If(german, "Spaltenfeld ", "Column Field "))
                Dim valueLabels = GroupingFieldLabels(values, hasHeaderRow, If(german, "Wert ", "Value "))
                Dim names = Enumerable.Range(0, functions.Length).Select(Function(index) GroupingLambdaLabel(functions(index), index)).ToArray()
                Dim outputRow = 0
                If showColumnFields Then
                    output(labelColumns) = String.Join("; ", columnLabels.Select(Function(item) TextValue(item)))
                    outputRow += 1
                End If
                For level = 0 To columnFields.Columns - 1
                    If showHeaders AndAlso vertical AndAlso level = columnFields.Columns - 1 Then
                        Array.Copy(rowLabels, 0, output, outputRow * width, rowLabels.Length)
                    End If
                    For index = 0 To columnGroups.Count - 1
                        For measure = 0 To repeatColumns - 1
                            output(outputRow * width + labelColumns + index * repeatColumns + measure) = columnGroups(index).Labels(level)
                        Next
                    Next
                    outputRow += 1
                Next
                If showLambdaLabels Then
                    For index = 0 To columnGroups.Count - 1
                        Array.Copy(names, 0, output, outputRow * width + labelColumns + index * repeatColumns, names.Length)
                    Next
                    outputRow += 1
                End If
                If showValueLabels Then
                    Array.Copy(rowLabels, 0, output, outputRow * width, rowLabels.Length)
                    For index = 0 To columnGroups.Count - 1
                        For measure = 0 To repeatColumns - 1
                            output(outputRow * width + labelColumns + index * repeatColumns + measure) = valueLabels(If(values.Columns = 1, 0, measure))
                        Next
                    Next
                    outputRow += 1
                End If

                For Each rowGroup In rowGroups
                    Dim displayRows = If(vertical, functions.Length, 1)
                    For layer = 0 To displayRows - 1
                        Dim at = (outputRow + layer) * width
                        Array.Copy(rowGroup.Labels, 0, output, at, rowGroup.Labels.Length)
                        If vertical Then output(at + rowFields.Columns) = names(layer)
                        If showVerticalMeasure Then output(at + rowFields.Columns + 1) = valueLabels(layer)
                    Next
                    For index = 0 To columnGroups.Count - 1
                        Dim columnGroup = columnGroups(index)
                        If columnFields.Columns > 0 AndAlso columnGroup.Level = 0 AndAlso values.Columns > 1 AndAlso functions.Length = 1 Then Continue For
                        Dim subset = rowGroup.Group.Rows.Intersect(columnGroup.Group.Rows).ToArray()
                        If subset.Length = 0 Then Continue For
                        Dim totalValues = allValues
                        If needsDenominator Then
                            Dim totals = PivotDenominatorRows(relative.Value, rowGroup, columnGroup, rowRoot)
                            totalValues = Enumerable.Range(0, values.Columns).
                                Select(Function(column) GroupValueArray(values, totals, column)).ToArray()
                        End If
                        Dim results = CalculateGroupValues(values, functions, subset, totalValues, formulaCell)
                        Dim outputColumn = labelColumns + index * repeatColumns
                        If vertical Then
                            For layer = 0 To results.Length - 1
                                output((outputRow + layer) * width + outputColumn) = results(layer)
                            Next
                        Else
                            Array.Copy(results, 0, output, outputRow * width + outputColumn, results.Length)
                        End If
                    Next
                    outputRow += displayRows
                Next
                Return New FormulaMatrix(height, width, output)
            End Function

            Private Function MakePivotAxis(root As WorksheetGroup, fields As FormulaMatrix, sort As Integer(), depth As Integer,
                                           aggregate As Func(Of Integer(), Object()),
                                           Optional valueSortFromLevel As Integer = 0) As List(Of PivotAxisMember)
                Dim members As New List(Of PivotAxisMember)
                If fields.Columns = 0 Then
                    members.Add(New PivotAxisMember With {.Group = root, .Parent = root, .Labels = Array.Empty(Of Object)(), .Level = 0})
                    Return members
                End If
                Dim comparer = StringComparer.Create(ConversionCulture, ignoreCase:=True)
                BuildGroupHierarchy(root, fields, 0, New UniqueSliceComparer(comparer), comparer, sort, aggregate, valueSortFromLevel)
                Dim german = ConversionCulture.TwoLetterISOLanguageName = "de"
                Dim grandTotal = If(Math.Abs(depth) <= 1, If(german, "Gesamt", "Total"), If(german, "Gesamtergebnis", "Grand Total"))
                ForEachGroupingAxisMember(root, root, fields, New Object(fields.Columns - 1) {}, 0, fields.Columns, depth,
                    Sub(node, parent, prefix, level)
                        Dim labels = GroupingLabels(prefix, level, grandTotal)
                        members.Add(New PivotAxisMember With {.Group = node, .Parent = parent, .Labels = labels, .Level = level})
                    End Sub)
                Return members
            End Function

            Private Shared Function PivotDenominatorRows(relative As Integer, rowGroup As PivotAxisMember, columnGroup As PivotAxisMember,
                                                         rowRoot As WorksheetGroup) As Integer()
                Select Case relative
                    Case 0 : Return columnGroup.Group.Rows
                    Case 1 : Return rowGroup.Group.Rows
                    Case 2 : Return rowRoot.Rows
                    Case 3 : Return rowGroup.Group.Rows.Intersect(columnGroup.Parent.Rows).ToArray()
                    Case Else : Return rowGroup.Parent.Rows.Intersect(columnGroup.Group.Rows).ToArray()
                End Select
            End Function

            Private Shared Function GroupingFieldLabels(fields As FormulaMatrix, headers As Boolean, label As String) As Object()
                Return Enumerable.Range(0, fields.Columns).
                    Select(Function(column) If(headers, If(fields.At(0, column), String.Empty),
                        CObj(label & (column + 1).ToString(CultureInfo.InvariantCulture)))).ToArray()
            End Function

        End Class

    End Class

End Namespace
