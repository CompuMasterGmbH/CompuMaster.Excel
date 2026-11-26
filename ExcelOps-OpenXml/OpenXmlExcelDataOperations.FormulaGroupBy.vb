Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private NotInheritable Class WorksheetGroup
                Friend ReadOnly Rows As Integer()
                Friend ReadOnly Results As Lazy(Of Object())
                Friend ReadOnly Children As New List(Of WorksheetGroup)
                Friend DisplayRow As Integer

                Friend Sub New(rows As Integer(), aggregate As Func(Of Integer(), Object()))
                    Me.Rows = rows
                    Results = New Lazy(Of Object())(Function() aggregate(rows))
                    DisplayRow = rows(0)
                End Sub
            End Class

            Private Function EvaluateGroupBy(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 3 OrElse arguments.Count > 8 Then Return New FormulaError("#VALUE!")
                Dim inputs = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                Dim failure = inputs.OfType(Of FormulaError)().FirstOrDefault()
                If failure IsNot Nothing Then Return failure
                Dim fields = AsMatrix(inputs(0))
                Dim values = AsMatrix(inputs(1))
                If fields.Rows <> values.Rows Then Return New FormulaError("#VALUE!")

                Dim functionMatrix = AsMatrix(inputs(2))
                Dim functions = functionMatrix.Values.OfType(Of FormulaLambda)().ToArray()
                If functions.Length <> functionMatrix.Values.Length OrElse (functionMatrix.Rows > 1 AndAlso functionMatrix.Columns > 1) OrElse
                    (values.Columns > 1 AndAlso functions.Length > 1 AndAlso values.Columns <> functions.Length) OrElse
                    functions.Any(Function(item) Not item.Accepts(1) AndAlso Not item.Accepts(2)) Then Return New FormulaError("#VALUE!")
                Dim vertical = functionMatrix.Rows > 1
                Dim outputValues = Math.Max(values.Columns, functions.Length)

                Dim headers = GroupOption(inputs.ElementAtOrDefault(3), -1, 0, 3)
                Dim depth = GroupOption(inputs.ElementAtOrDefault(4), 1, -fields.Columns, fields.Columns)
                Dim relationship = GroupOption(inputs.ElementAtOrDefault(7), 0, 0, 1)
                If Not headers.HasValue OrElse Not depth.HasValue OrElse Not relationship.HasValue OrElse
                    (relationship.Value = 1 AndAlso Math.Abs(depth.Value) > 1) Then Return New FormulaError("#VALUE!")
                Dim sort = ReadGroupingSort(inputs.ElementAtOrDefault(5), fields.Columns, outputValues)
                If sort Is Nothing Then Return New FormulaError("#VALUE!")
                Dim hasHeaderRow = headers.Value = 1 OrElse headers.Value = 3 OrElse (headers.Value = -1 AndAlso
                    values.Rows > 1 AndAlso TypeOf values.At(0, 0) Is String AndAlso TypeOf values.At(1, 0) Is Double)
                Dim keep = inputs.ElementAtOrDefault(6)
                Dim mask = If(keep Is Nothing, Nothing, AsMatrix(keep))
                If mask IsNot Nothing AndAlso (mask.Columns <> 1 OrElse mask.Rows <> fields.Rows OrElse
                    mask.Values.Any(Function(item) item IsNot Nothing AndAlso Not TypeOf item Is Boolean AndAlso Not TypeOf item Is Double)) Then
                    Return New FormulaError("#VALUE!")
                End If
                Dim rows = Enumerable.Range(If(hasHeaderRow, 1, 0), fields.Rows - If(hasHeaderRow, 1, 0)).
                    Where(Function(index) mask Is Nothing OrElse GroupingFilterIncludes(mask.Values(index))).ToArray()
                If rows.Length = 0 Then Return New FormulaError("#VALUE!")

                Dim allValues = Enumerable.Range(0, values.Columns).
                    Select(Function(column) GroupValueArray(values, rows, column)).ToArray()
                Dim aggregate As Func(Of Integer(), Object()) =
                    Function(members) CalculateGroupValues(values, functions, members, allValues, formulaCell)
                Dim root As New WorksheetGroup(rows, aggregate)
                Dim comparer = StringComparer.Create(ConversionCulture, ignoreCase:=True)
                Dim keys As New UniqueSliceComparer(comparer)
                If relationship.Value = 0 Then
                    BuildGroupHierarchy(root, fields, 0, keys, comparer, sort, aggregate)
                Else
                    fields = CanonicalizeTableGroupingFields(fields, rows, keys)
                    Dim groups = root.Rows.GroupBy(Function(index)
                                                      Return Enumerable.Range(0, fields.Columns).
                                                          Select(Function(column) GroupingEqualityKey(fields.At(index, column))).ToArray()
                                                  End Function, keys)
                    root.Children.AddRange(groups.Select(Function(group) New WorksheetGroup(group.ToArray(), aggregate)))
                    SortWorksheetGroups(root.Children, fields, sort, comparer)
                End If

                Dim rowsOut As New List(Of Object())
                Dim labels = Enumerable.Range(0, functions.Length).Select(Function(index) GroupingLambdaLabel(functions(index), index)).ToArray()
                Dim displayHeaders = headers.Value = 2 OrElse headers.Value = 3
                Dim german = ConversionCulture.TwoLetterISOLanguageName = "de"
                If displayHeaders Then
                    Dim fieldLabels = Enumerable.Range(0, fields.Columns).
                        Select(Function(column) If(hasHeaderRow, If(fields.At(0, column), String.Empty),
                            CObj(If(german, "Zeilenfeld ", "Row Field ") & (column + 1).ToString(CultureInfo.InvariantCulture))))
                    Dim valueLabels = Enumerable.Range(0, If(vertical, 1, outputValues)).
                        Select(Function(column) If(hasHeaderRow, If(values.At(0, Math.Min(column, values.Columns - 1)), String.Empty),
                            CObj(If(german, "Wert ", "Value ") & (column + 1).ToString(CultureInfo.InvariantCulture))))
                    If vertical Then fieldLabels = fieldLabels.Concat({CObj(String.Empty)})
                    rowsOut.Add(fieldLabels.Concat(valueLabels).ToArray())
                End If
                If Not vertical AndAlso functions.Length > 1 Then
                    rowsOut.Add(Enumerable.Repeat(CObj(String.Empty), fields.Columns).Concat(labels.Cast(Of Object)()).ToArray())
                End If
                Dim levels = If(relationship.Value = 1, 1, fields.Columns)
                Dim totalLabel = If(Math.Abs(depth.Value) <= 1, If(german, "Gesamt", "Total"), If(german, "Gesamtergebnis", "Grand Total"))
                EmitWorksheetGroup(root, fields, New Object(fields.Columns - 1) {}, 0, levels, depth.Value,
                                   totalLabel, vertical, labels, rowsOut)
                Dim columns = fields.Columns + If(vertical, 2, outputValues)
                Dim flattened = AllocateMatrixValues(rowsOut.Count, columns)
                For row = 0 To rowsOut.Count - 1
                    Array.Copy(rowsOut(row), 0, flattened, row * columns, columns)
                Next
                Return New FormulaMatrix(rowsOut.Count, columns, flattened)
            End Function

            Private Function GroupOption(input As Object, defaultValue As Integer, minimum As Integer, maximum As Integer) As Integer?
                If input Is Nothing Then Return defaultValue
                If IsMultiValue(input) Then Return Nothing
                Dim value = NumberValue(Scalar(input))
                If Not TypeOf value Is Double Then Return Nothing
                Dim number = DirectCast(value, Double)
                If number <> Math.Truncate(number) OrElse number < minimum OrElse number > maximum Then Return Nothing
                Return CInt(number)
            End Function

            Private Function ReadGroupingSort(input As Object, fields As Integer, values As Integer) As Integer()
                If input Is Nothing Then Return Enumerable.Range(1, fields).ToArray()
                Dim provided = AsMatrix(input)
                If provided.Rows > 1 AndAlso provided.Columns > 1 Then Return Nothing
                Dim maximum = fields + If(provided.Values.Length = 1, values, 0)
                Dim keys As New List(Of Integer)
                For Each value In provided.Values
                    Dim optionValue = GroupOption(value, 0, -maximum, maximum)
                    If Not optionValue.HasValue OrElse optionValue.Value = 0 Then Return Nothing
                    keys.Add(optionValue.Value)
                Next
                Return keys.ToArray()
            End Function

            Private Shared Function GroupingFilterIncludes(value As Object) As Boolean
                If TypeOf value Is Boolean Then Return DirectCast(value, Boolean)
                Return TypeOf value Is Double AndAlso DirectCast(value, Double) <> 0.0
            End Function

            Private Shared Function GroupingEqualityKey(value As Object) As Object
                If TypeOf value Is FormulaError Then Return ValueTuple.Create(DirectCast(value, FormulaError).Code)
                Return NormalizeSortKey(value)
            End Function

            Private Shared Function CanonicalizeTableGroupingFields(fields As FormulaMatrix, rows As Integer(),
                                                                    comparer As UniqueSliceComparer) As FormulaMatrix
                Dim values = DirectCast(fields.Values.Clone(), Object())
                For column = 0 To fields.Columns - 1
                    Dim originals As New Dictionary(Of Object(), Object)(comparer)
                    For Each row In rows
                        Dim position = row * fields.Columns + column
                        Dim key = {GroupingEqualityKey(values(position))}
                        Dim original As Object = Nothing
                        If originals.TryGetValue(key, original) Then
                            values(position) = original
                        Else
                            originals.Add(key, values(position))
                        End If
                    Next
                Next
                Return New FormulaMatrix(fields.Rows, fields.Columns, values)
            End Function

            Private Sub BuildGroupHierarchy(parent As WorksheetGroup, fields As FormulaMatrix, column As Integer, keys As UniqueSliceComparer,
                                            comparer As StringComparer, sort As Integer(), aggregate As Func(Of Integer(), Object()))
                If column >= fields.Columns Then Return
                If Not HasFormulaStackSpace() Then Throw New NotSupportedException("The GROUPBY hierarchy exceeds the supported evaluation depth.")
                Dim groups = parent.Rows.GroupBy(Function(index) {GroupingEqualityKey(fields.At(index, column))}, keys)
                parent.Children.AddRange(groups.Select(Function(group) New WorksheetGroup(group.ToArray(), aggregate)))
                For Each child In parent.Children
                    BuildGroupHierarchy(child, fields, column + 1, keys, comparer, sort, aggregate)
                Next
                Dim valueSort = sort.Length = 1 AndAlso Math.Abs(sort(0)) > fields.Columns
                Dim fieldSort = sort.FirstOrDefault(Function(key) Math.Abs(key) = column + 1)
                Dim keyOrder = If(valueSort, sort(0), If(fieldSort = 0, column + 1, fieldSort))
                If valueSort Then
                    SortWorksheetGroups(parent.Children, fields, {column + 1}, comparer)
                    parent.DisplayRow = parent.Children(0).DisplayRow
                End If
                SortWorksheetGroups(parent.Children, fields, {keyOrder}, comparer)
                If Not valueSort Then parent.DisplayRow = parent.Children(0).DisplayRow
            End Sub

            Private Shared Sub SortWorksheetGroups(groups As List(Of WorksheetGroup), fields As FormulaMatrix,
                                                  order As IReadOnlyList(Of Integer), comparer As StringComparer)
                groups.Sort(Function(left, right)
                                For Each key In order
                                    Dim index = Math.Abs(key) - 1
                                    Dim a = If(index < fields.Columns, fields.At(left.DisplayRow, index), left.Results.Value(index - fields.Columns))
                                    Dim b = If(index < fields.Columns, fields.At(right.DisplayRow, index), right.Results.Value(index - fields.Columns))
                                    If (a Is Nothing) <> (b Is Nothing) Then Return If(a Is Nothing, 1, -1)
                                    If a Is Nothing Then Continue For
                                    Dim comparison = CompareSortValues(NormalizeSortKey(a), NormalizeSortKey(b), comparer)
                                    If comparison <> 0 Then Return Math.Sign(comparison) * Math.Sign(key)
                                Next
                                Return left.Rows(0).CompareTo(right.Rows(0))
                            End Function)
            End Sub

            Private Shared Function GroupValueArray(values As FormulaMatrix, rows As Integer(), column As Integer) As FormulaMatrix
                Return New FormulaMatrix(rows.Length, 1, rows.Select(Function(index) If(values.At(index, column), 0.0)).ToArray())
            End Function

            Private Function CalculateGroupValues(values As FormulaMatrix, functions As FormulaLambda(), rows As Integer(),
                                                  all As FormulaMatrix(), formulaCell As ExcelCell) As Object()
                Dim size = Math.Max(values.Columns, functions.Length)
                Dim result(size - 1) As Object
                For index = 0 To size - 1
                    Dim column = If(values.Columns = 1, 0, index)
                    Dim lambda = functions(If(functions.Length = 1, 0, index))
                    Dim subset = GroupValueArray(values, rows, column)
                    Dim requiresTotal = If(lambda.NativeFunction Is Nothing, lambda.Parameters.Count = 2,
                                           lambda.NativeFunction.Equals("PERCENTOF", StringComparison.OrdinalIgnoreCase))
                    Dim args = If(requiresTotal, {CObj(subset), all(column)}, {CObj(subset)})
                    Dim value = CallFormulaLambda(lambda, args, formulaCell)
                    result(index) = If(IsMultiValue(value) OrElse TypeOf value Is FormulaLambda,
                        CObj(New FormulaError("#VALUE!")), If(Scalar(value), 0.0))
                Next
                Return result
            End Function

            Private Function GroupingLambdaLabel(lambda As FormulaLambda, index As Integer) As String
                If lambda.NativeFunction IsNot Nothing Then
                    Dim germanName As String = Nothing
                    If ConversionCulture.TwoLetterISOLanguageName = "de" AndAlso GermanFormulaNames.Value.TryGetValue(lambda.NativeFunction, germanName) Then Return germanName
                    Return lambda.NativeFunction.ToUpperInvariant()
                End If
                Return If(ConversionCulture.TwoLetterISOLanguageName = "de", "BENUTZERDEFINIERT", "CUSTOM") &
                    (index + 1).ToString(CultureInfo.InvariantCulture)
            End Function

            Private Sub EmitWorksheetGroup(node As WorksheetGroup, fields As FormulaMatrix, prefix As Object(), level As Integer,
                                           levels As Integer, depth As Integer, totalLabel As String, vertical As Boolean,
                                           labels As String(), output As List(Of Object()))
                Dim emit = level = levels OrElse level < Math.Abs(depth)
                If emit AndAlso (depth < 0 OrElse level = levels) Then
                    AddGroupedRow(node, prefix, level, totalLabel, vertical, labels, output)
                End If
                For Each child In node.Children
                    Dim keys = DirectCast(prefix.Clone(), Object())
                    For i = level To If(levels = 1, fields.Columns - 1, level)
                        keys(i) = If(fields.At(child.DisplayRow, i), 0.0)
                    Next
                    EmitWorksheetGroup(child, fields, keys, level + 1, levels, depth, totalLabel, vertical, labels, output)
                Next
                If emit AndAlso depth > 0 AndAlso level < levels Then
                    AddGroupedRow(node, prefix, level, totalLabel, vertical, labels, output)
                End If
            End Sub

            Private Shared Sub AddGroupedRow(node As WorksheetGroup, prefix As Object(), level As Integer, totalLabel As String, vertical As Boolean,
                                             labels As String(), output As List(Of Object()))
                Dim fields = prefix.Select(Function(item) If(item, String.Empty)).ToArray()
                If level = 0 Then fields(0) = totalLabel
                If vertical Then
                    For i = 0 To labels.Length - 1
                        output.Add(fields.Concat({CObj(labels(i)), node.Results.Value(i)}).ToArray())
                    Next
                Else
                    output.Add(fields.Concat(node.Results.Value).ToArray())
                End If
            End Sub

        End Class

    End Class

End Namespace
