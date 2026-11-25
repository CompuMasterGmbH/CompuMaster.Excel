Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly SortableErrors As String() = {
                "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"
            }

            Private Function EvaluateSortFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                  formulaCell As ExcelCell) As Object
                Dim minimum = If(name = "SORTBY", 2, 1)
                Dim maximum = If(name = "SORTBY", 255, If(name = "UNIQUE", 3, 4))
                If arguments.Count < minimum OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")

                Dim input = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf input Is FormulaError Then Return input
                Dim matrix = AsMatrix(input)
                Select Case name
                    Case "SORT"
                        Dim key = NumberValue(OptionalSortArgument(arguments, 1, formulaCell, 1.0))
                        If TypeOf key Is FormulaError Then Return key
                        Dim order = NumberValue(OptionalSortArgument(arguments, 2, formulaCell, 1.0))
                        If TypeOf order Is FormulaError Then Return order
                        Dim direction = DirectCast(order, Double)
                        If direction <> 1.0 AndAlso direction <> -1.0 Then Return New FormulaError("#VALUE!")
                        Dim axis = LogicalOptionValue(OptionalSortArgument(arguments, 3, formulaCell, False))
                        If TypeOf axis Is FormulaError Then Return axis
                        Dim byRow = Not DirectCast(axis, Boolean)
                        Dim position = Math.Truncate(DirectCast(key, Double))
                        If position < 1 OrElse position > If(byRow, matrix.Columns, matrix.Rows) Then Return New FormulaError("#VALUE!")
                        Dim fixedIndex = CInt(position) - 1
                        Dim keys = Enumerable.Range(0, If(byRow, matrix.Rows, matrix.Columns)).
                            Select(Function(index) matrix.At(If(byRow, index, fixedIndex), If(byRow, fixedIndex, index))).ToArray()
                        Return SortMatrix(matrix, byRow, {(Values:=keys, Direction:=CInt(direction))})
                    Case "SORTBY"
                        Return SortMatrixByKeys(matrix, arguments, formulaCell)
                    Case Else
                        Dim axis = LogicalOptionValue(OptionalSortArgument(arguments, 1, formulaCell, False))
                        If TypeOf axis Is FormulaError Then Return axis
                        Dim exactlyOnce = LogicalOptionValue(OptionalSortArgument(arguments, 2, formulaCell, False))
                        If TypeOf exactlyOnce Is FormulaError Then Return exactlyOnce
                        Return UniqueMatrix(matrix, byRow:=Not DirectCast(axis, Boolean), exactlyOnce:=DirectCast(exactlyOnce, Boolean))
                End Select
            End Function

            Private Function OptionalSortArgument(arguments As IReadOnlyList(Of FormulaExpression), index As Integer,
                                                  formulaCell As ExcelCell, defaultValue As Object) As Object
                If index >= arguments.Count Then Return defaultValue
                Dim value = EvaluateExpression(arguments(index), formulaCell)
                Return If(value Is Nothing, defaultValue, Scalar(value))
            End Function

            Private Function SortMatrixByKeys(matrix As FormulaMatrix, arguments As IReadOnlyList(Of FormulaExpression),
                                             formulaCell As ExcelCell) As Object
                Dim keys As New List(Of (Values As Object(), Direction As Integer))
                Dim byRow As Boolean
                For index = 1 To arguments.Count - 1 Step 2
                    Dim input = EvaluateExpression(arguments(index), formulaCell)
                    If TypeOf input Is FormulaError Then Return input
                    Dim key = AsMatrix(input)
                    If index = 1 Then byRow = key.Columns = 1 AndAlso key.Rows = matrix.Rows
                    If (byRow AndAlso (key.Columns <> 1 OrElse key.Rows <> matrix.Rows)) OrElse
                       (Not byRow AndAlso (key.Rows <> 1 OrElse key.Columns <> matrix.Columns)) Then Return New FormulaError("#VALUE!")

                    Dim order = NumberValue(OptionalSortArgument(arguments, index + 1, formulaCell, 1.0))
                    If TypeOf order Is FormulaError Then Return order
                    Dim direction = Math.Floor(DirectCast(order, Double))
                    If direction <> 1.0 AndAlso direction <> -1.0 Then Return New FormulaError("#VALUE!")
                    keys.Add((key.Values, CInt(direction)))
                Next
                Return SortMatrix(matrix, byRow, keys)
            End Function

            Private Function SortMatrix(matrix As FormulaMatrix, byRow As Boolean,
                                        keys As IReadOnlyList(Of (Values As Object(), Direction As Integer))) As FormulaMatrix
                Dim textComparer = StringComparer.Create(ConversionCulture, ignoreCase:=True)
                Dim sortableKeys = keys.Select(Function(key) (Values:=key.Values.Select(AddressOf NormalizeSortKey).ToArray(),
                                                               key.Direction)).ToArray()
                Dim positions = Enumerable.Range(0, If(byRow, matrix.Rows, matrix.Columns)).ToArray()
                Array.Sort(positions, Function(left, right)
                                          For Each key In sortableKeys
                                              Dim first = key.Values(left)
                                              Dim second = key.Values(right)
                                              ' True blank cells follow all values in both sort directions.
                                              If (first Is Nothing) <> (second Is Nothing) Then Return If(first Is Nothing, 1, -1)
                                              If first Is Nothing Then Continue For
                                              Dim comparison = CompareSortValues(first, second, textComparer)
                                              If comparison <> 0 Then Return Math.Sign(comparison) * key.Direction
                                          Next
                                          Return left.CompareTo(right)
                                      End Function)
                Return SelectMatrixSlices(matrix, positions, byRow)
            End Function

            Private Shared Function CompareSortValues(first As Object, second As Object, textComparer As StringComparer) As Integer
                Dim firstError = TryCast(first, FormulaError)
                Dim secondError = TryCast(second, FormulaError)
                If firstError IsNot Nothing OrElse secondError IsNot Nothing Then
                    If firstError Is Nothing Then Return -1
                    If secondError Is Nothing Then Return 1
                    Return Array.IndexOf(SortableErrors, firstError.Code).CompareTo(Array.IndexOf(SortableErrors, secondError.Code))
                End If

                Dim rank = ComparisonRank(first)
                Dim comparison = rank.CompareTo(ComparisonRank(second))
                If comparison <> 0 Then Return comparison
                Select Case rank
                    Case 0 : Return DirectCast(first, Double).CompareTo(DirectCast(second, Double))
                    Case 1 : Return textComparer.Compare(DirectCast(first, String), DirectCast(second, String))
                    Case Else : Return DirectCast(first, Boolean).CompareTo(DirectCast(second, Boolean))
                End Select
            End Function

            Private Function UniqueMatrix(matrix As FormulaMatrix, byRow As Boolean, exactlyOnce As Boolean) As Object
                Dim failure = matrix.Values.OfType(Of FormulaError)().FirstOrDefault()
                If failure IsNot Nothing Then Return failure
                Dim comparer As New UniqueSliceComparer(StringComparer.Create(ConversionCulture, ignoreCase:=True))
                Dim slices As New Dictionary(Of Object(), (Position As Integer, Count As Integer))(comparer)
                Dim size = If(byRow, matrix.Rows, matrix.Columns)
                Dim width = If(byRow, matrix.Columns, matrix.Rows)
                For index = 0 To size - 1
                    Dim key(width - 1) As Object
                    For offset = 0 To width - 1
                        key(offset) = NormalizeSortKey(matrix.At(If(byRow, index, offset), If(byRow, offset, index)))
                    Next
                    Dim previous As (Position As Integer, Count As Integer)
                    If slices.TryGetValue(key, previous) Then
                        slices(key) = (previous.Position, previous.Count + 1)
                    Else
                        slices.Add(key, (index, 1))
                    End If
                Next

                Dim retained = slices.Values.Where(Function(item) Not exactlyOnce OrElse item.Count = 1).
                    Select(Function(item) item.Position).OrderBy(Function(index) index).ToArray()
                If retained.Length = 0 Then Return New FormulaError(CalcErrorCode)
                Return SelectMatrixSlices(matrix, retained, byRow)
            End Function

            Private Shared Function NormalizeSortKey(value As Object) As Object
                Dim failure = TryCast(value, FormulaError)
                If failure IsNot Nothing AndAlso Array.IndexOf(SortableErrors, failure.Code) < 0 Then
                    Throw New NotSupportedException("Ordering this Excel error type is outside the supported calculation scope.")
                End If
                Dim text = TryCast(value, String)
                If text Is Nothing Then Return value
                ' Excel's collation treats these ligatures and their expansions as equal.
                Return text.Replace("ß", "ss").Replace("ẞ", "SS").Replace("æ", "ae").Replace("Æ", "AE")
            End Function

            Private NotInheritable Class UniqueSliceComparer
                Implements IEqualityComparer(Of Object())

                Private ReadOnly _textComparer As StringComparer

                Friend Sub New(textComparer As StringComparer)
                    _textComparer = textComparer
                End Sub

                Private Function EqualsSlice(first As Object(), second As Object()) As Boolean Implements IEqualityComparer(Of Object()).Equals
                    For index = 0 To first.Length - 1
                        Dim left = first(index)
                        Dim right = second(index)
                        If TypeOf left Is String AndAlso TypeOf right Is String Then
                            If Not _textComparer.Equals(DirectCast(left, String), DirectCast(right, String)) Then Return False
                        ElseIf Not Object.Equals(left, right) Then
                            Return False
                        End If
                    Next
                    Return True
                End Function

                Private Function SliceHash(slice As Object()) As Integer Implements IEqualityComparer(Of Object()).GetHashCode
                    Dim hash As Integer
                    For Each element In slice
                        Dim elementHash = If(TypeOf element Is String, _textComparer.GetHashCode(DirectCast(element, String)),
                                             If(element?.GetHashCode(), 0))
                        hash = CInt((CLng(hash) * 31 + elementHash) And Integer.MaxValue)
                    Next
                    Return hash
                End Function
            End Class

        End Class

    End Class

End Namespace
