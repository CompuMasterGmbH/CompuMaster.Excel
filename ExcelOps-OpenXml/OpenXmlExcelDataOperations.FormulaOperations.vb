Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Partial Private NotInheritable Class FormulaCalculation

            Friend Function EvaluateUnary(operation As String, operand As Object) As Object
                If IsMultiValue(operand) Then
                    Dim matrix = AsMatrix(operand)
                    Return New FormulaMatrix(matrix.Rows, matrix.Columns,
                        matrix.Values.Select(Function(item) EvaluateUnary(operation, item)).ToArray())
                End If
                Dim value = NumberValue(Scalar(operand))
                If TypeOf value Is FormulaError Then Return value
                Dim number = DirectCast(value, Double)
                Select Case operation
                    Case "+" : Return number
                    Case "-" : Return -number
                    Case "%" : Return number / 100.0
                    Case Else : Throw New NotSupportedException("Unsupported unary operator " & operation & ".")
                End Select
            End Function

            Friend Function EvaluateBinary(operation As String, first As Object, second As Object) As Object
                If IsMultiValue(first) OrElse IsMultiValue(second) Then Return EvaluateArrayBinary(operation, AsMatrix(first), AsMatrix(second))
                first = Scalar(first)
                second = Scalar(second)
                If TypeOf first Is FormulaError Then Return first
                If TypeOf second Is FormulaError Then Return second
                Select Case operation
                    Case "&"
                        Return TextValue(first) & TextValue(second)
                    Case "=", "<>", "<", "<=", ">", ">="
                        Return EvaluateComparison(operation, first, second)
                End Select

                Dim left = NumberValue(first)
                Dim right = NumberValue(second)
                If TypeOf left Is FormulaError Then Return left
                If TypeOf right Is FormulaError Then Return right
                Dim a = DirectCast(left, Double)
                Dim b = DirectCast(right, Double)
                Select Case operation
                    Case "+" : Return FiniteNumber(a + b)
                    Case "-" : Return FiniteNumber(a - b)
                    Case "*" : Return FiniteNumber(a * b)
                    Case "/"
                        If b = 0.0 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber(a / b)
                    Case "^"
                        If a = 0.0 AndAlso b < 0.0 Then Return New FormulaError("#DIV/0!")
                        If a = 0.0 AndAlso b = 0.0 Then Return New FormulaError("#NUM!")
                        Return FiniteNumber(Math.Pow(a, b))
                    Case Else : Throw New NotSupportedException("Unsupported binary operator " & operation & ".")
                End Select
            End Function

            Friend Function EvaluateFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Select Case name.ToUpperInvariant()
                    Case "IF"
                        If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                        Dim condition = Scalar(EvaluateExpression(arguments(0), formulaCell))
                        If TypeOf condition Is String Then Return New FormulaError("#VALUE!")
                        Dim truth = NumberValue(condition)
                        If TypeOf truth Is FormulaError Then Return truth
                        If DirectCast(truth, Double) <> 0.0 Then Return EvaluateExpression(arguments(1), formulaCell)
                        If arguments.Count = 3 Then Return EvaluateExpression(arguments(2), formulaCell)
                        Return False
                    Case "ABS"
                        If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                        Dim operand = NumberValue(Scalar(EvaluateExpression(arguments(0), formulaCell)))
                        If TypeOf operand Is FormulaError Then Return operand
                        Return Math.Abs(DirectCast(operand, Double))
                    Case "ISREF"
                        If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                        Dim reference = TryCast(EvaluateExpression(arguments(0), formulaCell), FormulaArea)
                        Return reference IsNot Nothing AndAlso ResolveReference(reference)._worksheets.ContainsKey(reference.Sheet)
                    Case "ROW", "COLUMN", "ROWS", "COLUMNS"
                        Return EvaluateReferenceFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "DATE", "TIME", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "TODAY", "NOW"
                        Return EvaluateDateFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC", "INT", "MOD", "SIGN", "SQRT", "POWER"
                        Return EvaluateNumericFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "TRIM", "UPPER", "LOWER", "LEN", "LEFT", "RIGHT", "MID", "CONCATENATE"
                        Return EvaluateTextFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "IFERROR", "IFNA", "_XLFN.IFNA", "NOT", "AND", "OR", "ISBLANK", "ISNUMBER", "ISTEXT", "ISLOGICAL", "ISERROR", "ISNA", "NA"
                        Return EvaluateLogicalFunction(If(name.Equals("_xlfn.IFNA", StringComparison.OrdinalIgnoreCase), "IFNA", name.ToUpperInvariant()), arguments, formulaCell)
                    Case "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS"
                        Return EvaluateCriteriaFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "MATCH", "INDEX", "VLOOKUP", "XLOOKUP", "_XLFN.XLOOKUP", "XMATCH", "_XLFN.XMATCH"
                        Return EvaluateLookupFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "SEQUENCE", "_XLFN.SEQUENCE", "TRANSPOSE", "FILTER", "_XLFN.FILTER", "_XLFN._XLWS.FILTER"
                        Return EvaluateArrayFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "COUNTA", "COUNTBLANK"
                        Return CountValues(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "PRODUCT"
                        Return Aggregate(name.ToUpperInvariant(), arguments, formulaCell)
                    Case Else
                        Return EvaluateRegisteredFunction(name, arguments, formulaCell)
                End Select
            End Function

            Private Function EvaluateReferenceFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count = 0 Then
                    If name = "ROW" Then Return CDbl(formulaCell.RowIndex + 1)
                    If name = "COLUMN" Then Return CDbl(formulaCell.ColumnIndex + 1)
                End If
                If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")

                ' Keep the reference as an area so address queries do not read cell values.
                Dim argument = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf argument Is FormulaError Then Return argument
                Dim area = TryCast(argument, FormulaArea)
                If area Is Nothing Then
                    Dim matrix = TryCast(argument, FormulaMatrix)
                    If name = "ROWS" Then Return If(matrix Is Nothing, 1.0, CDbl(matrix.Rows))
                    If name = "COLUMNS" Then Return If(matrix Is Nothing, 1.0, CDbl(matrix.Columns))
                    Return New FormulaError("#VALUE!")
                End If
                If Not ResolveReference(area)._worksheets.ContainsKey(area.Sheet) Then Return New FormulaError("#REF!")

                Select Case name
                    Case "ROW" : Return CDbl(area.FirstRow + 1)
                    Case "COLUMN" : Return CDbl(area.FirstColumn + 1)
                    Case "ROWS" : Return CDbl(area.LastRow - area.FirstRow + 1)
                    Case Else : Return CDbl(area.LastColumn - area.FirstColumn + 1)
                End Select
            End Function

            Private Function Aggregate(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count = 0 Then Return New FormulaError("#VALUE!")
                Dim sum As Double
                Dim product = 1.0
                Dim minimum As Double = Double.PositiveInfinity
                Dim maximum As Double = Double.NegativeInfinity
                Dim count As Integer
                Dim firstError As FormulaError = Nothing
                For Each argument In arguments
                    Dim value = EvaluateExpression(argument, formulaCell)
                    Dim area = TryCast(value, FormulaArea)
                    Dim matrix = TryCast(value, FormulaMatrix)
                    Dim values As IEnumerable(Of Object)
                    If area IsNot Nothing Then
                        values = RangeValues(area)
                    ElseIf matrix IsNot Nothing Then
                        values = matrix.Values
                    Else
                        values = {value}
                    End If
                    For Each item In values
                        If (area IsNot Nothing OrElse matrix IsNot Nothing) AndAlso
                            Not TypeOf item Is Double AndAlso Not TypeOf item Is FormulaError Then Continue For
                        If item Is Nothing Then Continue For
                        Dim number = NumberValue(item)
                        If TypeOf number Is FormulaError Then
                            If name <> "COUNT" AndAlso firstError Is Nothing Then firstError = DirectCast(number, FormulaError)
                            Continue For
                        End If
                        Dim numeric = DirectCast(number, Double)
                        count += 1
                        sum += numeric
                        product *= numeric
                        minimum = Math.Min(minimum, numeric)
                        maximum = Math.Max(maximum, numeric)
                    Next
                Next

                If firstError IsNot Nothing Then Return firstError
                Select Case name
                    Case "SUM" : Return FiniteNumber(sum)
                    Case "PRODUCT" : Return If(count = 0, 0.0, FiniteNumber(product))
                    Case "AVERAGE"
                        If count = 0 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber(sum / count)
                    Case "MIN" : Return If(count = 0, 0.0, minimum)
                    Case "MAX" : Return If(count = 0, 0.0, maximum)
                    Case Else : Return CDbl(count)
                End Select
            End Function

            Friend Shared Function FiniteNumber(number As Double) As Object
                If Double.IsInfinity(number) OrElse Double.IsNaN(number) Then Return New FormulaError("#NUM!")
                Return number
            End Function

            Private ReadOnly Property ConversionCulture As CultureInfo
                Get
                    Return If(_owner.LoadOptions.FormulaCulture, CultureInfo.InvariantCulture)
                End Get
            End Property

            Private Function NumberValue(value As Object) As Object
                If value Is Nothing Then Return 0.0
                If TypeOf value Is FormulaError OrElse TypeOf value Is Double Then Return value
                If TypeOf value Is Boolean Then Return If(DirectCast(value, Boolean), 1.0, 0.0)
                If TypeOf value Is String Then
                    Dim parsed As Double
                    If Double.TryParse(DirectCast(value, String), NumberStyles.Float, ConversionCulture, parsed) Then Return FiniteNumber(parsed)
                    Return New FormulaError("#VALUE!")
                End If
                Throw New NotSupportedException("The referenced value type is outside the supported calculation scope.")
            End Function

            Private Function TextValue(value As Object) As String
                If value Is Nothing Then Return String.Empty
                If TypeOf value Is Boolean Then
                    If ConversionCulture.TwoLetterISOLanguageName = "de" Then Return If(DirectCast(value, Boolean), "WAHR", "FALSCH")
                    Return If(DirectCast(value, Boolean), "TRUE", "FALSE")
                End If
                If TypeOf value Is Double Then Return DirectCast(value, Double).ToString("G15", ConversionCulture)
                Return Convert.ToString(value, ConversionCulture)
            End Function

            Private Shared Function EvaluateComparison(operation As String, first As Object, second As Object) As Boolean
                If first Is Nothing Then first = ComparableBlank(second)
                If second Is Nothing Then second = ComparableBlank(first)
                Dim firstType = ComparisonRank(first)
                Dim secondType = ComparisonRank(second)
                Dim comparison As Integer
                If firstType <> secondType Then
                    comparison = firstType.CompareTo(secondType)
                ElseIf firstType = 0 Then
                    comparison = DirectCast(first, Double).CompareTo(DirectCast(second, Double))
                ElseIf firstType = 1 Then
                    comparison = String.Compare(DirectCast(first, String), DirectCast(second, String), StringComparison.OrdinalIgnoreCase)
                Else
                    comparison = DirectCast(first, Boolean).CompareTo(DirectCast(second, Boolean))
                End If
                Select Case operation
                    Case "=" : Return comparison = 0
                    Case "<>" : Return comparison <> 0
                    Case "<" : Return comparison < 0
                    Case "<=" : Return comparison <= 0
                    Case ">" : Return comparison > 0
                    Case Else : Return comparison >= 0
                End Select
            End Function

            Private Shared Function ComparableBlank(other As Object) As Object
                If TypeOf other Is String Then Return String.Empty
                If TypeOf other Is Boolean Then Return False
                Return 0.0
            End Function

            Private Shared Function ComparisonRank(value As Object) As Integer
                If TypeOf value Is Double Then Return 0
                If TypeOf value Is String Then Return 1
                If TypeOf value Is Boolean Then Return 2
                Throw New NotSupportedException("This value type is not supported in formula comparisons.")
            End Function

        End Class

    End Class

End Namespace
