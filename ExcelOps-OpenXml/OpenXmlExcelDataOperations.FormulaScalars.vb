Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation
            Private Shared ReadOnly DuplicateSpaces As New Regex(" {2,}", RegexOptions.CultureInvariant)

            Private Function ReadNumbers(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell,
                                         ByRef values As Double()) As FormulaError
                values = New Double(arguments.Count - 1) {}
                Dim failure As FormulaError = Nothing
                For index = 0 To arguments.Count - 1
                    Dim item = NumberValue(Scalar(EvaluateExpression(arguments(index), formulaCell)))
                    If TypeOf item Is FormulaError Then
                        failure = If(failure, DirectCast(item, FormulaError))
                    Else
                        values(index) = DirectCast(item, Double)
                    End If
                Next
                Return failure
            End Function

            Private Function EvaluateNumericFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                Select Case name
                    Case "INT", "SIGN", "SQRT"
                        If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                    Case "TRUNC"
                        If arguments.Count < 1 OrElse arguments.Count > 2 Then Return New FormulaError("#VALUE!")
                    Case Else
                        If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                End Select
                Dim values As Double() = Nothing
                Dim failure = ReadNumbers(arguments, formulaCell, values)
                If failure IsNot Nothing Then Return failure
                Dim number = values(0)
                Select Case name
                    Case "INT" : Return Math.Floor(number)
                    Case "SIGN" : Return CDbl(Math.Sign(number))
                    Case "SQRT" : Return FiniteNumber(Math.Sqrt(number))
                    Case "POWER" : Return EvaluateBinary("^", number, values(1))
                    Case "MOD"
                        Dim divisor = values(1)
                        If divisor = 0.0 Then Return New FormulaError("#DIV/0!")
                        Dim remainder = number Mod divisor
                        If remainder <> 0.0 AndAlso Math.Sign(remainder) <> Math.Sign(divisor) Then remainder += divisor
                        Return FiniteNumber(remainder)
                    Case Else
                        Return RoundAt(number, If(values.Length = 2, Math.Truncate(values(1)), 0.0), name)
                End Select
            End Function

            Private Shared Function RoundAt(number As Double, decimals As Double, name As String) As Object
                If number = 0.0 OrElse decimals >= 309.0 Then Return number
                If decimals <= -309.0 Then Return If(name = "ROUNDUP", CType(New FormulaError("#NUM!"), Object), 0.0)
                Dim scale = Math.Pow(10, -decimals)
                Dim scaled = number / scale
                If Double.IsInfinity(scaled) Then Return number
                If scaled = 0.0 AndAlso name = "ROUNDUP" Then Return FiniteNumber(Math.Sign(number) * scale)
                ' Excel resolves decimal midpoint artifacts such as ROUND(1.005, 2).
                scaled = Double.Parse(scaled.ToString("G15", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
                Select Case name
                    Case "ROUND" : scaled = Math.Round(scaled, MidpointRounding.AwayFromZero)
                    Case "ROUNDUP" : scaled = Math.Sign(scaled) * Math.Ceiling(Math.Abs(scaled))
                    Case Else : scaled = Math.Truncate(scaled)
                End Select
                Return FiniteNumber(scaled * scale)
            End Function

            Private Function EvaluateTextFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                  formulaCell As ExcelCell) As Object
                If name = "CONCATENATE" Then
                    If arguments.Count = 0 Then Return New FormulaError("#VALUE!")
                    Dim combined As New StringBuilder()
                    For Each argument In arguments
                        Dim item = Scalar(EvaluateExpression(argument, formulaCell))
                        If TypeOf item Is FormulaError Then Return item
                        combined.Append(TextValue(item))
                        If combined.Length > 32767 Then Return New FormulaError("#VALUE!")
                    Next
                    Return combined.ToString()
                End If

                Dim minimum = If(name = "MID", 3, 1)
                Dim maximum = If(name = "LEFT" OrElse name = "RIGHT", 2, minimum)
                If arguments.Count < minimum OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")
                Dim value = Scalar(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf value Is FormulaError Then Return value
                Dim contents = TextValue(value)
                Select Case name
                    Case "TRIM" : Return DuplicateSpaces.Replace(contents.Trim(" "c), " ")
                    Case "UPPER" : Return contents.ToUpperInvariant()
                    Case "LOWER" : Return contents.ToLowerInvariant()
                    Case "LEN" : Return CDbl(contents.Length)
                End Select

                Dim position As Double = 0.0
                Dim length As Double = 1.0
                If name = "MID" Then
                    Dim start = NumberValue(Scalar(EvaluateExpression(arguments(1), formulaCell)))
                    If TypeOf start Is FormulaError Then Return start
                    position = Math.Truncate(DirectCast(start, Double)) - 1.0
                    If position < 0.0 Then Return New FormulaError("#VALUE!")
                End If
                If arguments.Count >= 2 Then
                    Dim amount = NumberValue(Scalar(EvaluateExpression(arguments(arguments.Count - 1), formulaCell)))
                    If TypeOf amount Is FormulaError Then Return amount
                    length = Math.Truncate(DirectCast(amount, Double))
                    If length < 0.0 Then Return New FormulaError("#VALUE!")
                End If
                If name = "RIGHT" Then position = Math.Max(0.0, contents.Length - length)
                If position >= contents.Length Then Return String.Empty
                Return contents.Substring(CInt(position), CInt(Math.Min(length, contents.Length - position)))
            End Function

            Private Function EvaluateLogicalFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                If name = "NA" Then Return New FormulaError(If(arguments.Count = 0, "#N/A", "#VALUE!"))
                If name = "TRUE" OrElse name = "FALSE" Then
                    If arguments.Count <> 0 Then Return New FormulaError("#VALUE!")
                    Return name = "TRUE"
                End If
                If name = "AND" OrElse name = "OR" OrElse name = "XOR" Then Return CombineLogical(name, arguments, formulaCell)
                Dim expectsFallback = name = "IFERROR" OrElse name = "IFNA"
                If arguments.Count <> If(expectsFallback, 2, 1) Then Return New FormulaError("#VALUE!")
                Dim value = Scalar(EvaluateExpression(arguments(0), formulaCell))
                Dim failure = TryCast(value, FormulaError)
                Select Case name
                    Case "IFERROR", "IFNA"
                        If failure Is Nothing OrElse (name = "IFNA" AndAlso failure.Code <> "#N/A") Then Return value
                        Return EvaluateExpression(arguments(1), formulaCell)
                    Case "ISBLANK" : Return value Is Nothing
                    Case "ISNUMBER" : Return TypeOf value Is Double
                    Case "ISTEXT" : Return TypeOf value Is String
                    Case "ISLOGICAL" : Return TypeOf value Is Boolean
                    Case "ISERROR" : Return failure IsNot Nothing
                    Case "ISNA" : Return failure?.Code = "#N/A"
                    Case Else
                        If TypeOf value Is String Then Return New FormulaError("#VALUE!")
                        Dim truth = NumberValue(value)
                        If TypeOf truth Is FormulaError Then Return truth
                        Return DirectCast(truth, Double) = 0.0
                End Select
            End Function

            Private Function LogicalOptionValue(value As Object) As Object
                If TypeOf value Is String Then
                    Dim text = DirectCast(value, String)
                    If text.Equals(TextValue(True), StringComparison.OrdinalIgnoreCase) Then Return True
                    If text.Equals(TextValue(False), StringComparison.OrdinalIgnoreCase) Then Return False
                    Return New FormulaError("#VALUE!")
                End If
                Dim number = NumberValue(value)
                If TypeOf number Is FormulaError Then Return number
                Return DirectCast(number, Double) <> 0.0
            End Function

            Private Function CombineLogical(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                            formulaCell As ExcelCell) As Object
                Dim result = name = "AND"
                Dim exclusive = name = "XOR"
                Dim found = False
                Dim firstError As FormulaError = Nothing
                For Each argument In arguments
                    Dim value = EvaluateExpression(argument, formulaCell)
                    Dim area = TryCast(value, FormulaArea)
                    Dim matrix = If(exclusive, TryCast(value, FormulaMatrix), Nothing)
                    Dim values As IEnumerable(Of Object)
                    If area IsNot Nothing Then
                        values = RangeValues(area)
                    ElseIf matrix IsNot Nothing Then
                        values = matrix.Values
                    Else
                        values = {value}
                    End If
                    For Each item In values
                        If (area IsNot Nothing OrElse matrix IsNot Nothing) AndAlso (item Is Nothing OrElse TypeOf item Is String) Then Continue For
                        Dim truth As Object
                        If exclusive Then
                            truth = LogicalOptionValue(item)
                            If TypeOf item Is String AndAlso TypeOf truth Is FormulaError Then Continue For
                        Else
                            Dim number = If(TypeOf item Is String, New FormulaError("#VALUE!"), NumberValue(item))
                            truth = If(TypeOf number Is FormulaError, number, DirectCast(number, Double) <> 0.0)
                        End If
                        If TypeOf truth Is FormulaError Then
                            firstError = If(firstError, DirectCast(truth, FormulaError))
                            Continue For
                        End If
                        Dim isTrue = DirectCast(truth, Boolean)
                        Select Case name
                            Case "AND" : result = result AndAlso isTrue
                            Case "OR" : result = result OrElse isTrue
                            Case Else : result = result Xor isTrue
                        End Select
                        found = True
                    Next
                Next
                If firstError IsNot Nothing Then Return firstError
                Return If(found, CType(result, Object), New FormulaError("#VALUE!"))
            End Function

            Private Function CountValues(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                         formulaCell As ExcelCell) As Object
                If arguments.Count = 0 OrElse (name = "COUNTBLANK" AndAlso arguments.Count <> 1) Then Return New FormulaError("#VALUE!")
                Dim count As Double
                For Each argument In arguments
                    Dim value = EvaluateExpression(argument, formulaCell)
                    Dim area = TryCast(value, FormulaArea)
                    If area Is Nothing Then
                        If name = "COUNTBLANK" Then Return If(TypeOf value Is FormulaError, value, New FormulaError("#VALUE!"))
                        count += 1.0
                    Else
                        If Not ResolveReference(area)._worksheets.ContainsKey(area.Sheet) Then Return New FormulaError("#REF!")
                        Dim values = RangeValues(area)
                        If name = "COUNTA" Then
                            count += values.Count(Function(item) item IsNot Nothing)
                        Else
                            Dim nonBlank = values.Count(Function(item) item IsNot Nothing AndAlso (Not TypeOf item Is String OrElse DirectCast(item, String).Length > 0))
                            count += CDbl(area.LastRow - area.FirstRow + 1) * (area.LastColumn - area.FirstColumn + 1) - nonBlank
                        End If
                    End If
                Next
                Return count
            End Function
        End Class

    End Class

End Namespace
