Imports System.Text

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const MaximumTextLength As Integer = 32767

            Private Function EvaluateTextJoin(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                              formulaCell As ExcelCell) As Object
                Dim joining = name = "TEXTJOIN"
                Dim firstValue = If(joining, 2, 0)
                If arguments.Count <= firstValue OrElse arguments.Count > If(joining, 254, 253) Then Return New FormulaError("#VALUE!")

                Dim separators As String() = {String.Empty}
                Dim skipEmpty As Boolean
                If joining Then
                    Dim delimiters = AsMatrix(EvaluateExpression(arguments(0), formulaCell)).Values
                    Dim failure = delimiters.OfType(Of FormulaError)().FirstOrDefault()
                    If failure IsNot Nothing Then Return failure
                    separators = delimiters.Select(Function(item) TextValue(item)).ToArray()

                    Dim optionValue = LogicalOptionValue(Scalar(EvaluateExpression(arguments(1), formulaCell)))
                    If TypeOf optionValue Is FormulaError Then Return optionValue
                    skipEmpty = DirectCast(optionValue, Boolean)
                End If

                Dim separatorIndex As Integer
                Dim hasValue As Boolean
                Dim result As New StringBuilder()
                For index = firstValue To arguments.Count - 1
                    Dim values = AsMatrix(EvaluateExpression(arguments(index), formulaCell)).Values
                    For Each value In values
                        If TypeOf value Is FormulaError Then Return value
                        Dim text = TextValue(value)
                        If skipEmpty AndAlso text.Length = 0 Then Continue For
                        Dim separator = If(hasValue, separators(separatorIndex), String.Empty)
                        If result.Length + separator.Length + text.Length > MaximumTextLength Then Return New FormulaError(CalcErrorCode)
                        If hasValue Then separatorIndex = (separatorIndex + 1) Mod separators.Length
                        result.Append(separator).Append(text)
                        hasValue = True
                    Next
                Next
                Return result.ToString()
            End Function

            Private Function EvaluateTextTransform(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                    formulaCell As ExcelCell) As Object
                Select Case name
                    Case "EXACT", "REPT"
                        If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                    Case "REPLACE"
                        If arguments.Count <> 4 Then Return New FormulaError("#VALUE!")
                    Case "SUBSTITUTE"
                        If arguments.Count < 3 OrElse arguments.Count > 4 Then Return New FormulaError("#VALUE!")
                End Select

                Dim values(arguments.Count - 1) As Object
                For index = 0 To values.Length - 1
                    values(index) = Scalar(EvaluateExpression(arguments(index), formulaCell))
                    If TypeOf values(index) Is FormulaError Then Return values(index)
                Next
                Dim contents = TextValue(values(0))
                Select Case name
                    Case "EXACT"
                        Return String.Equals(contents, TextValue(values(1)), StringComparison.Ordinal)
                    Case "REPT"
                        Return RepeatText(contents, NumberValue(values(1)))
                    Case "REPLACE"
                        Return ReplaceTextAt(contents, NumberValue(values(1)), NumberValue(values(2)), TextValue(values(3)))
                    Case Else
                        Dim occurrence = If(values.Length = 4, NumberValue(values(3)), Nothing)
                        Return SubstituteText(contents, TextValue(values(1)), TextValue(values(2)), occurrence)
                End Select
            End Function

            Private Shared Function RepeatText(contents As String, repetitions As Object) As Object
                If TypeOf repetitions Is FormulaError Then Return repetitions
                Dim rawCount = DirectCast(repetitions, Double)
                Dim count = Math.Truncate(rawCount)
                If rawCount < 0 OrElse count > Integer.MaxValue OrElse count * contents.Length > MaximumTextLength Then
                    Return New FormulaError("#VALUE!")
                End If
                If contents.Length = 0 OrElse count = 0 Then Return String.Empty
                Dim result As New StringBuilder(CInt(count) * contents.Length)
                For index = 1 To CInt(count)
                    result.Append(contents)
                Next
                Return result.ToString()
            End Function

            Private Shared Function ReplaceTextAt(contents As String, position As Object, length As Object,
                                                  replacement As String) As Object
                If TypeOf position Is FormulaError Then Return position
                If TypeOf length Is FormulaError Then Return length
                Dim start = Math.Truncate(DirectCast(position, Double))
                Dim rawLength = DirectCast(length, Double)
                Dim count = Math.Truncate(rawLength)
                If start < 1 OrElse start > MaximumTextLength OrElse rawLength < 0 OrElse count > MaximumTextLength Then
                    Return New FormulaError("#VALUE!")
                End If
                Dim offset = CInt(Math.Min(contents.Length, start - 1))
                Dim removed = Math.Min(contents.Length - offset, CInt(count))
                If contents.Length - removed + replacement.Length > MaximumTextLength Then Return New FormulaError("#VALUE!")
                Return contents.Substring(0, offset) & replacement & contents.Substring(offset + removed)
            End Function

            Private Shared Function SubstituteText(contents As String, oldText As String, newText As String,
                                                   occurrence As Object) As Object
                If TypeOf occurrence Is FormulaError Then Return occurrence
                Dim target As Integer
                If occurrence IsNot Nothing Then
                    Dim index = Math.Truncate(DirectCast(occurrence, Double))
                    If index < 1 OrElse index >= Integer.MaxValue Then Return New FormulaError("#VALUE!")
                    target = CInt(index)
                End If
                If oldText.Length = 0 Then Return contents

                Dim result As New StringBuilder()
                Dim start As Integer
                Dim match = contents.IndexOf(oldText, StringComparison.Ordinal)
                While match >= 0
                    If target = 0 Then
                        If result.Length + match - start + newText.Length > MaximumTextLength Then Return New FormulaError("#VALUE!")
                        result.Append(contents, start, match - start).Append(newText)
                        start = match + oldText.Length
                    Else
                        target -= 1
                        If target = 0 Then Return ReplaceTextAt(contents, CDbl(match) + 1, CDbl(oldText.Length), newText)
                    End If
                    ' A selected occurrence counts overlapping matches. Replacing every match consumes the matched text.
                    Dim advance = If(occurrence Is Nothing, oldText.Length, 1)
                    match = contents.IndexOf(oldText, match + advance, StringComparison.Ordinal)
                End While
                If result.Length + contents.Length - start > MaximumTextLength Then Return New FormulaError("#VALUE!")
                Return result.Append(contents, start, contents.Length - start).ToString()
            End Function

        End Class

    End Class

End Namespace
