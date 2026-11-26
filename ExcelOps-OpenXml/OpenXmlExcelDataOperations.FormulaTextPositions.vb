Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared Function TextCharacterCount(contents As String, unicodePositions As Boolean,
                                                       Optional untilOffset As Integer = -1) As Integer
                Dim finish = If(untilOffset < 0, contents.Length, untilOffset)
                Dim count = finish
                If unicodePositions Then
                    For index = 1 To finish - 1
                        If Char.IsSurrogatePair(contents, index - 1) Then count -= 1
                    Next
                End If
                Return count
            End Function

            Private Shared Function Utf16CharacterOffset(contents As String, position As Double, unicodePositions As Boolean) As Integer
                If position >= contents.Length Then Return contents.Length
                If Not unicodePositions Then Return CInt(position)

                Dim index = 0
                Dim remaining = CInt(position)
                While remaining > 0 AndAlso index < contents.Length
                    index += If(Char.IsSurrogatePair(contents, index), 2, 1)
                    remaining -= 1
                End While
                Return index
            End Function

            Private Function EvaluateTextSearch(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                               formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim needle = AsMatrix(EvaluateExpression(arguments(0), formulaCell))
                Dim subjects = AsMatrix(EvaluateExpression(arguments(1), formulaCell))
                Dim starts = AsMatrix(If(arguments.Count = 3, EvaluateExpression(arguments(2), formulaCell), 1.0))
                Dim rows = Math.Max(needle.Rows, Math.Max(subjects.Rows, starts.Rows))
                Dim columns = Math.Max(needle.Columns, Math.Max(subjects.Columns, starts.Columns))
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim pattern = needle.BroadcastAt(row, column)
                        Dim subject = subjects.BroadcastAt(row, column)
                        Dim position = NumberValue(starts.BroadcastAt(row, column))
                        Dim result As Object
                        If TypeOf pattern Is FormulaError Then
                            result = pattern
                        ElseIf TypeOf subject Is FormulaError Then
                            result = subject
                        ElseIf TypeOf position Is FormulaError Then
                            result = position
                        Else
                            result = FindTextPosition(name, TextValue(pattern), TextValue(subject), DirectCast(position, Double))
                        End If
                        values(row * columns + column) = result
                    Next
                Next
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Function FindTextPosition(name As String, pattern As String, subject As String, start As Double) As Object
                Dim position = Math.Truncate(start)
                Dim last = TextCharacterCount(subject, _unicodeTextPositions)
                Dim literal = name = "FIND"
                If literal AndAlso pattern.Length = 0 Then last += 1
                If position < 1.0 OrElse position > last Then Return New FormulaError("#VALUE!")
                Dim offset = Utf16CharacterOffset(subject, position - 1.0, _unicodeTextPositions)
                Dim index As Integer
                If literal Then
                    index = subject.IndexOf(pattern, offset, StringComparison.Ordinal)
                Else
                    Dim wildcard = BuildWildcard(FoldSearchCase(pattern), searchWithinText:=True)
                    Dim searchable = FoldSearchCase(subject)
                    Try
                        Dim found = wildcard.Match(searchable, offset)
                        ' Excel matches wildcard units as UTF-16 even in version 2. Only the reported positions must start at a full character.
                        While found.Success AndAlso _unicodeTextPositions AndAlso found.Index > 0 AndAlso
                            Char.IsSurrogatePair(subject, found.Index - 1)
                            found = wildcard.Match(searchable, found.Index + 1)
                        End While
                        index = If(found.Success, found.Index, -1)
                    Catch ex As RegexMatchTimeoutException
                        Throw New NotSupportedException("The wildcard search exceeds the supported complexity.", ex)
                    End Try
                End If
                If index < 0 Then Return New FormulaError("#VALUE!")
                Return CDbl(TextCharacterCount(subject, _unicodeTextPositions, index) + 1)
            End Function

            Private Shared Function FoldSearchCase(text As String) As String
                ' Normalize final sigma explicitly so SEARCH agrees across both Windows and ICU casing implementations.
                Return text.Replace(ChrW(&H3C2), ChrW(&H3C3)).ToUpperInvariant()
            End Function

        End Class

    End Class

End Namespace
