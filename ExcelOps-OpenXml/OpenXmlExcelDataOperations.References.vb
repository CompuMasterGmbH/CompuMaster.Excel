Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Const RowCount As Integer = 1048576
        Private Const ColumnCount As Integer = 16384

        Private Shared ReadOnly FormulaString As New Regex("\G""(?:[^""]|"""")*""", RegexOptions.CultureInvariant)
        Private Shared ReadOnly FormulaSheet As New Regex("\G(?<name>'(?:[^']|'')+'|[\p{L}\p{M}\p{Nd}_.]+)!", RegexOptions.CultureInvariant)
        Private Shared ReadOnly FormulaName As New Regex("\G[\p{L}_\\][\p{L}\p{M}\p{Nd}_.\\?]*", RegexOptions.CultureInvariant)
        Private Shared ReadOnly FormulaNumber As New Regex("\G(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?", RegexOptions.CultureInvariant)
        Private Shared ReadOnly FormulaError As New Regex("\G#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!|BUSY!|PYTHON!)", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
        Private Shared ReadOnly A1Endpoint As New Regex(
            "\G(?:(?<column>\$?[A-Za-z]{1,3})(?<row>\$?0*[1-9][0-9]{0,6})?|(?<row>\$?0*[1-9][0-9]{0,6}))(?![\p{L}\p{M}\p{Nd}_.\\?$])",
            RegexOptions.CultureInvariant)
        Private Shared ReadOnly UnsupportedSheetQualifierCharacters As Char() = {":"c, "["c, "]"c}
        Private Shared ReadOnly BareSheetName As New Regex("\A[\p{L}_][\p{L}\p{M}\p{Nd}_.]*\z", RegexOptions.CultureInvariant)
        Private Shared ReadOnly R1C1Name As New Regex("\A(?:R|C|R[0-9]+C[0-9]+)\z", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

        Private NotInheritable Class ReferenceEndpoint
            Friend ReadOnly ColumnText As String
            Friend ReadOnly RowText As String
            Friend ReadOnly Column As Integer
            Friend ReadOnly Row As Integer

            Friend Sub New(columnText As String, rowText As String)
                Me.ColumnText = columnText
                Me.RowText = rowText
                Column = If(columnText.Length = 0, -1, ExcelCell.CalculateColumnIndex(columnText.TrimStart("$"c)))
                Row = If(rowText.Length = 0, -1, Integer.Parse(rowText.TrimStart("$"c), CultureInfo.InvariantCulture) - 1)
            End Sub

            Friend ReadOnly Property IsCell As Boolean
                Get
                    Return Column >= 0 AndAlso Row >= 0
                End Get
            End Property

            Friend ReadOnly Property IsAbsolute As Boolean
                Get
                    Return (Column < 0 OrElse ColumnText(0) = "$"c) AndAlso (Row < 0 OrElse RowText(0) = "$"c)
                End Get
            End Property

            Friend Function Format(vertical As Boolean, index As Integer) As String
                Return Format(If(vertical, index, Row), If(vertical, Column, index))
            End Function

            Friend Function Format(rowIndex As Integer, columnIndex As Integer) As String
                Dim rowPart = RowText
                Dim columnPart = ColumnText
                If rowIndex <> Row Then
                    rowPart = If(RowText(0) = "$"c, "$", "") & (rowIndex + 1).ToString(CultureInfo.InvariantCulture)
                End If
                If columnIndex <> Column Then
                    Dim letters = ExcelCell.LocalCellAddress(0, columnIndex).TrimEnd("1"c)
                    columnPart = If(ColumnText(0) = "$"c, "$", "") & letters
                End If
                Return columnPart & rowPart
            End Function
        End Class

        Private Shared Function ReadEndpoint(value As String, start As Integer, ByRef nextIndex As Integer) As ReferenceEndpoint
            Dim token = A1Endpoint.Match(value, start)
            If Not token.Success Then Return Nothing
            Dim endpoint As New ReferenceEndpoint(token.Groups("column").Value, token.Groups("row").Value)
            If endpoint.Row >= RowCount OrElse endpoint.Column >= ColumnCount Then Return Nothing
            nextIndex = token.Index + token.Length
            Return endpoint
        End Function

        Private Shared Function SkipWhitespace(value As String, start As Integer) As Integer
            While start < value.Length AndAlso Char.IsWhiteSpace(value(start))
                start += 1
            End While
            Return start
        End Function

        Private Shared Function StartsRangeExpression(value As String, start As Integer) As Boolean
            If start >= value.Length Then Return False
            If value(start) = "("c Then Return True
            Dim name = FormulaName.Match(value, start)
            If Not name.Success Then Return False
            Dim nameEnd = start + name.Length
            Dim afterName = SkipWhitespace(value, nameEnd)
            If afterName >= value.Length OrElse value(afterName) <> "("c Then Return False
            ' A cell reference followed by a space and a group uses the intersection operator.
            If afterName > nameEnd AndAlso Not String.Equals(name.Value, "LOG10", StringComparison.OrdinalIgnoreCase) Then
                Dim endpointEnd As Integer
                Dim endpoint = ReadEndpoint(value, start, endpointEnd)
                If endpoint IsNot Nothing AndAlso endpoint.IsCell Then Return False
            End If
            Return True
        End Function

        Private Shared Function IsExpressionRangeOperator(value As String, index As Integer, lastReferenceEnd As Integer) As Boolean
            Dim operandStart = SkipWhitespace(value, index + 1)
            If operandStart >= value.Length Then Return False
            Dim previous = index - 1
            While previous >= 0 AndAlso Char.IsWhiteSpace(value(previous))
                previous -= 1
            End While
            ' A closing parenthesis ends a grouped reference or a reference-returning call.
            If previous >= 0 AndAlso value(previous) = ")"c Then Return True
            Return lastReferenceEnd >= 0 AndAlso SkipWhitespace(value, lastReferenceEnd) = index AndAlso
                StartsRangeExpression(value, operandStart)
        End Function

        Private Shared Function SheetNameFromToken(token As Match) As String
            Dim name = token.Groups("name").Value
            If name(0) = "'"c Then name = name.Substring(1, name.Length - 2).Replace("''", "'")
            If name.IndexOfAny(UnsupportedSheetQualifierCharacters) >= 0 Then Throw UnsupportedReferenceSyntax()
            Return name
        End Function

        Private Shared Function FormatSheetName(name As String) As String
            Dim endpointEnd As Integer
            Dim endpoint = ReadEndpoint(name, 0, endpointEnd)
            If BareSheetName.IsMatch(name) AndAlso Not R1C1Name.IsMatch(name) AndAlso
               Not (endpoint IsNot Nothing AndAlso endpoint.IsCell AndAlso endpointEnd = name.Length) Then
                Return name & "!"
            End If
            Return "'" & name.Replace("'", "''") & "'!"
        End Function

        Private Shared Function UnsupportedReferenceSyntax() As NotSupportedException
            Return New NotSupportedException("This structural edit requires formula references supported by the Open XML SDK engine. External references, 3-D references, spill references, references to unknown tables, and some compound ranges cannot be rewritten.")
        End Function

        Private Shared Function ReadTableSelector(value As String, start As Integer) As Integer
            Dim depth = 0
            Dim index = start
            While index < value.Length
                Select Case value(index)
                    Case "'"c
                        If index + 1 < value.Length AndAlso "[]#'@".IndexOf(value(index + 1)) >= 0 Then index += 1
                    Case "["c
                        depth += 1
                    Case "]"c
                        depth -= 1
                        If depth = 0 Then Return index + 1
                End Select
                index += 1
            End While
            Throw UnsupportedReferenceSyntax()
        End Function

        Private Shared Function TranslateCopiedIndex(index As Integer, text As String, offset As Integer) As Integer
            If index < 0 OrElse text(0) = "$"c Then Return index
            Return index + offset
        End Function

        Private Shared Function FormatCopiedEndpoint(endpoint As ReferenceEndpoint, rowOffset As Integer, columnOffset As Integer) As String
            Dim row = TranslateCopiedIndex(endpoint.Row, endpoint.RowText, rowOffset)
            Dim column = TranslateCopiedIndex(endpoint.Column, endpoint.ColumnText, columnOffset)
            If (endpoint.Row >= 0 AndAlso row < 0) OrElse row >= RowCount OrElse
               (endpoint.Column >= 0 AndAlso column < 0) OrElse column >= ColumnCount Then Return Nothing
            Return endpoint.Format(row, column)
        End Function

        ' Identify reference operands without interpreting functions or touching string literals.
        Private Shared Function RewriteFormula(value As String, scopeSheet As String, editedSheet As String,
                                               shift As StructuralShift, renamedSheet As String, isDefinedName As Boolean,
                                               Optional tableNames As HashSet(Of String) = Nothing, Optional copyRowOffset As Integer = 0,
                                               Optional copyColumnOffset As Integer = 0) As String
            Dim result As New StringBuilder(value.Length)
            Dim index = 0
            Dim lastReferenceEnd = -1
            While index < value.Length
                Dim literal = FormulaString.Match(value, index)
                If literal.Success Then
                    result.Append(literal.Value)
                    index += literal.Length
                    Continue While
                End If

                If value(index) = "["c Then
                    If tableNames Is Nothing OrElse tableNames.Count = 0 Then Throw UnsupportedReferenceSyntax()
                    Dim selectorEnd = ReadTableSelector(value, index)
                    If selectorEnd < value.Length AndAlso (value(selectorEnd) = "!"c OrElse FormulaSheet.Match(value, selectorEnd).Success) Then
                        Throw UnsupportedReferenceSyntax()
                    End If
                    result.Append(value, index, selectorEnd - index)
                    index = selectorEnd
                    Continue While
                End If

                If value(index) = ":"c Then
                    If Not IsExpressionRangeOperator(value, index, lastReferenceEnd) Then Throw UnsupportedReferenceSyntax()
                    result.Append(":"c)
                    index += 1
                    Continue While
                End If

                Dim tokenStart = index
                Dim sheetToken = FormulaSheet.Match(value, index)
                Dim explicitSheet As String = Nothing
                Dim qualifier As String = Nothing
                If sheetToken.Success Then
                    explicitSheet = SheetNameFromToken(sheetToken)
                    qualifier = If(renamedSheet IsNot Nothing AndAlso String.Equals(explicitSheet, editedSheet, StringComparison.OrdinalIgnoreCase),
                                   FormatSheetName(renamedSheet), sheetToken.Value)
                    index += sheetToken.Length
                End If

                Dim referenceEnd As Integer
                Dim first = ReadEndpoint(value, index, referenceEnd)
                Dim second As ReferenceEndpoint = Nothing
                Dim separator As String = Nothing
                If first IsNot Nothing Then
                    Dim colon = SkipWhitespace(value, referenceEnd)
                    If colon < value.Length AndAlso value(colon) = ":"c Then
                        Dim secondStart = SkipWhitespace(value, colon + 1)
                        If Not first.IsCell OrElse Not StartsRangeExpression(value, secondStart) Then
                            Dim secondQualifier = FormulaSheet.Match(value, secondStart)
                            If secondQualifier.Success Then
                                If Not String.Equals(explicitSheet, SheetNameFromToken(secondQualifier), StringComparison.OrdinalIgnoreCase) Then
                                    Throw UnsupportedReferenceSyntax()
                                End If
                                secondStart += secondQualifier.Length
                            End If
                            Dim secondEnd As Integer
                            second = ReadEndpoint(value, secondStart, secondEnd)
                            If second Is Nothing OrElse (first.Row < 0) <> (second.Row < 0) OrElse
                               (first.Column < 0) <> (second.Column < 0) Then Throw UnsupportedReferenceSyntax()
                            separator = value.Substring(referenceEnd, secondStart - referenceEnd)
                            If secondQualifier.Success AndAlso renamedSheet IsNot Nothing AndAlso
                               String.Equals(explicitSheet, editedSheet, StringComparison.OrdinalIgnoreCase) Then
                                separator = value.Substring(referenceEnd, secondQualifier.Index - referenceEnd) & qualifier
                            End If
                            referenceEnd = secondEnd
                        End If
                    ElseIf explicitSheet Is Nothing AndAlso colon < value.Length AndAlso value(colon) = "("c AndAlso
                           (colon = referenceEnd OrElse String.Equals(value.Substring(index, referenceEnd - index), "LOG10", StringComparison.OrdinalIgnoreCase)) Then
                        first = Nothing
                    End If
                End If

                If first IsNot Nothing AndAlso (first.IsCell OrElse second IsNot Nothing) Then
                    result.Append(qualifier)
                    Dim unchanged = value.Substring(index, referenceEnd - index)
                    If shift IsNot Nothing AndAlso isDefinedName AndAlso
                       (explicitSheet Is Nothing AndAlso scopeSheet Is Nothing OrElse Not first.IsAbsolute OrElse
                        (second IsNot Nothing AndAlso Not second.IsAbsolute)) Then
                        Throw New NotSupportedException("Structural shifts require every A1 reference in a defined name to be absolute and to have an explicit or worksheet-local scope.")
                    End If
                    If copyRowOffset <> 0 OrElse copyColumnOffset <> 0 Then
                        Dim firstAddress = FormatCopiedEndpoint(first, copyRowOffset, copyColumnOffset)
                        Dim secondAddress = If(second Is Nothing, firstAddress, FormatCopiedEndpoint(second, copyRowOffset, copyColumnOffset))
                        If firstAddress Is Nothing OrElse secondAddress Is Nothing Then
                            result.Append("#REF!")
                        Else
                            result.Append(firstAddress)
                            If second IsNot Nothing Then result.Append(separator).Append(secondAddress)
                        End If
                    ElseIf shift IsNot Nothing AndAlso String.Equals(If(explicitSheet, scopeSheet), editedSheet, StringComparison.OrdinalIgnoreCase) Then
                        result.Append(shift.RewriteReference(first, second, separator))
                    ElseIf second IsNot Nothing AndAlso renamedSheet IsNot Nothing Then
                        result.Append(first.ColumnText).Append(first.RowText).Append(separator).Append(second.ColumnText).Append(second.RowText)
                    Else
                        result.Append(unchanged)
                    End If
                    index = referenceEnd
                    lastReferenceEnd = index
                    Continue While
                End If

                Dim token = FormulaError.Match(value, index)
                If Not token.Success Then token = FormulaName.Match(value, index)
                If Not token.Success AndAlso explicitSheet Is Nothing Then token = FormulaNumber.Match(value, index)
                If token.Success Then
                    If String.Equals(token.Value, "_xlfn.ANCHORARRAY", StringComparison.OrdinalIgnoreCase) Then Throw UnsupportedReferenceSyntax()
                    result.Append(qualifier).Append(token.Value)
                    index += token.Length
                    If String.Equals(token.Value, "#REF!", StringComparison.OrdinalIgnoreCase) Then lastReferenceEnd = index
                    If index < value.Length AndAlso value(index) = "["c Then
                        If explicitSheet IsNot Nothing OrElse tableNames Is Nothing OrElse Not tableNames.Contains(token.Value) Then
                            Throw UnsupportedReferenceSyntax()
                        End If
                        Dim selectorEnd = ReadTableSelector(value, index)
                        result.Append(value, index, selectorEnd - index)
                        index = selectorEnd
                    End If
                ElseIf explicitSheet Is Nothing AndAlso
                       ("+-*/^&=<>%(),;{}@".IndexOf(value(tokenStart)) >= 0 OrElse Char.IsWhiteSpace(value(tokenStart))) Then
                    result.Append(value(tokenStart))
                    index += 1
                Else
                    Throw UnsupportedReferenceSyntax()
                End If
            End While
            Return result.ToString()
        End Function

    End Class

End Namespace
