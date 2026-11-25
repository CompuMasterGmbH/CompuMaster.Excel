Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Const RowCount As Integer = 1048576
        Private Const ColumnCount As Integer = 16384

        Private Shared ReadOnly FormulaString As New Regex("\G""(?:[^""]|"""")*""", RegexOptions.CultureInvariant)
        Private Shared ReadOnly FormulaSheet As New Regex("\G(?<name>'(?:[^']|'')+'|(?:\[[1-9][0-9]*\])?[\p{L}\p{M}\p{Nd}_.]+)!", RegexOptions.CultureInvariant)
        Private Shared ReadOnly FormulaSheetSpan As New Regex(
            "\G(?:'(?<first>(?:[^':\[\]]|'')+):(?<last>(?:[^':\[\]]|'')+)'|(?<first>[\p{L}\p{M}\p{Nd}_.]+):(?<last>[\p{L}\p{M}\p{Nd}_.]+))!",
            RegexOptions.CultureInvariant)
        Private Shared ReadOnly FormulaName As New Regex("\G[\p{L}_\\][\p{L}\p{M}\p{Nd}_.\\?]*", RegexOptions.CultureInvariant)
        Private Shared ReadOnly FormulaNumber As New Regex("\G(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?", RegexOptions.CultureInvariant)
        Private Shared ReadOnly FormulaErrorToken As New Regex("\G#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!|BUSY!|PYTHON!)", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
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

        Private Shared Function SheetNameFromToken(book As WorkbookPart, token As Match, ByRef externalBook As Boolean) As String
            Dim name = token.Groups("name").Value
            If name(0) = "'"c Then name = name.Substring(1, name.Length - 2).Replace("''", "'")
            externalBook = name.StartsWith("[", StringComparison.Ordinal)
            If externalBook Then
                Dim bracket = name.IndexOf("]"c)
                Dim index As Integer
                Dim references = book.Workbook.ExternalReferences
                If bracket < 2 OrElse bracket = name.Length - 1 OrElse
                   Not Integer.TryParse(name.Substring(1, bracket - 1), NumberStyles.None, CultureInfo.InvariantCulture, index) OrElse
                   index < 1 OrElse references Is Nothing OrElse index > references.ChildElements.Count Then
                    Throw UnsupportedReferenceSyntax()
                End If
                Dim reference = TryCast(references.ChildElements(index - 1), ExternalReference)
                Dim part As OpenXmlPart = Nothing
                If reference?.Id Is Nothing OrElse Not book.TryGetPartById(reference.Id.Value, part) OrElse
                   Not TypeOf part Is ExternalWorkbookPart OrElse name.IndexOfAny(UnsupportedSheetQualifierCharacters, bracket + 1) >= 0 Then
                    Throw UnsupportedReferenceSyntax()
                End If
            ElseIf name.IndexOfAny(UnsupportedSheetQualifierCharacters) >= 0 Then
                Throw UnsupportedReferenceSyntax()
            End If
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

        Private Shared Function LinkedStructuralSheet(qualifier As String, links As HashSet(Of Integer), ByRef prefix As String) As String
            Dim bracket = qualifier.IndexOf("]"c)
            Dim link = Integer.Parse(qualifier.Substring(1, bracket - 1), CultureInfo.InvariantCulture)
            If Not links.Contains(link) Then Return Nothing
            prefix = qualifier.Substring(0, bracket + 1)
            Return qualifier.Substring(bracket + 1)
        End Function

        Private Shared Function FormatLinkedSheetName(prefix As String, name As String) As String
            Return "'" & prefix & name.Replace("'", "''") & "'!"
        End Function

        Private Shared Function RewriteSheetSpan(span As Match, book As WorkbookPart, editedSheet As String, newSheetName As String) As String
            Dim first = span.Groups("first").Value.Replace("''", "'")
            Dim last = span.Groups("last").Value.Replace("''", "'")
            Dim names = book.Workbook.Sheets.Elements(Of Sheet)().Select(Function(sheet) sheet.Name.Value)
            If String.Equals(first, last, StringComparison.OrdinalIgnoreCase) OrElse
               Not names.Contains(first, StringComparer.OrdinalIgnoreCase) OrElse
               Not names.Contains(last, StringComparer.OrdinalIgnoreCase) Then Throw UnsupportedReferenceSyntax()
            If newSheetName Is Nothing Then Return span.Value

            If String.Equals(first, editedSheet, StringComparison.OrdinalIgnoreCase) Then
                first = newSheetName
            ElseIf String.Equals(last, editedSheet, StringComparison.OrdinalIgnoreCase) Then
                last = newSheetName
            Else
                Return span.Value
            End If
            If FormatSheetName(first).StartsWith("'", StringComparison.Ordinal) OrElse
               FormatSheetName(last).StartsWith("'", StringComparison.Ordinal) Then
                Return "'" & first.Replace("'", "''") & ":" & last.Replace("'", "''") & "'!"
            End If
            Return first & ":" & last & "!"
        End Function

        Private Shared Function UnsupportedReferenceSyntax() As NotSupportedException
            Return New NotSupportedException("This structural edit requires formula references supported by the Open XML SDK engine. Some external, 3-D, spill, structured and compound references cannot be rewritten.")
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

        Private Shared Function TableSelectorUsesColumn(value As String, start As Integer, finish As Integer,
                                                       columns As HashSet(Of String)) As Boolean
            Dim leafStart = -1
            Dim index = start
            While index < finish
                Select Case value(index)
                    Case "'"c
                        If index + 1 < finish AndAlso "[]#'@".IndexOf(value(index + 1)) >= 0 Then index += 1
                    Case "["c
                        leafStart = index + 1
                    Case "]"c
                        If leafStart >= 0 AndAlso leafStart < index Then
                            Dim name = value.Substring(leafStart, index - leafStart)
                            If name(0) = "@"c Then name = name.Substring(1)
                            If name.Length > 0 AndAlso name(0) <> "#"c AndAlso
                               columns.Contains(Regex.Replace(name, "'([\[\]#'@])", "$1")) Then Return True
                        End If
                        leafStart = -1
                End Select
                index += 1
            End While
            Return False
        End Function

        Private Shared Function FormatCopiedEndpoint(endpoint As ReferenceEndpoint, rowOffset As Integer, columnOffset As Integer) As String
            Dim row = TranslateCopiedIndex(endpoint.Row, endpoint.RowText, rowOffset)
            Dim column = TranslateCopiedIndex(endpoint.Column, endpoint.ColumnText, columnOffset)
            If (endpoint.Row >= 0 AndAlso row < 0) OrElse row >= RowCount OrElse
               (endpoint.Column >= 0 AndAlso column < 0) OrElse column >= ColumnCount Then Return Nothing
            Return endpoint.Format(row, column)
        End Function

        ' Identify reference operands without interpreting functions or touching string literals.
        Private Function RewriteFormula(value As String, scopeSheet As String, editedSheet As String,
                                               shift As StructuralShift, renamedSheet As String, isDefinedName As Boolean,
                                               Optional tableNames As HashSet(Of String) = Nothing, Optional copyRowOffset As Integer = 0,
                                               Optional copyColumnOffset As Integer = 0,
                                        Optional rewriteSheetQualifier As Func(Of String, String, String) = Nothing,
                                        Optional formulaWorkbook As WorkbookPart = Nothing,
                                        Optional structuralExternalLinks As HashSet(Of Integer) = Nothing,
                                        Optional contextualTable As StructuralTable = Nothing,
                                        Optional observeName As Action(Of String, String) = Nothing,
                                        Optional allowSpillReferences As Boolean = False,
                                        Optional requireAbsoluteNameOperands As Boolean = False,
                                        Optional validateInvokedName As Action(Of String, String) = Nothing,
                                        Optional rewriteNameOperand As Func(Of String, String, String) = Nothing) As String
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

                If value(index) = "["c AndAlso Not FormulaSheet.Match(value, index).Success Then
                    If tableNames Is Nothing OrElse tableNames.Count = 0 Then Throw UnsupportedReferenceSyntax()
                    Dim selectorEnd = ReadTableSelector(value, index)
                    If selectorEnd < value.Length AndAlso (value(selectorEnd) = "!"c OrElse FormulaSheet.Match(value, selectorEnd).Success) Then
                        Throw UnsupportedReferenceSyntax()
                    End If
                    If structuralExternalLinks Is Nothing AndAlso shift?.ReferencesDeletedTableColumn(value, index, selectorEnd, Nothing, contextualTable, isDefinedName) = True Then
                        result.Append("#REF!")
                    Else
                        result.Append(value, index, selectorEnd - index)
                    End If
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
                Dim sheetSpan = FormulaSheetSpan.Match(value, index)
                Dim sheetToken = FormulaSheet.Match(value, index)
                Dim explicitSheet As String = Nothing
                Dim externalBook As Boolean
                Dim linkedSheet As String = Nothing
                Dim linkedPrefix As String = Nothing
                Dim qualifier As String = Nothing
                If sheetSpan.Success Then
                    If rewriteSheetQualifier IsNot Nothing Then Throw UnsupportedReferenceSyntax()
                    qualifier = RewriteSheetSpan(sheetSpan, If(formulaWorkbook, WorkbookPart), editedSheet,
                                                 If(structuralExternalLinks Is Nothing, renamedSheet, Nothing))
                    index += sheetSpan.Length
                ElseIf sheetToken.Success Then
                    explicitSheet = SheetNameFromToken(If(formulaWorkbook, WorkbookPart), sheetToken, externalBook)
                    If externalBook AndAlso structuralExternalLinks IsNot Nothing Then
                        linkedSheet = LinkedStructuralSheet(explicitSheet, structuralExternalLinks, linkedPrefix)
                    End If
                    If rewriteSheetQualifier IsNot Nothing Then
                        qualifier = rewriteSheetQualifier(explicitSheet, sheetToken.Value)
                    ElseIf linkedSheet IsNot Nothing AndAlso renamedSheet IsNot Nothing AndAlso
                           String.Equals(linkedSheet, editedSheet, StringComparison.OrdinalIgnoreCase) Then
                        qualifier = FormatLinkedSheetName(linkedPrefix, renamedSheet)
                    Else
                        qualifier = If(Not externalBook AndAlso structuralExternalLinks Is Nothing AndAlso
                                       renamedSheet IsNot Nothing AndAlso String.Equals(explicitSheet, editedSheet, StringComparison.OrdinalIgnoreCase),
                                       FormatSheetName(renamedSheet), sheetToken.Value)
                    End If
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
                        If sheetSpan.Success AndAlso StartsRangeExpression(value, secondStart) Then Throw UnsupportedReferenceSyntax()
                        If Not first.IsCell OrElse Not StartsRangeExpression(value, secondStart) Then
                            Dim secondQualifier = FormulaSheet.Match(value, secondStart)
                            If secondQualifier.Success Then
                                If sheetSpan.Success Then Throw UnsupportedReferenceSyntax()
                                Dim secondExternalBook As Boolean
                                Dim secondSheet = SheetNameFromToken(If(formulaWorkbook, WorkbookPart), secondQualifier, secondExternalBook)
                                If externalBook <> secondExternalBook OrElse Not String.Equals(explicitSheet, secondSheet, StringComparison.OrdinalIgnoreCase) Then
                                    Throw UnsupportedReferenceSyntax()
                                End If
                                secondStart += secondQualifier.Length
                            End If
                            Dim secondEnd As Integer
                            second = ReadEndpoint(value, secondStart, secondEnd)
                            If second Is Nothing OrElse (first.Row < 0) <> (second.Row < 0) OrElse
                               (first.Column < 0) <> (second.Column < 0) Then Throw UnsupportedReferenceSyntax()
                            separator = value.Substring(referenceEnd, secondStart - referenceEnd)
                            If secondQualifier.Success AndAlso rewriteSheetQualifier IsNot Nothing Then
                                separator = value.Substring(referenceEnd, secondQualifier.Index - referenceEnd) &
                                    rewriteSheetQualifier(explicitSheet, secondQualifier.Value)
                            ElseIf secondQualifier.Success AndAlso qualifier <> sheetToken.Value Then
                                separator = value.Substring(referenceEnd, secondQualifier.Index - referenceEnd) & qualifier
                            End If
                            referenceEnd = secondEnd
                        End If
                    ElseIf Not sheetSpan.Success AndAlso explicitSheet Is Nothing AndAlso colon < value.Length AndAlso value(colon) = "("c AndAlso
                           (colon = referenceEnd OrElse String.Equals(value.Substring(index, referenceEnd - index), "LOG10", StringComparison.OrdinalIgnoreCase)) Then
                        first = Nothing
                    End If
                End If

                If first IsNot Nothing AndAlso (first.IsCell OrElse second IsNot Nothing) Then
                    result.Append(qualifier)
                    Dim unchanged = value.Substring(index, referenceEnd - index)
                    If (requireAbsoluteNameOperands OrElse (shift IsNot Nothing AndAlso isDefinedName AndAlso Not externalBook)) AndAlso
                       (Not sheetSpan.Success AndAlso explicitSheet Is Nothing AndAlso scopeSheet Is Nothing OrElse Not first.IsAbsolute OrElse
                        (second IsNot Nothing AndAlso Not second.IsAbsolute)) Then
                        Throw New NotSupportedException("This edit requires every A1 reference in a defined name to be absolute and to have an explicit or worksheet-local scope.")
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
                    ElseIf shift IsNot Nothing AndAlso Not sheetSpan.Success AndAlso
                           String.Equals(If(externalBook OrElse structuralExternalLinks IsNot Nothing, linkedSheet, If(explicitSheet, scopeSheet)),
                                         editedSheet, StringComparison.OrdinalIgnoreCase) Then
                        result.Append(shift.RewriteReference(first, second, separator))
                    ElseIf second IsNot Nothing AndAlso (renamedSheet IsNot Nothing OrElse rewriteSheetQualifier IsNot Nothing) Then
                        result.Append(first.ColumnText).Append(first.RowText).Append(separator).Append(second.ColumnText).Append(second.RowText)
                    Else
                        result.Append(unchanged)
                    End If
                    index = referenceEnd
                    lastReferenceEnd = index
                    Continue While
                End If

                If sheetSpan.Success Then Throw UnsupportedReferenceSyntax()
                Dim token = FormulaErrorToken.Match(value, index)
                Dim isName = False
                If Not token.Success Then
                    token = FormulaName.Match(value, index)
                    isName = token.Success
                End If
                If Not token.Success AndAlso explicitSheet Is Nothing Then token = FormulaNumber.Match(value, index)
                If token.Success Then
                    If Not allowSpillReferences AndAlso String.Equals(token.Value, "_xlfn.ANCHORARRAY", StringComparison.OrdinalIgnoreCase) Then Throw UnsupportedReferenceSyntax()
                    index += token.Length
                    If String.Equals(token.Value, "#REF!", StringComparison.OrdinalIgnoreCase) Then lastReferenceEnd = index
                    If index < value.Length AndAlso value(index) = "["c Then
                        If explicitSheet IsNot Nothing OrElse tableNames Is Nothing OrElse Not tableNames.Contains(token.Value) Then
                            Throw UnsupportedReferenceSyntax()
                        End If
                        Dim selectorEnd = ReadTableSelector(value, index)
                        If structuralExternalLinks Is Nothing AndAlso shift?.ReferencesDeletedTableColumn(value, index, selectorEnd, token.Value, Nothing, isDefinedName) = True Then
                            result.Append("#REF!")
                        Else
                            result.Append(token.Value).Append(value, index, selectorEnd - index)
                        End If
                        index = selectorEnd
                    Else
                        Dim nameOperand = token.Value
                        If isName AndAlso (observeName IsNot Nothing OrElse validateInvokedName IsNot Nothing OrElse rewriteNameOperand IsNot Nothing) Then
                            Dim afterName = SkipWhitespace(value, index)
                            If afterName < value.Length AndAlso value(afterName) = "("c Then
                                validateInvokedName?.Invoke(token.Value, explicitSheet)
                            Else
                                observeName?.Invoke(token.Value, explicitSheet)
                                If rewriteNameOperand IsNot Nothing Then nameOperand = rewriteNameOperand(token.Value, explicitSheet)
                            End If
                        End If
                        result.Append(qualifier).Append(nameOperand)
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
