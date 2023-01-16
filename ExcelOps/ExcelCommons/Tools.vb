Option Explicit On
Option Strict On

Namespace ExcelOps
    Public NotInheritable Class Tools

        Public Shared Function ReplaceWholeValue(value As String, searchValue As String, replacementValue As String) As String
            If value = searchValue Then
                Return replacementValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        ''' Formula without cell references contain just simple mathematic operations such as *, /, +, -, ^, (, )
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaWithoutCellReferences(formula As String) As Boolean
            Dim Parts As String() = formula.Split(New Char() {":"c, "+"c, "-"c, "*"c, "/"c, "^"c, "("c, ")"c, ","c})
            For Each Part As String In Parts
                If ExcelCell.IsValidAddress(Part, True) = True Then
                    Return False
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' Formula without cell references contain just simple mathematic operations such as *, /, +, -, ^, (, )
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaWithoutCellReferencesOrCellReferenceInSameRow(formula As String, rowIndex As Integer) As Boolean
            Dim Parts As String() = formula.Split(New Char() {":"c, "+"c, "-"c, "*"c, "/"c, "^"c, "("c, ")"c, ","c})
            For Each Part As String In Parts
                If ExcelCell.IsValidAddress(Part, True) = True Then
                    Dim Cell As New ExcelCell(Part, ExcelCell.ValueTypes.All)
                    If Cell.SheetName <> Nothing Then
                        Return False
                    ElseIf Cell.RowIndex <> rowIndex Then
                        Return False
                    End If
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' Formula is just a simple cell reference like "Grunddaten!B4"
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaSimpleCellReference(formula As String) As Boolean
            Return ExcelCell.IsValidAddress(formula, True)
        End Function

        ''' <summary>
        ''' Formula is just a simple SUM function with a cell range
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaSimpleSumFunction(formula As String) As Boolean
            If formula.StartsWith("SUM(") = False Then
                Return False
            End If
            Dim Parts As String() = formula.Substring(3).Split(New Char() {":"c, "("c, ")"c}, StringSplitOptions.RemoveEmptyEntries)
            If Parts.Length = 0 OrElse Parts.Length > 2 Then
                Return False
            End If
            For Each Part As String In Parts
                If ExcelCell.IsValidAddress(Part, True) = False Then
                    Return False
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaWithSheetReference(formula As String, checkForSheetName As String) As Boolean
            If formula = Nothing Then
                Return False
            Else
                Return formula.Contains(checkForSheetName & "!") OrElse formula.Contains("'" & checkForSheetName & "'!") OrElse formula.Contains("]" & checkForSheetName & "'!")
            End If
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formulas"></param>
        ''' <returns></returns>
        Public Shared Function FormulasWithSheetReferencesToSheet(formulas As List(Of TextTableCell), checkForSheetNameInFormula As String, dontSearchInThisSheetName As String) As List(Of TextTableCell)
            Dim Result As New List(Of TextTableCell)
            For MyFormulaCounter As Integer = 0 To formulas.Count - 1
                If dontSearchInThisSheetName = Nothing OrElse dontSearchInThisSheetName <> formulas(MyFormulaCounter).SheetName Then
                    If ExcelOps.Tools.IsFormulaWithSheetReference(formulas(MyFormulaCounter).CellContent, checkForSheetNameInFormula) = True Then
                        Result.Add(formulas(MyFormulaCounter))
                    End If
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formulas"></param>
        ''' <returns></returns>
        Public Shared Function ContainsFormulasWithSheetReferencesToSheet(formulas As List(Of TextTableCell), checkForSheetNameInFormula As String, dontSearchInThisSheetName As String) As Boolean
            For MyFormulaCounter As Integer = 0 To formulas.Count - 1
                If dontSearchInThisSheetName = Nothing OrElse dontSearchInThisSheetName <> formulas(MyFormulaCounter).SheetName Then
                    If ExcelOps.Tools.IsFormulaWithSheetReference(formulas(MyFormulaCounter).CellContent, checkForSheetNameInFormula) = True Then
                        Return True
                    End If
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formulas"></param>
        ''' <returns></returns>
        Public Shared Function FormulasWithSheetReferencesFromSheet(searchInThisSheetName As String, formulas As List(Of TextTableCell), checkForSheetNamesInFormula As String()) As List(Of TextTableCell)
            Dim Result As New List(Of TextTableCell)
            If searchInThisSheetName = Nothing Then Throw New ArgumentNullException(NameOf(searchInThisSheetName))
            For MyFormulaCounter As Integer = 0 To formulas.Count - 1
                If searchInThisSheetName = formulas(MyFormulaCounter).SheetName Then
                    For SearchForSheetNamesCounter As Integer = 0 To checkForSheetNamesInFormula.Length - 1
                        If ExcelOps.Tools.IsFormulaWithSheetReference(formulas(MyFormulaCounter).CellContent, checkForSheetNamesInFormula(SearchForSheetNamesCounter)) = True Then
                            Result.Add(formulas(MyFormulaCounter))
                        End If
                    Next
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formulas"></param>
        ''' <returns></returns>
        Public Shared Function ContainsFormulasWithSheetReferencesFromSheet(searchInThisSheetName As String, formulas As List(Of TextTableCell), checkForSheetNamesInFormula As String()) As Boolean
            If searchInThisSheetName = Nothing Then Throw New ArgumentNullException(NameOf(searchInThisSheetName))
            For MyFormulaCounter As Integer = 0 To formulas.Count - 1
                If searchInThisSheetName = formulas(MyFormulaCounter).SheetName Then
                    For SearchForSheetNamesCounter As Integer = 0 To checkForSheetNamesInFormula.Length - 1
                        If ExcelOps.Tools.IsFormulaWithSheetReference(formulas(MyFormulaCounter).CellContent, checkForSheetNamesInFormula(SearchForSheetNamesCounter)) = True Then
                            Return True
                        End If
                    Next
                End If
            Next
            Return False
        End Function

        Public Shared Function ParseToDoubleCultureSafe(s As String) As Double
            Dim Result As Double
            If s = Nothing Then
                Throw New ArgumentNullException(NameOf(s))
            ElseIf TryParseToDoubleCultureSafe(s, Result) = True Then
                Return Result
            Else
                Throw New FormatException("Not a parsable value: " & s)
            End If
        End Function

        Public Shared Function TryParseToDoubleCultureSafe(s As String, ByRef result As Double) As Boolean
            Dim sb As New System.Text.StringBuilder(s.Trim)

            'Find separators and shortcut exit with failure if invalid chars are present
            Dim SeparatorsFromRightDirection As List(Of KeyValuePair(Of Char, Integer))
            SeparatorsFromRightDirection = FindSeparatorsFromRightDirection(sb)
            If SeparatorsFromRightDirection Is Nothing Then
                'FindSeparatorsFromRightDirection failed
                Return False
            End If

            'Analyse findings
            If SeparatorsFromRightDirection.Count = 0 Then
                'Integer value -> just continue
            ElseIf SeparatorsFromRightDirection.Count = 1 Then
                'Either a decimal or a bigger integer between 1000 and 999999
                Select Case SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehind(sb, SeparatorsFromRightDirection(0))
                    Case SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator
                    'continue with next steps
                    Case SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.UnsafeDecision
                        Return False
                    Case Else
                        Throw New NotImplementedException
                End Select
            ElseIf SeparatorsFromRightDirection.Count = 2 Then
                If SeparatorsFromRightDirection(0).Key = SeparatorsFromRightDirection(1).Key Then
                    If SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehind(sb, SeparatorsFromRightDirection(0)) = SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator Then
                        'there can't be multiple decimal separators
                        Return False
                    ElseIf HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(sb, SeparatorsFromRightDirection(1)) = False Then
                        'Throw New FormatException("Rule mismatch: HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator: " & s)
                        Return False
                    Else
                        'same char twice -> both are thousands separator -> just remove all of them
                        For Each Finding In SeparatorsFromRightDirection
                            sb.Remove(Finding.Value, 1)
                        Next
                    End If
                Else
                    'different chars -> first is decimal separator, 2nd is thousands separator
                    sb.Remove(SeparatorsFromRightDirection(1).Value, 1)
                End If
            Else 'If SeparatorsFromRightDirection.Count >= 2 Then
                Dim FindingsStartIndex As Integer
                If SeparatorsFromRightDirection(0).Key = SeparatorsFromRightDirection(1).Key Then
                    'same char twice -> both are thousands separator -> just remove all of them
                    If SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehind(sb, SeparatorsFromRightDirection(0)) = SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator Then
                        'there can't be multiple decimal separators
                        Return False
                    Else
                        FindingsStartIndex = 0
                    End If
                Else
                    'different chars -> first is decimal separator, 2nd is thousands separator
                    FindingsStartIndex = 1
                End If
                For FindingsCounter As Integer = FindingsStartIndex To SeparatorsFromRightDirection.Count - 1
                    If SeparatorsFromRightDirection(FindingsCounter).Key <> SeparatorsFromRightDirection(FindingsStartIndex).Key Then
                        'Throw New FormatException("Too many decimal separators: " & s)
                        Return False
                    ElseIf HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(sb, SeparatorsFromRightDirection(FindingsCounter)) = False Then
                        'Throw New FormatException("Rule mismatch: HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator: " & s)
                        Return False
                    End If
                Next
                For FindingsCounter As Integer = FindingsStartIndex To SeparatorsFromRightDirection.Count - 1
                    sb.Remove(SeparatorsFromRightDirection(FindingsCounter).Value, 1)
                Next
            End If

            'Refresh separators list
            SeparatorsFromRightDirection = FindSeparatorsFromRightDirection(sb)
            If SeparatorsFromRightDirection Is Nothing Then
                'FindSeparatorsFromRightDirection failed
                Return False
            End If

            'Re-check again and convert decimal separator "," to "." to match InvariantCulture
            If SeparatorsFromRightDirection.Count = 0 Then
                'Ok - nothing to do, here
            ElseIf SeparatorsFromRightDirection.Count = 1 Then
                'Ok - just convert decimal separator "," to "." to match InvariantCulture (if applicable)
                If SeparatorsFromRightDirection(0).Key = ","c Then
                    sb.Chars(SeparatorsFromRightDirection(0).Value) = "."c
                End If
            Else 'If SeparatorsFromRightDirection.Count > 1 Then
                'Failure: only 1 separator allowed
                Return False
            End If

            'Try parsing safe value
            Return Double.TryParse(sb.ToString, Globalization.NumberStyles.AllowDecimalPoint Or Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, result)
        End Function

        Friend Shared Function HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(s As String, foundSeparator As KeyValuePair(Of Char, Integer)) As Boolean
            Return HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(New System.Text.StringBuilder(s), foundSeparator)
        End Function

        Friend Shared Function HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(sb As System.Text.StringBuilder, foundSeparator As KeyValuePair(Of Char, Integer)) As Boolean
            If foundSeparator.Value = 0 Then
                'no char before separator
                Return False
            ElseIf IsDigitChar(sb.Chars(foundSeparator.Value - 1)) = False Then
                'not a digit
                Return False
            ElseIf sb.Length - 1 < foundSeparator.Value + 3 Then
                'no 3 further chars
                Return False
            ElseIf IsDigitChar(sb.Chars(foundSeparator.Value + 1)) = False Then
                'not a digit
                Return False
            ElseIf IsDigitChar(sb.Chars(foundSeparator.Value + 2)) = False Then
                'not a digit
                Return False
            ElseIf IsDigitChar(sb.Chars(foundSeparator.Value + 3)) = False Then
                'not a digit
                Return False
            ElseIf sb.Length - 1 >= foundSeparator.Value + 4 AndAlso IsSeparatorChar(sb.Chars(foundSeparator.Value + 4)) = False Then
                'not a digit
                Return False
            Else
                'all checks successful
                Return True
            End If
        End Function

        Private Shared Function IsSeparatorChar(value As Char) As Boolean
            Select Case value
                Case ","c, "."c
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsDigitChar(value As Char) As Boolean
            Select Case value
                Case "0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Enum SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult As Byte
            MustBeDecimalSeparator = 1
            UnsafeDecision = 2
        End Enum

        Private Shared Function SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehind(sb As System.Text.StringBuilder, foundSeparator As KeyValuePair(Of Char, Integer)) As SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult
            Dim NumberOfDigitsBehind As Integer = sb.Length - 1 - foundSeparator.Value
            Select Case NumberOfDigitsBehind
                Case 0, 1, 2
                    'it's a decimal separator!
                    Return SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator
                Case 3
                    'it might be a decimal separator or a thousands separator -> method must fail for unsafety reasons
                    Return SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.UnsafeDecision
                Case Else
                    '4 or more digits -> must be a decimal separator
                    Return SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator
            End Select
        End Function

        Private Shared Function FindSeparatorsFromRightDirection(sb As System.Text.StringBuilder) As List(Of KeyValuePair(Of Char, Integer))
            Dim SeparatorsFromRightDirection As New List(Of KeyValuePair(Of Char, Integer))
            For MyCounter As Integer = sb.Length - 1 To 0 Step -1
                Dim CurrentChar As Char = sb.Chars(MyCounter)
                Select Case CurrentChar
                    Case "0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c, "-"c
                    'OK
                    Case ","c, "."c
                        SeparatorsFromRightDirection.Add(New KeyValuePair(Of Char, Integer)(CurrentChar, MyCounter))
                    Case " "c, "+"c
                        'Not allowed
                        Return Nothing
                    Case Else
                        'Throw New FormatException("Invalid number value for conversion to Double: " & s)
                        Return Nothing
                End Select
            Next
            Return SeparatorsFromRightDirection
        End Function

    End Class

End Namespace