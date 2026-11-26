Imports System.Text.RegularExpressions
Imports PCRE

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly WorksheetRegexLimits As New PcreMatchSettings With {
                .MatchLimit = 100000UI,
                .DepthLimit = 1000UI,
                .HeapLimit = 8192UI
            }

            Private Shared ReadOnly BackslashReplacementGroups As New Regex(
                "\\Q.*?(?:\\E|\z)|\\\\|\\(?<group>[0-9]+)", RegexOptions.Singleline Or RegexOptions.CultureInvariant)

            Private Function SelectRegexLookupValue(needle As Object, keys As Object(), reverse As Boolean,
                                                    selectValue As Func(Of Integer, Object)) As Object
                Dim found = WithWorksheetRegex(TextValue(needle), ignoreCase:=False,
                    Function(regex)
                        For offset = 0 To keys.Length - 1
                            Dim candidate = If(reverse, keys.Length - 1 - offset, offset)
                            If TypeOf keys(candidate) Is FormulaError Then Continue For
                            If RegexMatch(regex, TextValue(keys(candidate))).Success Then Return candidate
                        Next
                        Return -1
                    End Function)
                If TypeOf found Is FormulaError Then Return found
                Return selectValue(DirectCast(found, Integer))
            End Function

            Private Function EvaluateRegexFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                   formulaCell As ExcelCell) As Object
                Dim isReplacement = name = "REGEXREPLACE"
                Dim maximum = If(isReplacement, 5, If(name = "REGEXEXTRACT", 4, 3))
                If arguments.Count < If(isReplacement, 3, 2) OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")

                Dim inputs = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                If Not inputs.Any(AddressOf IsMultiValue) Then
                    Return EvaluateRegexScalar(name, inputs.Select(Function(value) Scalar(value)).ToArray())
                End If

                Dim matrices = inputs.Select(Function(value) AsMatrix(value)).ToArray()
                Dim rows = matrices.Max(Function(matrix) matrix.Rows)
                Dim columns = matrices.Max(Function(matrix) matrix.Columns)
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim indices = (Row:=row, Column:=column)
                        Dim values = matrices.Select(Function(matrix) matrix.BroadcastAt(indices.Row, indices.Column)).ToArray()
                        Dim result = EvaluateRegexScalar(name, values)
                        ' An array of inputs takes the first value of each extraction.
                        Dim extraction = TryCast(result, FormulaMatrix)
                        results(row * columns + column) = If(extraction Is Nothing, result, extraction.At(0, 0))
                    Next
                Next
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function EvaluateRegexScalar(name As String, arguments As Object()) As Object
                Dim failure = arguments.OfType(Of FormulaError)().FirstOrDefault()
                If failure IsNot Nothing Then Return failure
                Dim isReplacement = name = "REGEXREPLACE"
                Dim caseIndex = If(isReplacement, 4, If(name = "REGEXEXTRACT", 3, 2))
                Dim caseValue = NumberValue(If(arguments.Length > caseIndex, arguments(caseIndex), Nothing))
                If TypeOf caseValue Is FormulaError Then Return caseValue
                Dim caseMode = Math.Truncate(DirectCast(caseValue, Double))
                If caseMode <> 0 AndAlso caseMode <> 1 Then Return New FormulaError("#VALUE!")

                Dim subject = TextValue(arguments(0))
                Dim pattern = TextValue(arguments(1))
                If name = "REGEXTEST" Then
                    Return WithWorksheetRegex(pattern, caseMode = 1, Function(regex) RegexMatch(regex, subject).Success)
                End If

                Dim modeIndex = If(isReplacement, 3, 2)
                Dim modeValue = NumberValue(If(arguments.Length > modeIndex, arguments(modeIndex), Nothing))
                If TypeOf modeValue Is FormulaError Then Return modeValue
                Dim mode = Math.Truncate(DirectCast(modeValue, Double))
                If isReplacement Then
                    Dim replacement = TextValue(arguments(2))
                    Return WithWorksheetRegex(pattern, caseMode = 1, Function(regex) ReplaceRegexMatches(regex, subject, replacement, mode))
                End If
                If mode < 0 OrElse mode > 2 Then Return New FormulaError("#VALUE!")
                Return WithWorksheetRegex(pattern, caseMode = 1, Function(regex) ExtractRegexMatches(regex, subject, CInt(mode)))
            End Function

            Private Shared Function ExtractRegexMatches(regex As PcreRegex, subject As String, mode As Integer) As Object
                If mode = 1 Then
                    Dim matches = RegexMatches(regex, subject).Select(Function(match) CType(match.Value, Object)).ToArray()
                    If matches.Length = 0 Then Return New FormulaError("#N/A")
                    If matches.Length > ColumnCount Then Return New FormulaError(SpillErrorCode)
                    Return New FormulaMatrix(1, matches.Length, matches)
                End If

                Dim groups = CInt(regex.PatternInfo.CaptureCount)
                If mode = 2 AndAlso groups = 0 Then Return New FormulaError("#VALUE!")
                Dim first = RegexMatch(regex, subject)
                If Not first.Success Then Return New FormulaError("#N/A")
                If mode = 0 Then Return first.Value

                Dim values = AllocateMatrixValues(1, groups)
                For index = 1 To groups
                    If index <= first.CaptureCount AndAlso first(index).Success Then
                        values(index - 1) = first(index).Value
                    Else
                        values(index - 1) = New FormulaError("#N/A")
                    End If
                Next
                Return New FormulaMatrix(1, groups, values)
            End Function

            Private Shared Function ReplaceRegexMatches(regex As PcreRegex, subject As String, replacement As String, occurrence As Double) As Object
                Dim options = PcreSubstituteOptions.SubstituteGlobal Or PcreSubstituteOptions.SubstituteExtended Or
                              PcreSubstituteOptions.SubstituteUnsetEmpty
                Dim selectOccurrence As PcreSubstituteCalloutFunc = Nothing
                If occurrence <> 0 Then
                    Dim total = RegexMatches(regex, subject).Count()
                    Dim selected = If(occurrence > 0, occurrence, total + occurrence + 1)
                    If selected < 1 OrElse selected > total Then Return subject
                    selectOccurrence = Function(context) If(context.SubstitutionCount = selected,
                                                           PcreSubstituteCalloutResult.Pass, PcreSubstituteCalloutResult.Fail)
                End If
                replacement = BackslashReplacementGroups.Replace(replacement,
                    Function(token) If(token.Groups("group").Success, "${" & token.Groups("group").Value & "}", token.Value))
                Dim result = regex.Substitute(subject, replacement, 0, options, selectOccurrence, WorksheetRegexLimits)
                If result.Length > MaximumTextLength Then Return New FormulaError("#VALUE!")
                Return result
            End Function

            Private Shared Function RegexMatch(regex As PcreRegex, subject As String) As PcreMatch
                Return regex.Match(subject, 0, PcreMatchOptions.None,
                    DirectCast(Nothing, Func(Of PcreCallout, PcreCalloutResult)), WorksheetRegexLimits)
            End Function

            Private Shared Function RegexMatches(regex As PcreRegex, subject As String) As IEnumerable(Of PcreMatch)
                Return regex.Matches(subject, 0, PcreMatchOptions.None,
                    DirectCast(Nothing, Func(Of PcreCallout, PcreCalloutResult)), WorksheetRegexLimits)
            End Function

            Private Shared Function WithWorksheetRegex(pattern As String, ignoreCase As Boolean, evaluate As Func(Of PcreRegex, Object)) As Object
                Dim options = PcreOptions.Utf Or PcreOptions.Ucp
                If ignoreCase Then options = options Or PcreOptions.IgnoreCase
                Try
                    Return evaluate(New PcreRegex(pattern, options))
                Catch ex As PcrePatternException
                    Return New FormulaError("#VALUE!")
                Catch ex As PcreException When ex.ErrorCode = PcreErrorCode.MatchLimit OrElse ex.ErrorCode = PcreErrorCode.DepthLimit OrElse
                                               ex.ErrorCode = PcreErrorCode.HeapLimit
                    Throw New NotSupportedException("The regular expression exceeds the supported calculation resource limits.", ex)
                Catch ex As PcreException
                    Return New FormulaError("#VALUE!")
                Catch ex As DllNotFoundException
                    Throw New NotSupportedException("PCRE2 is unavailable on this platform.", ex)
                Catch ex As BadImageFormatException
                    Throw New NotSupportedException("The installed PCRE2 library is incompatible with this process.", ex)
                End Try
            End Function

        End Class

    End Class

End Namespace
