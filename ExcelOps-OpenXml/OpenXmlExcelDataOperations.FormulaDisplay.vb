Imports System.Globalization
Imports System.Text.RegularExpressions
Imports ExcelNumberFormat

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly NumberFormatTokens As New Regex(
                """(?:[^""]|"""")*""|\\.|[_*].|\[[^\]]*\]|AM/PM|A/P|[A-Z]+|[,.]",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly GermanFormatColors As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"[Schwarz]", "[Black]"}, {"[Blau]", "[Blue]"}, {"[Zyan]", "[Cyan]"},
                {"[Grün]", "[Green]"}, {"[Magenta]", "[Magenta]"}, {"[Rot]", "[Red]"},
                {"[Weiß]", "[White]"}, {"[Gelb]", "[Yellow]"}
            }

            Private Function EvaluateDisplayFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                     formulaCell As ExcelCell) As Object
                Select Case name
                    Case "ARRAYTOTEXT", "VALUETOTEXT"
                        If arguments.Count < 1 OrElse arguments.Count > 2 Then Return New FormulaError("#VALUE!")
                        Dim raw = EvaluateExpression(arguments(0), formulaCell)
                        Dim data = AsMatrix(raw)
                        Dim options = If(arguments.Count = 2, EvaluateExpression(arguments(1), formulaCell), 0.0)
                        If name = "ARRAYTOTEXT" Then
                            Dim optionMatrix = AsMatrix(options)
                            Dim results = optionMatrix.Values.Select(Function(optionValue) ArrayToWorksheetText(data, optionValue)).ToArray()
                            If results.Length = 1 Then Return results(0)
                            Return New FormulaMatrix(optionMatrix.Rows, optionMatrix.Columns, results)
                        End If
                        Dim format = AsMatrix(options)
                        Dim rows = Math.Max(data.Rows, format.Rows), columns = Math.Max(data.Columns, format.Columns)
                        Dim values = AllocateMatrixValues(rows, columns)
                        For row = 0 To rows - 1
                            For column = 0 To columns - 1
                                Dim strict = WorksheetTextFormat(format.BroadcastAt(row, column))
                                values(row * columns + column) = If(TypeOf strict Is FormulaError, strict,
                                    WorksheetValueText(data.BroadcastAt(row, column), DirectCast(strict, Boolean)))
                            Next
                        Next
                        If values.Length = 1 Then Return values(0)
                        Return New FormulaMatrix(rows, columns, values)
                    Case "TEXT"
                        If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                        Return BroadcastReferenceArguments(arguments, formulaCell, AddressOf FormatWorksheetText)
                    Case Else
                        If arguments.Count = 0 OrElse arguments.Count > If(name = "DOLLAR", 2, 3) Then Return New FormulaError("#VALUE!")
                        Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) FixedWorksheetText(name, values))
                End Select
            End Function

            Private Function ArrayToWorksheetText(data As FormulaMatrix, format As Object) As Object
                Dim selector = WorksheetTextFormat(format)
                If TypeOf selector Is FormulaError Then Return selector
                If data.Values.Length = 1 Then
                    If data.Values(0) Is Nothing Then Return New FormulaError("#VALUE!")
                    If TypeOf data.Values(0) Is FormulaError Then Return data.Values(0)
                End If
                Dim strict = DirectCast(selector, Boolean)
                Dim german = ConversionCulture.TwoLetterISOLanguageName = "de"
                Dim separator = If(german, ";", ",")
                Dim labels = data.Values.Select(Function(value) WorksheetValueText(value, strict)).ToArray()
                If Not strict Then Return CheckedWorksheetText(String.Join(separator & " ", labels))
                Dim rows As New List(Of String)(data.Rows)
                For row = 0 To data.Rows - 1
                    rows.Add(String.Join(If(german, ".", ","), labels, row * data.Columns, data.Columns))
                Next
                Return CheckedWorksheetText("{" & String.Join(";", rows) & "}")
            End Function

            Private Function WorksheetTextFormat(value As Object) As Object
                Dim numeric = NumberValue(value)
                If TypeOf numeric Is FormulaError Then Return numeric
                Dim mode = Math.Truncate(DirectCast(numeric, Double))
                If mode <> 0 AndAlso mode <> 1 Then Return New FormulaError("#VALUE!")
                Return mode = 1
            End Function

            Private Function WorksheetValueText(value As Object, strict As Boolean) As String
                Dim text = TryCast(value, String)
                If text IsNot Nothing AndAlso strict Then Return """" & text.Replace("""", """""") & """"
                Dim failure = TryCast(value, FormulaError)
                If failure IsNot Nothing Then Return If(ConversionCulture.TwoLetterISOLanguageName = "de", DisplayGermanError(failure.Code), failure.Code)
                If TypeOf value Is Double Then
                    Dim number = DirectCast(value, Double)
                    If number <> 0.0 AndAlso Math.Abs(number) < 0.0001 AndAlso Math.Abs(number) >= 0.00000000000001 Then
                        Return number.ToString("0.#############################", ConversionCulture)
                    End If
                End If
                Return TextValue(value)
            End Function

            Private Shared Function CheckedWorksheetText(text As String) As Object
                If text.Length > 32767 Then Return New FormulaError("#VALUE!")
                Return text
            End Function

            Private Function FixedWorksheetText(name As String, arguments As Object()) As Object
                Dim number = NumberValue(arguments(0))
                If TypeOf number Is FormulaError Then Return number
                Dim decimals = NumberValue(If(arguments.Length >= 2, arguments(1), 2.0))
                If TypeOf decimals Is FormulaError Then Return decimals
                Dim withoutGroups = If(arguments.Length >= 3, LogicalOptionValue(arguments(2)), False)
                If TypeOf withoutGroups Is FormulaError Then Return withoutGroups
                Dim count = Math.Truncate(DirectCast(decimals, Double))
                If count > 127 Then Return New FormulaError("#VALUE!")
                Dim rounded = RoundAt(DirectCast(number, Double), count, "ROUND")
                If TypeOf rounded Is FormulaError Then Return rounded
                Dim value = DirectCast(rounded, Double)
                Dim places = CInt(Math.Max(0.0, count))
                Dim displayedPlaces = Math.Min(places, 30)
                Dim result As String
                If name = "DOLLAR" Then
                    Dim format = DirectCast(ConversionCulture.NumberFormat.Clone(), NumberFormatInfo)
                    If Equals(ConversionCulture, CultureInfo.InvariantCulture) Then format = CultureInfo.GetCultureInfo("en-US").NumberFormat
                    Dim formatCode = "C" & displayedPlaces.ToString(CultureInfo.InvariantCulture)
                    If value = 0.0 AndAlso count >= 0.0 AndAlso DirectCast(number, Double) < 0.0 Then
                        ' Retain the negative currency pattern even on runtimes that drop the sign of negative zero.
                        result = (-1.0).ToString(formatCode, format)
                        Dim digit = result.IndexOf("1"c)
                        If digit >= 0 Then result = result.Remove(digit, 1).Insert(digit, "0")
                    Else
                        result = value.ToString(formatCode, format)
                    End If
                    If places > displayedPlaces Then
                        Dim at = result.LastIndexOf(format.CurrencyDecimalSeparator, StringComparison.Ordinal) + format.CurrencyDecimalSeparator.Length + displayedPlaces
                        result = result.Insert(at, New String("0"c, places - displayedPlaces))
                    End If
                Else
                    Dim format = If(DirectCast(withoutGroups, Boolean), "F", "N") & displayedPlaces.ToString(CultureInfo.InvariantCulture)
                    result = If(value = 0.0, 0.0, value).ToString(format, ConversionCulture) & New String("0"c, places - displayedPlaces)
                End If
                Return result
            End Function

            Private _excelDateTextCulture As CultureInfo

            Private Function FormatWorksheetText(arguments As Object()) As Object
                Dim german = ConversionCulture.TwoLetterISOLanguageName = "de"
                Dim pattern = TextValue(arguments(1))
                If german Then
                    If pattern.Equals("General", StringComparison.OrdinalIgnoreCase) Then Return New FormulaError("#VALUE!")
                    If pattern.Equals("Standard", StringComparison.OrdinalIgnoreCase) Then pattern = "General"
                End If
                Dim value = If(arguments(0), 0.0)
                If TypeOf value Is String Then
                    Dim parsed = NumberValue(value)
                    If TypeOf parsed Is Double Then value = parsed
                End If
                If german AndAlso Not pattern.Equals("General", StringComparison.OrdinalIgnoreCase) Then
                    Dim serial = If(TypeOf value Is Double, DirectCast(value, Double), 0.0)
                    pattern = InvariantGermanNumberFormat(pattern, serial)
                    If pattern Is Nothing Then Return New FormulaError("#VALUE!")
                End If
                Dim format As New NumberFormat(pattern)
                If Not format.IsValid Then Return New FormulaError("#VALUE!")
                If TypeOf value Is Boolean OrElse pattern = "@" OrElse pattern.Equals("General", StringComparison.OrdinalIgnoreCase) Then
                    Return WorksheetValueText(value, strict:=False)
                End If
                If TypeOf value Is Double Then
                    Dim number = DirectCast(value, Double)
                    If number < 0.0 AndAlso (format.IsDateTimeFormat OrElse format.IsTimeSpanFormat) Then Return New FormulaError("#VALUE!")
                    Dim numeric = Regex.Match(pattern, "\A[0#?,]+(?:\.(?<fraction>[0#?]+))?(?<percent>%)?\z", RegexOptions.CultureInvariant)
                    If numeric.Success Then
                        Dim places = numeric.Groups("fraction").Value.Length
                        If numeric.Groups("percent").Success Then places += 2
                        value = RoundAt(number, places, "ROUND")
                        If TypeOf value Is FormulaError Then Return value
                    End If
                End If
                If TypeOf value Is Double AndAlso DirectCast(value, Double) = 0.0 AndAlso
                    Regex.IsMatch(pattern, "\A[0#]+(?:\.[0#]+)?E[+-]0+\z", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
                    Return 0.0.ToString(pattern, ConversionCulture)
                End If
                Try
                    Dim culture = If(format.IsDateTimeFormat, ExcelDateTextCulture(), ConversionCulture)
                    Return CheckedWorksheetText(format.Format(value, culture, Uses1904Dates))
                Catch ex As ArgumentException
                    Return New FormulaError("#VALUE!")
                End Try
            End Function

            Private Function ExcelDateTextCulture() As CultureInfo
                If _excelDateTextCulture IsNot Nothing Then Return _excelDateTextCulture
                Dim culture = ConversionCulture
                If String.Equals(culture.Name, "de-DE", StringComparison.OrdinalIgnoreCase) Then
                    Dim months = culture.DateTimeFormat.AbbreviatedMonthNames
                    If Not String.Equals(months(2), "Mrz", StringComparison.Ordinal) Then
                        culture = DirectCast(culture.Clone(), CultureInfo)
                        ' Excel's German MMM token abbreviates March as Mrz.
                        months(2) = "Mrz"
                        culture.DateTimeFormat.AbbreviatedMonthNames = months
                    End If
                End If
                _excelDateTextCulture = culture
                Return culture
            End Function

            Private Shared Function InvariantGermanNumberFormat(format As String, serial As Double) As String
                Dim tokens = NumberFormatTokens.Matches(format).Cast(Of Match)().Select(Function(part) part.Value).ToArray()
                Dim localizedDate = tokens.Any(Function(part) IsFormatLetter(part, "t"c) OrElse IsFormatLetter(part, "j"c))
                Dim foreignDate = tokens.Any(Function(part) IsFormatLetter(part, "d"c) OrElse IsFormatLetter(part, "y"c))
                If foreignDate AndAlso Not localizedDate AndAlso serial < 0.0 Then Return Nothing
                Dim converted = NumberFormatTokens.Replace(format,
                    Function(token)
                        Dim code = token.Value
                        If code = "." Then Return If(localizedDate OrElse foreignDate, """.""", ",")
                        If code = "," Then Return "."
                        Dim color As String = Nothing
                        If GermanFormatColors.TryGetValue(code, color) Then Return color
                        If IsFormatLetter(code, "t"c) Then Return New String("d"c, code.Length)
                        If IsFormatLetter(code, "j"c) Then Return New String("y"c, code.Length)
                        If IsFormatLetter(code, "d"c) OrElse IsFormatLetter(code, "y"c) Then Return """" & code & """"
                        If foreignDate AndAlso Not localizedDate AndAlso IsFormatLetter(code, "m"c) Then
                            Dim second = Math.Round((serial - Math.Floor(serial)) * 86400.0, MidpointRounding.AwayFromZero)
                            Dim minute = Math.Floor(second / 60.0) Mod 60.0
                            Return """" & minute.ToString(If(code.Length > 1, "00", "0"), CultureInfo.InvariantCulture) & """"
                        End If
                        Return code
                    End Function)
                If Regex.IsMatch(converted, "\A[0#?]+\.[0#?,]+\z", RegexOptions.CultureInvariant) Then
                    Return converted.Replace(",", "")
                End If
                Return converted
            End Function

            Private Shared Function IsFormatLetter(token As String, letter As Char) As Boolean
                Return token.All(Function(character) Char.ToLowerInvariant(character) = letter)
            End Function

        End Class

    End Class

End Namespace
