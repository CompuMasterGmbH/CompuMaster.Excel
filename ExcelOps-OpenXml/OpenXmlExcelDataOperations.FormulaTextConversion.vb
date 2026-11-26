Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly WorksheetNumberPattern As New Regex(
                "\A[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?\z", RegexOptions.CultureInvariant)

            Private Function EvaluateTextConversion(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                   formulaCell As ExcelCell) As Object
                If arguments.Count = 0 OrElse arguments.Count > If(name = "NUMBERVALUE", 3, 1) Then Return New FormulaError("#VALUE!")
                Dim inputs = arguments.Select(Function(argument) AsMatrix(EvaluateExpression(argument, formulaCell))).ToArray()
                Dim rows = inputs.Max(Function(input) input.Rows)
                Dim columns = inputs.Max(Function(input) input.Columns)
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim value = inputs(0).BroadcastAt(row, column)
                        Dim result As Object
                        If TypeOf value Is FormulaError Then
                            result = value
                        ElseIf name = "NUMBERVALUE" Then
                            Dim format = ConversionCulture.NumberFormat
                            Dim decimalSeparator = If(inputs.Length > 1, inputs(1).BroadcastAt(row, column), format.NumberDecimalSeparator)
                            Dim groupSeparator = If(inputs.Length > 2, inputs(2).BroadcastAt(row, column), format.NumberGroupSeparator)
                            result = ConvertWorksheetNumber(value, decimalSeparator, groupSeparator, discardSpaces:=True)
                        Else
                            result = ConvertWorksheetText(name, value)
                        End If
                        values(row * columns + column) = result
                    Next
                Next
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Function ConvertWorksheetText(name As String, value As Object) As Object
                If name = "CHAR" OrElse name = "UNICHAR" Then Return WorksheetCharacter(value, name = "UNICHAR")
                If name = "VALUE" Then
                    If value Is Nothing Then Return 0.0
                    If TypeOf value Is Double Then Return value
                    Dim format = ConversionCulture.NumberFormat
                    Return ConvertWorksheetNumber(value, format.NumberDecimalSeparator, format.NumberGroupSeparator, discardSpaces:=False)
                End If

                Dim text = TextValue(value)
                Select Case name
                    Case "CLEAN"
                        Return New String(text.Where(Function(character)
                                                         Dim code = AscW(character) And &HFFFF
                                                         Return code >= 32 AndAlso (code < 128 OrElse code >= 160)
                                                     End Function).ToArray())
                    Case "PROPER"
                        Return WorksheetProperCase(text)
                    Case "ENCODEURL"
                        Return EncodeWorksheetUrl(text)
                    Case "CODE", "UNICODE"
                        If text.Length = 0 Then Return New FormulaError("#VALUE!")
                        Dim pair = Char.IsSurrogatePair(text, 0)
                        If name = "CODE" Then
                            Dim encoding = DirectCast(WorksheetCodePage().Clone(), Encoding)
                            encoding.EncoderFallback = New EncoderReplacementFallback("?")
                            Return CDbl(encoding.GetBytes(text.Substring(0, If(pair, 2, 1)))(0))
                        End If
                        If Char.IsSurrogate(text(0)) AndAlso Not pair Then Return New FormulaError("#VALUE!")
                        Return If(pair, CDbl(Char.ConvertToUtf32(text, 0)), CDbl(AscW(text(0)) And &HFFFF))
                    Case Else
                        Return New FormulaError("#VALUE!")
                End Select
            End Function

            Private Shared Function EncodeWorksheetUrl(text As String) As Object
                Const maxEncodedLength As Integer = 32766
                If text.Length > maxEncodedLength Then Return New FormulaError("#VALUE!")
                Dim encoded As New StringBuilder(text.Length)
                For Each character In Encoding.UTF8.GetBytes(text)
                    If (character >= 65 AndAlso character <= 90) OrElse (character >= 97 AndAlso character <= 122) OrElse
                       (character >= 48 AndAlso character <= 57) OrElse character = 45 OrElse character = 46 OrElse character = 95 Then
                        encoded.Append(ChrW(character))
                    Else
                        encoded.Append("%"c).Append(character.ToString("X2", CultureInfo.InvariantCulture))
                    End If
                    If encoded.Length > maxEncodedLength Then Return New FormulaError("#VALUE!")
                Next
                Return encoded.ToString()
            End Function

            Private Function WorksheetCharacter(value As Object, unicode As Boolean) As Object
                Dim number = NumberValue(value)
                If TypeOf number Is FormulaError Then Return number
                Dim whole = Math.Truncate(DirectCast(number, Double))
                If whole < 1.0 OrElse whole > If(unicode, &H10FFFF, 255) Then Return New FormulaError("#VALUE!")
                Dim code = CInt(whole)
                If Not unicode Then Return WorksheetCodePage().GetString({CByte(code)})
                If (code >= &HD800 AndAlso code <= &HDFFF) OrElse (code >= &HFDD0 AndAlso code <= &HFDEF) OrElse
                    (code And &HFFFF) >= &HFFFE Then Return New FormulaError("#N/A")
                Return Char.ConvertFromUtf32(code)
            End Function

            Private Function WorksheetCodePage() As Encoding
                Dim codePage = ConversionCulture.TextInfo.ANSICodePage
#If NETFRAMEWORK Then
                Return Encoding.GetEncoding(If(codePage = 0, 1252, codePage))
#Else
                Return If(CodePagesEncodingProvider.Instance.GetEncoding(codePage), CodePagesEncodingProvider.Instance.GetEncoding(1252))
#End If
            End Function

            Private Function WorksheetProperCase(value As String) As String
                Dim result As New StringBuilder(value.Length)
                Dim newWord = True
                Dim culture = ConversionCulture
                Dim turkic = culture.TwoLetterISOLanguageName = "tr" OrElse culture.TwoLetterISOLanguageName = "az"
                For index = 0 To value.Length - 1
                    Dim character = value(index)
                    If Char.IsLetter(character) Then
                        If character = ChrW(&H130) AndAlso Not turkic Then
                            result.Append(If(newWord, "I"c, "i"c))
                        ElseIf Not newWord AndAlso character = ChrW(&H3A3) AndAlso
                            (index = value.Length - 1 OrElse Not Char.IsLetter(value(index + 1))) Then
                            result.Append(ChrW(&H3C2))
                        Else
                            result.Append(If(newWord, Char.ToUpper(character, culture), Char.ToLower(character, culture)))
                        End If
                        newWord = False
                    Else
                        result.Append(character)
                        newWord = True
                    End If
                Next
                Return result.ToString()
            End Function

            Private Function ConvertWorksheetNumber(value As Object, decimalArgument As Object, groupArgument As Object,
                                                    discardSpaces As Boolean) As Object
                If TypeOf decimalArgument Is FormulaError Then Return decimalArgument
                If TypeOf groupArgument Is FormulaError Then Return groupArgument
                If TypeOf value Is Boolean Then Return New FormulaError("#VALUE!")
                Dim decimalText = TextValue(decimalArgument)
                Dim groupText = TextValue(groupArgument)
                If decimalText.Length = 0 OrElse (discardSpaces AndAlso groupText.Length = 0) Then Return New FormulaError("#VALUE!")
                Dim decimalSeparator = decimalText.Substring(0, 1)
                Dim groupSeparator = If(groupText.Length = 0, "", groupText.Substring(0, 1))
                If decimalSeparator = groupSeparator Then Return New FormulaError("#VALUE!")

                Dim text = TextValue(value)
                text = If(discardSpaces, text.Replace(" ", ""), text.Trim())
                Dim number = ReadWorksheetNumber(text, decimalSeparator, groupSeparator, discardSpaces)
                If number.HasValue Then Return number.Value
                Dim serial = ReadDateTimeText(text)
                If serial.HasValue Then Return serial.Value
                Return New FormulaError("#VALUE!")
            End Function

            Private Function ReadWorksheetNumber(text As String, decimalSeparator As String, groupSeparator As String,
                                                 looseGrouping As Boolean) As Double?
                Dim percentCount As Integer
                While text.EndsWith("%", StringComparison.Ordinal)
                    percentCount += 1
                    If Not looseGrouping AndAlso percentCount > 1 Then Return Nothing
                    text = text.Substring(0, text.Length - 1).TrimEnd()
                End While

                Dim parentheses = text.StartsWith("(", StringComparison.Ordinal) AndAlso text.EndsWith(")", StringComparison.Ordinal)
                If parentheses Then
                    If text.Length < 3 Then Return Nothing
                    text = text.Substring(1, text.Length - 2).Trim()
                End If
                Dim negative = parentheses
                Dim signSeen = parentheses
                If text.StartsWith("+", StringComparison.Ordinal) OrElse text.StartsWith("-", StringComparison.Ordinal) Then
                    If signSeen Then Return Nothing
                    signSeen = True
                    negative = text(0) = "-"c
                    text = text.Substring(1)
                End If
                Dim currency = ConversionCulture.NumberFormat.CurrencySymbol
                If currency.Length > 0 AndAlso text.StartsWith(currency, StringComparison.Ordinal) Then
                    text = text.Substring(currency.Length).TrimStart()
                    If text.StartsWith("+", StringComparison.Ordinal) OrElse text.StartsWith("-", StringComparison.Ordinal) Then
                        If signSeen Then Return Nothing
                        signSeen = True
                        negative = text(0) = "-"c
                        text = text.Substring(1)
                    End If
                ElseIf currency.Length > 0 AndAlso text.EndsWith(currency, StringComparison.Ordinal) Then
                    text = text.Substring(0, text.Length - currency.Length).TrimEnd()
                End If

                Dim decimalIndex = text.IndexOf(decimalSeparator, StringComparison.Ordinal)
                Dim grouped = If(decimalIndex < 0, text, text.Substring(0, decimalIndex))
                Dim suffix = If(decimalIndex < 0, "", text.Substring(decimalIndex + 1))
                If groupSeparator.Length > 0 Then
                    If suffix.Contains(groupSeparator) Then Return Nothing
                    If Not looseGrouping AndAlso Not HasValidValueGroups(grouped, groupSeparator) Then Return Nothing
                    grouped = grouped.Replace(groupSeparator, "")
                End If
                Dim normalized = If(decimalIndex < 0, grouped, grouped & "." & suffix)
                If normalized.Length = 0 AndAlso looseGrouping AndAlso Not parentheses AndAlso Not signSeen Then Return 0.0
                If Not WorksheetNumberPattern.IsMatch(normalized) Then Return Nothing

                Dim parsed As Double
                If Not Double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, parsed) OrElse
                    Double.IsInfinity(parsed) OrElse Double.IsNaN(parsed) Then Return Nothing
                Return If(negative, -parsed, parsed) / Math.Pow(100.0, percentCount)
            End Function

            Private Function HasValidValueGroups(text As String, separator As String) As Boolean
                Dim exponentIndex = text.IndexOfAny({"E"c, "e"c})
                If exponentIndex >= 0 AndAlso text.Substring(exponentIndex).Contains(separator) Then Return False
                Dim integerPart = If(exponentIndex >= 0, text.Substring(0, exponentIndex), text)
                If Not integerPart.Contains(separator) Then Return True
                Dim groups = integerPart.Split({separator}, StringSplitOptions.None)
                Dim sizes = ConversionCulture.NumberFormat.NumberGroupSizes
                If sizes.Length = 0 Then Return False
                Dim size = sizes(0)
                For index = groups.Length - 1 To 1 Step -1
                    If size = 0 OrElse groups(index).Length <> size Then Return False
                    size = sizes(Math.Min(groups.Length - index, sizes.Length - 1))
                Next
                Return groups(0).Length >= 1 AndAlso (size = 0 OrElse groups(0).Length <= size)
            End Function

        End Class

    End Class

End Namespace
