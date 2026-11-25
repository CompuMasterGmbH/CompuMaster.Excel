Imports System.Globalization
Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly DateTextTime As New Regex(
                "\A(?:(?<date>.+?)\s+)?(?<hour>[0-9]{1,8})\s*:\s*(?<minute>[0-9]{1,8})" &
                "(?:\s*:\s*(?<second>[0-9]{1,8}))?(?:(?<separator>[.,])(?<fraction>[0-9]+))?(?:\s+(?<period>[AP]M))?\z",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly DateTextNumeric As New Regex(
                "\A(?<first>[0-9]{1,4})\s*(?<separator>[-./])\s*(?<second>[0-9]{1,2})(?:\s*\k<separator>\s*(?<third>[0-9]{1,4}))?\z",
                RegexOptions.CultureInvariant)
            Private Shared ReadOnly DateTextNamed As New Regex(
                "\A(?:[0-9]{1,4}[./ -]+)?[\p{L}\p{M}]+\.?(?:[./ -]+[0-9]{1,4}){0,2}\z",
                RegexOptions.CultureInvariant)
            Private Shared ReadOnly DateTextParts As New Regex("[0-9]+|[\p{L}\p{M}]+", RegexOptions.CultureInvariant)

            Private Function EvaluateDateTextFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                      formulaCell As ExcelCell) As Object
                If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                Dim value = Scalar(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf value Is FormulaError Then Return value
                Dim text = TryCast(value, String)?.Trim()
                Dim serial = ReadDateTimeText(text)
                If Not serial.HasValue Then Return New FormulaError("#VALUE!")
                Dim day = Math.Floor(serial.Value)
                Return If(name = "DATEVALUE", day, serial.Value - day)
            End Function

            Private Function ReadDateTimeText(text As String) As Double?
                If String.IsNullOrEmpty(text) Then Return Nothing
                Dim time = DateTextTime.Match(text)
                Dim serial As Double
                If time.Success Then
                    Dim duration = ReadTextTime(time)
                    If Not duration.HasValue Then Return Nothing
                    If time.Groups("date").Success Then
                        Dim start = ReadTextDate(time.Groups("date").Value.TrimEnd())
                        If Not start.HasValue Then Return Nothing
                        serial = start.Value
                    End If
                    serial += duration.Value
                Else
                    Dim start = ReadTextDate(text)
                    If Not start.HasValue Then Return Nothing
                    serial = start.Value
                End If

                Return serial
            End Function

            Private Function ReadTextTime(time As Match) As Double?
                Dim hours = ReadTimePart(time, "hour")
                Dim minutes = ReadTimePart(time, "minute")
                Dim seconds = ReadTimePart(time, "second")
                Dim period = time.Groups("period")
                If period.Success Then
                    If hours > 12 OrElse minutes > 59 OrElse seconds > 59 Then Return Nothing
                    hours = (hours Mod 12) + If(period.Value.Equals("PM", StringComparison.OrdinalIgnoreCase), 12, 0)
                ElseIf hours = 0 AndAlso time.Groups("hour").Length > 2 Then
                    Return Nothing
                End If

                Dim fraction = time.Groups("fraction")
                Dim milliseconds As Double
                If fraction.Success Then
                    If time.Groups("separator").Value <> ConversionCulture.NumberFormat.NumberDecimalSeparator OrElse period.Success Then Return Nothing
                    ' A fractional final component is seconds. With two components, Excel treats the first as minutes.
                    If Not time.Groups("second").Success Then
                        seconds = minutes
                        minutes = hours
                        hours = 0
                    End If
                    If seconds > 59 Then Return Nothing
                    Dim digits = fraction.Value
                    milliseconds = Double.Parse("0." & digits.Substring(0, Math.Min(4, digits.Length)), CultureInfo.InvariantCulture) * 1000.0
                    milliseconds = Math.Round(milliseconds, MidpointRounding.AwayFromZero)
                End If

                Return (hours * 3600.0 + minutes * 60.0 + seconds + milliseconds / 1000.0) / SecondsPerDay
            End Function

            Private Shared Function ReadTimePart(time As Match, name As String) As Integer
                Dim part = time.Groups(name)
                Return If(part.Success, Integer.Parse(part.Value, CultureInfo.InvariantCulture), 0)
            End Function

            Private Function ReadTextDate(text As String) As Double?
                Dim numeric = DateTextNumeric.Match(text)
                If numeric.Success Then
                    Dim first = numeric.Groups("first")
                    Dim second = ReadTimePart(numeric, "second")
                    Dim third = numeric.Groups("third")
                    If first.Length = 4 Then
                        If Not third.Success OrElse third.Length > 2 Then Return Nothing
                        Return DateSerialFromParts(ReadTimePart(numeric, "first"), second, ReadTimePart(numeric, "third"))
                    End If
                    If first.Length > 2 Then Return Nothing
                    Dim dateOrder = ConversionCulture.DateTimeFormat.ShortDatePattern
                    Dim monthFirst = dateOrder.IndexOf("M"c) < dateOrder.IndexOf("d"c) AndAlso numeric.Groups("separator").Value <> "."
                    Dim initial = ReadTimePart(numeric, "first")
                    Dim month = If(monthFirst, initial, second)
                    Dim day = If(monthFirst, second, initial)
                    Dim year = If(third.Success, ReadTextYear(third.Value), CurrentTextDateYear())
                    Return DateSerialFromParts(year, month, day)
                End If

                If Not DateTextNamed.IsMatch(text) Then Return Nothing
                Dim parts = DateTextParts.Matches(text).Cast(Of Match)().Select(Function(part) part.Value).ToArray()
                Dim monthText = parts.Single(Function(part) Char.IsLetter(part(0)))
                Dim monthNumber = FindTextMonth(monthText)
                If monthNumber = 0 Then Return Nothing
                Dim numbers = parts.Where(Function(part) Not Char.IsLetter(part(0))).ToArray()
                Select Case numbers.Length
                    Case 1
                        Dim value = Integer.Parse(numbers(0), CultureInfo.InvariantCulture)
                        If numbers(0).Length = 4 Then Return DateSerialFromParts(value, monthNumber, 1)
                        If numbers(0).Length > 2 Then Return Nothing
                        Dim dateInCurrentYear = DateSerialFromParts(CurrentTextDateYear(), monthNumber, value)
                        Return If(dateInCurrentYear, DateSerialFromParts(ReadTextYear(numbers(0)), monthNumber, 1))
                    Case 2
                        If numbers(0).Length = 4 Then
                            Return DateSerialFromParts(ReadTextYear(numbers(0)), monthNumber, Integer.Parse(numbers(1), CultureInfo.InvariantCulture))
                        End If
                        Return DateSerialFromParts(ReadTextYear(numbers(1)), monthNumber, Integer.Parse(numbers(0), CultureInfo.InvariantCulture))
                    Case Else
                        Return Nothing
                End Select
            End Function

            Private Function CurrentTextDateYear() As Integer
                MarkVolatile()
                Return _run.LocalTime.Value.Year
            End Function

            Private Shared Function ReadTextYear(value As String) As Integer
                Dim year = Integer.Parse(value, CultureInfo.InvariantCulture)
                If value.Length > 2 Then Return year
                Return year + If(year < 30, 2000, 1900)
            End Function

            Private Function FindTextMonth(text As String) As Integer
                Dim format = ConversionCulture.DateTimeFormat
                Dim comparer = ConversionCulture.CompareInfo
                For index = 0 To 11
                    If comparer.Compare(text, format.MonthNames(index), CompareOptions.IgnoreCase) = 0 OrElse
                       comparer.Compare(text, format.AbbreviatedMonthNames(index).TrimEnd("."c), CompareOptions.IgnoreCase) = 0 OrElse
                       (text.Length = 3 AndAlso comparer.IsPrefix(format.MonthNames(index), text, CompareOptions.IgnoreCase)) Then
                        Return index + 1
                    End If
                Next
                Return 0
            End Function

            Private Function DateSerialFromParts(year As Integer, month As Integer, day As Integer) As Double?
                If year < If(Uses1904Dates, 1904, 1900) OrElse year > 9999 OrElse month < 1 OrElse month > 12 OrElse day < 1 Then Return Nothing
                If year = 1900 AndAlso month = 2 AndAlso day = 29 Then Return 60.0
                If day > DateTime.DaysInMonth(year, month) Then Return Nothing
                Return ToExcelDateSerial(New DateTime(year, month, day))
            End Function

        End Class

    End Class

End Namespace
