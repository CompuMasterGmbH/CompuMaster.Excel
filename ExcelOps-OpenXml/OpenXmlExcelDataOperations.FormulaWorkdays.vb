Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const StandardWeekend As Integer = (1 << 5) Or (1 << 6)
            Private Const FullWeekend As Integer = (1 << 7) - 1

            Private Function EvaluateWorkdayFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                      formulaCell As ExcelCell) As Object
                Dim international = name.EndsWith(".INTL", StringComparison.Ordinal)
                Dim holidayIndex = If(international, 3, 2)
                If arguments.Count < 2 OrElse arguments.Count > holidayIndex + 1 Then Return New FormulaError("#VALUE!")

                Dim countDays = name.StartsWith("NETWORKDAYS", StringComparison.Ordinal)
                Dim start = Scalar(EvaluateExpression(arguments(0), formulaCell))
                Dim first = ReadWorkdayDate(start, shiftText:=True)
                If TypeOf first Is FormulaError Then Return first

                Dim lastOrDays = Scalar(EvaluateExpression(arguments(1), formulaCell))
                Dim second As Object
                If countDays Then
                    second = ReadWorkdayDate(lastOrDays, shiftText:=True)
                ElseIf TypeOf lastOrDays Is Boolean Then
                    second = New FormulaError("#VALUE!")
                Else
                    second = NumberValue(lastOrDays)
                End If
                If TypeOf second Is FormulaError Then Return second

                Dim startDay = CInt(DirectCast(first, Double))
                Dim endOrDays = Math.Floor(DirectCast(second, Double))
                Dim maximum = CInt(ToExcelDateSerial(DateTime.MaxValue.Date))
                If Not countDays Then
                    ' In a 1904 workbook Excel bounds a numeric WORKDAY start after adding the date-system offset.
                    If Uses1904Dates AndAlso (startDay > maximum - DateSystemsOffset OrElse
                        startDay + endOrDays > maximum - DateSystemsOffset) Then Return New FormulaError("#NUM!")
                    If endOrDays = 0 Then Return CDbl(startDay)
                    If startDay < 0 OrElse Math.Abs(endOrDays) > maximum Then Return New FormulaError("#NUM!")
                End If

                Dim weekend As Object = StandardWeekend
                If international AndAlso arguments.Count > 2 Then
                    Dim input = EvaluateExpression(arguments(2), formulaCell)
                    If input IsNot Nothing Then weekend = ReadWeekendMask(Scalar(input), countDays)
                End If
                If TypeOf weekend Is FormulaError Then Return weekend
                Dim mask = DirectCast(weekend, Integer)
                If Not countDays AndAlso mask = FullWeekend Then Return New FormulaError("#VALUE!")

                Dim suppliedHolidays = If(arguments.Count > holidayIndex, EvaluateExpression(arguments(holidayIndex), formulaCell), Nothing)
                Dim holidays = ReadWorkdayHolidays(suppliedHolidays, mask, countDays)
                If TypeOf holidays Is FormulaError Then Return holidays
                Dim excluded = DirectCast(holidays, Integer())
                Dim daysPerWeek = 7 - Enumerable.Range(0, 7).Count(Function(day) (mask And (1 << day)) <> 0)
                If countDays Then
                    Dim endDay = CInt(endOrDays)
                    Dim count = CountWorkdays(Math.Min(startDay, endDay), Math.Max(startDay, endDay), mask, daysPerWeek, excluded)
                    Return CDbl(If(endDay < startDay, -count, count))
                End If
                Return FindWorkday(startDay, CInt(endOrDays), maximum, mask, daysPerWeek, excluded)
            End Function

            Private Function ReadWorkdayDate(value As Object, shiftText As Boolean) As Object
                If TypeOf value Is Boolean Then Return New FormulaError("#VALUE!")
                Dim input = ReadCalendarSerial(value)
                If TypeOf input Is FormulaError Then Return input
                Dim serial = DirectCast(input, Double)
                If Not IsCalendarSerialInRange(serial) Then Return New FormulaError("#NUM!")
                Return Math.Floor(serial) - If(shiftText AndAlso Uses1904Dates AndAlso TypeOf value Is String, DateSystemsOffset, 0)
            End Function

            Private Shared Function ReadWeekendMask(value As Object, countDays As Boolean) As Object
                Dim text = TryCast(value, String)
                If text IsNot Nothing Then
                    If text.Length <> 7 OrElse text.Any(Function(day) day <> "0"c AndAlso day <> "1"c) Then
                        Return New FormulaError("#VALUE!")
                    End If
                    Dim mask As Integer
                    For day As Integer = 0 To 6
                        If text(day) = "1"c Then mask = mask Or (1 << day)
                    Next
                    Return mask
                End If

                If TypeOf value Is FormulaError Then Return New FormulaError(If(countDays, "#VALUE!", "#NUM!"))
                Dim code = Math.Floor(If(value Is Nothing, 0.0, Convert.ToDouble(value, Globalization.CultureInfo.InvariantCulture)))
                Select Case code
                    Case 1 To 7
                        Return (1 << ((CInt(code) + 4) Mod 7)) Or (1 << ((CInt(code) + 5) Mod 7))
                    Case 11 To 17
                        Return 1 << ((CInt(code) - 5) Mod 7)
                    Case Else
                        Return New FormulaError("#NUM!")
                End Select
            End Function

            Private Function ReadWorkdayHolidays(value As Object, weekend As Integer, countDays As Boolean) As Object
                Dim area = TryCast(value, FormulaArea)
                Dim matrix = TryCast(value, FormulaMatrix)
                Dim values As IEnumerable(Of Object)
                If area IsNot Nothing Then
                    If Not countDays AndAlso Not IsMultiValue(area) Then
                        values = {If(Scalar(area), 0.0)}
                    Else
                        values = RangeValues(area)
                    End If
                ElseIf matrix IsNot Nothing Then
                    values = matrix.Values
                Else
                    values = {value}
                End If

                ' Excel converts literal text in NETWORKDAYS and array text in WORKDAY through the legacy date base.
                Dim shiftText = If(countDays, area Is Nothing AndAlso matrix Is Nothing, matrix IsNot Nothing)
                Dim holidays As New HashSet(Of Integer)
                For Each entry In values
                    If entry Is Nothing Then Continue For
                    Dim dateValue = ReadWorkdayDate(entry, shiftText)
                    If TypeOf dateValue Is FormulaError Then Return dateValue
                    Dim day = CInt(DirectCast(dateValue, Double))
                    If Not IsWeekend(day, weekend) Then holidays.Add(day)
                Next
                Dim result = holidays.ToArray()
                Array.Sort(result)
                Return result
            End Function

            Private Function IsWeekend(day As Integer, weekend As Integer) As Boolean
                Return (weekend And (1 << WeekdayOffset(day, 1))) <> 0
            End Function

            Private Function CountWorkdays(first As Integer, last As Integer, weekend As Integer, daysPerWeek As Integer,
                                           holidays As Integer()) As Integer
                If last < first Then Return 0
                Dim length = last - first + 1
                Dim result = (length \ 7) * daysPerWeek
                For day As Integer = last - (length Mod 7) + 1 To last
                    If Not IsWeekend(day, weekend) Then result += 1
                Next
                Return result - (HolidayIndex(holidays, last + 1) - HolidayIndex(holidays, first))
            End Function

            Private Shared Function HolidayIndex(holidays As Integer(), day As Integer) As Integer
                Dim index = Array.BinarySearch(holidays, day)
                Return If(index < 0, Not index, index)
            End Function

            Private Function FindWorkday(start As Integer, days As Integer, maximum As Integer, weekend As Integer,
                                         daysPerWeek As Integer, holidays As Integer()) As Object
                Dim forward = days > 0
                Dim target = Math.Abs(days)
                Dim low = 1
                Dim high = If(forward, maximum - start, start)
                Dim count As Func(Of Integer, Integer) =
                    If(forward,
                       Function(distance As Integer) CountWorkdays(start + 1, start + distance, weekend, daysPerWeek, holidays),
                       Function(distance As Integer) CountWorkdays(start - distance, start - 1, weekend, daysPerWeek, holidays))
                If high < low OrElse count(high) < target Then Return New FormulaError("#NUM!")
                While low < high
                    Dim middle = low + (high - low) \ 2
                    If count(middle) < target Then
                        low = middle + 1
                    Else
                        high = middle
                    End If
                End While
                Return CDbl(If(forward, start + low, start - low))
            End Function

        End Class

    End Class

End Namespace
