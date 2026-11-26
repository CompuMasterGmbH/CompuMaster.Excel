Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const DateSystemsOffset As Integer = 1462

            Private Function EvaluateCalendarFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                      formulaCell As ExcelCell) As Object
                Dim optionIsOptional = name = "WEEKDAY" OrElse name = "WEEKNUM"
                Dim required = If(optionIsOptional OrElse name = "ISOWEEKNUM", 1, 2)
                Dim maximum = If(optionIsOptional, 2, required)
                If arguments.Count < required OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")

                Dim input = Scalar(EvaluateExpression(arguments(0), formulaCell))
                Dim serial = ReadCalendarSerial(input)
                If TypeOf serial Is FormulaError Then Return serial
                Dim first = DirectCast(serial, Double)

                Dim supplied = If(arguments.Count = 2, EvaluateExpression(arguments(1), formulaCell), Nothing)
                Dim second = Scalar(supplied)
                Select Case name
                    Case "DAYS"
                        Dim starting = ReadCalendarSerial(second)
                        If TypeOf starting Is FormulaError Then Return starting
                        Dim last = DirectCast(starting, Double)
                        If Not IsCalendarSerialInRange(first) OrElse Not IsCalendarSerialInRange(last) Then Return New FormulaError("#NUM!")
                        Return Math.Floor(first) - Math.Floor(last)
                    Case "EDATE", "EOMONTH"
                        If first < 0 Then Return New FormulaError("#NUM!")
                        If supplied Is Nothing Then Return New FormulaError("#N/A")
                        Dim amount = NumberValue(second)
                        If TypeOf amount Is FormulaError Then Return amount
                        ' Excel interprets text start dates in the 1900 system for month arithmetic.
                        Dim legacy = If(Uses1904Dates AndAlso Not TypeOf input Is String, first + DateSystemsOffset, first)
                        Return ShiftCalendarMonth(legacy, Math.Truncate(DirectCast(amount, Double)), name = "EOMONTH")
                    Case Else
                        If Not IsCalendarSerialInRange(first) Then Return New FormulaError("#NUM!")
                        Dim whole = CInt(Math.Floor(first))
                        If name = "ISOWEEKNUM" Then Return IsoWeekNumber(whole)
                        Dim optionValue = If(arguments.Count = 1 OrElse (name = "WEEKNUM" AndAlso supplied Is Nothing), 1.0,
                                             NumberValue(second))
                        If TypeOf optionValue Is FormulaError Then Return optionValue
                        Return CalendarWeekResult(name, whole, Math.Truncate(DirectCast(optionValue, Double)))
                End Select
            End Function

            Private Function ReadCalendarSerial(value As Object) As Object
                Dim numeric = NumberValue(value)
                If Not TypeOf value Is String OrElse Not TypeOf numeric Is FormulaError Then Return numeric
                Dim text = DirectCast(value, String).Trim()
                If Not Uses1904Dates AndAlso (text = "1900-02-29" OrElse
                    (ConversionCulture.TwoLetterISOLanguageName = "de" AndAlso text = "29.02.1900")) Then Return 60.0
                Dim dateValue As DateTime
                If Not DateTime.TryParse(text, ConversionCulture, DateTimeStyles.AllowWhiteSpaces, dateValue) Then Return numeric
                Return ToExcelDateSerial(dateValue.Date)
            End Function

            Private Function IsCalendarSerialInRange(serial As Double) As Boolean
                Return serial >= 0 AndAlso serial < ToExcelDateSerial(DateTime.MaxValue.Date) + 1.0
            End Function

            Private Function ShiftCalendarMonth(legacySerial As Double, months As Double, endOfMonth As Boolean) As Object
                If legacySerial < 0 OrElse legacySerial >= To1900DateSerial(DateTime.MaxValue.Date) + 1.0 OrElse Math.Abs(months) > 120000 Then
                    Return New FormulaError("#NUM!")
                End If
                Dim whole = CInt(Math.Floor(legacySerial))
                Dim start = Date1900Origin.AddDays(If(whole >= 60, whole - 1, whole))
                Dim firstOfMonth As DateTime
                Dim day As Integer
                If whole = 0 Then
                    firstOfMonth = New DateTime(1900, 1, 1)
                Else
                    firstOfMonth = New DateTime(start.Year, start.Month, 1)
                    day = If(whole = 60, 29, start.Day)
                End If

                Try
                    Dim target = firstOfMonth.AddMonths(CInt(months))
                    If target.Year < 1900 Then Return New FormulaError("#NUM!")
                    Dim lastDay = DateTime.DaysInMonth(target.Year, target.Month)
                    Dim result = target.AddDays(If(endOfMonth, lastDay, Math.Min(day, lastDay)) - 1)
                    Return To1900DateSerial(result) - If(Uses1904Dates, DateSystemsOffset, 0)
                Catch ex As ArgumentOutOfRangeException
                    Return New FormulaError("#NUM!")
                End Try
            End Function

            Private Function CalendarWeekResult(name As String, serial As Integer, returnType As Double) As Object
                If name = "WEEKNUM" AndAlso returnType = 21 Then Return IsoWeekNumber(serial)

                Dim firstDay As Integer
                Select Case returnType
                    Case 1 : firstDay = 0
                    Case 2 : firstDay = 1
                    Case 3
                        If name <> "WEEKDAY" Then Return New FormulaError("#NUM!")
                        Return CDbl(WeekdayOffset(serial, 1))
                    Case 11 To 17 : firstDay = CInt(returnType) - 10
                    Case Else : Return New FormulaError("#NUM!")
                End Select

                If name = "WEEKDAY" Then Return CDbl(WeekdayOffset(serial, firstDay) + 1)
                Dim year = CInt(DirectCast(DateComponent("YEAR", serial), Double))
                Dim first = CInt(ToExcelDateSerial(New DateTime(year, 1, 1)))
                Return 1.0 + Math.Floor((serial - first + WeekdayOffset(first, firstDay)) / 7.0)
            End Function

            Private Function IsoWeekNumber(serial As Integer) As Double
                If Not Uses1904Dates AndAlso serial < 2 Then Return 52.0
                Dim year = CInt(DirectCast(DateComponent("YEAR", serial), Double))
                Dim first = FirstIsoWeek(year)
                If serial < first Then
                    first = FirstIsoWeek(year - 1)
                ElseIf year < DateTime.MaxValue.Year AndAlso serial >= FirstIsoWeek(year + 1) Then
                    Return 1.0
                End If
                Return 1.0 + Math.Floor((serial - first) / 7.0)
            End Function

            Private Function FirstIsoWeek(year As Integer) As Integer
                Dim fourth = CInt(ToExcelDateSerial(New DateTime(year, 1, 4)))
                Return fourth - WeekdayOffset(fourth, 1)
            End Function

            Private Function WeekdayOffset(serial As Integer, firstDay As Integer) As Integer
                Dim zeroBased = (serial + If(Uses1904Dates, 5, 6) - firstDay) Mod 7
                Return (zeroBased + 7) Mod 7
            End Function

        End Class

    End Class

End Namespace
