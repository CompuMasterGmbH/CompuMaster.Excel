Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateYearFraction(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim start = ReadYearFractionDate(Scalar(EvaluateExpression(arguments(0), formulaCell)))
                If TypeOf start Is FormulaError Then Return start
                Dim finish = ReadYearFractionDate(Scalar(EvaluateExpression(arguments(1), formulaCell)))
                If TypeOf finish Is FormulaError Then Return finish
                Dim input = If(arguments.Count = 3, Scalar(EvaluateExpression(arguments(2), formulaCell)), Nothing)
                If TypeOf input Is Boolean Then Return New FormulaError("#VALUE!")
                Dim basis = NumberValue(input)
                If TypeOf basis Is FormulaError Then Return basis
                Dim mode = Math.Floor(DirectCast(basis, Double))
                If mode < 0 OrElse mode > 4 Then Return New FormulaError("#NUM!")

                Dim first = Math.Min(DirectCast(start, Integer), DirectCast(finish, Integer))
                Dim last = Math.Max(DirectCast(start, Integer), DirectCast(finish, Integer))
                Dim days = last - first
                If mode = 2 Then Return days / 360.0
                If mode = 3 Then Return days / 365.0
                If Not IsCalendarSerialInRange(last) Then Return New FormulaError("#NUM!")
                If days = 0 Then Return 0.0
                If mode = 4 Then Return Days360(first, last, european:=True) / 360.0

                Dim firstDate = DateParts(first)
                Dim lastDate = DateParts(last)
                If mode = 1 Then Return days / ActualYearLength(firstDate, lastDate)
                Return YearFractionUsDays(firstDate, lastDate) / 360.0
            End Function

            Private Function ReadYearFractionDate(input As Object) As Object
                If TypeOf input Is Boolean Then Return New FormulaError("#VALUE!")
                Dim value = ReadDateDifferenceInput(input)
                If TypeOf value Is FormulaError Then Return value
                Dim serial = DirectCast(value, Double)
                If Not IsCalendarSerialInRange(serial) Then Return New FormulaError("#NUM!")

                ' Excel adds the 1904 offset again for numeric inputs and leaves text conversions on their original date base.
                Dim offset = If(Uses1904Dates AndAlso Not TypeOf input Is String, DateSystemsOffset, 0)
                Return CInt(Math.Floor(serial)) + offset
            End Function

            Private Shared Function ActualYearLength(first As (Year As Integer, Month As Integer, Day As Integer),
                                                     last As (Year As Integer, Month As Integer, Day As Integer)) As Double
                If first.Year = last.Year Then Return If(DateTime.IsLeapYear(first.Year), 366.0, 365.0)
                If last.Year = first.Year + 1 AndAlso
                    (last.Month < first.Month OrElse (last.Month = first.Month AndAlso last.Day <= first.Day)) Then
                    Dim leapDateInFirstYear = DateTime.IsLeapYear(first.Year) AndAlso first.Month <= 2
                    Dim leapDateInLastYear = DateTime.IsLeapYear(last.Year) AndAlso (last.Month > 2 OrElse (last.Month = 2 AndAlso last.Day = 29))
                    Return If(leapDateInFirstYear OrElse leapDateInLastYear, 366.0, 365.0)
                End If
                Dim duration = (New DateTime(last.Year, 12, 31) - New DateTime(first.Year, 1, 1)).Days + 1
                Return duration / CDbl(last.Year - first.Year + 1)
            End Function

            Private Shared Function YearFractionUsDays(first As (Year As Integer, Month As Integer, Day As Integer),
                                                        last As (Year As Integer, Month As Integer, Day As Integer)) As Integer
                Dim startDay = Math.Min(first.Day, 30)
                Dim endDay = If(startDay = 30, Math.Min(last.Day, 30), last.Day)
                ' YEARFRAC applies the February adjustment after the 31st-day rule and uses the Gregorian leap calendar.
                If first.Month = 2 AndAlso first.Day = DateTime.DaysInMonth(first.Year, 2) Then
                    startDay = 30
                    If last.Month = 2 AndAlso last.Day = DateTime.DaysInMonth(last.Year, 2) Then endDay = 30
                End If
                Return (last.Year - first.Year) * 360 + (last.Month - first.Month) * 30 + endDay - startDay
            End Function

        End Class

    End Class

End Namespace
