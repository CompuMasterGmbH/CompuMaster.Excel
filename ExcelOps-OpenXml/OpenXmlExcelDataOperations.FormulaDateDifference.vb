Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateDateDifferenceFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                            formulaCell As ExcelCell) As Object
                Dim isDateDif = name = "DATEDIF"
                If arguments.Count < If(isDateDif, 3, 2) OrElse arguments.Count > 3 Then Return New FormulaError("#VALUE!")
                Dim start = ReadDateDifferenceInput(Scalar(EvaluateExpression(arguments(0), formulaCell)))
                If TypeOf start Is FormulaError Then Return start
                Dim finish = ReadDateDifferenceInput(Scalar(EvaluateExpression(arguments(1), formulaCell)))
                If TypeOf finish Is FormulaError Then Return finish
                If Not IsCalendarSerialInRange(DirectCast(start, Double)) OrElse Not IsCalendarSerialInRange(DirectCast(finish, Double)) Then
                    Return New FormulaError("#NUM!")
                End If
                Dim first = CInt(Math.Floor(DirectCast(start, Double)))
                Dim last = CInt(Math.Floor(DirectCast(finish, Double)))

                Dim optionValue = If(arguments.Count = 3, Scalar(EvaluateExpression(arguments(2), formulaCell)), Nothing)
                If TypeOf optionValue Is FormulaError Then Return optionValue
                If isDateDif Then
                    If last < first Then Return New FormulaError("#NUM!")
                    Return DateDifference(first, last, TextValue(optionValue).ToUpperInvariant())
                End If

                Dim european As Boolean
                Dim text = TryCast(optionValue, String)
                If text IsNot Nothing Then
                    If String.Equals(text, TextValue(True), StringComparison.OrdinalIgnoreCase) Then
                        european = True
                    ElseIf Not String.Equals(text, TextValue(False), StringComparison.OrdinalIgnoreCase) Then
                        Return New FormulaError("#VALUE!")
                    End If
                Else
                    european = DirectCast(NumberValue(optionValue), Double) <> 0.0
                End If
                Return Days360(first, last, european)
            End Function

            Private Function ReadDateDifferenceInput(value As Object) As Object
                Dim serial = ReadCalendarSerial(value)
                If TypeOf value Is String AndAlso TypeOf serial Is Double AndAlso DirectCast(serial, Double) < 0 Then
                    Dim numeric = NumberValue(value)
                    If TypeOf numeric Is FormulaError Then Return numeric
                End If
                Return serial
            End Function

            Private Function DateDifference(first As Integer, last As Integer, unit As String) As Object
                If unit = "D" Then Return CDbl(last - first)
                Dim start = DateParts(first)
                Dim finish = DateParts(last)
                Dim months = (finish.Year - start.Year) * 12 + finish.Month - start.Month
                If finish.Day < start.Day Then months -= 1
                Dim years = months \ 12
                Select Case unit
                    Case "Y" : Return CDbl(years)
                    Case "M" : Return CDbl(months)
                    Case "YM" : Return CDbl(months Mod 12)
                    Case "MD", "YD"
                        Dim remaining = finish.Day - start.Day
                        If remaining < 0 Then
                            Dim previous = New DateTime(finish.Year, finish.Month, 1).AddMonths(-1)
                            remaining += DaysInExcelMonth(previous.Year, previous.Month)
                        End If
                        If unit = "MD" Then Return CDbl(remaining)

                        ' Count the remaining whole months from the original date. Excel borrows the partial month from the end year.
                        Dim remainingMonth = New DateTime(start.Year, start.Month, 1).AddMonths(months Mod 12)
                        Return ToExcelDateSerial(remainingMonth) + start.Day - 1.0 - first + remaining
                    Case Else : Return New FormulaError("#NUM!")
                End Select
            End Function

            Private Function Days360(first As Integer, last As Integer, european As Boolean) As Double
                Dim start = DateParts(first)
                Dim finish = DateParts(last)
                Dim startDay = start.Day
                Dim endDay = finish.Day
                If european Then
                    startDay = Math.Min(startDay, 30)
                    endDay = Math.Min(endDay, 30)
                Else
                    If startDay = DaysInExcelMonth(start.Year, start.Month) Then startDay = 30
                    If endDay = 31 AndAlso startDay >= 30 Then endDay = 30
                End If
                Return (finish.Year - start.Year) * 360.0 + (finish.Month - start.Month) * 30.0 + endDay - startDay
            End Function

            Private Function DateParts(serial As Integer) As (Year As Integer, Month As Integer, Day As Integer)
                If Not Uses1904Dates AndAlso serial = 0 Then Return (1900, 1, 0)
                If Not Uses1904Dates AndAlso serial = 60 Then Return (1900, 2, 29)
                Dim dateValue = If(Uses1904Dates, Date1904Origin, Date1900Origin).AddDays(
                    If(Not Uses1904Dates AndAlso serial > 60, serial - 1, serial))
                Return (dateValue.Year, dateValue.Month, dateValue.Day)
            End Function

            Private Function DaysInExcelMonth(year As Integer, month As Integer) As Integer
                If Not Uses1904Dates AndAlso year = 1900 AndAlso month = 2 Then Return 29
                Return DateTime.DaysInMonth(year, month)
            End Function

        End Class

    End Class

End Namespace
