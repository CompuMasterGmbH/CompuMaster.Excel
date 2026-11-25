Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation
            Private Const SecondsPerDay As Integer = 86400
            Private Shared ReadOnly Date1900Origin As New DateTime(1899, 12, 31)
            Private Shared ReadOnly Date1904Origin As New DateTime(1904, 1, 1)
            Private Shared ReadOnly FirstRealDateAfterExcelLeapDay As New DateTime(1900, 3, 1)

            Private ReadOnly Property Uses1904Dates As Boolean
                Get
                    Return If(_owner.WorkbookPart.Workbook.WorkbookProperties?.Date1904?.Value, False)
                End Get
            End Property

            Private Function ToExcelDateSerial(value As DateTime) As Double
                If Uses1904Dates Then Return (value - Date1904Origin).TotalDays
                Dim serial = (value - Date1900Origin).TotalDays
                If value >= FirstRealDateAfterExcelLeapDay Then serial += 1
                Return serial
            End Function

            Private Function EvaluateDateFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                Select Case name
                    Case "NOW", "TODAY"
                        If arguments.Count <> 0 Then Return New FormulaError("#VALUE!")
                        MarkVolatile()
                        Dim localTime = _run.LocalTime.Value
                        Return ToExcelDateSerial(If(name = "TODAY", localTime.Date, localTime))
                    Case "DATE", "TIME"
                        If arguments.Count <> 3 Then Return New FormulaError("#VALUE!")
                        Dim values As Double() = Nothing
                        Dim failure = ReadNumbers(arguments, formulaCell, values)
                        If failure IsNot Nothing Then Return failure
                        If name = "DATE" Then Return ComposeDate(values)
                        If values.Any(Function(value) Math.Abs(Math.Truncate(value)) > 32767.0) Then Return New FormulaError("#NUM!")
                        Dim seconds = Math.Truncate(values(0)) * 3600.0 + Math.Truncate(values(1)) * 60.0 + Math.Truncate(values(2))
                        If seconds < 0 Then Return New FormulaError("#NUM!")
                        Return (seconds Mod SecondsPerDay) / SecondsPerDay
                    Case Else
                        If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                        Dim input = NumberValue(Scalar(EvaluateExpression(arguments(0), formulaCell)))
                        If TypeOf input Is FormulaError Then Return input
                        Dim serial = DirectCast(input, Double)
                        If serial < 0 Then Return New FormulaError("#NUM!")
                        Return DateComponent(name, serial)
                End Select
            End Function

            Private Function ComposeDate(values As Double()) As Object
                Dim year = Math.Truncate(values(0))
                Dim month = Math.Truncate(values(1))
                Dim day = Math.Truncate(values(2))
                If year < 0 OrElse year >= 10000 OrElse Math.Abs(month) > 120000 OrElse Math.Abs(day) > 4000000 Then
                    Return New FormulaError("#NUM!")
                End If
                If year < 1900 Then year += 1900
                Try
                    Dim first = New DateTime(CInt(year), 1, 1).AddMonths(CInt(month) - 1)
                    Dim serial = ToExcelDateSerial(first) + day - 1.0
                    If serial < 0 OrElse serial > ToExcelDateSerial(DateTime.MaxValue.Date) Then Return New FormulaError("#NUM!")
                    Return serial
                Catch ex As ArgumentOutOfRangeException
                    Return New FormulaError("#NUM!")
                End Try
            End Function

            Private Function DateComponent(name As String, serial As Double) As Object
                Select Case name
                    Case "HOUR", "MINUTE", "SECOND"
                        Dim seconds = CInt(Math.Round((serial - Math.Floor(serial)) * SecondsPerDay, MidpointRounding.AwayFromZero)) Mod SecondsPerDay
                        If name = "HOUR" Then Return CDbl(seconds \ 3600)
                        If name = "MINUTE" Then Return CDbl((seconds \ 60) Mod 60)
                        Return CDbl(seconds Mod 60)
                End Select

                Dim wholeDays = Math.Floor(serial)
                If wholeDays > ToExcelDateSerial(DateTime.MaxValue.Date) Then Return New FormulaError("#NUM!")
                If Not Uses1904Dates AndAlso wholeDays = 0 Then
                    If name = "YEAR" Then Return 1900.0
                    Return If(name = "MONTH", 1.0, 0.0)
                End If
                If Not Uses1904Dates AndAlso wholeDays = 60 Then
                    If name = "YEAR" Then Return 1900.0
                    Return If(name = "MONTH", 2.0, 29.0)
                End If

                Dim origin = If(Uses1904Dates, Date1904Origin, Date1900Origin)
                If Not Uses1904Dates AndAlso wholeDays >= 61 Then wholeDays -= 1
                Dim value = origin.AddDays(wholeDays)
                If name = "YEAR" Then Return CDbl(value.Year)
                Return If(name = "MONTH", CDbl(value.Month), CDbl(value.Day))
            End Function
        End Class

    End Class

End Namespace
