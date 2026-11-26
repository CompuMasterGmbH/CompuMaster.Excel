Imports System.Globalization
Imports System.Text

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly ThaiDigits As String() = {
                "ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"
            }
            Private Shared ReadOnly ThaiPlaces As String() = {"", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"}

            Private Function EvaluateBahtText(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) BahtTextValue(values(0)))
            End Function

            Private Function BahtTextValue(value As Object) As Object
                Dim numeric = NumberValue(value)
                If TypeOf numeric Is FormulaError Then Return numeric
                Dim number = DirectCast(numeric, Double)
                Dim rounded = RoundAt(Math.Abs(number), 2.0, "ROUND")
                If TypeOf rounded Is FormulaError Then Return rounded
                Dim formatted = DirectCast(rounded, Double).ToString("F2", CultureInfo.InvariantCulture)
                Dim decimalPoint = formatted.LastIndexOf("."c)
                Dim baht = formatted.Substring(0, decimalPoint)
                Dim satang = formatted.Substring(decimalPoint + 1)
                Dim output As New StringBuilder()
                If number < 0.0 Then output.Append("ลบ")
                If baht <> "0" OrElse satang = "00" Then
                    output.Append(If(baht = "0", ThaiDigits(0), ReadThaiDigits(baht))).Append("บาท")
                End If
                If satang = "00" Then
                    output.Append("ถ้วน")
                Else
                    output.Append(ReadThaiDigits(satang)).Append("สตางค์")
                End If
                Return output.ToString()
            End Function

            Private Shared Function ReadThaiDigits(digits As String) As String
                Dim output As New StringBuilder()
                Dim hasPreviousDigit = False
                For index = 0 To digits.Length - 1
                    Dim digit = AscW(digits(index)) - AscW("0"c)
                    Dim power = (digits.Length - index - 1) Mod 6
                    If digit <> 0 Then
                        If power = 1 Then
                            If digit = 2 Then output.Append("ยี่")
                            If digit > 2 Then output.Append(ThaiDigits(digit))
                        Else
                            output.Append(If(power = 0 AndAlso digit = 1 AndAlso hasPreviousDigit, "เอ็ด", ThaiDigits(digit)))
                        End If
                        output.Append(ThaiPlaces(power))
                        hasPreviousDigit = True
                    End If
                    If power = 0 AndAlso index <> digits.Length - 1 Then output.Append("ล้าน")
                Next
                Return output.ToString()
            End Function

        End Class

    End Class

End Namespace
