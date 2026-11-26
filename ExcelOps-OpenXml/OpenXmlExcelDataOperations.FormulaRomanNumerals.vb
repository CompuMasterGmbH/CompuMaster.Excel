Imports System.Text

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const RomanSymbols As String = "MDCLXVI"
            Private Shared ReadOnly RomanValues As Integer() = {1000, 500, 100, 50, 10, 5, 1}

            Private Function EvaluateRomanNumerals(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                  formulaCell As ExcelCell) As Object
                Dim roman = name = "ROMAN"
                If arguments.Count < 1 OrElse arguments.Count > If(roman, 2, 1) Then Return New FormulaError("#VALUE!")
                Dim inputs = AsMatrix(EvaluateExpression(arguments(0), formulaCell))
                Dim forms = AsMatrix(If(arguments.Count = 2, EvaluateExpression(arguments(1), formulaCell), Nothing))
                Dim rows = Math.Max(inputs.Rows, forms.Rows)
                Dim columns = Math.Max(inputs.Columns, forms.Columns)
                Dim values = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim input = inputs.BroadcastAt(row, column)
                        Dim result = input
                        If Not TypeOf input Is FormulaError Then
                            result = If(roman, ConvertToRomanNumeral(input, forms.BroadcastAt(row, column)),
                                        ConvertFromRomanNumeral(TextValue(input)))
                        End If
                        values(row * columns + column) = result
                    Next
                Next
                If values.Length = 1 Then Return values(0)
                Return New FormulaMatrix(rows, columns, values)
            End Function

            Private Function ConvertToRomanNumeral(input As Object, style As Object) As Object
                Dim number = NumberValue(input)
                If TypeOf number Is FormulaError Then Return number
                Dim form = If(TypeOf style Is Boolean, If(DirectCast(style, Boolean), 0.0, 4.0), NumberValue(style))
                If TypeOf form Is FormulaError Then Return form
                Dim value = Math.Truncate(DirectCast(number, Double))
                Dim format = Math.Truncate(DirectCast(form, Double))
                If value < 0.0 OrElse value > 3999.0 OrElse format < 0.0 OrElse format > 4.0 Then Return New FormulaError("#VALUE!")

                Dim remainder = CInt(value)
                Dim result As New StringBuilder()
                For index = 0 To RomanSymbols.Length - 1
                    Dim threshold = RomanValues(index)
                    result.Append(RomanSymbols(index), remainder \ threshold)
                    remainder = remainder Mod threshold
                    Dim firstSubtractor = index + If(index Mod 2 = 0, 2, 1)
                    Dim lastSubtractor = Math.Min(RomanSymbols.Length - 1, firstSubtractor + CInt(format))
                    For subtractor = lastSubtractor To firstSubtractor Step -1
                        Dim difference = threshold - RomanValues(subtractor)
                        If remainder < difference Then Continue For
                        result.Append(RomanSymbols(subtractor)).Append(RomanSymbols(index))
                        remainder -= difference
                        Exit For
                    Next
                Next
                Return result.ToString()
            End Function

            Private Shared Function ConvertFromRomanNumeral(input As String) As Object
                If input.Length > 255 Then Return New FormulaError("#VALUE!")
                Dim text = input.Trim(" "c, ChrW(9), ChrW(10), ChrW(11), ChrW(12), ChrW(13))
                If text.Length = 0 Then Return 0.0
                Dim negative = text(0) = "-"c
                Dim start = If(negative, 1, 0)
                If start = text.Length Then Return New FormulaError("#NUM!")

                Dim largest = 0
                Dim total = 0
                For index = text.Length - 1 To start Step -1
                    Dim symbol = text(index)
                    If symbol >= "a"c AndAlso symbol <= "z"c Then symbol = ChrW(AscW(symbol) - AscW("a"c) + AscW("A"c))
                    Dim digit = RomanSymbols.IndexOf(symbol)
                    If digit < 0 Then Return New FormulaError("#VALUE!")
                    Dim value = RomanValues(digit)
                    total += If(value < largest, -value, value)
                    largest = Math.Max(largest, value)
                Next
                Return CDbl(If(negative, -total, total))
            End Function

        End Class

    End Class

End Namespace
