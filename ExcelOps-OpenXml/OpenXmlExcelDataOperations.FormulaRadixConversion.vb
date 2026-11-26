Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const RadixDigits As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Private Const BaseExclusiveLimit As Double = 9007199254740992.0

            Private Function EvaluateRadixConversion(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                       formulaCell As ExcelCell) As Object
                Dim general = name = "BASE" OrElse name = "DECIMAL"
                Dim decimalOutput = Not general AndAlso name.EndsWith("DEC", StringComparison.Ordinal)
                Dim minimum = If(general, 2, 1)
                Dim maximum = If(general, If(name = "BASE", 3, 2), If(decimalOutput, 1, 2))
                If arguments.Count < minimum OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")

                Dim first = AsMatrix(EvaluateExpression(arguments(0), formulaCell))
                Dim secondInput = If(arguments.Count > 1, EvaluateExpression(arguments(1), formulaCell), Nothing)
                Dim second = AsMatrix(secondInput)
                Dim third = AsMatrix(If(arguments.Count > 2, EvaluateExpression(arguments(2), formulaCell), Nothing))
                Dim rows = Math.Max(first.Rows, Math.Max(second.Rows, third.Rows))
                Dim columns = Math.Max(first.Columns, Math.Max(second.Columns, third.Columns))
                Dim source = If(general, 0, LegacyRadix(name(0)))
                Dim destination = If(general, 0, LegacyRadix(name(4)))
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim value = first.BroadcastAt(row, column)
                        Dim parameter = second.BroadcastAt(row, column)
                        Dim result As Object
                        If general Then
                            result = GeneralRadixValue(name, value, parameter, third.BroadcastAt(row, column))
                        Else
                            result = LegacyRadixValue(source, destination, value, parameter, secondInput IsNot Nothing)
                        End If
                        results(row * columns + column) = result
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function GeneralRadixValue(name As String, input As Object, radixInput As Object, lengthInput As Object) As Object
                If TypeOf input Is FormulaError Then Return input
                Dim encode = name = "BASE"
                Dim number As Object = Nothing
                If encode Then
                    number = NumberValue(input)
                    If TypeOf number Is FormulaError Then Return number
                End If
                Dim rawRadix = NumberValue(radixInput)
                If TypeOf rawRadix Is FormulaError Then Return rawRadix
                Dim rawLength = NumberValue(lengthInput)
                If TypeOf rawLength Is FormulaError Then Return rawLength

                Dim radix = Math.Truncate(DirectCast(rawRadix, Double))
                If radix < 2.0 OrElse radix > 36.0 Then Return New FormulaError("#NUM!")
                If Not encode Then Return DecodeGeneralRadix(TextValue(input), CInt(radix))

                Dim value = DirectCast(number, Double)
                Dim exact As Decimal
                If TypeOf input Is String AndAlso value >= 1.0E15 AndAlso value < 1.0E16 AndAlso
                    Decimal.TryParse(DirectCast(input, String), NumberStyles.Float, ConversionCulture, exact) Then
                    ' The worksheet text-to-number conversion retains fifteen significant digits.
                    value = CDbl(Decimal.Floor(exact / 10D) * 10D)
                End If
                Dim length = DirectCast(rawLength, Double)
                If value < 0.0 OrElse value >= BaseExclusiveLimit OrElse length < 0.0 OrElse Math.Truncate(length) > 255.0 Then
                    Return New FormulaError("#NUM!")
                End If
                Return EncodeRadix(CLng(Math.Truncate(value)), CInt(radix)).PadLeft(CInt(Math.Truncate(length)), "0"c)
            End Function

            Private Shared Function DecodeGeneralRadix(input As String, radix As Integer) As Object
                If input.Length > 255 Then Return New FormulaError("#VALUE!")
                Dim digits = input.TrimStart(" "c, ChrW(9))
                Dim hasPrefix = radix = 16 AndAlso digits.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
                If hasPrefix Then digits = digits.Substring(2)
                Dim hasSuffix = (radix = 16 AndAlso digits.EndsWith("h", StringComparison.OrdinalIgnoreCase)) OrElse
                    (radix = 2 AndAlso digits.EndsWith("b", StringComparison.OrdinalIgnoreCase))
                If hasSuffix Then digits = digits.Substring(0, digits.Length - 1)
                Dim value = DecodeRadixDigits(digits, radix)
                If value < 0.0 OrElse (value = 0.0 AndAlso hasSuffix AndAlso Not hasPrefix) Then Return New FormulaError("#NUM!")
                Return FiniteNumber(value)
            End Function

            Private Function LegacyRadixValue(source As Integer, destination As Integer, input As Object,
                                             placesInput As Object, hasPlaces As Boolean) As Object
                Dim width = 0
                If hasPlaces Then
                    Dim number = LegacyNumericArgument(placesInput)
                    If TypeOf number Is FormulaError Then Return number
                    Dim places = Math.Truncate(DirectCast(number, Double))
                    If places < 1.0 OrElse places > 10.0 Then Return New FormulaError("#NUM!")
                    width = CInt(places)
                End If

                Dim value As Double
                If TypeOf input Is FormulaError Then Return input
                If source = 10 Then
                    Dim number = LegacyNumericArgument(input)
                    If TypeOf number Is FormulaError Then Return number
                    value = Math.Truncate(DirectCast(number, Double))
                Else
                    If TypeOf input Is Boolean Then Return New FormulaError("#VALUE!")
                    Dim digits = TextValue(input)
                    If digits.Length > 10 Then Return New FormulaError("#NUM!")
                    value = DecodeRadixDigits(digits, source)
                    If value < 0.0 Then Return New FormulaError("#NUM!")
                    Dim signBit = LegacySignBit(source)
                    If value >= signBit Then value -= 2L * signBit
                End If

                If destination = 10 Then Return value
                Dim limit = LegacySignBit(destination)
                If value < -limit OrElse value >= limit Then Return New FormulaError("#NUM!")
                If value < 0 Then Return EncodeRadix(CLng(value) + 2L * limit, destination)
                Dim result = EncodeRadix(CLng(value), destination)
                If hasPlaces AndAlso result.Length > width Then Return New FormulaError("#NUM!")
                Return result.PadLeft(width, "0"c)
            End Function

            Private Function LegacyNumericArgument(value As Object) As Object
                If TypeOf value Is Boolean Then Return New FormulaError("#VALUE!")
                Return NumberValue(value)
            End Function

            Private Shared Function LegacyRadix(specifier As Char) As Integer
                Select Case specifier
                    Case "B"c : Return 2
                    Case "O"c : Return 8
                    Case "H"c : Return 16
                    Case Else : Return 10
                End Select
            End Function

            Private Shared Function LegacySignBit(radix As Integer) As Long
                Return 1L << If(radix = 2, 9, If(radix = 8, 29, 39))
            End Function

            Private Shared Function DecodeRadixDigits(text As String, radix As Integer) As Double
                Dim value = 0.0
                For Each character In text
                    Dim digit = RadixDigits.IndexOf(character)
                    If character >= "a"c AndAlso character <= "z"c Then digit = AscW(character) - AscW("a"c) + 10
                    If digit < 0 OrElse digit >= radix Then Return -1.0
                    value = value * radix + digit
                Next
                Return value
            End Function

            Private Shared Function EncodeRadix(value As Long, radix As Integer) As String
                Dim buffer(63) As Char
                Dim offset = buffer.Length
                Do
                    offset -= 1
                    buffer(offset) = RadixDigits(CInt(value Mod radix))
                    value = value \ radix
                Loop While value <> 0
                Return New String(buffer, offset, buffer.Length - offset)
            End Function

        End Class

    End Class

End Namespace
