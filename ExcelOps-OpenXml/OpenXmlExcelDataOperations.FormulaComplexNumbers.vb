Imports System.Globalization
Imports System.Numerics

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Structure ComplexOperand
                Friend ReadOnly Value As Complex
                Friend ReadOnly Unit As Char

                Friend Sub New(value As Complex, unit As Char)
                    Me.Value = value
                    Me.Unit = unit
                End Sub
            End Structure

            Private Function EvaluateComplexFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                    formulaCell As ExcelCell) As Object
                If name = "IMSUM" OrElse name = "IMPRODUCT" Then Return AggregateComplexValues(name, arguments, formulaCell)
                Dim construct = name = "COMPLEX"
                Dim binary = name = "IMDIV" OrElse name = "IMSUB" OrElse name = "IMPOWER"
                Dim minimum = If(construct OrElse binary, 2, 1)
                If arguments.Count < minimum OrElse arguments.Count > If(construct, 3, minimum) Then Return New FormulaError("#VALUE!")

                Dim inputs = arguments.Select(Function(argument) EvaluateExpression(argument, formulaCell)).ToArray()
                If inputs.Any(Function(input) TypeOf input Is FormulaArea AndAlso IsMultiValue(input)) Then Return New FormulaError("#VALUE!")
                Dim matrices = inputs.Select(Function(input) AsMatrix(input)).ToArray()
                Dim rows = matrices.Max(Function(matrix) matrix.Rows)
                Dim columns = matrices.Max(Function(matrix) matrix.Columns)
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim first = matrices(0).BroadcastAt(row, column)
                        Dim second = If(matrices.Length > 1, matrices(1).BroadcastAt(row, column), Nothing)
                        Dim result As Object
                        If construct Then
                            Dim unit = If(matrices.Length > 2, matrices(2).BroadcastAt(row, column), Nothing)
                            result = ConstructComplexValue(first, second, unit)
                        Else
                            result = CalculateComplexValue(name, first, second)
                        End If
                        results(row * columns + column) = result
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function ConstructComplexValue(real As Object, imaginary As Object, rawUnit As Object) As Object
                If TypeOf real Is FormulaError Then Return real
                If TypeOf imaginary Is FormulaError Then Return imaginary
                If TypeOf rawUnit Is FormulaError Then Return rawUnit
                Dim realPart = ConvertWorksheetText("VALUE", real)
                If TypeOf realPart Is FormulaError Then Return realPart
                Dim imaginaryPart = ConvertWorksheetText("VALUE", imaginary)
                If TypeOf imaginaryPart Is FormulaError Then Return imaginaryPart
                Dim unit = TextValue(rawUnit)
                If unit.Length <> 0 AndAlso unit <> "i" AndAlso unit <> "j" Then Return New FormulaError("#VALUE!")
                Return FormatComplexValue(New Complex(DirectCast(realPart, Double), DirectCast(imaginaryPart, Double)),
                                          If(unit.Length = 0, "i"c, unit(0)))
            End Function

            Private Function CalculateComplexValue(name As String, first As Object, second As Object) As Object
                Dim firstValue = ReadComplexOperand(first)
                If TypeOf firstValue Is FormulaError Then Return firstValue
                Dim operand = DirectCast(firstValue, ComplexOperand)
                Dim value = operand.Value
                Select Case name
                    Case "IMABS" : Return FiniteNumber(Complex.Abs(value))
                    Case "IMAGINARY" : Return value.Imaginary
                    Case "IMREAL" : Return value.Real
                    Case "IMARGUMENT"
                        If value = Complex.Zero Then Return New FormulaError("#DIV/0!")
                        Return value.Phase
                    Case "IMCONJUGATE" : Return FormatComplexValue(Complex.Conjugate(value), operand.Unit)
                    Case "IMSUB", "IMDIV"
                    Case Else : Return CalculateComplexMath(name, operand, second)
                End Select

                Dim secondValue = ReadComplexOperand(second)
                If TypeOf secondValue Is FormulaError Then Return secondValue
                Dim other = DirectCast(secondValue, ComplexOperand)
                If Not ComplexUnitsMatch(operand.Unit, other.Unit) Then Return New FormulaError("#NUM!")
                Dim resultUnit = If(operand.Unit = ChrW(0), other.Unit, operand.Unit)
                If name = "IMSUB" Then Return FormatComplexValue(value - other.Value, resultUnit)
                If other.Value = Complex.Zero Then Return New FormulaError("#NUM!")
                Return FormatComplexValue(value / other.Value, resultUnit)
            End Function

            Private Function AggregateComplexValues(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                    formulaCell As ExcelCell) As Object
                If arguments.Count = 0 OrElse arguments.Count > 255 Then Return New FormulaError("#VALUE!")
                Dim sum = name = "IMSUM"
                Dim result = If(sum, Complex.Zero, Complex.One)
                Dim unit = ChrW(0)
                For Each argument In arguments
                    Dim input = EvaluateExpression(argument, formulaCell)
                    If input Is Nothing Then Return New FormulaError("#N/A")
                    For Each item In AsMatrix(input).Values
                        Dim parsed = ReadComplexOperand(item)
                        If TypeOf parsed Is FormulaError Then Return parsed
                        Dim operand = DirectCast(parsed, ComplexOperand)
                        If Not ComplexUnitsMatch(unit, operand.Unit) Then Return New FormulaError("#VALUE!")
                        If unit = ChrW(0) Then unit = operand.Unit
                        result = If(sum, result + operand.Value, result * operand.Value)
                    Next
                Next
                Return FormatComplexValue(result, unit)
            End Function

            Private Shared Function ComplexUnitsMatch(first As Char, second As Char) As Boolean
                Return first = ChrW(0) OrElse second = ChrW(0) OrElse first = second
            End Function

            Private Function ReadComplexOperand(input As Object) As Object
                If TypeOf input Is FormulaError Then Return input
                If TypeOf input Is Boolean Then Return New FormulaError("#VALUE!")
                If input Is Nothing Then Return New ComplexOperand(Complex.Zero, ChrW(0))
                If TypeOf input Is Double Then
                    Return New ComplexOperand(New Complex(NormalizeWorksheetUnderflow(DirectCast(input, Double)), 0.0), ChrW(0))
                End If

                Dim text = TextValue(input).Trim(" "c)
                If text.Length = 0 OrElse text.Any(AddressOf Char.IsControl) Then Return New FormulaError("#NUM!")
                Dim suffix = text(text.Length - 1)
                If suffix <> "i"c AndAlso suffix <> "j"c Then
                    Dim number = ReadComplexComponent(text)
                    If Not number.HasValue Then Return New FormulaError("#NUM!")
                    Return New ComplexOperand(New Complex(number.Value, 0.0), ChrW(0))
                End If

                Dim body = text.Substring(0, text.Length - 1)
                Dim split = 0
                For index = 1 To body.Length - 1
                    If body(index) <> "+"c AndAlso body(index) <> "-"c Then Continue For
                    If body(index - 1) = "E"c OrElse body(index - 1) = "e"c Then Continue For
                    split = index
                    Exit For
                Next
                Dim real As Double? = 0.0
                Dim sign = 1.0
                Dim coefficient = body
                If split > 0 Then
                    real = ReadComplexComponent(body.Substring(0, split))
                    sign = If(body(split) = "-"c, -1.0, 1.0)
                    coefficient = body.Substring(split + 1)
                End If
                Dim imaginary As Double?
                If coefficient.Length = 0 OrElse coefficient = "+" Then
                    imaginary = 1.0
                ElseIf coefficient = "-" Then
                    imaginary = -1.0
                Else
                    imaginary = ReadComplexComponent(coefficient)
                End If
                If Not real.HasValue OrElse Not imaginary.HasValue Then Return New FormulaError("#NUM!")
                Return New ComplexOperand(New Complex(real.Value, NormalizeWorksheetUnderflow(sign * imaginary.Value)), suffix)
            End Function

            Private Function ReadComplexComponent(input As String) As Double?
                Dim text = input.Trim(" "c)
                Dim sign = 1.0
                If text.StartsWith("+", StringComparison.Ordinal) OrElse text.StartsWith("-", StringComparison.Ordinal) Then
                    sign = If(text(0) = "-"c, -1.0, 1.0)
                    text = text.Substring(1).TrimStart(" "c)
                    If text.StartsWith("+", StringComparison.Ordinal) OrElse text.StartsWith("-", StringComparison.Ordinal) Then Return Nothing
                End If
                Dim number = ConvertWorksheetText("VALUE", text)
                If Not TypeOf number Is Double Then Return Nothing
                Dim magnitude = DirectCast(number, Double)
                If Math.Abs(magnitude) >= 1.0E308 Then Return Nothing
                If magnitude = 0.0 Then
                    Dim marker = text.IndexOfAny({"e"c, "E"c})
                    Dim exponent As Integer
                    If marker > 0 AndAlso Integer.TryParse(text.Substring(marker + 1), NumberStyles.Integer, CultureInfo.InvariantCulture, exponent) Then
                        If exponent < -323 OrElse exponent > 308 OrElse text.Take(marker).Any(Function(ch) ch >= "1"c AndAlso ch <= "9"c) Then Return Nothing
                    End If
                End If
                Return NormalizeWorksheetUnderflow(sign * magnitude)
            End Function

            Private Function FormatComplexValue(value As Complex, unit As Char) As Object
                If Double.IsNaN(value.Real) OrElse Double.IsNaN(value.Imaginary) OrElse
                    Double.IsInfinity(value.Real) OrElse Double.IsInfinity(value.Imaginary) Then Return New FormulaError("#NUM!")

                Dim real = NormalizeWorksheetUnderflow(value.Real)
                Dim imaginary = NormalizeWorksheetUnderflow(value.Imaginary)
                If imaginary = 0.0 Then Return FormatComplexNumber(real)
                Dim suffix = If(unit = ChrW(0), "i"c, unit)
                Dim coefficient = If(Math.Abs(imaginary) = 1.0, "", FormatComplexNumber(Math.Abs(imaginary)))
                Dim sign = If(imaginary < 0.0, "-", If(real = 0.0, "", "+"))
                Return If(real = 0.0, "", FormatComplexNumber(real)) & sign & coefficient & suffix
            End Function

            Private Function FormatComplexNumber(value As Double) As String
                Dim sign = If(value < 0.0, ConversionCulture.NumberFormat.NegativeSign, "")
                Dim formatted = Math.Abs(value).ToString("G15", ConversionCulture)
                Dim exponentStart = formatted.IndexOf("E"c)
                If exponentStart < 0 Then Return sign & formatted

                Dim exponent = Integer.Parse(formatted.Substring(exponentStart + 1), NumberStyles.Integer, CultureInfo.InvariantCulture)
                Dim decimalSeparator = ConversionCulture.NumberFormat.NumberDecimalSeparator
                Dim digits = formatted.Substring(0, exponentStart).Replace(decimalSeparator, "")
                Dim decimalPosition = 1 + exponent
                ' Excel switches to scientific notation when a component needs more than 21 characters.
                If decimalPosition <= 0 Then
                    If 1 + decimalSeparator.Length + digits.Length - decimalPosition <= 21 Then
                        Return sign & "0" & decimalSeparator & New String("0"c, -decimalPosition) & digits
                    End If
                ElseIf decimalPosition >= digits.Length Then
                    If decimalPosition <= 21 Then Return sign & digits.PadRight(decimalPosition, "0"c)
                ElseIf digits.Length + decimalSeparator.Length <= 21 Then
                    Return sign & digits.Insert(decimalPosition, decimalSeparator)
                End If
                Return sign & formatted
            End Function

        End Class

    End Class

End Namespace
