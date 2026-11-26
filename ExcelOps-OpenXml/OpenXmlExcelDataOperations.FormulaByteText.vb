Imports System.Globalization
Imports System.Text
Imports OpenXmlCompositeElement = DocumentFormat.OpenXml.OpenXmlCompositeElement
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly NarrowKatakana As New Lazy(Of Dictionary(Of String, String))(AddressOf ReadNarrowKatakana)

            Private ReadOnly Property UsesDoubleByteText As Boolean
                Get
                    Select Case ConversionCulture.TwoLetterISOLanguageName
                        Case "ja", "zh", "ko" : Return True
                        Case Else : Return False
                    End Select
                End Get
            End Property

            Private Function EvaluateWidthFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                   formulaCell As ExcelCell) As Object
                If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                If name = "PHONETIC" Then Return ReadPhoneticReference(EvaluateExpression(arguments(0), formulaCell))
                Return BroadcastReferenceArguments(arguments, formulaCell,
                    Function(values)
                        If name = "JIS" AndAlso ConversionCulture.TwoLetterISOLanguageName <> "ja" Then Return New FormulaError("#NAME?")
                        Dim input = TextValue(values(0))
                        If Not UsesDoubleByteText Then Return input
                        Return ConvertTextWidth(input, name <> "ASC")
                    End Function)
            End Function

            Private Function ReadPhoneticReference(value As Object) As Object
                If TypeOf value Is FormulaError Then Return value
                Dim area = TryCast(value, FormulaArea)
                If area Is Nothing Then Return New FormulaError("#N/A")
                Dim source = ResolveReference(area)
                Dim referencedValue = source.ReadCell(area.Sheet, area.FirstRow, area.FirstColumn)
                If TypeOf referencedValue Is FormulaError Then Return referencedValue
                Dim sheet As WorksheetSnapshot = Nothing
                If Not source._worksheets.TryGetValue(area.Sheet, sheet) Then Return New FormulaError("#REF!")
                Dim cell As Cell = Nothing
                If Not sheet.Cells.TryGetValue((area.FirstRow, area.FirstColumn), cell) Then Return New FormulaError("#N/A")
                Dim container As OpenXmlCompositeElement = cell.InlineString
                If cell.DataType?.Value = CellValues.SharedString Then
                    Dim index As Integer
                    If Integer.TryParse(cell.CellValue?.Text, NumberStyles.None, CultureInfo.InvariantCulture, index) Then
                        container = source._owner.WorkbookPart.SharedStringTablePart?.SharedStringTable?.Elements(Of SharedStringItem)().ElementAtOrDefault(index)
                    End If
                End If
                Dim reading = If(container Is Nothing, "", String.Concat(container.Elements(Of PhoneticRun)().Select(Function(run) DecodeText(run.InnerText))))
                If reading.Length > 0 Then Return reading
                Return New FormulaError("#N/A")
            End Function

            Private Shared Function ReadNarrowKatakana() As Dictionary(Of String, String)
                Dim forms As New Dictionary(Of String, String)(StringComparer.Ordinal)
                For code = &HFF61 To &HFF9F
                    Dim narrow = ChrW(code).ToString()
                    forms(narrow.Normalize(NormalizationForm.FormKC)) = narrow
                    For mark = &HFF9E To &HFF9F
                        Dim pair = narrow & ChrW(mark)
                        Dim wide = pair.Normalize(NormalizationForm.FormKC)
                        If wide.Length = 1 Then forms(wide) = pair
                    Next
                Next
                Return forms
            End Function

            Private Shared Function ConvertTextWidth(input As String, wide As Boolean) As String
                Dim output As New StringBuilder(input.Length)
                Dim index = 0
                While index < input.Length
                    Dim c = input(index)
                    Dim code = AscW(c) And &HFFFF
                    If wide Then
                        If c = " "c Then
                            output.Append(ChrW(&H3000))
                        ElseIf code >= &H21 AndAlso code <= &H7E Then
                            output.Append(ChrW(code + &HFEE0))
                        ElseIf code >= &HFF61 AndAlso code <= &HFF9F Then
                            Dim length = If(index + 1 < input.Length AndAlso (input(index + 1) = ChrW(&HFF9E) OrElse input(index + 1) = ChrW(&HFF9F)), 2, 1)
                            output.Append(input.Substring(index, length).Normalize(NormalizationForm.FormKC))
                            index += length - 1
                        Else
                            output.Append(c)
                        End If
                    ElseIf code = &H3000 Then
                        output.Append(" "c)
                    ElseIf code >= &HFF01 AndAlso code <= &HFF5E Then
                        output.Append(ChrW(code - &HFEE0))
                    Else
                        Dim narrow As String = Nothing
                        output.Append(If(NarrowKatakana.Value.TryGetValue(c.ToString(), narrow), narrow, c.ToString()))
                    End If
                    index += 1
                End While
                Return output.ToString()
            End Function

            Private Function EvaluateByteFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                  formulaCell As ExcelCell) As Object
                Dim minimum = If(name = "MIDB", 3, If(name = "REPLACEB", 4, 1))
                Dim maximum = If(name = "LEFTB" OrElse name = "RIGHTB", 2, minimum)
                If arguments.Count < minimum OrElse arguments.Count > maximum Then Return New FormulaError("#VALUE!")
                Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) ByteTextValue(name, values))
            End Function

            Private Function ByteTextValue(name As String, values As Object()) As Object
                Dim contents = TextValue(values(0))
                Dim doubleByte = UsesDoubleByteText
                Dim encoding = If(doubleByte, WorksheetCodePage(), Nothing)
                Dim count = If(doubleByte, encoding.GetByteCount(contents), contents.Length)
                If name = "LENB" Then Return CDbl(count)
                Dim start As Double = 0.0
                Dim length As Double = 1.0
                If name = "MIDB" OrElse name = "REPLACEB" Then
                    Dim position = NumberValue(values(1))
                    If TypeOf position Is FormulaError Then Return position
                    start = Math.Truncate(DirectCast(position, Double)) - 1.0
                    If start < 0.0 Then Return New FormulaError("#VALUE!")
                End If
                Dim replace = name = "REPLACEB"
                If values.Length >= 2 Then
                    Dim amount = NumberValue(values(If(replace, 2, values.Length - 1)))
                    If TypeOf amount Is FormulaError Then Return amount
                    length = Math.Truncate(DirectCast(amount, Double))
                    If length < 0.0 Then Return New FormulaError("#VALUE!")
                End If
                Dim unicodePositions = Not doubleByte AndAlso (name = "LEFTB" OrElse name = "RIGHTB")
                If unicodePositions Then count = TextCharacterCount(contents, True)
                If name = "RIGHTB" Then start = Math.Max(0.0, count - length)
                Dim before = Math.Min(start, count)
                Dim after = Math.Min(start + length, count)
                If doubleByte Then
                    If replace Then
                        Return ByteSlice(contents, encoding, 0, CInt(before)) & TextValue(values(3)) &
                            ByteSlice(contents, encoding, CInt(after), count)
                    End If
                    Return ByteSlice(contents, encoding, CInt(before), CInt(after))
                End If
                Dim first = Utf16CharacterOffset(contents, before, unicodePositions)
                Dim last = Utf16CharacterOffset(contents, after, unicodePositions)
                Dim result = If(replace, contents.Substring(0, first) & TextValue(values.Last()) & contents.Substring(last),
                                contents.Substring(first, last - first))
                Return CheckedWorksheetText(result)
            End Function

            Private Shared Function ByteSlice(text As String, encoding As Encoding, first As Integer, last As Integer) As String
                Dim output As New StringBuilder()
                Dim byteOffset = 0
                Dim offset = 0
                While offset < text.Length AndAlso byteOffset < last
                    Dim width = If(Char.IsSurrogatePair(text, offset), 2, 1)
                    Dim element = text.Substring(offset, width)
                    Dim size = encoding.GetByteCount(element)
                    Dim overlap = Math.Min(byteOffset + size, last) - Math.Max(byteOffset, first)
                    If overlap = size Then
                        output.Append(element)
                    ElseIf overlap > 0 Then
                        output.Append(" "c, overlap)
                    End If
                    offset += width
                    byteOffset += size
                End While
                Return output.ToString()
            End Function

        End Class

    End Class

End Namespace
