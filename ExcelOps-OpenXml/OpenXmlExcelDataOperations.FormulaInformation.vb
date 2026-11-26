Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateInformationFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                         formulaCell As ExcelCell) As Object
                If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                Dim input = EvaluateExpression(arguments(0), formulaCell)
                Dim area = TryCast(input, FormulaArea)
                Dim matrix = TryCast(input, FormulaMatrix)
                If name = "TYPE" Then
                    If matrix IsNot Nothing OrElse IsMultiValue(area) Then Return 64.0
                    Return InformationValue(name, Scalar(input))
                End If

                If area IsNot Nothing AndAlso IsMultiValue(area) Then
                    Select Case name
                        Case "ISEVEN", "ISODD"
                            Return New FormulaError("#VALUE!")
                        Case "N", "T"
                            Dim first = ResolveReference(area).ReadCell(area.Sheet, area.FirstRow, area.FirstColumn)
                            Return InformationValue(name, first)
                        Case Else
                            matrix = AsMatrix(area)
                    End Select
                End If
                If matrix Is Nothing OrElse matrix.Values.Length = 1 Then Return InformationValue(name, Scalar(input))
                Dim values = matrix.Values.Select(Function(value) InformationValue(name, value)).ToArray()
                Return New FormulaMatrix(matrix.Rows, matrix.Columns, values)
            End Function

            Private Function InformationValue(name As String, value As Object) As Object
                Dim failure = TryCast(value, FormulaError)
                Select Case name
                    Case "ISERR" : Return failure IsNot Nothing AndAlso failure.Code <> "#N/A"
                    Case "ISNONTEXT" : Return Not TypeOf value Is String
                    Case "TYPE"
                        If TypeOf value Is FormulaLambda Then Return 128.0
                        If TypeOf value Is String Then Return 2.0
                        If TypeOf value Is Boolean Then Return 4.0
                        If failure IsNot Nothing Then Return 16.0
                        Return 1.0
                    Case "ERROR.TYPE" : Return InformationErrorType(failure)
                    Case "N"
                        If TypeOf value Is String Then Return 0.0
                        Return NumberValue(value)
                    Case "T"
                        If failure IsNot Nothing OrElse TypeOf value Is String Then Return value
                        Return String.Empty
                    Case Else
                        Dim number = InformationParityNumber(value)
                        If TypeOf number Is FormulaError Then Return number
                        Dim even = Math.Truncate(DirectCast(number, Double)) Mod 2.0 = 0.0
                        Return even = (name = "ISEVEN")
                End Select
            End Function

            Private Function InformationParityNumber(value As Object) As Object
                If TypeOf value Is Boolean Then Return New FormulaError("#VALUE!")
                Dim number = NumberValue(value)
                If Not TypeOf value Is String OrElse Not TypeOf number Is FormulaError Then Return number

                Dim text = DirectCast(value, String).Trim()
                Dim serial = ReadDateTimeText(text)
                If serial.HasValue Then Return serial.Value
                Dim divisor = 1.0
                If text.EndsWith("%", StringComparison.Ordinal) Then
                    divisor = 100.0
                    text = text.Substring(0, text.Length - 1)
                End If
                Dim parsed As Double
                Dim style = NumberStyles.Float Or NumberStyles.AllowThousands Or NumberStyles.AllowCurrencySymbol
                If Double.TryParse(text, style, ConversionCulture, parsed) Then Return FiniteNumber(parsed / divisor)
                Return number
            End Function

            Private Shared Function InformationErrorType(errorValue As FormulaError) As Object
                Select Case errorValue?.Code
                    Case "#NULL!" : Return 1.0
                    Case "#DIV/0!" : Return 2.0
                    Case "#VALUE!" : Return 3.0
                    Case "#REF!" : Return 4.0
                    Case "#NAME?" : Return 5.0
                    Case "#NUM!" : Return 6.0
                    Case "#N/A" : Return 7.0
                    Case "#GETTING_DATA" : Return 8.0
                    Case SpillErrorCode : Return 9.0
                    Case CalcErrorCode : Return 14.0
                    Case Else : Return New FormulaError("#N/A")
                End Select
            End Function

        End Class

    End Class

End Namespace
