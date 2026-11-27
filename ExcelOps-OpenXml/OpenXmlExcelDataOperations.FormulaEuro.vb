Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly EuroRates As New Dictionary(Of String, (UnitsPerEuro As Double, Decimals As Integer))(StringComparer.OrdinalIgnoreCase) From {
                {"EUR", (1.0, 2)}, {"ATS", (13.7603, 2)}, {"BEF", (40.3399, 0)}, {"DEM", (1.95583, 2)},
                {"ESP", (166.386, 0)}, {"FIM", (5.94573, 2)}, {"FRF", (6.55957, 2)}, {"GRD", (340.75, 0)},
                {"IEP", (0.787564, 2)}, {"ITL", (1936.27, 0)}, {"LUF", (40.3399, 0)}, {"NLG", (2.20371, 2)},
                {"PTE", (200.482, 0)}, {"SIT", (239.64, 2)}
            }

            Private Function EvaluateEuroConvert(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 3 OrElse arguments.Count > 5 Then Return New FormulaError("#VALUE!")
                Dim values(arguments.Count - 1) As Object
                For index = 0 To arguments.Count - 1
                    Dim input = EvaluateExpression(arguments(index), formulaCell)
                    If IsMultiValue(input) Then Return New FormulaError("#VALUE!")
                    values(index) = Scalar(input)
                    If TypeOf values(index) Is FormulaError Then Return values(index)
                Next
                Dim converted = NumberValue(values(0))
                If TypeOf converted Is FormulaError Then Return converted
                Dim sourceName = TextValue(values(1))
                Dim targetName = TextValue(values(2))
                Dim source, target As (UnitsPerEuro As Double, Decimals As Integer)
                If Not EuroRates.TryGetValue(sourceName, source) OrElse Not EuroRates.TryGetValue(targetName, target) Then
                    Return New FormulaError("#VALUE!")
                End If
                Dim precision = If(values.Length >= 4, LogicalOptionValue(values(3)), False)
                If TypeOf precision Is FormulaError Then Return precision
                Dim number = DirectCast(converted, Double)
                If String.Equals(sourceName, targetName, StringComparison.OrdinalIgnoreCase) Then Return number
                Dim euro = number / source.UnitsPerEuro
                If values.Length >= 5 AndAlso values(4) IsNot Nothing Then
                    Dim places = NumberValue(values(4))
                    If TypeOf places Is FormulaError Then Return places
                    Dim significant = Math.Truncate(DirectCast(places, Double))
                    If significant < 3.0 Then Return New FormulaError("#VALUE!")
                    If euro <> 0.0 AndAlso source.UnitsPerEuro <> 1.0 Then
                        Dim rounded = RoundAt(euro, significant - 1.0 - Math.Floor(Math.Log10(Math.Abs(euro))), "ROUND")
                        If TypeOf rounded Is FormulaError Then Return rounded
                        euro = DirectCast(rounded, Double)
                    End If
                End If
                Dim result = FiniteNumber(euro * target.UnitsPerEuro)
                If TypeOf result Is FormulaError OrElse DirectCast(precision, Boolean) Then Return result
                Return RoundAt(DirectCast(result, Double), target.Decimals, "ROUND")
            End Function

        End Class

    End Class

End Namespace
