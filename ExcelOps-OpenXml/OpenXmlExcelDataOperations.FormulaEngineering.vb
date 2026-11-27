Imports System.Globalization
Imports System.IO
Imports MathNet.Numerics

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Structure WorksheetUnit
                Friend ReadOnly Group As String
                Friend ReadOnly Symbol As String
                Friend ReadOnly Factor As Double
                Friend ReadOnly PrefixPower As Integer

                Friend Sub New(group As String, symbol As String, factor As Double, prefixPower As Integer)
                    Me.Group = group
                    Me.Symbol = symbol
                    Me.Factor = factor
                    Me.PrefixPower = prefixPower
                End Sub
            End Structure

            Private Shared ReadOnly WorksheetUnits As New Lazy(Of Dictionary(Of String, WorksheetUnit))(AddressOf ReadWorksheetUnits)
            Private Shared ReadOnly SiUnitPrefixes As New Dictionary(Of String, Integer)(StringComparer.Ordinal) From {
                {"Y", 24}, {"Z", 21}, {"E", 18}, {"P", 15}, {"T", 12}, {"G", 9}, {"M", 6}, {"k", 3}, {"h", 2},
                {"da", 1}, {"e", 1}, {"d", -1}, {"c", -2}, {"m", -3}, {"u", -6}, {"n", -9}, {"p", -12},
                {"f", -15}, {"a", -18}, {"z", -21}, {"y", -24}
            }
            Private Shared ReadOnly BinaryUnitPrefixes As New Dictionary(Of String, Integer)(StringComparer.Ordinal) From {
                {"Yi", 80}, {"Zi", 70}, {"Ei", 60}, {"Pi", 50}, {"Ti", 40}, {"Gi", 30}, {"Mi", 20}, {"ki", 10}
            }

            Private Function EvaluateEngineeringFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                         formulaCell As ExcelCell) As Object
                If name = "CONVERT" Then
                    If arguments.Count <> 3 Then Return New FormulaError("#VALUE!")
                    Return BroadcastReferenceArguments(arguments, formulaCell, AddressOf ConvertWorksheetUnit)
                End If
                If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                Return BroadcastReferenceArguments(arguments, formulaCell, Function(values) WorksheetBessel(name, values))
            End Function

            Private Function WorksheetBessel(name As String, values As Object()) As Object
                If values.Any(Function(value) TypeOf value Is Boolean) Then Return New FormulaError("#VALUE!")
                Dim converted = NumberValue(values(0))
                If TypeOf converted Is FormulaError Then Return converted
                Dim n = NumberValue(values(1))
                If TypeOf n Is FormulaError Then Return n
                Dim x = DirectCast(converted, Double)
                Dim rank = DirectCast(n, Double)
                If rank < 0.0 OrElse ((name = "BESSELK" OrElse name = "BESSELY") AndAlso x <= 0.0) Then Return New FormulaError("#NUM!")
                rank = Math.Truncate(rank)
                If x = 0.0 Then Return If(rank = 0.0, 1.0, 0.0)
                Dim sign = If(x < 0.0 AndAlso rank Mod 2.0 <> 0.0, -1.0, 1.0)
                Dim magnitude = Math.Abs(x)
                Select Case name
                    Case "BESSELI" : Return FiniteNumber(sign * ModifiedBesselI(rank, magnitude))
                    Case "BESSELJ" : Return FiniteNumber(sign * SpecialFunctions.BesselJ(rank, magnitude))
                    Case "BESSELK" : Return FiniteNumber(ModifiedBesselK(rank, magnitude))
                    Case Else : Return FiniteNumber(SpecialFunctions.BesselY(rank, magnitude))
                End Select
            End Function

            Private Shared Function ReadWorksheetUnits() As Dictionary(Of String, WorksheetUnit)
                Dim units As New Dictionary(Of String, WorksheetUnit)(StringComparer.Ordinal)
                Using stream = GetType(OpenXmlExcelDataOperations).Assembly.GetManifestResourceStream("CompuMaster.Excel.OpenXml.ExcelUnits.tsv")
                    Using reader As New StreamReader(stream)
                        While Not reader.EndOfStream
                            Dim fields = reader.ReadLine().Split(ChrW(9))
                            units.Add(fields(0), New WorksheetUnit(fields(1), fields(0),
                                Double.Parse(fields(2), CultureInfo.InvariantCulture),
                                Integer.Parse(fields(3), CultureInfo.InvariantCulture)))
                        End While
                    End Using
                End Using
                Return units
            End Function

            Private Shared Function TryWorksheetUnit(input As String, ByRef unit As WorksheetUnit, ByRef multiplier As Double) As Boolean
                multiplier = 1.0
                If WorksheetUnits.Value.TryGetValue(input, unit) Then Return True
                For Each length In {2, 1}
                    If input.Length <= length Then Continue For
                    Dim prefix = input.Substring(0, length)
                    Dim symbol = input.Substring(length)
                    If Not WorksheetUnits.Value.TryGetValue(symbol, unit) OrElse unit.PrefixPower = 0 Then Continue For
                    Dim exponent As Integer
                    If SiUnitPrefixes.TryGetValue(prefix, exponent) Then
                        multiplier = Math.Pow(10.0, exponent * unit.PrefixPower)
                        Return True
                    End If
                    If unit.Group = "bit" AndAlso BinaryUnitPrefixes.TryGetValue(prefix, exponent) Then
                        multiplier = Math.Pow(2.0, exponent)
                        Return True
                    End If
                Next
                Return False
            End Function

            Private Function ConvertWorksheetUnit(values As Object()) As Object
                Dim converted = NumberValue(values(0))
                If TypeOf converted Is FormulaError Then Return converted
                If values.Any(Function(value) TypeOf value Is Boolean) Then Return New FormulaError("#VALUE!")
                Dim from As WorksheetUnit
                Dim target As WorksheetUnit
                Dim prefixFrom, prefixTo As Double
                Dim fromName = TextValue(values(1))
                Dim targetName = TextValue(values(2))
                If Not TryWorksheetUnit(fromName, from, prefixFrom) OrElse Not TryWorksheetUnit(targetName, target, prefixTo) OrElse
                    from.Group <> target.Group Then Return New FormulaError("#N/A")
                Dim number = DirectCast(converted, Double)
                If fromName = targetName Then Return number
                If from.Group = "C" Then
                    Return FiniteNumber(ConvertTemperature(ConvertTemperature(number * prefixFrom, from.Symbol, True), target.Symbol, False) / prefixTo)
                End If
                Return FiniteNumber(number * (from.Factor / target.Factor) * (prefixFrom / prefixTo))
            End Function

            Private Shared Function ConvertTemperature(value As Double, symbol As String, toCelsius As Boolean) As Double
                Select Case symbol
                    Case "F", "fah" : Return If(toCelsius, (value - 32.0) * 5.0 / 9.0, value * 9.0 / 5.0 + 32.0)
                    Case "K", "kel" : Return If(toCelsius, value - 273.15, value + 273.15)
                    Case "Rank" : Return If(toCelsius, value * 5.0 / 9.0 - 273.15, (value + 273.15) * 9.0 / 5.0)
                    Case "Reau" : Return If(toCelsius, value * 1.25, value * 0.8)
                    Case Else : Return value
                End Select
            End Function

        End Class

    End Class

End Namespace
