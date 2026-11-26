Imports System.IO
Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly FormulaTextTokens As New Regex(
                """(?:[^""]|"""")*""|'(?:[^']|'')*'|" &
                "(?<name>[\p{L}_][\p{L}\p{M}\p{Nd}_.]*)(?=\s*\()|(?<prefix>_xlpm\.)|" &
                "(?<error>#[A-Z][A-Z0-9/?!._]*)|(?<decimal>(?<![\p{L}\p{M}\p{Nd}_.])(?:[0-9]+\.[0-9]*|\.[0-9]+))|" &
                "(?<![\p{L}\p{Nd}_.])(?:TRUE|FALSE)(?![\p{L}\p{Nd}_.])|[{},;]",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly GermanFormulaNames As New Lazy(Of IReadOnlyDictionary(Of String, String))(AddressOf LoadGermanFormulaNames)

            Private Shared Function LoadGermanFormulaNames() As IReadOnlyDictionary(Of String, String)
                Dim names As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Using stream = GetType(OpenXmlExcelDataOperations).Assembly.GetManifestResourceStream("CompuMaster.Excel.OpenXml.GermanFormulaNames.tsv")
                    Using reader As New StreamReader(stream)
                        While Not reader.EndOfStream
                            Dim line = reader.ReadLine()
                            Dim separator = line.IndexOf(ChrW(9))
                            names.Add(line.Substring(0, separator), line.Substring(separator + 1))
                        End While
                    End Using
                End Using
                Return names
            End Function

            Private Function DisplayFormula(formula As String) As String
                Dim german = ConversionCulture.TwoLetterISOLanguageName = "de"
                Dim arrayDepth = 0
                Return "=" & FormulaTextTokens.Replace(formula,
                    Function(match)
                        Dim token = match.Value
                        If match.Groups("name").Success Then
                            For Each prefix In {"_xlfn.", "_xlws."}
                                If token.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then token = token.Substring(prefix.Length)
                            Next
                            Dim localized As String = Nothing
                            If german AndAlso GermanFormulaNames.Value.TryGetValue(token, localized) Then Return localized
                            Return token
                        End If
                        If match.Groups("prefix").Success Then Return String.Empty
                        If Not german Then Return token
                        If token = "{" Then arrayDepth += 1
                        If token = "}" Then arrayDepth -= 1
                        If token = "," Then Return If(arrayDepth > 0, ".", ";")
                        If token.Equals("TRUE", StringComparison.OrdinalIgnoreCase) Then Return "WAHR"
                        If token.Equals("FALSE", StringComparison.OrdinalIgnoreCase) Then Return "FALSCH"
                        If match.Groups("decimal").Success Then Return If(token(0) = "."c, "0", "") & token.Replace("."c, ","c)
                        If match.Groups("error").Success Then Return DisplayGermanError(token)
                        Return token
                    End Function)
            End Function

            Private Shared Function DisplayGermanError(errorCode As String) As String
                Select Case errorCode.ToUpperInvariant()
                    Case "#VALUE!" : Return "#WERT!"
                    Case "#REF!" : Return "#BEZUG!"
                    Case "#NUM!" : Return "#ZAHL!"
                    Case "#N/A" : Return "#NV"
                    Case "#SPILL!" : Return "#ÜBERLAUF!"
                    Case Else : Return errorCode
                End Select
            End Function

        End Class

    End Class

End Namespace
