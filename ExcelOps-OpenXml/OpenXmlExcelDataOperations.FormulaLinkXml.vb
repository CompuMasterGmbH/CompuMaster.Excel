Imports System.Globalization
Imports System.IO
Imports System.Xml
Imports System.Xml.XPath

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateLinkOrXml(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                               formulaCell As ExcelCell) As Object
                If name = "FILTERXML" Then
                    If arguments.Count <> 2 Then Return New FormulaError("#VALUE!")
                    Dim result = BroadcastReferenceArguments(arguments, formulaCell, AddressOf FilterWorksheetXml)
                    Dim matrix = TryCast(result, FormulaMatrix)
                    If matrix IsNot Nothing AndAlso matrix.Values.Any(Function(item) TypeOf item Is FormulaMatrix) Then
                        Return New FormulaError(CalcErrorCode)
                    End If
                    Return result
                End If
                If arguments.Count < 1 OrElse arguments.Count > 2 Then Return New FormulaError("#VALUE!")
                Dim links = AsMatrix(EvaluateExpression(arguments(0), formulaCell))
                Dim labels = If(arguments.Count = 2, AsMatrix(EvaluateExpression(arguments(1), formulaCell)), Nothing)
                Dim rows = If(labels Is Nothing, links.Rows, Math.Max(links.Rows, labels.Rows))
                Dim columns = If(labels Is Nothing, links.Columns, Math.Max(links.Columns, labels.Columns))
                Dim results = AllocateMatrixValues(rows, columns)
                For row = 0 To rows - 1
                    For column = 0 To columns - 1
                        Dim link = links.BroadcastAt(row, column)
                        Dim result As Object
                        If TypeOf link Is FormulaError Then
                            result = link
                        ElseIf labels Is Nothing Then
                            result = TextValue(link)
                        Else
                            result = If(labels.At(0, 0), 0.0)
                        End If
                        results(row * columns + column) = result
                    Next
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private Function FilterWorksheetXml(arguments As Object()) As Object
                Dim xmlText = TextValue(arguments(0))
                Dim query = TextValue(arguments(1))
                If xmlText.Length = 0 OrElse query.Length = 0 Then Return New FormulaError("#VALUE!")
                Dim settings As New XmlReaderSettings With {
                    .DtdProcessing = DtdProcessing.Prohibit,
                    .XmlResolver = Nothing,
                    .MaxCharactersInDocument = 32767
                }
                Try
                    Dim doc As New XmlDocument With {.XmlResolver = Nothing}
                    Using contents As New StringReader(xmlText), reader = XmlReader.Create(contents, settings)
                        doc.Load(reader)
                    End Using
                    ' Excel resolves unprefixed paths without a default namespace.
                    For Each element In doc.SelectNodes("//*[namespace-uri()!='' and string-length(substring-before(name(), ':'))=0]").OfType(Of XmlElement)().ToArray()
                        Dim replacement = doc.CreateElement(element.LocalName)
                        For Each attribute In element.Attributes.OfType(Of XmlAttribute)()
                            If attribute.Name <> "xmlns" Then replacement.Attributes.Append(DirectCast(attribute.CloneNode(True), XmlAttribute))
                        Next
                        While element.HasChildNodes
                            replacement.AppendChild(element.FirstChild)
                        End While
                        element.ParentNode.ReplaceChild(replacement, element)
                    Next
                    Dim matches = doc.SelectNodes(query)
                    If matches.Count = 0 Then Return New FormulaError("#VALUE!")
                    Dim values = matches.OfType(Of XmlNode)().Select(AddressOf ReadWorksheetXmlNode).ToArray()
                    If values.Length = 1 Then Return values(0)
                    Return New FormulaMatrix(values.Length, 1, values)
                Catch ex As XmlException
                    Return New FormulaError("#VALUE!")
                Catch ex As XPathException
                    Return New FormulaError("#VALUE!")
                End Try
            End Function

            Private Function ReadWorksheetXmlNode(node As XmlNode) As Object
                Dim text As String
                Select Case node.NodeType
                    Case XmlNodeType.Attribute
                        text = node.Value
                    Case XmlNodeType.Element
                        text = node.FirstChild?.Value
                    Case Else
                        Return New FormulaError("#VALUE!")
                End Select
                text = text?.Trim()
                If String.IsNullOrEmpty(text) Then Return New FormulaError("#VALUE!")
                Dim value As Double
                If Double.TryParse(text, NumberStyles.Float Or NumberStyles.AllowThousands, ConversionCulture, value) AndAlso
                    Not Double.IsInfinity(value) AndAlso Not Double.IsNaN(value) AndAlso
                    Not text.Contains(ConversionCulture.NumberFormat.NumberGroupSeparator) Then Return value
                Dim serial = ReadDateTimeText(text)
                If serial.HasValue Then Return serial.Value
                Return text
            End Function

        End Class

    End Class

End Namespace
