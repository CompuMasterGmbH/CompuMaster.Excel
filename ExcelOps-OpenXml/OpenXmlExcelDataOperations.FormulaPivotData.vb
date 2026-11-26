Imports System.Globalization
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluatePivotData(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count < 2 OrElse arguments.Count Mod 2 <> 0 OrElse arguments.Count > 254 Then Return New FormulaError("#VALUE!")
                Dim measure = Scalar(EvaluateExpression(arguments(0), formulaCell))
                If TypeOf measure Is FormulaError Then Return measure
                Dim reference = EvaluateExpression(arguments(1), formulaCell)
                If TypeOf reference Is FormulaError Then Return reference
                Dim area = TryCast(reference, FormulaArea)
                If area Is Nothing Then Return New FormulaError("#REF!")
                Dim source = ResolveReference(area)
                Dim snapshot As WorksheetSnapshot = Nothing
                If Not source._worksheets.TryGetValue(area.Sheet, snapshot) Then Return New FormulaError("#REF!")

                Dim requested As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For i = 2 To arguments.Count - 1 Step 2
                    Dim field = Scalar(EvaluateExpression(arguments(i), formulaCell))
                    If TypeOf field Is FormulaError Then Return field
                    Dim item = Scalar(EvaluateExpression(arguments(i + 1), formulaCell))
                    If TypeOf item Is FormulaError Then Return item
                    Dim fieldName = TextValue(field)
                    If requested.ContainsKey(fieldName) Then Return New FormulaError("#REF!")
                    requested.Add(fieldName, TextValue(item))
                Next

                For Each table In snapshot.Part.PivotTableParts
                    Dim definition = table.PivotTableDefinition
                    If String.IsNullOrEmpty(definition?.Location?.Reference?.Value) Then Continue For
                    Dim bounds As New ExcelRange(snapshot.Name, definition.Location.Reference.Value)
                    If area.FirstRow > bounds.AddressEnd.RowIndex OrElse area.LastRow < bounds.AddressStart.RowIndex OrElse
                       area.FirstColumn > bounds.AddressEnd.ColumnIndex OrElse area.LastColumn < bounds.AddressStart.ColumnIndex Then Continue For
                    Return ReadPivotData(source, table, bounds, TextValue(measure), requested)
                Next
                Return New FormulaError("#REF!")
            End Function

            Private Function ReadPivotData(source As FormulaCalculation, table As PivotTablePart, bounds As ExcelRange,
                                           measure As String, requested As Dictionary(Of String, String)) As Object
                Dim definition = table.PivotTableDefinition
                Dim cache = table.PivotTableCacheDefinitionPart?.PivotCacheDefinition
                If cache Is Nothing OrElse cache.CacheSource Is Nothing OrElse Not (cache.CacheSource.Type?.Value = SourceValues.Worksheet) OrElse
                   definition.PageFields?.HasChildren = True Then
                    Throw New NotSupportedException("GETPIVOTDATA requires a saved worksheet-based pivot table without report-filter fields.")
                End If
                If cache.CacheFields Is Nothing OrElse definition.PivotFields Is Nothing Then
                    Throw New IO.InvalidDataException("The pivot table is missing its field definitions.")
                End If
                Dim cacheFields = cache.CacheFields.Elements(Of CacheField)().ToArray()
                Dim dataFields = If(definition.DataFields?.Elements(Of DataField)().ToArray(), Array.Empty(Of DataField)())
                Dim measureIndex = Array.FindIndex(dataFields, Function(field) String.Equals(field.Name?.Value, measure, StringComparison.OrdinalIgnoreCase))
                If measureIndex < 0 Then
                    measureIndex = Array.FindIndex(dataFields, Function(field) String.Equals(
                        cacheFields.ElementAtOrDefault(PivotAttributeIndex(field, "fld"))?.Name?.Value, measure, StringComparison.OrdinalIgnoreCase))
                End If
                If measureIndex < 0 Then Return New FormulaError("#REF!")

                Dim pivotFields = definition.PivotFields.Elements(Of PivotField)().ToArray()
                Dim rowFields = PivotAxisFields(definition.RowFields)
                Dim columnFields = PivotAxisFields(definition.ColumnFields)
                Dim rowRequests As New Dictionary(Of Integer, String)
                Dim columnRequests As New Dictionary(Of Integer, String)
                For Each request In requested
                    Dim index = Array.FindIndex(cacheFields, Function(field) String.Equals(field.Name?.Value, request.Key, StringComparison.OrdinalIgnoreCase))
                    If rowFields.Contains(index) AndAlso index >= 0 Then
                        rowRequests.Add(index, request.Value)
                    ElseIf columnFields.Contains(index) AndAlso index >= 0 Then
                        columnRequests.Add(index, request.Value)
                    Else
                        Return New FormulaError("#REF!")
                    End If
                    If cacheFields(index).GetFirstChild(Of FieldGroup)() IsNot Nothing Then
                        Throw New NotSupportedException("GETPIVOTDATA with grouped pivot fields is outside the supported calculation scope.")
                    End If
                Next

                Dim rows = MatchPivotAxis(definition.RowItems, rowFields, rowRequests, measureIndex, pivotFields, cacheFields)
                Dim columns = MatchPivotAxis(definition.ColumnItems, columnFields, columnRequests, measureIndex, pivotFields, cacheFields)
                Dim firstRow = bounds.AddressStart.RowIndex + CInt(If(definition.Location.FirstDataRow?.Value, 0UI))
                Dim firstColumn = bounds.AddressStart.ColumnIndex + CInt(If(definition.Location.FirstDataColumn?.Value, 0UI))
                For Each row In rows
                    For Each column In columns
                        If firstRow + row > bounds.AddressEnd.RowIndex OrElse firstColumn + column > bounds.AddressEnd.ColumnIndex Then Continue For
                        Dim address As New ExcelCell(bounds.AddressStart.SheetName, firstRow + row, firstColumn + column, ExcelCell.ValueTypes.All)
                        Dim value = ReadMetadataCell(source, address)
                        If value IsNot Nothing Then Return value
                    Next
                Next
                Return New FormulaError("#REF!")
            End Function

            Private Shared Function PivotAxisFields(fields As OpenXmlElement) As Integer()
                Return If(fields?.ChildElements.Select(Function(field) PivotAttributeIndex(field, "x")).ToArray(), Array.Empty(Of Integer)())
            End Function

            Private Shared Function MatchPivotAxis(axis As OpenXmlElement, fields As Integer(), requested As Dictionary(Of Integer, String),
                                                   measureIndex As Integer, pivotFields As PivotField(), cacheFields As CacheField()) As Integer()
                If axis Is Nothing OrElse Not axis.HasChildren Then Return If(requested.Count = 0, {0}, Array.Empty(Of Integer)())
                Dim found As New List(Of Integer)
                Dim previous As Integer() = Array.Empty(Of Integer)()
                For offset = 0 To axis.ChildElements.Count - 1
                    Dim item = axis.ChildElements(offset)
                    Dim repeat = PivotAttributeIndex(item, "r")
                    Dim values = previous.Take(repeat).Concat(item.ChildElements.Select(Function(entry) PivotAttributeIndex(entry, "v"))).ToArray()
                    previous = values
                    Dim kind = PivotAttribute(item, "t")
                    If kind IsNot Nothing AndAlso kind <> "data" AndAlso kind <> "default" AndAlso kind <> "grand" Then Continue For
                    If fields.Contains(-2) AndAlso PivotAttributeIndex(item, "i") <> measureIndex Then Continue For

                    Dim matched As Integer = 0
                    Dim differs As Boolean = False
                    If kind <> "grand" Then
                        For dimension = 0 To System.Math.Min(fields.Length, values.Length) - 1
                            Dim field = fields(dimension)
                            If field < 0 Then Continue For
                            Dim requestedValue As String = Nothing
                            If Not requested.TryGetValue(field, requestedValue) Then
                                differs = True
                                Exit For
                            End If
                            Dim caption = PivotItemCaption(pivotFields.ElementAtOrDefault(field), cacheFields.ElementAtOrDefault(field), values(dimension))
                            If Not String.Equals(caption, requestedValue, StringComparison.OrdinalIgnoreCase) Then
                                differs = True
                                Exit For
                            End If
                            matched += 1
                        Next
                    End If
                    If Not differs AndAlso matched = requested.Count Then found.Add(offset)
                Next
                Return found.ToArray()
            End Function

            Private Shared Function PivotItemCaption(field As PivotField, cacheField As CacheField, index As Integer) As String
                Dim item = field?.GetFirstChild(Of Items)()?.ChildElements.ElementAtOrDefault(index)
                If item Is Nothing Then Return Nothing
                Dim label = PivotAttribute(item, "n")
                If label IsNot Nothing Then Return label
                Dim cacheItem = cacheField?.GetFirstChild(Of SharedItems)()?.ChildElements.ElementAtOrDefault(PivotAttributeIndex(item, "x"))
                Return If(cacheItem Is Nothing, Nothing, If(PivotAttribute(cacheItem, "v"), String.Empty))
            End Function

            Private Shared Function PivotAttribute(element As OpenXmlElement, name As String) As String
                Return element.GetAttributes().FirstOrDefault(Function(attribute) attribute.NamespaceUri.Length = 0 AndAlso attribute.LocalName = name).Value
            End Function

            Private Shared Function PivotAttributeIndex(element As OpenXmlElement, name As String) As Integer
                Dim index As Integer
                Integer.TryParse(PivotAttribute(element, name), NumberStyles.Integer, CultureInfo.InvariantCulture, index)
                Return index
            End Function

        End Class

    End Class

End Namespace
