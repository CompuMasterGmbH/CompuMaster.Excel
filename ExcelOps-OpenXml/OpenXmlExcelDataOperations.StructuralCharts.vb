Imports System.Globalization
Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports ChartXml = DocumentFormat.OpenXml.Drawing.Charts
Imports DrawingXml = DocumentFormat.OpenXml.Drawing.Spreadsheet
Imports ExtendedChartXml = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Const ExtendedChartNamespace As String = "http://schemas.microsoft.com/office/drawing/2014/chartex"
        Private ReadOnly _changedCharts As New Dictionary(Of ChartPart, ChartXml.ChartSpace)
        Private ReadOnly _changedDrawings As New Dictionary(Of DrawingsPart, DrawingXml.WorksheetDrawing)

        Private Function StageChartChanges(sheets As IEnumerable(Of Sheet), editedPart As WorksheetPart, editedSheet As String,
                                           shift As StructuralShift, newSheetName As String,
                                           Optional book As WorkbookPart = Nothing, Optional links As HashSet(Of Integer) = Nothing) As (
                                           Charts As Dictionary(Of ChartPart, ChartXml.ChartSpace),
                                           Drawings As Dictionary(Of DrawingsPart, DrawingXml.WorksheetDrawing))
            Dim charts As New Dictionary(Of ChartPart, ChartXml.ChartSpace)
            Dim drawings As New Dictionary(Of DrawingsPart, DrawingXml.WorksheetDrawing)
            book = If(book, WorkbookPart)
            For Each sheet In sheets
                Dim owner = book.GetPartById(sheet.Id.Value)
                Dim worksheet = TryCast(owner, WorksheetPart)
                Dim chartSheet = TryCast(owner, ChartsheetPart)
                Dim drawing = If(worksheet IsNot Nothing, worksheet?.DrawingsPart, chartSheet?.DrawingsPart)
                If chartSheet IsNot Nothing AndAlso
                   (drawing Is Nothing OrElse chartSheet.Parts.Any(
                        Function(part) part.OpenXmlPart IsNot drawing AndAlso Not TypeOf part.OpenXmlPart Is SpreadsheetPrinterSettingsPart)) Then
                    Throw UnsupportedOperation("Updating references on an unsupported chart sheet")
                End If
                If drawing Is Nothing Then Continue For
                Dim extendedCharts = drawing.Parts.Select(Function(item) item.OpenXmlPart).OfType(Of ExtendedChartPart)().ToArray()
                Dim allowExtendedCharts = extendedCharts.Length > 0
                If drawing.WorksheetDrawing Is Nothing OrElse drawing.GetParentParts().Count() <> 1 OrElse drawing.Parts.Any(
                    Function(part) Not TypeOf part.OpenXmlPart Is ChartPart AndAlso Not TypeOf part.OpenXmlPart Is ImagePart AndAlso
                        Not (allowExtendedCharts AndAlso TypeOf part.OpenXmlPart Is ExtendedChartPart)) Then
                    Throw UnsupportedOperation("Updating references in unsupported drawing parts")
                End If
                If allowExtendedCharts Then ValidateExtendedChartNames(extendedCharts, book, editedSheet, shift, links)
                Dim placement As DrawingXml.WorksheetDrawing = Nothing
                If Not _changedDrawings.TryGetValue(drawing, placement) Then placement = drawing.WorksheetDrawing
                Dim changedPlacement = StageChartAnchors(placement, If(worksheet IsNot Nothing AndAlso worksheet Is editedPart, shift, Nothing),
                                                         allowExtendedCharts)
                If changedPlacement IsNot Nothing Then drawings.Add(drawing, changedPlacement)

                For Each chart In drawing.ChartParts
                    Dim chartSpace As ChartXml.ChartSpace = Nothing
                    If Not _changedCharts.TryGetValue(chart, chartSpace) Then chartSpace = chart.ChartSpace
                    Dim changedChart = StageChartReferences(chart, chartSpace, editedSheet, shift, newSheetName, book, links)
                    If changedChart IsNot Nothing Then charts.Add(chart, changedChart)
                Next
            Next
            Return (charts, drawings)
        End Function

        Private Sub ValidateExtendedChartNames(charts As ExtendedChartPart(), book As WorkbookPart, editedSheet As String,
                                               shift As StructuralShift, links As HashSet(Of Integer))
            Dim names As New Dictionary(Of String, DefinedName)(StringComparer.OrdinalIgnoreCase)
            If book.Workbook.DefinedNames IsNot Nothing Then
                For Each definition In book.Workbook.DefinedNames.Elements(Of DefinedName)()
                    If definition.LocalSheetId Is Nothing AndAlso definition.Name.Value.StartsWith("_xlchart", StringComparison.OrdinalIgnoreCase) Then
                        names.Add(definition.Name.Value, definition)
                    End If
                Next
            End If
            For Each chart In charts
                Dim source = chart.ChartSpace
                If source Is Nothing OrElse chart.EmbeddedPackagePart IsNot Nothing OrElse chart.ExternalRelationships.Any() OrElse
                   source.Descendants().Any(Function(element) element.LocalName = "externalData" OrElse
                       (element.LocalName = "f" AndAlso Not TypeOf element Is ExtendedChartXml.Formula)) OrElse
                   source.Descendants(Of ExtendedChartXml.Formula)().Any(Function(formula) Not names.ContainsKey(formula.Text)) Then
                    Throw UnsupportedOperation("Updating a ChartEx chart with unsupported sources")
                End If
                If shift Is Nothing Then Continue For
                If links IsNot Nothing Then Throw UnsupportedOperation("Shifting incoming workbook links in ChartEx data")
                ' Excel stores the uncached data references in the workbook's hidden names.
                For Each formula In source.Descendants(Of ExtendedChartXml.Formula)()
                    Dim definition = names(formula.Text)
                    If Not IsDirectAbsoluteName(definition) Then Throw UnsupportedOperation("Shifting a ChartEx source that is not a direct absolute range")
                    Dim qualifier = FormulaSheet.Match(definition.Text, 0)
                    Dim external As Boolean
                    Dim sheet = SheetNameFromToken(book, qualifier, external)
                    If external OrElse Not String.Equals(sheet, editedSheet, StringComparison.OrdinalIgnoreCase) Then Continue For
                    If shift.ClearsTableDataBody Then Throw UnsupportedOperation("Clearing a table's data body on a ChartEx source worksheet")
                    Dim rewritten = RewriteFormula(definition.Text, Nothing, editedSheet, shift, Nothing, True, formulaWorkbook:=book)
                    If rewritten = definition.Text Then Continue For
                    If rewritten.EndsWith("#REF!", StringComparison.Ordinal) OrElse
                       formula.Parent.ChildElements.Any(Function(item) item.LocalName = "lvl" OrElse item.LocalName = "numLvl" OrElse item.LocalName = "strLvl") Then
                        Throw UnsupportedOperation("Deleting an entire ChartEx source or shifting cached ChartEx data")
                    End If
                Next
            Next
        End Sub

        Private Function StageChartReferences(chart As ChartPart, source As ChartXml.ChartSpace, editedSheet As String,
                                              shift As StructuralShift, newSheetName As String, book As WorkbookPart,
                                              links As HashSet(Of Integer)) As ChartXml.ChartSpace
            If source Is Nothing OrElse chart.GetParentParts().Count() <> 1 OrElse source.Descendants(Of ChartXml.ExternalData)().Any() OrElse
               source.Descendants(Of ChartXml.PivotSource)().Any() OrElse chart.Parts.Any(
                    Function(part) Not TypeOf part.OpenXmlPart Is ChartStylePart AndAlso
                        Not TypeOf part.OpenXmlPart Is ChartColorStylePart AndAlso Not TypeOf part.OpenXmlPart Is ThemeOverridePart AndAlso
                        Not TypeOf part.OpenXmlPart Is ImagePart) Then
                Throw UnsupportedOperation("Updating references in an unsupported chart")
            End If
            If source.Descendants().Any(Function(element) Not TypeOf element Is ChartXml.Formula AndAlso
                (element.LocalName = "f" OrElse element.LocalName = "formula" OrElse element.LocalName = "formulaRef")) Then
                Throw UnsupportedOperation("Updating chart formulas in an extended chart element")
            End If

            Dim revised = DirectCast(source.CloneNode(True), ChartXml.ChartSpace)
            Dim changed = False
            For Each formula In revised.Descendants(Of ChartXml.Formula)().ToArray()
                Dim qualifier = FormulaSheet.Match(formula.Text, 0)
                If Not qualifier.Success Then Throw UnsupportedOperation("Updating a chart source without an explicit sheet")
                Dim external As Boolean
                Dim sheet = SheetNameFromToken(book, qualifier, external)
                Dim externalPrefix As String = Nothing
                If external <> (links IsNot Nothing) Then Continue For
                If external Then sheet = LinkedStructuralSheet(sheet, links, externalPrefix)
                If Not String.Equals(sheet, editedSheet, StringComparison.OrdinalIgnoreCase) Then Continue For
                If shift?.ClearsTableDataBody = True Then Throw UnsupportedOperation("Clearing a table's data body on a chart source worksheet")
                Dim newQualifier = qualifier.Value
                If newSheetName IsNot Nothing Then
                    newQualifier = If(external, FormatLinkedSheetName(externalPrefix, newSheetName), FormatSheetName(newSheetName))
                End If
                Dim address = formula.Text.Substring(qualifier.Length)
                If address = "#REF!" Then
                    If newSheetName IsNot Nothing Then
                        formula.Text = newQualifier & address
                        changed = True
                    End If
                    Continue For
                End If
                Dim original = ReadChartRange(address)
                Dim rewritten = If(shift Is Nothing, address, shift.RewriteReference(original.First, original.Last, ":"))
                If rewritten = address AndAlso newSheetName Is Nothing Then Continue For
                If rewritten = "#REF!" AndAlso revised.Descendants().Count(Function(element) element.LocalName = "ser") > 1 Then
                    Throw UnsupportedOperation("Deleting an entire source range in a chart with multiple series")
                End If

                Dim reference = formula.Parent
                Dim role = reference.Parent
                If Not (TypeOf role Is ChartXml.CategoryAxisData OrElse TypeOf role Is ChartXml.Values OrElse
                        TypeOf role Is ChartXml.XValues OrElse TypeOf role Is ChartXml.YValues OrElse
                        TypeOf role Is ChartXml.SeriesText OrElse
                        (TypeOf role Is ChartXml.ChartText AndAlso TypeOf role.Parent Is ChartXml.Title)) OrElse
                   (TypeOf role Is ChartXml.CategoryAxisData AndAlso reference.GetFirstChild(Of ChartXml.MultiLevelStringCache)() IsNot Nothing) Then
                    Throw UnsupportedOperation("Updating this chart reference role or cache")
                End If

                changed = True
                formula.Text = newQualifier & rewritten
                If rewritten = "#REF!" AndAlso (TypeOf role Is ChartXml.CategoryAxisData OrElse TypeOf role Is ChartXml.SeriesText) Then
                    role.Remove()
                ElseIf shift IsNot Nothing AndAlso rewritten <> address Then
                    RemapChartCache(reference, original, rewritten, shift)
                End If
            Next
            Return If(changed, revised, Nothing)
        End Function

        Private Shared Function ReadChartRange(address As String) As (First As ReferenceEndpoint, Last As ReferenceEndpoint)
            Dim cursor As Integer
            Dim first = ReadEndpoint(address, 0, cursor)
            If first Is Nothing OrElse Not first.IsCell Then Throw UnsupportedReferenceSyntax()
            If cursor = address.Length Then Return (first, Nothing)
            If address(cursor) <> ":"c Then Throw UnsupportedReferenceSyntax()
            Dim last = ReadEndpoint(address, cursor + 1, cursor)
            If last Is Nothing OrElse Not last.IsCell OrElse cursor <> address.Length OrElse
               first.Row > last.Row OrElse first.Column > last.Column OrElse
               (first.Row <> last.Row AndAlso first.Column <> last.Column) Then Throw UnsupportedReferenceSyntax()
            Return (first, last)
        End Function

        Private Shared Sub RemapChartCache(reference As OpenXmlElement, original As (First As ReferenceEndpoint, Last As ReferenceEndpoint),
                                           rewritten As String, shift As StructuralShift)
            Dim numberCache = reference.GetFirstChild(Of ChartXml.NumberingCache)()
            Dim stringCache = reference.GetFirstChild(Of ChartXml.StringCache)()
            Dim cache As OpenXmlCompositeElement = If(DirectCast(numberCache, OpenXmlCompositeElement), stringCache)
            If cache Is Nothing Then Return

            If rewritten = "#REF!" Then
                cache.RemoveAllChildren(Of ChartXml.NumericPoint)()
                cache.RemoveAllChildren(Of ChartXml.StringPoint)()
                Dim count = cache.GetFirstChild(Of ChartXml.PointCount)()
                If count Is Nothing Then
                    count = New ChartXml.PointCount()
                    cache.AddChild(count)
                End If
                If TypeOf reference.Parent Is ChartXml.Values OrElse TypeOf reference.Parent Is ChartXml.YValues Then
                    If numberCache Is Nothing Then Throw UnsupportedReferenceSyntax()
                    count.Val = 1UI
                    cache.AddChild(New ChartXml.NumericPoint(New ChartXml.NumericValue("1")) With {.Index = 0UI})
                Else
                    count.Val = 0UI
                End If
                Return
            End If

            Dim newRange = ReadChartRange(rewritten)
            Dim first = original.First
            Dim last = If(original.Last, first)
            Dim newFirst = newRange.First
            Dim newLast = If(newRange.Last, newFirst)
            Dim byRow = first.Row <> last.Row
            Dim newPointCount = newLast.Row - newFirst.Row + newLast.Column - newFirst.Column + 1
            Dim pointCount = cache.GetFirstChild(Of ChartXml.PointCount)()
            If pointCount IsNot Nothing Then pointCount.Val = CUInt(newPointCount)
            For Each point In cache.ChildElements.ToArray()
                Dim numericPoint = TryCast(point, ChartXml.NumericPoint)
                Dim stringPoint = TryCast(point, ChartXml.StringPoint)
                Dim index = If(numericPoint?.Index?.Value, stringPoint?.Index?.Value)
                If Not index.HasValue Then Continue For
                If index.Value > CUInt(last.Row - first.Row + last.Column - first.Column) Then Throw UnsupportedReferenceSyntax()
                Dim row = first.Row + If(byRow, CInt(index.Value), 0)
                Dim column = first.Column + If(byRow, 0, CInt(index.Value))
                If shift.Vertical AndAlso shift.Includes(column) Then
                    row = shift.MoveIndex(row)
                ElseIf Not shift.Vertical AndAlso shift.Includes(row) Then
                    column = shift.MoveIndex(column)
                End If
                If row < newFirst.Row OrElse row > newLast.Row OrElse column < newFirst.Column OrElse column > newLast.Column Then
                    point.Remove()
                Else
                    Dim newIndex = CUInt(row - newFirst.Row + column - newFirst.Column)
                    If numericPoint IsNot Nothing Then
                        numericPoint.Index = newIndex
                    Else
                        stringPoint.Index = newIndex
                    End If
                End If
            Next
        End Sub

        Private Function StageChartAnchors(source As DrawingXml.WorksheetDrawing, shift As StructuralShift,
                                           Optional allowExtendedCharts As Boolean = False) As DrawingXml.WorksheetDrawing
            Dim revised = DirectCast(source.CloneNode(True), DrawingXml.WorksheetDrawing)
            Dim changed = False
            For Each anchor In revised.ChildElements
                Dim extendedFrame = allowExtendedCharts AndAlso anchor.Descendants().Any(
                    Function(element) element.LocalName = "chart" AndAlso element.NamespaceUri = ExtendedChartNamespace)
                Dim picture = anchor.GetFirstChild(Of DrawingXml.Picture)()
                Dim frame = anchor.GetFirstChild(Of DrawingXml.GraphicFrame)()
                If Not extendedFrame AndAlso picture Is Nothing AndAlso
                   (frame Is Nothing OrElse Not frame.Descendants(Of ChartXml.ChartReference)().Any()) Then
                    Throw UnsupportedOperation("Updating a drawing with unsupported worksheet objects")
                End If
                If shift Is Nothing OrElse TypeOf anchor Is DrawingXml.AbsoluteAnchor Then Continue For
                If picture IsNot Nothing Then Throw UnsupportedOperation("Shifting cells in a worksheet with a cell-anchored picture")
                Dim two = TryCast(anchor, DrawingXml.TwoCellAnchor)
                Dim one = TryCast(anchor, DrawingXml.OneCellAnchor)
                If two Is Nothing AndAlso one Is Nothing Then Throw UnsupportedOperation("Updating an unsupported drawing anchor")
                If two?.EditAs?.Value = DrawingXml.EditAsValues.Absolute Then Continue For
                Dim first = If(two?.FromMarker, one?.FromMarker)
                Dim startIndex = ReadChartMarker(first, shift.Vertical)
                If Not shift.WholeAxis Then
                    If shift.Includes(ReadChartMarker(first, Not shift.Vertical)) AndAlso startIndex >= shift.Start Then
                        Throw UnsupportedOperation("Shifting a partial cell range through a drawing anchor")
                    End If
                    Continue For
                End If
                Dim movedStart = ShiftChartMarker(startIndex, shift)
                changed = SetChartMarker(first, shift.Vertical, movedStart, Not shift.Inserting AndAlso startIndex >= shift.Start AndAlso
                                         startIndex < shift.Start + shift.Count) OrElse changed
                If two Is Nothing Then Continue For
                Dim lastIndex = ReadChartMarker(two.ToMarker, shift.Vertical)
                Dim movesWithoutResizing = If(two.EditAs?.Value, DrawingXml.EditAsValues.TwoCell) = DrawingXml.EditAsValues.OneCell
                Dim movedEnd = If(movesWithoutResizing,
                                  lastIndex + movedStart - startIndex, ShiftChartMarker(lastIndex, shift))
                changed = SetChartMarker(two.ToMarker, shift.Vertical, movedEnd, Not shift.Inserting AndAlso
                                         Not movesWithoutResizing AndAlso
                                         lastIndex >= shift.Start AndAlso lastIndex < shift.Start + shift.Count) OrElse changed
            Next
            Return If(changed, revised, Nothing)
        End Function

        Private Shared Function ShiftChartMarker(index As Integer, shift As StructuralShift) As Integer
            Dim moved = shift.MoveIndex(index)
            Return If(moved < 0, shift.Start, moved)
        End Function

        Private Shared Function ReadChartMarker(marker As DrawingXml.MarkerType, vertical As Boolean) As Integer
            Dim text = If(vertical, marker?.GetFirstChild(Of DrawingXml.RowId)()?.Text, marker?.GetFirstChild(Of DrawingXml.ColumnId)()?.Text)
            Dim result As Integer
            If Not Integer.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, result) OrElse
               result >= If(vertical, RowCount, ColumnCount) Then Throw New InvalidDataException("Invalid drawing marker.")
            Return result
        End Function

        Private Shared Function SetChartMarker(marker As DrawingXml.MarkerType, vertical As Boolean, index As Integer, resetOffset As Boolean) As Boolean
            If index < 0 OrElse index >= If(vertical, RowCount, ColumnCount) Then Throw UnsupportedReferenceSyntax()
            Dim coordinate As OpenXmlLeafTextElement = If(vertical, DirectCast(marker.GetFirstChild(Of DrawingXml.RowId)(), OpenXmlLeafTextElement),
                                                         marker.GetFirstChild(Of DrawingXml.ColumnId)())
            Dim value = index.ToString(CultureInfo.InvariantCulture)
            If Not resetOffset AndAlso coordinate.Text = value Then Return False
            coordinate.Text = value
            If resetOffset Then
                Dim offset As OpenXmlLeafTextElement = If(vertical, DirectCast(marker.GetFirstChild(Of DrawingXml.RowOffset)(), OpenXmlLeafTextElement),
                                                         marker.GetFirstChild(Of DrawingXml.ColumnOffset)())
                offset.Text = "0"
            End If
            Return True
        End Function

        Private Sub SaveChangedCharts(source As WorkbookPart, target As WorkbookPart)
            For Each change In _changedDrawings
                Dim owner = change.Key.GetParentParts().Single()
                Dim targetOwner = target.GetPartById(source.GetIdOfPart(owner))
                Dim targetDrawing = targetOwner.GetPartById(owner.GetIdOfPart(change.Key))
                Using stream = targetDrawing.GetStream(FileMode.Create, FileAccess.Write)
                    change.Value.Save(stream)
                End Using
            Next
            For Each change In _changedCharts
                Dim drawing = change.Key.GetParentParts().OfType(Of DrawingsPart)().Single()
                Dim owner = drawing.GetParentParts().Single()
                Dim targetOwner = target.GetPartById(source.GetIdOfPart(owner))
                Dim targetDrawing = targetOwner.GetPartById(owner.GetIdOfPart(drawing))
                Dim targetChart = targetDrawing.GetPartById(drawing.GetIdOfPart(change.Key))
                Using stream = targetChart.GetStream(FileMode.Create, FileAccess.Write)
                    change.Value.Save(stream)
                End Using
            Next
        End Sub

    End Class

End Namespace
