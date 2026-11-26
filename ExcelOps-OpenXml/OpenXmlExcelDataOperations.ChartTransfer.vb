Imports System.Globalization
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports ChartXml = DocumentFormat.OpenXml.Drawing.Charts
Imports DrawingXml = DocumentFormat.OpenXml.Drawing.Spreadsheet
Imports ExtendedChartXml = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Class TransferredChartNames
            Private ReadOnly _destination As Workbook
            Private ReadOnly _transferredSheets As Dictionary(Of String, String)
            Private ReadOnly _sourcePositions As Dictionary(Of String, Integer)
            Private ReadOnly _assigned As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Private ReadOnly _taken As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Private _nextId As Integer

            Friend ReadOnly Imported As New List(Of DefinedName)

            Friend Sub New(destination As Workbook, transferredSheets As Dictionary(Of String, String), orderedSheets As IReadOnlyList(Of String))
                _destination = destination
                _transferredSheets = transferredSheets
                _sourcePositions = orderedSheets.Select(Function(name, index) (Name:=name, Index:=index)).
                    ToDictionary(Function(item) item.Name, Function(item) item.Index, StringComparer.OrdinalIgnoreCase)
                If destination.DefinedNames IsNot Nothing Then
                    _taken.UnionWith(destination.DefinedNames.Elements(Of DefinedName)().Select(Function(item) item.Name.Value))
                End If
            End Sub

            Friend ReadOnly Property SourceNames As IEnumerable(Of String)
                Get
                    Return _assigned.Keys
                End Get
            End Property

            Friend Function IsExcelMoveDifference(owner As String, referencedSheet As String, isChartSheet As Boolean) As Boolean
                Dim targetName As String = Nothing
                Return _transferredSheets.TryGetValue(referencedSheet, targetName) AndAlso
                    _sourcePositions(owner) < _sourcePositions(referencedSheet) AndAlso
                    (isChartSheet OrElse Not String.Equals(referencedSheet, targetName, StringComparison.OrdinalIgnoreCase))
            End Function

            Friend Function Import(definition As DefinedName, formulaMode As CopyWorksheetFormulaMode,
                                   rewrite As Func(Of String, String)) As String
                Dim name = definition.Name.Value
                Dim assigned As String = Nothing
                If _assigned.TryGetValue(name, assigned) Then Return assigned
                assigned = name
                If Not _taken.Add(assigned) Then
                    If formulaMode = CopyWorksheetFormulaMode.PreserveFormulas Then
                        Dim existing = ReadChartSourceName(_destination, name)
                        If existing.Text <> definition.Text Then
                            Throw New NotSupportedException("The destination already defines a different ChartEx source with the same name.")
                        End If
                        _assigned.Add(name, assigned)
                        Return assigned
                    End If
                    Do
                        assigned = "_xlchart.v1." & _nextId.ToString(CultureInfo.InvariantCulture)
                        _nextId += 1
                    Loop Until _taken.Add(assigned)
                End If
                Dim copy = DirectCast(definition.CloneNode(True), DefinedName)
                copy.Name = assigned
                If formulaMode = CopyWorksheetFormulaMode.ExcelCompatible Then copy.Text = rewrite(definition.Text)
                Imported.Add(copy)
                _assigned.Add(name, assigned)
                Return assigned
            End Function
        End Class

        Private Shared Function ReadChartSourceName(workbook As Workbook, name As String) As DefinedName
            Dim definition = workbook.DefinedNames?.Elements(Of DefinedName)().SingleOrDefault(
                Function(item) item.LocalSheetId Is Nothing AndAlso String.Equals(item.Name.Value, name, StringComparison.OrdinalIgnoreCase))
            If definition Is Nothing OrElse Not name.StartsWith("_xlchart", StringComparison.OrdinalIgnoreCase) Then
                Throw New NotSupportedException("The ChartEx formula does not use a workbook-scoped chart source name.")
            End If
            Dim qualifier = FormulaSheet.Match(definition.Text, 0)
            If Not IsDirectAbsoluteName(definition) AndAlso
               Not (qualifier.Success AndAlso definition.Text.Substring(qualifier.Length) = "#REF!") Then
                Throw New NotSupportedException("The ChartEx source must refer to a direct absolute range.")
            End If
            Return definition
        End Function

        Private Shared Function IsExtendedChartFallback(placement As DrawingXml.WorksheetDrawing, relationId As String) As Boolean
            Return placement.Descendants(Of AlternateContent)().Any(Function(alternative)
                Return alternative.Elements(Of AlternateContentChoice)().Any(Function(choice) choice.Descendants().Any(
                    Function(item) item.NamespaceUri = ExtendedChartNamespace AndAlso item.LocalName = "chart")) AndAlso
                    alternative.Elements(Of AlternateContentFallback)().Any(Function(fallback) fallback.Descendants(Of ChartXml.ChartReference)().Any(
                    Function(reference) reference.Id?.Value = relationId))
            End Function)
        End Function

        Private Sub ValidateExtendedChartFallback(chart As ChartPart)
            ValidateChartCopyRelationships(chart)
            If chart.ChartSpace Is Nothing OrElse chart.ChartSpace.Descendants().Any(
                Function(item) {"f", "formula", "formulaRef", "externalData"}.Contains(item.LocalName)) OrElse
               chart.Parts.Any(Function(item) Not TypeOf item.OpenXmlPart Is ChartDrawingPart) Then
                Throw UnsupportedOperation("Updating an extended chart with an unsupported classic fallback")
            End If
            If chart.ChartDrawingPart IsNot Nothing Then
                ValidateChartCopyRelationships(chart.ChartDrawingPart)
                If chart.ChartDrawingPart.Parts.Any() Then
                    Throw UnsupportedOperation("Updating an extended chart with linked fallback shapes")
                End If
            End If
        End Sub

        Private Sub StageTransferredExtendedCharts(drawing As DrawingsPart, charts As ExtendedChartPart(), sheet As String,
                                                   formulaMode As CopyWorksheetFormulaMode, resolveQualifier As Func(Of String, String, String),
                                                   move As Boolean, retained As Boolean, formulaWorkbook As WorkbookPart,
                                                   names As TransferredChartNames, result As Dictionary(Of String, OpenXmlPartRootElement))
            Dim book = If(formulaWorkbook, WorkbookPart)
            Dim isChartSheet = drawing.GetParentParts().OfType(Of ChartsheetPart)().Any()
            For Each chart In charts
                ValidateChartCopyRelationships(chart)
                Dim source As ExtendedChartXml.ChartSpace = Nothing
                If Not _changedExtendedCharts.TryGetValue(chart, source) Then source = chart.ChartSpace
                If source Is Nothing OrElse chart.GetParentParts().Count() <> 1 OrElse chart.EmbeddedPackagePart IsNot Nothing OrElse
                   source.Descendants().Any(Function(item) item.LocalName = "externalData" OrElse
                       (item.LocalName = "f" AndAlso Not TypeOf item Is ExtendedChartXml.Formula) OrElse
                       (item.NamespaceUri = ExtendedChartNamespace AndAlso {"lvl", "numLvl", "strLvl"}.Contains(item.LocalName))) Then
                    Throw UnsupportedOperation("Transferring an extended chart with unsupported sources or cached data")
                End If
                If retained Then
                    ' The workbook name rewrite updates the retained chart's sources without changing its chart part.
                    For Each formula In source.Descendants(Of ExtendedChartXml.Formula)()
                        ReadChartSourceName(book.Workbook, formula.Text)
                    Next
                    Continue For
                End If

                Dim copy = DirectCast(source.CloneNode(True), ExtendedChartXml.ChartSpace)
                For Each formula In copy.Descendants(Of ExtendedChartXml.Formula)()
                    ValidateCopiedFormulaContext(move)
                    Dim definition = ReadChartSourceName(book.Workbook, formula.Text)
                    If move AndAlso Not LoadOptions.AllowKnownStructuralReferenceDifferences Then
                        Dim external As Boolean
                        Dim referencedSheet = SheetNameFromToken(book, FormulaSheet.Match(definition.Text, 0), external)
                        If Not external AndAlso names.IsExcelMoveDifference(sheet, referencedSheet, isChartSheet) Then
                            Throw New StructuralReferenceCompatibilityException(EngineName,
                                "Excel removes ChartEx series when moving this chart and its source sheet together. This engine preserves the series")
                        End If
                    End If
                    formula.Text = names.Import(definition, formulaMode,
                        Function(expression) RewriteFormula(expression, sheet, sheet, Nothing, Nothing, True,
                                                            rewriteSheetQualifier:=resolveQualifier, formulaWorkbook:=formulaWorkbook))
                Next
                result.Add(drawing.GetIdOfPart(chart), copy)
            Next
        End Sub

        Private Shared Function RemoveMovedChartSourceNames(book As WorkbookPart, moved As Dictionary(Of String, String),
                                                            sourceNames As IEnumerable(Of String)) As HashSet(Of String)
            Dim unused As New HashSet(Of String)(sourceNames, StringComparer.OrdinalIgnoreCase)
            If unused.Count = 0 Then Return unused

            For Each sheet In book.Workbook.Sheets.Elements(Of Sheet)()
                If moved.ContainsKey(sheet.Name.Value) Then Continue For
                Dim parent = book.GetPartById(sheet.Id.Value)
                For Each drawing In parent.Parts.Select(Function(item) item.OpenXmlPart).OfType(Of DrawingsPart)()
                    For Each chart In drawing.Parts.Select(Function(item) item.OpenXmlPart).OfType(Of ExtendedChartPart)()
                        If chart.ChartSpace IsNot Nothing Then
                            unused.ExceptWith(chart.ChartSpace.Descendants(Of ExtendedChartXml.Formula)().Select(Function(formula) formula.Text))
                        End If
                    Next
                Next
            Next
            Dim definitions = book.Workbook.DefinedNames
            For Each definition In definitions.Elements(Of DefinedName)().ToArray()
                If definition.LocalSheetId Is Nothing AndAlso unused.Contains(definition.Name.Value) Then definition.Remove()
            Next
            If Not definitions.HasChildren Then definitions.Remove()
            Return unused
        End Function

    End Class

End Namespace
