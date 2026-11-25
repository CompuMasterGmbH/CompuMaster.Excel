Imports System.IO
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Class TransferNamePlan
            Private ReadOnly _assigned As Dictionary(Of String, String)
            Private ReadOnly _globalSource As New Dictionary(Of String, DefinedName)(StringComparer.OrdinalIgnoreCase)
            Private ReadOnly _local As New Dictionary(Of String, Dictionary(Of String, DefinedName))(StringComparer.OrdinalIgnoreCase)
            Private ReadOnly _usedGlobals As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            Private ReadOnly _destinationGlobals As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Private ReadOnly _importedGlobals As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Friend Sub New(source As Workbook, destination As Workbook, sheets As List(Of Sheet), assigned As Dictionary(Of String, String))
                _assigned = assigned
                If source.DefinedNames IsNot Nothing Then
                    For Each definition In source.DefinedNames.Elements(Of DefinedName)()
                        If definition.LocalSheetId Is Nothing Then
                            _globalSource.Add(definition.Name.Value, definition)
                        Else
                            If definition.LocalSheetId.Value >= sheets.Count Then Throw New InvalidDataException("A defined name has an invalid worksheet scope.")
                            Dim sheet = sheets(CInt(definition.LocalSheetId.Value)).Name.Value
                            Dim names As Dictionary(Of String, DefinedName) = Nothing
                            If Not _local.TryGetValue(sheet, names) Then
                                names = New Dictionary(Of String, DefinedName)(StringComparer.OrdinalIgnoreCase)
                                _local.Add(sheet, names)
                            End If
                            names.Add(definition.Name.Value, definition)
                        End If
                    Next
                End If
                If destination.DefinedNames IsNot Nothing Then
                    _destinationGlobals.UnionWith(destination.DefinedNames.Elements(Of DefinedName)().
                        Where(Function(definition) definition.LocalSheetId Is Nothing).Select(Function(definition) definition.Name.Value))
                End If
            End Sub

            Friend Sub Observe(sheet As String, name As String, qualifier As String)
                If qualifier IsNot Nothing Then
                    If Not _assigned.ContainsKey(qualifier) OrElse Not ContainsLocal(qualifier, name) Then
                        Throw New NotSupportedException("Transferring a qualified defined-name reference outside the transferred sheets is unsupported.")
                    End If
                    Return
                End If
                If ContainsLocal(sheet, name) OrElse Not _globalSource.ContainsKey(name) Then Return
                Dim used As HashSet(Of String) = Nothing
                If Not _usedGlobals.TryGetValue(sheet, used) Then
                    used = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    _usedGlobals.Add(sheet, used)
                End If
                used.Add(name)
            End Sub

            Friend Sub ValidateRetainedReference(name As String, qualifier As String)
                If qualifier IsNot Nothing AndAlso _assigned.ContainsKey(qualifier) AndAlso ContainsLocal(qualifier, name) Then
                    Throw New NotSupportedException("A retained formula explicitly refers to a sheet-scoped name on a moved sheet.")
                End If
            End Sub

            Private Function ContainsLocal(sheet As String, name As String) As Boolean
                Dim names As Dictionary(Of String, DefinedName) = Nothing
                Return _local.TryGetValue(sheet, names) AndAlso names.ContainsKey(name)
            End Function

            Friend Iterator Function DefinitionsFor(sheet As String, targetIndex As UInteger) As IEnumerable(Of (Definition As DefinedName, TargetScope As UInteger?))
                Dim local As Dictionary(Of String, DefinedName) = Nothing
                If _local.TryGetValue(sheet, local) Then
                    For Each definition In local.Values
                        Yield (definition, targetIndex)
                    Next
                End If

                Dim used As HashSet(Of String) = Nothing
                If Not _usedGlobals.TryGetValue(sheet, used) Then Return
                For Each name In used
                    If _destinationGlobals.Contains(name) Then
                        Yield (_globalSource(name), targetIndex)
                    ElseIf _importedGlobals.Add(name) Then
                        Yield (_globalSource(name), Nothing)
                    End If
                Next
            End Function
        End Class

        Private Function PrepareTransferredNames(plan As TransferNamePlan, sheetNames As IReadOnlyList(Of String), firstTargetIndex As Integer,
                                                 resolveQualifier As Func(Of String, String, String), formulaMode As CopyWorksheetFormulaMode,
                                                 move As Boolean) As List(Of DefinedName)
            Dim result As New List(Of DefinedName)
            For index = 0 To sheetNames.Count - 1
                Dim sheet = sheetNames(index)
                For Each item In plan.DefinitionsFor(sheet, CUInt(firstTargetIndex + index))
                    ValidateCopiedFormulaContext(move)
                    Dim definition = DirectCast(item.Definition.CloneNode(True), DefinedName)
                    definition.LocalSheetId = If(item.TargetScope.HasValue, New DocumentFormat.OpenXml.UInt32Value(item.TargetScope.Value), Nothing)
                    If formulaMode = CopyWorksheetFormulaMode.ExcelCompatible Then
                        If Not IsDirectAbsoluteName(item.Definition) Then
                            Throw New NotSupportedException("Excel-compatible sheet transfer supports defined names that refer to a direct absolute cell or range.")
                        End If
                        definition.Text = RewriteFormula(definition.Text, sheet, sheet, Nothing, Nothing, True, rewriteSheetQualifier:=resolveQualifier)
                    End If
                    result.Add(definition)
                Next
            Next
            Return result
        End Function

        Private Shared Function IsDirectAbsoluteName(definition As DefinedName) As Boolean
            Dim value = definition.Text
            Dim sheet = FormulaSheet.Match(value, 0)
            If Not sheet.Success AndAlso definition.LocalSheetId Is Nothing Then Return False
            Dim endpointEnd As Integer
            Dim first = ReadEndpoint(value, sheet.Length, endpointEnd)
            If first Is Nothing OrElse Not first.IsCell OrElse Not first.IsAbsolute Then Return False
            If endpointEnd = value.Length Then Return True
            If value(endpointEnd) <> ":"c Then Return False
            Dim last = ReadEndpoint(value, endpointEnd + 1, endpointEnd)
            Return last IsNot Nothing AndAlso last.IsCell AndAlso last.IsAbsolute AndAlso endpointEnd = value.Length
        End Function

    End Class

End Namespace
