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

            Friend Iterator Function ReferencedDefinitions(sheet As String) As IEnumerable(Of DefinedName)
                Dim local As Dictionary(Of String, DefinedName) = Nothing
                If _local.TryGetValue(sheet, local) Then
                    For Each definition In local.Values
                        Yield definition
                    Next
                End If
                Dim used As HashSet(Of String) = Nothing
                If _usedGlobals.TryGetValue(sheet, used) Then
                    For Each name In used
                        Yield _globalSource(name)
                    Next
                End If
            End Function

            Friend Function ObserveDefinitionDependency(sheet As String, name As String, qualifier As String) As (Definition As DefinedName, Operand As String)
                If qualifier IsNot Nothing OrElse ContainsLocal(sheet, name) Then
                    Throw New NotSupportedException("Transferring a defined name that depends on a worksheet-local name is unsupported.")
                End If
                Dim definition As DefinedName = Nothing
                If Not _globalSource.TryGetValue(name, definition) Then
                    If String.Equals(name, "TRUE", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(name, "FALSE", StringComparison.OrdinalIgnoreCase) Then Return (Nothing, name)
                    Throw New NotSupportedException("A transferred name contains an unresolved defined-name reference.")
                End If
                Observe(sheet, name, Nothing)
                Return (definition, If(_destinationGlobals.Contains(name), FormatSheetName(_assigned(sheet)) & name, name))
            End Function

            Friend Function IsWorkbookImport(definition As DefinedName) As Boolean
                Return definition.LocalSheetId Is Nothing AndAlso Not _destinationGlobals.Contains(definition.Name.Value)
            End Function

            Friend Sub ValidateInvokedName(sheet As String, name As String, qualifier As String)
                If qualifier IsNot Nothing OrElse ContainsLocal(sheet, name) OrElse _globalSource.ContainsKey(name) Then
                    Throw New NotSupportedException("Transferring a defined name that calls a named function is unsupported.")
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
            Dim globalBindings As New Dictionary(Of DefinedName, String)
            For index = 0 To sheetNames.Count - 1
                Dim sheet = sheetNames(index)
                Dim rewritten As Dictionary(Of DefinedName, String) = Nothing
                If formulaMode = CopyWorksheetFormulaMode.ExcelCompatible Then
                    rewritten = RewriteTransferredNameDefinitions(plan, sheet, resolveQualifier, move)
                    For Each definition In rewritten.Keys.Where(AddressOf plan.IsWorkbookImport)
                        Dim previous As String = Nothing
                        If globalBindings.TryGetValue(definition, previous) Then
                            If previous <> rewritten(definition) Then
                                Throw New NotSupportedException("Transferring a workbook name with different worksheet-local dependencies is unsupported.")
                            End If
                        Else
                            globalBindings.Add(definition, rewritten(definition))
                        End If
                    Next
                End If
                For Each item In plan.DefinitionsFor(sheet, CUInt(firstTargetIndex + index))
                    If rewritten Is Nothing Then ValidateCopiedFormulaContext(move)
                    Dim definition = DirectCast(item.Definition.CloneNode(True), DefinedName)
                    definition.LocalSheetId = If(item.TargetScope.HasValue, New DocumentFormat.OpenXml.UInt32Value(item.TargetScope.Value), Nothing)
                    If rewritten IsNot Nothing Then definition.Text = rewritten(item.Definition)
                    result.Add(definition)
                Next
            Next
            Return result
        End Function

        Private Function RewriteTransferredNameDefinitions(plan As TransferNamePlan, sheet As String,
                                                           resolveQualifier As Func(Of String, String, String),
                                                           move As Boolean) As Dictionary(Of DefinedName, String)
            Dim scheduled As New HashSet(Of DefinedName)(plan.ReferencedDefinitions(sheet))
            Dim pending As New Queue(Of DefinedName)(scheduled)
            Dim rewritten As New Dictionary(Of DefinedName, String)
            While pending.Count > 0
                ValidateCopiedFormulaContext(move)
                Dim definition = pending.Dequeue()
                Dim scope = If(definition.LocalSheetId IsNot Nothing, sheet, Nothing)
                Dim text = RewriteFormula(definition.Text, scope, sheet, Nothing, Nothing, True, rewriteSheetQualifier:=resolveQualifier,
                    validateTransferredNameOperands:=True,
                    validateInvokedName:=Sub(name, qualifier) plan.ValidateInvokedName(sheet, name, qualifier),
                    rewriteNameOperand:=Function(name, qualifier)
                                            Dim dependency = plan.ObserveDefinitionDependency(sheet, name, qualifier)
                                            If dependency.Definition IsNot Nothing AndAlso scheduled.Add(dependency.Definition) Then pending.Enqueue(dependency.Definition)
                                            Return dependency.Operand
                                        End Function)
                rewritten.Add(definition, text)
            End While
            Return rewritten
        End Function

        Private Shared Sub InstallTransferredNames(workbook As Workbook, copied As List(Of DefinedName), formulaMode As CopyWorksheetFormulaMode)
            If copied.Count = 0 Then Return
            If workbook.DefinedNames Is Nothing Then workbook.DefinedNames = New DefinedNames()
            Dim names = workbook.DefinedNames
            For Each definition In copied
                Dim sameGlobal As DefinedName = Nothing
                If definition.LocalSheetId IsNot Nothing AndAlso formulaMode = CopyWorksheetFormulaMode.ExcelCompatible Then
                    sameGlobal = names.Elements(Of DefinedName)().FirstOrDefault(
                        Function(item) item.LocalSheetId Is Nothing AndAlso
                            String.Equals(item.Name.Value, definition.Name.Value, StringComparison.OrdinalIgnoreCase))
                End If
                If sameGlobal Is Nothing Then
                    names.Append(definition)
                Else
                    ' Excel resolves qualified name operands to the global name on load if it precedes this local name.
                    names.InsertBefore(definition, sameGlobal)
                End If
            Next
        End Sub

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
