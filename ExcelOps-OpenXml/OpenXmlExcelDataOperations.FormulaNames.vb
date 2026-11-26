Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation
            Private ReadOnly _activeNames As New HashSet(Of DefinedName)
            Private _localValues As New Dictionary(Of Object, Lazy(Of Object))

            Friend Function EvaluateLet(bindings As IReadOnlyList(Of (Key As Object, Value As FormulaExpression)),
                                        result As FormulaExpression, formulaCell As ExcelCell) As Object
                Try
                    For Each binding In bindings
                        Dim expression = binding.Value
                        _localValues.Add(binding.Key, New Lazy(Of Object)(Function() EvaluateExpression(expression, formulaCell)))
                    Next
                    Return EvaluateExpression(result, formulaCell)
                Finally
                    For Each binding In bindings
                        _localValues.Remove(binding.Key)
                    Next
                End Try
            End Function

            Friend Function ReadLocalName(key As Object) As Object
                Return _localValues(key).Value
            End Function

            Friend Function EvaluateDefinedName(name As String, sheetName As String, formulaCell As ExcelCell) As Object
                If Not IsValidDefinedNameToken(name) Then Throw New NotSupportedException("Unsupported worksheet reference or name " & name & ".")
                Dim lookupSheet = If(sheetName, formulaCell.SheetName)
                Dim sheetIndex = DefinedNameSheetIndex(lookupSheet)
                If sheetIndex < 0 Then Return New FormulaError("#REF!")
                Dim definition = FindDefinedName(name, sheetIndex)
                If definition Is Nothing Then Return New FormulaError("#NAME?")
                Dim context = If(sheetName Is Nothing, formulaCell,
                    New ExcelCell(lookupSheet, formulaCell.RowIndex, formulaCell.ColumnIndex, ExcelCell.ValueTypes.All))
                Return EvaluateDefinedNameValue(definition, context)
            End Function

            Private Function TryInvokeDefinedLambda(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                   formulaCell As ExcelCell, ByRef result As Object) As Boolean
                If Not IsValidDefinedNameToken(name) OrElse _owner.WorkbookPart.Workbook.DefinedNames Is Nothing Then Return False
                Dim definition = FindDefinedName(name, DefinedNameSheetIndex(formulaCell.SheetName))
                If definition Is Nothing Then Return False
                result = InvokeFormulaLambda(EvaluateDefinedNameValue(definition, formulaCell), arguments, formulaCell)
                Return True
            End Function

            Private Function DefinedNameSheetIndex(sheetName As String) As Integer
                Return _owner.WorkbookSheets().FindIndex(Function(sheet) String.Equals(sheet.Name.Value, sheetName, StringComparison.OrdinalIgnoreCase))
            End Function

            Private Function FindDefinedName(name As String, sheetIndex As Integer) As DefinedName
                Dim definitions = _owner.WorkbookPart.Workbook.DefinedNames
                If definitions Is Nothing Then Return Nothing
                Dim candidates = definitions.Elements(Of DefinedName)().Where(
                    Function(item) String.Equals(item.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
                Dim local = If(sheetIndex < 0, Nothing, candidates.FirstOrDefault(
                    Function(item) item.LocalSheetId IsNot Nothing AndAlso item.LocalSheetId.Value = CUInt(sheetIndex)))
                Return If(local, candidates.FirstOrDefault(Function(item) item.LocalSheetId Is Nothing))
            End Function

            Private Function EvaluateDefinedNameValue(definition As DefinedName, formulaCell As ExcelCell) As Object
                If Not _activeNames.Add(definition) Then Throw New CircularReferenceException(formulaCell, definition.Text)
                Try
                    Dim expression = FormulaParser.Parse(definition.Text, requireAbsoluteReferences:=True)
                    Return EvaluateExpression(expression, formulaCell)
                Finally
                    _activeNames.Remove(definition)
                End Try
            End Function

        End Class

        Private Shared Function IsValidDefinedNameToken(name As String) As Boolean
            Return name.Length > 0 AndAlso (Char.IsLetter(name(0)) OrElse name(0) = "_"c) AndAlso
                name.All(Function(character) Char.IsLetterOrDigit(character) OrElse character = "_"c OrElse character = "."c)
        End Function

    End Class

End Namespace
