Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation
            Private ReadOnly _activeNames As New HashSet(Of DefinedName)

            Friend Function EvaluateDefinedName(name As String, sheetName As String, formulaCell As ExcelCell) As Object
                If Not IsValidDefinedNameToken(name) Then Throw New NotSupportedException("Unsupported worksheet reference or name " & name & ".")
                Dim lookupSheet = If(sheetName, formulaCell.SheetName)
                Dim sheetIndex = _owner.WorkbookSheets().FindIndex(
                    Function(sheet) String.Equals(sheet.Name.Value, lookupSheet, StringComparison.OrdinalIgnoreCase))
                If sheetIndex < 0 Then Return New FormulaError("#REF!")

                Dim definitions = _owner.WorkbookPart.Workbook.DefinedNames
                If definitions Is Nothing Then Return New FormulaError("#NAME?")
                Dim candidates = definitions.Elements(Of DefinedName)().Where(
                    Function(item) String.Equals(item.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
                Dim definition = candidates.FirstOrDefault(Function(item) item.LocalSheetId IsNot Nothing AndAlso item.LocalSheetId.Value = CUInt(sheetIndex))
                If definition Is Nothing Then definition = candidates.FirstOrDefault(Function(item) item.LocalSheetId Is Nothing)
                If definition Is Nothing Then Return New FormulaError("#NAME?")
                If Not _activeNames.Add(definition) Then Throw New CircularReferenceException(formulaCell, definition.Text)
                Try
                    Dim expression = FormulaParser.Parse(definition.Text, requireAbsoluteReferences:=True)
                    Dim context = If(sheetName Is Nothing, formulaCell,
                        New ExcelCell(lookupSheet, formulaCell.RowIndex, formulaCell.ColumnIndex, ExcelCell.ValueTypes.All))
                    Return EvaluateExpression(expression, context)
                Finally
                    _activeNames.Remove(definition)
                End Try
            End Function

            Private Shared Function IsValidDefinedNameToken(name As String) As Boolean
                Return name.Length > 0 AndAlso (Char.IsLetter(name(0)) OrElse name(0) = "_"c) AndAlso
                    name.All(Function(character) Char.IsLetterOrDigit(character) OrElse character = "_"c OrElse character = "."c)
            End Function
        End Class

    End Class

End Namespace
