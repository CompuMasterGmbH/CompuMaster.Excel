Imports System.IO
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Function ExpandSharedFormulas(source As SheetData, sheetName As String, tableNames As HashSet(Of String)) As SheetData
            If Not source.Descendants(Of CellFormula)().Any(Function(formula) If(formula.FormulaType?.Value, CellFormulaValues.Normal) = CellFormulaValues.Shared) Then Return source
            Dim data = DirectCast(source.CloneNode(True), SheetData)
            Dim groups As New Dictionary(Of UInteger, (First As ReferenceEndpoint, Last As ReferenceEndpoint, Formula As String))
            Dim followers As New List(Of (Cell As Cell, Formula As CellFormula))
            For Each cell In data.Descendants(Of Cell)()
                Dim formula = cell.CellFormula
                If If(formula?.FormulaType?.Value, CellFormulaValues.Normal) <> CellFormulaValues.Shared Then Continue For
                If formula.SharedIndex Is Nothing Then Throw New InvalidDataException("A shared formula has no group index.")
                If String.IsNullOrEmpty(formula.Text) Then
                    followers.Add((cell, formula))
                Else
                    Dim area = ReadStructuralRange(formula.Reference?.Value)
                    Dim anchor = ReadSharedCellAddress(cell)
                    If anchor.Row <> area.First.Row OrElse anchor.Column <> area.First.Column OrElse groups.ContainsKey(formula.SharedIndex.Value) Then
                        Throw New InvalidDataException("A shared formula has an invalid or duplicate anchor.")
                    End If
                    groups.Add(formula.SharedIndex.Value, (area.First, area.Last, formula.Text))
                    MakeIndividualFormula(formula, formula.Text)
                End If
            Next

            For Each follower In followers
                Dim group As (First As ReferenceEndpoint, Last As ReferenceEndpoint, Formula As String) = Nothing
                Dim address = ReadSharedCellAddress(follower.Cell)
                If Not groups.TryGetValue(follower.Formula.SharedIndex.Value, group) OrElse
                   address.Row < group.First.Row OrElse address.Row > group.Last.Row OrElse
                   address.Column < group.First.Column OrElse address.Column > group.Last.Column Then
                    Throw New InvalidDataException("A shared formula cell lies outside its group.")
                End If
                Dim text = RewriteFormula(group.Formula, sheetName, sheetName, Nothing, Nothing, False, tableNames,
                                          copyRowOffset:=address.Row - group.First.Row, copyColumnOffset:=address.Column - group.First.Column)
                MakeIndividualFormula(follower.Formula, text)
            Next
            Return data
        End Function

        Private Shared Function ReadSharedCellAddress(cell As Cell) As ReferenceEndpoint
            Dim reference = cell.CellReference?.Value
            If String.IsNullOrEmpty(reference) Then Throw New InvalidDataException("A shared formula cell is missing its address.")
            Dim tokenEnd As Integer
            Dim endpoint = ReadEndpoint(reference, 0, tokenEnd)
            If endpoint Is Nothing OrElse Not endpoint.IsCell OrElse tokenEnd <> reference.Length Then
                Throw New InvalidDataException("Invalid shared formula cell address: " & reference)
            End If
            Return endpoint
        End Function

        Private Shared Sub MakeIndividualFormula(formula As CellFormula, text As String)
            formula.FormulaType = Nothing
            formula.SharedIndex = Nothing
            formula.Reference = Nothing
            formula.Text = text
        End Sub

    End Class

End Namespace
