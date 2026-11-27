Imports System.IO
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Function ExpandSharedFormulas(source As SheetData, sheetName As String, tableNames As HashSet(Of String),
                                              Optional allowSpillReferences As Boolean = False) As SheetData
            If Not source.Descendants(Of CellFormula)().Any(Function(formula) If(formula.FormulaType?.Value, CellFormulaValues.Normal) = CellFormulaValues.Shared) Then Return source
            Dim data = DirectCast(source.CloneNode(True), SheetData)
            Dim groups = ReadSharedFormulaGroups(source.Descendants(Of Cell)())
            For Each cell In data.Descendants(Of Cell)()
                Dim formula = cell.CellFormula
                If If(formula?.FormulaType?.Value, CellFormulaValues.Normal) <> CellFormulaValues.Shared Then Continue For
                Dim address = ReadSharedCellAddress(cell)
                Dim text = ResolveSharedFormulaText(formula, groups, sheetName, address.Row, address.Column, tableNames, allowSpillReferences)
                MakeIndividualFormula(formula, text)
            Next
            Return data
        End Function

        Private NotInheritable Class SharedFormulaGroup
            Friend ReadOnly First As ReferenceEndpoint
            Friend ReadOnly Last As ReferenceEndpoint
            Friend ReadOnly Formula As String

            Friend Sub New(first As ReferenceEndpoint, last As ReferenceEndpoint, formula As String)
                Me.First = first
                Me.Last = last
                Me.Formula = formula
            End Sub
        End Class

        Private Shared Function ReadSharedFormulaGroups(cells As IEnumerable(Of Cell)) As Dictionary(Of UInteger, SharedFormulaGroup)
            Dim groups As New Dictionary(Of UInteger, SharedFormulaGroup)
            For Each cell In cells
                Dim formula = cell.CellFormula
                If If(formula?.FormulaType?.Value, CellFormulaValues.Normal) <> CellFormulaValues.Shared Then Continue For
                If formula.SharedIndex Is Nothing Then Throw New InvalidDataException("A shared formula has no group index.")
                If String.IsNullOrEmpty(formula.Text) Then Continue For
                Dim area = ReadStructuralRange(formula.Reference?.Value)
                Dim anchor = ReadSharedCellAddress(cell)
                If anchor.Row <> area.First.Row OrElse anchor.Column <> area.First.Column OrElse groups.ContainsKey(formula.SharedIndex.Value) Then
                    Throw New InvalidDataException("A shared formula has an invalid or duplicate anchor.")
                End If
                groups.Add(formula.SharedIndex.Value, New SharedFormulaGroup(area.First, area.Last, formula.Text))
            Next
            Return groups
        End Function

        Private Function ResolveSharedFormulaText(formula As CellFormula, groups As Dictionary(Of UInteger, SharedFormulaGroup),
                                                  sheetName As String, row As Integer, column As Integer, tableNames As HashSet(Of String),
                                                  allowSpillReferences As Boolean) As String
            Dim group As SharedFormulaGroup = Nothing
            If formula.SharedIndex Is Nothing OrElse Not groups.TryGetValue(formula.SharedIndex.Value, group) OrElse
                row < group.First.Row OrElse row > group.Last.Row OrElse column < group.First.Column OrElse column > group.Last.Column Then
                Throw New InvalidDataException("A shared formula cell lies outside its group.")
            End If
            If row = group.First.Row AndAlso column = group.First.Column Then Return group.Formula
            Return RewriteFormula(group.Formula, sheetName, sheetName, Nothing, Nothing, False, tableNames,
                                  copyRowOffset:=row - group.First.Row, copyColumnOffset:=column - group.First.Column,
                                  allowSpillReferences:=allowSpillReferences)
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
