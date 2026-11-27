Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateDatabaseFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                      formulaCell As ExcelCell) As Object
                If arguments.Count <> 3 Then Return New FormulaError("#VALUE!")
                Dim databaseInput = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf databaseInput Is FormulaError Then Return databaseInput
                Dim fieldInput = EvaluateExpression(arguments(1), formulaCell)
                Dim criteriaInput = EvaluateExpression(arguments(2), formulaCell)
                If TypeOf criteriaInput Is FormulaError Then Return criteriaInput
                If TypeOf fieldInput Is FormulaArea AndAlso IsMultiValue(fieldInput) Then Return New FormulaError("#VALUE!")
                Dim fields = AsMatrix(fieldInput)
                Dim invalidCriteria = TryCast(criteriaInput, FormulaMatrix)
                If invalidCriteria IsNot Nothing Then
                    Dim rows = Math.Max(fields.Rows, invalidCriteria.Rows)
                    Dim columns = Math.Max(fields.Columns, invalidCriteria.Columns)
                    Dim errors = AllocateMatrixValues(rows, columns)
                    For index = 0 To errors.Length - 1
                        errors(index) = New FormulaError("#VALUE!")
                    Next
                    If errors.Length = 1 Then Return errors(0)
                    Return New FormulaMatrix(rows, columns, errors)
                End If

                Dim databaseArea = TryCast(databaseInput, FormulaArea)
                Dim criteriaArea = TryCast(criteriaInput, FormulaArea)
                If databaseArea Is Nothing OrElse criteriaArea Is Nothing OrElse databaseArea.FirstRow = databaseArea.LastRow OrElse
                    criteriaArea.FirstRow = criteriaArea.LastRow Then Return New FormulaError("#VALUE!")
                Dim database = AsMatrix(databaseArea)
                Dim criteria = AsMatrix(criteriaArea)

                Dim filters As New List(Of List(Of Predicate(Of Integer)))
                For row = 1 To criteria.Rows - 1
                    Dim conditions As New List(Of Predicate(Of Integer))
                    For column = 0 To criteria.Columns - 1
                        Dim condition = criteria.At(row, column)
                        If condition Is Nothing Then Continue For
                        Dim field = DatabaseFieldIndex(database, criteria.At(0, column), headersOnly:=True)
                        If field < 0 Then
                            conditions.Add(DatabaseComputedCriterion(criteriaArea, row, column))
                        Else
                            conditions.Add(DatabaseColumnCriterion(database, field, condition))
                        End If
                    Next
                    filters.Add(conditions)
                Next

                Dim matched As New List(Of Integer)
                For row = 1 To database.Rows - 1
                    Dim record = row
                    If filters.Any(Function(conditions) conditions.All(Function(condition) condition(record))) Then matched.Add(row)
                Next

                Dim results = AllocateMatrixValues(fields.Rows, fields.Columns)
                For index = 0 To fields.Values.Length - 1
                    Dim field = fields.Values(index)
                    If TypeOf field Is FormulaError Then
                        results(index) = field
                    ElseIf fieldInput Is Nothing AndAlso (name = "DCOUNT" OrElse name = "DCOUNTA") Then
                        results(index) = CDbl(matched.Count)
                    Else
                        Dim fieldIndex = DatabaseFieldIndex(database, field)
                        results(index) = If(fieldIndex < 0, New FormulaError("#VALUE!"), DatabaseAggregate(name, database, fieldIndex, matched))
                    End If
                Next
                If results.Length = 1 Then Return results(0)
                Return New FormulaMatrix(fields.Rows, fields.Columns, results)
            End Function

            Private Function DatabaseColumnCriterion(database As FormulaMatrix, field As Integer, input As Object) As Predicate(Of Integer)
                Dim matches = MakeDatabaseCriterion(input)
                Return Function(record) matches(database.At(record, field))
            End Function

            Private Function DatabaseComputedCriterion(area As FormulaArea, row As Integer, column As Integer) As Predicate(Of Integer)
                Dim source = ResolveReference(area)
                Dim address As New ExcelCell(area.Sheet, area.FirstRow + row, area.FirstColumn + column, ExcelCell.ValueTypes.All)
                Dim cell As DocumentFormat.OpenXml.Spreadsheet.Cell = Nothing
                If Not source._worksheets(area.Sheet).Cells.TryGetValue((address.RowIndex, address.ColumnIndex), cell) OrElse
                    cell.CellFormula Is Nothing Then Return Function(record) False
                Dim formula = cell.CellFormula.Text
                Return Function(record)
                           Dim shifted = source._owner.RewriteFormula(formula, area.Sheet, area.Sheet, Nothing, Nothing, False, copyRowOffset:=record - 1)
                           Dim value = source.Scalar(source.EvaluateExpression(FormulaParser.Parse(shifted), address))
                           If TypeOf value Is Boolean Then Return DirectCast(value, Boolean)
                           Return TypeOf value Is Double AndAlso DirectCast(value, Double) <> 0.0
                       End Function
            End Function

            Private Shared Function DatabaseFieldIndex(database As FormulaMatrix, input As Object, Optional headersOnly As Boolean = False) As Integer
                If input Is Nothing Then Return -1
                If TypeOf input Is String Then
                    Dim label = DirectCast(input, String)
                    If label.Length = 0 Then Return -1
                    For index = 0 To database.Columns - 1
                        Dim header = TryCast(database.At(0, index), String)
                        If String.Equals(label, header, StringComparison.OrdinalIgnoreCase) Then Return index
                    Next
                    Return -1
                End If
                If Not headersOnly AndAlso (TypeOf input Is Double OrElse TypeOf input Is Boolean) Then
                    Dim index = Math.Truncate(If(TypeOf input Is Boolean, If(DirectCast(input, Boolean), 1.0, 0.0), DirectCast(input, Double)))
                    If index >= 1.0 AndAlso index <= database.Columns Then Return CInt(index) - 1
                End If
                Return -1
            End Function

            Private Function MakeDatabaseCriterion(input As Object) As Predicate(Of Object)
                Dim criterion = TryCast(input, String)
                If Not String.IsNullOrEmpty(criterion) AndAlso "=<>".IndexOf(criterion(0)) < 0 Then
                    Dim number As Double
                    If Not Double.TryParse(criterion, NumberStyles.Float, ConversionCulture, number) Then
                        Dim pattern = BuildWildcard(criterion & "*")
                        Return Function(value) TypeOf value Is String AndAlso MatchesWildcard(pattern, DirectCast(value, String))
                    End If
                End If
                Return MakeCriterion(input)
            End Function

            Private Shared Function DatabaseAggregate(name As String, database As FormulaMatrix, field As Integer,
                                                      matches As IReadOnlyList(Of Integer)) As Object
                If name = "DGET" Then
                    If matches.Count = 0 Then Return New FormulaError("#VALUE!")
                    If matches.Count > 1 Then Return New FormulaError("#NUM!")
                    Return If(database.At(matches(0), field), New FormulaError("#VALUE!"))
                End If
                If name = "DCOUNT" Then Return CDbl(matches.Where(Function(row) TypeOf database.At(row, field) Is Double).Count())
                If name = "DCOUNTA" Then Return CDbl(matches.Where(Function(row) database.At(row, field) IsNot Nothing).Count())

                Dim numbers As New List(Of Double)
                For Each row In matches
                    Dim value = database.At(row, field)
                    If TypeOf value Is FormulaError Then Return value
                    If TypeOf value Is Double Then numbers.Add(DirectCast(value, Double))
                Next
                Select Case name
                    Case "DSUM" : Return FiniteNumber(numbers.Sum())
                    Case "DPRODUCT" : Return If(numbers.Count = 0, 0.0, FiniteNumber(numbers.Aggregate(1.0, Function(product, value) product * value)))
                    Case "DMAX" : Return If(numbers.Count = 0, 0.0, numbers.Max())
                    Case "DMIN" : Return If(numbers.Count = 0, 0.0, numbers.Min())
                    Case "DAVERAGE"
                        If numbers.Count = 0 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber(numbers.Average())
                    Case Else : Return EvaluateDeviation(name.Substring(1), numbers)
                End Select
            End Function

        End Class

    End Class

End Namespace
