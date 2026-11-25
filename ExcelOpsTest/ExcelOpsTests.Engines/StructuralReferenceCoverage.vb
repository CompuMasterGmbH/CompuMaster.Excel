Imports System.IO
Imports System.Text.RegularExpressions
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOpsTests.Engines

    Friend NotInheritable Class StructuralReferenceCoverage

        Friend Const FixtureName As String = "StructuralReferenceCoverage.xlsx"
        Private Const MergeSheet As String = "O'Brien Merge"
        Private Const MergeQualifier As String = "'O''Brien Merge'!"
        Private Shared ReadOnly LiteralOrSimpleSheet As New Regex(
            """(?:[^""]|"""")*""|'(?:[^']|'')+'!|(?<![\p{L}\p{M}\p{Nd}_.])(?:Data|PlainTable|Records)!",
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

        Public Shared ReadOnly Cases As String() = {
            "MergeRowBefore", "MergeRowInside", "MergeRowAfter",
            "MergeColumnBefore", "MergeColumnInside", "MergeColumnAfter",
            "MergeSplit", "TableRowBefore", "TableHeader", "TableFirstData", "TableBody",
            "TableTotals", "TableAfter", "TableDownBody", "TableDownTotals", "TableSplit", "PlainFirstData", "PlainBody", "TableWithoutTotalsAfter"
        }

        Friend Shared Function CompareWithExcel(workbook As ExcelDataOperationsBase, scenario As String, outputPath As String) As List(Of String)
            Dim differences As New List(Of String)
            Dim expectedMerge As String = "B3:D5"
            Dim expectedMergeName As String = "$B$3:$D$5"
            Dim expectedTable As String = "B3:D7"
            Dim expectedDataRows As String = "B3:D6"
            Dim expectedFirstData As Integer = 4
            Dim expectedLastData As Integer = 6
            Dim expectedLastDataCheck As String = "C5:C6"
            Dim expectedWideRange As String = "C4:E6"
            Dim expectedDirectRow As Integer = 4
            Dim expectedPlainLastData As Integer = 6
            Dim plainAbsoluteRow As Integer = 10

            Try
                Select Case scenario
                    Case "MergeRowBefore"
                        workbook.AddRow(MergeSheet, 1, 1, True)
                        expectedMerge = "B4:D6"
                        expectedMergeName = "$B$4:$D$6"
                    Case "MergeRowInside"
                        workbook.AddRow(MergeSheet, 3, 1, True)
                        expectedMerge = "B3:D6"
                        expectedMergeName = "$B$3:$D$6"
                    Case "MergeRowAfter"
                        workbook.AddRow(MergeSheet, 5, 1, True)
                    Case "MergeColumnBefore"
                        workbook.AddColumn(MergeSheet, 0, 1, True)
                        expectedMerge = "C3:E5"
                        expectedMergeName = "$C$3:$E$5"
                    Case "MergeColumnInside"
                        workbook.AddColumn(MergeSheet, 2, 1, True)
                        expectedMerge = "B3:E5"
                        expectedMergeName = "$B$3:$E$5"
                    Case "MergeColumnAfter"
                        workbook.AddColumn(MergeSheet, 4, 1, True)
                    Case "MergeSplit"
                        workbook.AddCells(MergeSheet, 3, 1, 1, 2, CellInsertShiftDirection.ShiftCellsDown, True)
                        expectedMerge = Nothing
                    Case "TableRowBefore", "TableHeader"
                        workbook.AddRow("Data", If(scenario = "TableRowBefore", 1, 2), 1, True)
                        expectedTable = "B4:D8"
                        expectedDataRows = "B4:D7"
                        expectedFirstData = 5
                        expectedLastData = 7
                        expectedLastDataCheck = "C6:C7"
                        expectedWideRange = "C5:E7"
                        expectedDirectRow = 5
                    Case "TableFirstData", "TableBody", "TableTotals", "TableDownBody", "TableDownTotals"
                        Select Case scenario
                            Case "TableFirstData"
                                workbook.AddRow("Data", 3, 1, True)
                                expectedLastDataCheck = "C6:C7"
                                expectedWideRange = "C5:E7"
                                expectedDirectRow = 5
                            Case "TableBody"
                                workbook.AddRow("Data", 4, 1, True)
                                expectedLastDataCheck = "C6:C7"
                                expectedWideRange = "C4:E7"
                            Case "TableTotals"
                                workbook.AddRow("Data", 6, 1, True)
                            Case "TableDownBody"
                                workbook.AddCells("Data", 4, 1, 1, 3, CellInsertShiftDirection.ShiftCellsDown, True)
                                expectedLastDataCheck = "C6:C7"
                            Case "TableDownTotals"
                                workbook.AddCells("Data", 6, 1, 1, 3, CellInsertShiftDirection.ShiftCellsDown, True)
                        End Select
                        expectedTable = "B3:D8"
                        expectedDataRows = "B3:D7"
                        expectedLastData = 7
                    Case "TableAfter"
                        workbook.AddRow("Data", 7, 1, True)
                    Case "TableSplit"
                        workbook.AddCells("Data", 4, 2, 1, 2, CellInsertShiftDirection.ShiftCellsDown, True)
                        differences.Add("Excel rejects a cell shift that changes only part of the table width; the engine accepted it.")
                        Return differences
                    Case "PlainFirstData", "PlainBody", "TableWithoutTotalsAfter"
                        Dim row = If(scenario = "PlainFirstData", 3, If(scenario = "PlainBody", 4, 6))
                        workbook.AddRow("PlainTable", row, 1, True)
                        If row < 6 Then expectedPlainLastData = 7
                        plainAbsoluteRow = 11
                    Case Else
                        Throw New ArgumentException("Unknown structural test case.", NameOf(scenario))
                End Select
            Catch ex As Exception When TypeOf ex Is NotSupportedException OrElse TypeOf ex Is InvalidOperationException OrElse
                                       TypeOf ex Is System.Runtime.InteropServices.COMException
                If scenario <> "TableSplit" Then differences.Add("The edit was rejected (" & ex.GetType().Name & "): " & ex.Message)
                Return differences
            End Try

            Try
                workbook.SaveCopyAs(outputPath, SaveOptionsForDisabledCalculationEngines.NoReset)
            Catch ex As Exception When TypeOf ex Is NullReferenceException OrElse TypeOf ex Is IOException OrElse
                                       TypeOf ex Is NotSupportedException OrElse TypeOf ex Is InvalidOperationException
                differences.Add("The edited workbook could not be saved or reopened (" & ex.GetType().Name & ").")
                Return differences
            End Try
            workbook.Close()
            Using document = SpreadsheetDocument.Open(outputPath, False)
                Dim part = document.WorkbookPart
                Dim merge = WorksheetFor(part, MergeSheet)
                Dim checks = WorksheetFor(part, "Checks")
                Dim sales = WorksheetFor(part, "Data")
                Dim table = sales.TableDefinitionParts.SingleOrDefault()?.Table
                Dim plainSheet = WorksheetFor(part, "PlainTable")
                Dim plain = plainSheet.TableDefinitionParts.SingleOrDefault()?.Table
                Dim names = part.Workbook.DefinedNames?.Elements(Of DefinedName)().ToArray()

                Compare(differences, "merge", merge.Worksheet.GetFirstChild(Of MergeCells)()?.Elements(Of MergeCell)().SingleOrDefault()?.Reference?.Value, expectedMerge)
                Compare(differences, "workbook merge name", NameText(names, "MergedArea"), MergeQualifier & expectedMergeName)
                Dim mergeRef = If(expectedMerge, "B3:D5")
                Compare(differences, "direct merge formula", FormulaAt(checks, "A10"), "SUM(" & MergeQualifier & mergeRef & ")")
                Compare(differences, "formula using merge name", FormulaAt(checks, "A11"), "SUM(MergedArea)")
                Compare(differences, "Sales table", table?.Reference?.Value, expectedTable)
                Compare(differences, "Sales filter", table?.GetFirstChild(Of AutoFilter)()?.Reference?.Value, expectedDataRows)
                Compare(differences, "plain table", plain?.Reference?.Value, "B3:D" & expectedPlainLastData)
                Compare(differences, "plain filter", plain?.GetFirstChild(Of AutoFilter)()?.Reference?.Value, "B3:D" & expectedPlainLastData)
                For row = 4 To expectedPlainLastData
                    Dim absoluteRow = If(row = 4 AndAlso scenario = "PlainFirstData", 10, plainAbsoluteRow)
                    Compare(differences, "A1 calculated column D" & row, FormulaAt(plainSheet, "D" & row), "C" & row & "*$A$" & absoluteRow)
                Next
                Dim template = plain?.GetFirstChild(Of TableColumns)()?.Elements(Of TableColumn)().LastOrDefault()?.CalculatedColumnFormula?.Text
                Compare(differences, "A1 calculated column template", template, "C4*$A$" & If(scenario = "PlainFirstData", 10, plainAbsoluteRow))
                Compare(differences, "plain data range", NameText(names, "PlainCells"), "PlainTable!$B$4:$D$" & expectedPlainLastData)
                Compare(differences, "structured formula", FormulaAt(checks, "A1"), "SUM(Sales[Total])")
                Compare(differences, "structured totals formula", FormulaAt(checks, "A14"), "SUM(Sales[[#Totals],[Qty]])")
                Compare(differences, "direct data range", FormulaAt(checks, "A2"), "SUM(Data!C" & expectedFirstData & ":C" & expectedLastData & ")")
                Compare(differences, "partial data range", FormulaAt(checks, "A6"), "SUM(Data!" & expectedLastDataCheck & ")")
                Compare(differences, "range wider than the table", FormulaAt(checks, "A7"), "SUM(Data!" & expectedWideRange & ")")
                Compare(differences, "workbook data name", NameText(names, "SalesCells"), "Data!$B$" & expectedFirstData & ":$D$" & expectedLastData)
                Compare(differences, "local data name", NameText(names, "LocalSales"), "Data!$C$" & expectedFirstData & ":$C$" & expectedLastData)
                Compare(differences, "direct absolute R1C1 in Excel", FormulaAt(checks, "A5"), "Data!$C$" & expectedDirectRow)
                Compare(differences, "direct mixed R1C1 in Excel", FormulaAt(checks, "A8"), "Data!$C" & expectedDirectRow)
                Compare(differences, "R1C1 workbook setting", part.Workbook.CalculationProperties?.GetAttributes().
                        FirstOrDefault(Function(attribute) attribute.LocalName = "refMode").Value, "R1C1")
                Compare(differences, "indirect A1 address", FormulaAt(checks, "A3"), "INDIRECT(""Data!C4"")")
                Compare(differences, "indirect R1C1 address", FormulaAt(checks, "A4"), "INDIRECT(""Data!R4C3"",FALSE)")
                Compare(differences, "direct operand inside INDIRECT", FormulaAt(checks, "A9"), "INDIRECT(""Data!C""&ROW(Data!C" & expectedDirectRow & "))")
                For row = expectedFirstData To expectedLastData
                    Dim formula = FormulaAt(sales, "D" & row)
                    If formula IsNot Nothing Then formula = formula.Replace("Sales[[#This Row],[Qty]]", "[@Qty]")
                    Compare(differences, "calculated column D" & row, formula, "[@Qty]*2")
                Next
                Dim totalsFormula = FormulaAt(sales, "D" & (expectedLastData + 1))
                If totalsFormula = "SUM([Total])" Then totalsFormula = "SUM(Sales[Total])"
                Compare(differences, "totals row", totalsFormula, "SUM(Sales[Total])")
            End Using
            Return differences
        End Function

        Friend Shared Sub Compare(differences As List(Of String), label As String, actual As String, expected As String)
            If String.Equals(NormalizeSimpleSheetQualifiers(actual), NormalizeSimpleSheetQualifiers(expected), StringComparison.Ordinal) Then Return
            differences.Add(label & " expected [" & If(expected, "none") & "], got [" & If(actual, "none") & "].")
        End Sub

        Friend Shared Function WorksheetFor(part As WorkbookPart, name As String) As WorksheetPart
            Dim sheet = part.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = name)
            Return DirectCast(part.GetPartById(sheet.Id.Value), WorksheetPart)
        End Function

        Friend Shared Function NameText(names As DefinedName(), name As String) As String
            Return names?.SingleOrDefault(Function(item) item.Name.Value = name)?.Text
        End Function

        Private Shared Function NormalizeSimpleSheetQualifiers(value As String) As String
            If value Is Nothing Then Return Nothing
            ' Excel ignores case and optional quotes in these qualifiers. Leave formula strings intact.
            Return LiteralOrSimpleSheet.Replace(value,
                Function(token)
                    If token.Value(0) = """"c Then Return token.Value
                    Dim name = token.Value.TrimEnd("!"c).Trim("'"c)
                    If String.Equals(name, "Data", StringComparison.OrdinalIgnoreCase) Then Return "Data!"
                    If String.Equals(name, "PlainTable", StringComparison.OrdinalIgnoreCase) Then Return "PlainTable!"
                    If String.Equals(name, "Records", StringComparison.OrdinalIgnoreCase) Then Return "Records!"
                    Return token.Value
                End Function)
        End Function

        Friend Shared Function FormulaAt(sheet As WorksheetPart, address As String) As String
            Dim data = sheet.Worksheet.GetFirstChild(Of SheetData)()
            Dim formula = data.Descendants(Of Cell)().FirstOrDefault(Function(cell) cell.CellReference?.Value = address)?.CellFormula
            If formula?.SharedIndex IsNot Nothing AndAlso String.IsNullOrEmpty(formula.Text) Then
                Dim master = data.Descendants(Of CellFormula)().FirstOrDefault(
                    Function(other) other.SharedIndex IsNot Nothing AndAlso other.SharedIndex.Value = formula.SharedIndex.Value AndAlso Not String.IsNullOrEmpty(other.Text))
                If master Is Nothing Then Return Nothing
                ' This fixture shares vertical calculated columns. Only its plain table has a relative row to translate.
                Dim masterCell = DirectCast(master.Parent, Cell)
                Dim offset = New ExcelCell(Nothing, address, ExcelCell.ValueTypes.All).RowIndex -
                    New ExcelCell(Nothing, masterCell.CellReference.Value, ExcelCell.ValueTypes.All).RowIndex
                Return Regex.Replace(master.Text, "\AC([0-9]+)(\*\$A\$[0-9]+)\z",
                    Function(match) "C" & (Integer.Parse(match.Groups(1).Value, Globalization.CultureInfo.InvariantCulture) + offset).
                        ToString(Globalization.CultureInfo.InvariantCulture) & match.Groups(2).Value)
            End If
            Return formula?.Text
        End Function

    End Class

End Namespace
