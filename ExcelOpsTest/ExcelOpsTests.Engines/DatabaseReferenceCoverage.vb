Imports System.Globalization
Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOpsTests.Engines

    Friend NotInheritable Class DatabaseReferenceCoverage

        Private Const RecordsSheet As String = "Records"

        Friend Shared Sub CreateFixture(path As String, includeCachedResults As Boolean)
            Using document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)
                Dim part = document.AddWorkbookPart()
                Dim records = part.AddNewPart(Of WorksheetPart)()
                Dim checks = part.AddNewPart(Of WorksheetPart)()
                records.Worksheet = New Worksheet(New SheetData(
                    RowAt(3, 1, "Name", "Qty", "Price"),
                    RowAt(4, 1, "Alpha", 10, 5), RowAt(5, 1, "Beta", 20, 6),
                    RowAt(6, 1, "Gamma", 30, 7), RowAt(7, 1, "Alpha", 40, 8),
                    RowAt(10, 1, "Name", Nothing, Nothing, "Name"),
                    RowAt(11, 1, "Alpha", Nothing, Nothing, "Beta")))
                Dim formulas = DatabaseFormulas("D", 7, "E", 10)
                Dim excelResults = {6, 6, 6, 50, 2, 25, 8, 5, 400, 2, 225, 15, 13, 13, 50}
                checks.Worksheet = New Worksheet(New SheetData(formulas.Select(
                    Function(formula, row) New Row(New Cell With {
                        .CellReference = ExcelCell.LocalCellAddress(row, 0),
                        .CellFormula = New CellFormula(formula),
                        .CellValue = If(includeCachedResults, New CellValue(excelResults(row)), Nothing)
                    }) With {.RowIndex = CUInt(row + 1)})))
                part.Workbook = New Workbook(
                    New Sheets(New Sheet With {.Id = part.GetIdOfPart(records), .SheetId = 1UI, .Name = RecordsSheet},
                               New Sheet With {.Id = part.GetIdOfPart(checks), .SheetId = 2UI, .Name = "Checks"}),
                    New DefinedNames(New DefinedName("Records!$B$3:$D$7") With {.Name = "LookupDatabase"},
                                     New DefinedName("Records!$B$10:$B$11") With {.Name = "AlphaCriteria"}))
            End Using
        End Sub

        Friend Shared Function CompareWithExcel(workbook As ExcelDataOperationsBase, output As String, insertColumn As Boolean) As List(Of String)
            Dim differences As New List(Of String)
            Try
                workbook.AddRow(RecordsSheet, 4, 1, True)
                If insertColumn Then workbook.AddColumn(RecordsSheet, 3, 1, True)
            Catch ex As NotSupportedException
                differences.Add("The edit was rejected: " & ex.Message)
                Return differences
            End Try

            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
            workbook.Close()
            Using document = SpreadsheetDocument.Open(output, False)
                Dim part = document.WorkbookPart
                Dim checks = StructuralReferenceCoverage.WorksheetFor(part, "Checks")
                Dim expected = DatabaseFormulas(If(insertColumn, "E", "D"), 8, If(insertColumn, "F", "E"), 11)
                For row = 0 To expected.Length - 1
                    Dim address = ExcelCell.LocalCellAddress(row, 0)
                    StructuralReferenceCoverage.Compare(differences, "Checks!" & address,
                        StructuralReferenceCoverage.FormulaAt(checks, address), expected(row))
                Next

                Dim names = part.Workbook.DefinedNames?.Elements(Of DefinedName)().ToArray()
                StructuralReferenceCoverage.Compare(differences, "database name",
                    StructuralReferenceCoverage.NameText(names, "LookupDatabase"), "Records!$B$3:$" & If(insertColumn, "E", "D") & "$8")
                StructuralReferenceCoverage.Compare(differences, "criteria name",
                    StructuralReferenceCoverage.NameText(names, "AlphaCriteria"), "Records!$B$11:$B$12")
            End Using
            Return differences
        End Function

        Private Shared Function RowAt(row As Integer, firstColumn As Integer, ParamArray values As Object()) As Row
            Dim result As New Row With {.RowIndex = CUInt(row)}
            For index = 0 To values.Length - 1
                Dim value = values(index)
                If value Is Nothing Then Continue For
                Dim cell As New Cell With {.CellReference = ExcelCell.LocalCellAddress(row - 1, firstColumn + index)}
                If TypeOf value Is String Then
                    cell.DataType = CellValues.InlineString
                    cell.InlineString = New InlineString(New Text(CStr(value)))
                Else
                    cell.CellValue = New CellValue(CInt(value))
                End If
                result.Append(cell)
            Next
            Return result
        End Function

        Private Shared Function DatabaseFormulas(fieldColumn As String, lastDataRow As Integer, betaColumn As String, firstCriterionRow As Integer) As String()
            Dim last = lastDataRow.ToString(CultureInfo.InvariantCulture)
            Dim firstCriterion = firstCriterionRow.ToString(CultureInfo.InvariantCulture)
            Dim lastCriterion = (firstCriterionRow + 1).ToString(CultureInfo.InvariantCulture)
            Dim database = "Records!$B$3:$" & fieldColumn & "$" & last
            Dim alpha = "Records!$B$" & firstCriterion & ":$B$" & lastCriterion
            Dim beta = "Records!$" & betaColumn & "$" & firstCriterion & ":$" & betaColumn & "$" & lastCriterion
            Return {
                "VLOOKUP(""Beta""," & database & ",3,FALSE)",
                "INDEX(Records!$" & fieldColumn & "$4:$" & fieldColumn & "$" & last & ",MATCH(""Beta"",Records!$B$4:$B$" & last & ",0))",
                "DGET(" & database & ",""Price""," & beta & ")",
                "DSUM(" & database & ",""Qty""," & alpha & ")",
                "DCOUNT(" & database & ",""Qty""," & alpha & ")",
                "DAVERAGE(" & database & ",""Qty""," & alpha & ")",
                "DMAX(" & database & ",""Price""," & alpha & ")",
                "DMIN(" & database & ",""Price""," & alpha & ")",
                "DPRODUCT(" & database & ",""Qty""," & alpha & ")",
                "DCOUNTA(" & database & ",""Name""," & alpha & ")",
                "DVARP(" & database & ",""Qty""," & alpha & ")",
                "DSTDEVP(" & database & ",""Qty""," & alpha & ")",
                "DSUM(" & database & ",Records!$" & fieldColumn & "$3," & alpha & ")",
                "DSUM(" & database & ",3," & alpha & ")",
                "DSUM(LookupDatabase,""Qty"",AlphaCriteria)"
            }
        End Function

    End Class

End Namespace
