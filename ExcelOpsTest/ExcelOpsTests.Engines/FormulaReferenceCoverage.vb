Imports System.Globalization
Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

Namespace ExcelOpsTests.Engines

    Friend NotInheritable Class FormulaReferenceCoverage

        Friend Shared Function Report(engine As String, openWorkbook As Func(Of String, ExcelDataOperationsOptions, ExcelDataOperationsBase),
                                      readResult As Func(Of ExcelDataOperationsBase, ExcelCell, Object)) As (Matched As Integer, Total As Integer)
            Dim source = TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", "FormulaReference.xlsx")
            Dim folder = Path.GetDirectoryName(source)
            Dim temporary = Path.Combine(Path.GetTempPath(), "cm-excel-" & Guid.NewGuid().ToString("N"))
            Dim companion = "FormulaReferenceSource.xlsx"
            Directory.CreateDirectory(temporary)
            Dim previousCulture = CultureInfo.CurrentCulture
            Dim referenceCulture = CultureInfo.GetCultureInfo("de-DE")
            Dim matched As Integer
            Dim total As Integer
            Try
                CultureInfo.CurrentCulture = referenceCulture
                File.Copy(Path.Combine(folder, companion), Path.Combine(temporary, companion))
                For Each sourceFile In Directory.EnumerateFiles(folder, "*.xlsx").OrderBy(Function(item) item, StringComparer.Ordinal)
                    If Path.GetFileName(sourceFile).StartsWith("FormulaReferenceIteration", StringComparison.Ordinal) Then Continue For
                    Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly) With {
                        .DisableInitialCalculation = True, .DisableCalculationEngine = False, .FormulaCulture = referenceCulture
                    }
                    Dim working = Path.Combine(temporary, Path.GetFileName(sourceFile))
                    Try
                        File.Copy(sourceFile, working, overwrite:=True)
                        Dim result = Compare(sourceFile, working, Function(path) openWorkbook(path, options), readResult)
                        matched += result.Matched
                        total += result.Total
                        TestContext.Out.WriteLine(Path.GetFileName(sourceFile) & " " & result.Matched & "/" & result.Total)
                    Finally
                        File.Delete(working)
                    End Try
                Next
            Finally
                CultureInfo.CurrentCulture = previousCulture
                File.Delete(Path.Combine(temporary, companion))
                Directory.Delete(temporary)
            End Try
            Assert.That(total, [Is].GreaterThan(0))
            TestContext.Out.WriteLine("TOTAL " & engine & ": " & matched & "/" & total)
            Return (matched, total)
        End Function

        Private Shared Function Compare(source As String, working As String, openWorkbook As Func(Of String, ExcelDataOperationsBase),
                                        readResult As Func(Of ExcelDataOperationsBase, ExcelCell, Object)) As (Matched As Integer, Total As Integer)
            Dim cases As New List(Of (Address As ExcelCell, Expected As Object, Kind As CellValues))
            Dim readOnlyOptions As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly) With {
                .DisableInitialCalculation = True, .DisableCalculationEngine = True
            }
            Using reference As New OpenXmlExcelDataOperations(source, ExcelDataOperationsBase.OpenMode.OpenExistingFile, readOnlyOptions),
                  package = SpreadsheetDocument.Open(working, True)
                Dim part = package.WorkbookPart
                For Each sheet In part.Workbook.Sheets.Elements(Of Sheet)()
                    Dim name = sheet.Name.Value
                    If name = "External cases" Then Continue For
                    Dim worksheet = DirectCast(part.GetPartById(sheet.Id.Value), WorksheetPart).Worksheet
                    Dim cells = worksheet.Descendants(Of Cell)().ToArray()
                    Dim spillRanges = cells.Where(Function(cell) cell.CellMetaIndex IsNot Nothing AndAlso
                        If(cell.CellFormula?.FormulaType?.Value, CellFormulaValues.Normal) = CellFormulaValues.Array AndAlso cell.CellFormula.Reference IsNot Nothing).
                        Select(Function(cell) New ExcelRange(name, cell.CellFormula.Reference.Value)).ToArray()
                    For Each cell In cells
                        Dim formula = cell.CellFormula
                        If formula Is Nothing OrElse (String.IsNullOrEmpty(formula.Text) AndAlso
                            If(formula.CalculateCell?.Value, False) AndAlso If(formula.FormulaType?.Value, CellFormulaValues.Normal) = CellFormulaValues.Normal) Then
                            If spillRanges.Length = 0 Then Continue For
                            Dim position As New ExcelCell(name, cell.CellReference.Value, ExcelCell.ValueTypes.All)
                            If Not spillRanges.Any(Function(range) position.RowIndex >= range.AddressStart.RowIndex AndAlso
                                position.RowIndex <= range.AddressEnd.RowIndex AndAlso position.ColumnIndex >= range.AddressStart.ColumnIndex AndAlso
                                position.ColumnIndex <= range.AddressEnd.ColumnIndex) Then Continue For
                            cell.CellValue = Nothing
                            cell.DataType = Nothing
                            cell.ValueMetaIndex = Nothing
                            Continue For
                        End If
                        Dim address As New ExcelCell(name, cell.CellReference.Value, ExcelCell.ValueTypes.All)
                        Assert.That(cell.CellValue, [Is].Not.Null, address.Address(True))
                        Dim value = reference.LookupCellValueAsObject(address, FormulaCalculationMode.UseCachedValueOnly)
                        cases.Add((address, value, If(cell.DataType?.Value, CellValues.Number)))
                        cell.CellValue = Nothing
                        cell.ValueMetaIndex = Nothing
                        cell.DataType = Nothing
                    Next
                Next
            End Using

            Dim matched As Integer
            Dim workbook = openWorkbook(working)
            Try
                If TypeOf workbook Is MsExcelDataOperations Then workbook.RecalculateAll()
                For Each sheet In cases.GroupBy(Function(item) item.Address.SheetName)
                    Try
                        workbook.RecalculateSheet(sheet.Key)
                    Catch ex As Exception
                        TestContext.Out.WriteLine(IO.Path.GetFileName(source) & " " & sheet.Key & " recalculation: " & ex.GetType().Name)
                    End Try

                    For Each item In sheet
                        Dim actual As Object
                        Try
                            actual = readResult(workbook, item.Address)
                        Catch ex As Exception
                            actual = "[" & ex.GetType().Name & "]"
                        End Try
                        If Equivalent(item.Expected, actual, item.Kind) Then
                            matched += 1
                        Else
                            TestContext.Out.WriteLine(IO.Path.GetFileName(source) & " " & item.Address.Address(True) &
                                                      " expected " & Display(item.Expected) & ", actual " & Display(actual))
                        End If
                    Next
                Next
            Finally
                If Not workbook.IsClosed Then workbook.Close()
            End Try
            Return (matched, cases.Count)
        End Function

        Private Shared Function Equivalent(expected As Object, actual As Object, kind As CellValues) As Boolean
            If kind = CellValues.Error Then
                Return actual IsNot Nothing AndAlso String.Equals(Convert.ToString(expected, CultureInfo.InvariantCulture),
                                                                  Convert.ToString(actual, CultureInfo.InvariantCulture), StringComparison.Ordinal)
            End If
            If TypeOf expected Is Double AndAlso (TypeOf actual Is Double OrElse TypeOf actual Is Integer OrElse
               TypeOf actual Is Long OrElse TypeOf actual Is Single OrElse TypeOf actual Is Decimal) Then
                Dim number = Convert.ToDouble(actual, CultureInfo.InvariantCulture)
                Return Math.Abs(DirectCast(expected, Double) - number) <= 1.0E-12
            End If
            Return Object.Equals(expected, actual)
        End Function

        Private Shared Function Display(value As Object) As String
            Dim text = If(value Is Nothing, "<null>", Convert.ToString(value, CultureInfo.InvariantCulture))
            Return If(text.Length <= 120, text, text.Substring(0, 120) & "...")
        End Function

    End Class

End Namespace
