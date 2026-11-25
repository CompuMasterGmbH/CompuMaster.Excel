Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlAutomaticFormulaTest

    <Test>
    Public Sub UpdatesFiveStagesAndVolatileValuesButLeavesUnrelatedFunctionsAlone()
        Dim session As New ExcelFormulaSession()
        Dim independentCalls = 0
        session.RegisterFunction("APP.CHECK",
            Function(context, arguments)
                independentCalls += 1
                Return ExcelFormulaValue.FromNumber(independentCalls)
            End Function)
        Dim current As New DateTime(2011, 7, 5, 12, 0, 0)
        Dim options = MakeOptions(session)
        options.FormulaLocalClock = Function() current
        options.FormulaWorkbookName = "Local"

        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, options)
            workbook.WriteCellFormula("Sheet1", 0, 1, "A1+1", False)
            workbook.WriteCellFormula("Sheet1", 0, 2, "B1+1", False)
            workbook.WriteCellFormula("Sheet1", 0, 3, "C1+1", False)
            workbook.WriteCellFormula("Sheet1", 0, 4, "D1+1", False)
            workbook.WriteCellFormula("Sheet1", 0, 5, "E1+NOW()", False)
            workbook.WriteCellFormula("Sheet1", 0, 7, "APP.CHECK()", False)
            workbook.WriteCellFormula("Sheet1", 0, 8, "NOW()-TODAY()", False)
            workbook.WriteCellValue("Sheet1", 0, 0, 1.0)
            Assert.That(Result(workbook, 5), [Is].EqualTo(40734.5))
            Assert.That(Result(workbook, 8), [Is].EqualTo(0.5))
            Assert.That(independentCalls, [Is].EqualTo(1))

            current = current.AddHours(6)
            workbook.WriteCellValue("Sheet1", 0, 0, 2.0)
            Assert.That(Result(workbook, 5), [Is].EqualTo(40735.75))
            Assert.That(Result(workbook, 8), [Is].EqualTo(0.75))
            Assert.That(independentCalls, [Is].EqualTo(1))

            workbook.AutoCalculationEnabledWorkbookSetting = False
            workbook.WriteCellValue("Sheet1", 0, 0, 4.0)
            Assert.That(Result(workbook, 5), [Is].EqualTo(40735.75))
            workbook.RecalculateCell("Sheet1", 0, 5)
            Assert.That(Result(workbook, 5), [Is].EqualTo(40737.75))
            Assert.That(independentCalls, [Is].EqualTo(1))

            workbook.AutoCalculationEnabledWorkbookSetting = True
            workbook.WriteCellValue("Sheet1", 0, 0, 6.0)
            Assert.That(Result(workbook, 5), [Is].EqualTo(40739.75))
        End Using
    End Sub

    ' The normalized results were verified with Excel's real clock and calculation modes.
    <Test>
    Public Sub MatchesExcelAcrossFiveWorksheetDependenciesInBothCalculationModes()
        Dim path = TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", "FormulaReference1900.xlsx")
        Dim options = MakeOptions()
        Dim clock As New DateTime(2026, 11, 25, 11, 0, 0)
        options.FormulaLocalClock = Function() clock
        options.FormulaCulture = System.Globalization.CultureInfo.GetCultureInfo("de-DE")
        Const row As Integer = 49
        Using workbook As New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, options)
            workbook.WriteCellFormula("Standard", row, 1, "A50+NOW()", False)
            workbook.WriteCellFormula("Dates", row, 0, "Standard!B50+1", False)
            For column = 1 To 4
                Dim precedingAddress = ExcelCell.LocalCellAddress(row, column - 1)
                workbook.WriteCellFormula("Dates", row, column, precedingAddress & "+1", False)
            Next
            workbook.WriteCellFormula("Dates", row, 5, "NOW()", False)

            Dim output As New ExcelCell("Dates", row, 4, ExcelCell.ValueTypes.All)
            Dim observedTime As New ExcelCell("Dates", row, 5, ExcelCell.ValueTypes.All)
            Dim value As Func(Of ExcelCell, Double) =
                Function(cell) workbook.LookupCellValue(Of Double)(cell, FormulaCalculationMode.UseCachedValueOnly)

            workbook.WriteCellValue("Standard", row, 0, 1.0)
            Assert.That(value(output) - value(observedTime), [Is].EqualTo(6.0))

            clock = clock.AddHours(1)
            workbook.WriteCellValue("Standard", row, 0, 2.0)
            Assert.That(value(output) - value(observedTime), [Is].EqualTo(7.0))
            Dim previous = value(output)

            workbook.AutoCalculationEnabledWorkbookSetting = False
            clock = clock.AddHours(1)
            workbook.WriteCellValue("Standard", row, 0, 4.0)
            Assert.That(value(output), [Is].EqualTo(previous))

            workbook.RecalculateAll()
            Assert.That(value(output) - value(observedTime), [Is].EqualTo(9.0))
        End Using
    End Sub

    <Test>
    Public Sub TracksCellsThatWereEmptyWhenAnAggregateWasLastEvaluated()
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, MakeOptions())
            workbook.WriteCellFormula("Sheet1", 0, 3, "SUM(A1:A3)", False)
            workbook.WriteCellValue("Sheet1", 0, 0, 1)
            Assert.That(Result(workbook, 3), [Is].EqualTo(1.0))
            workbook.WriteCellValue("Sheet1", 1, 0, 2)
            Assert.That(Result(workbook, 3), [Is].EqualTo(3.0))
        End Using
    End Sub

    <Test>
    Public Sub ReusesTheSessionDependencyGraphWhenExternalInputsChange()
        Dim session As New ExcelFormulaSession()
        Dim targetPath = TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", "FormulaReference.xlsx")
        Dim sourcePath = TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", "FormulaReferenceSource.xlsx")
        Dim sourceOptions = MakeOptions(session)
        sourceOptions.CalculateOnCellEdit = False
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, sourceOptions),
              target As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, MakeOptions(session))
            source.WriteCellValue("Source Data", 1, 0, 45.0)
            Assert.That(ExternalValues(target), [Is].EqualTo({45.0, 106.0, 49.0}))
            source.WriteCellValue("Source Data", 1, 0, 47.0)
            Assert.That(ExternalValues(target), [Is].EqualTo({47.0, 110.0, 51.0}))
        End Using
    End Sub

    <Test>
    Public Sub RebuildsDependenciesAfterAReferencedWorksheetIsRenamedAndShifted()
        Dim formulas As New ExcelFormulaSession()
        Dim targetPath = TestEnvironment.FullPathOfExistingTestFile("test_data", "WorkbookReferenceTarget.xlsx")
        Dim sourcePath = TestEnvironment.FullPathOfExistingTestFile("test_data", "WorkbookReferenceSource.xlsx")
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, MakeOptions(formulas)),
              target As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, MakeOptions(formulas)),
              references As New OpenXmlWorkbookSession()
            references.Attach(source)
            references.Attach(target)
            Dim result As Func(Of Double) =
                Function() target.LookupCellValue(Of Double)("Checks", 0, 0, FormulaCalculationMode.UseCachedValueOnly)

            source.WriteCellValue("Remote Data", 1, 1, 45.0)
            Assert.That(result(), [Is].EqualTo(45.0))

            source.RenameWorksheet("Remote Data", "Renamed source", True)
            source.AddRow("Renamed source", 1, 1, True)
            source.WriteCellValue("Renamed source", 2, 1, 47.0)
            Assert.That(result(), [Is].EqualTo(47.0))

            source.WriteCellValue("Renamed source", 1, 1, 99.0)
            Assert.That(result(), [Is].EqualTo(47.0))
        End Using
    End Sub

    <Test>
    Public Sub SerializesConcurrentEditsWithinTheSameFormulaSession()
        Dim session As New ExcelFormulaSession()
        Dim targetPath = TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", "FormulaReference.xlsx")
        Dim sourcePath = TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", "FormulaReferenceSource.xlsx")
        Using source As New OpenXmlExcelDataOperations(sourcePath, OpenMode.OpenExistingFile, MakeOptions(session)),
              target As New OpenXmlExcelDataOperations(targetPath, OpenMode.OpenExistingFile, MakeOptions(session))
            target.WriteCellFormula("External cases", 0, 7, "Cases!A50+C2", False)
            System.Threading.Tasks.Parallel.Invoke(
                Sub()
                    For value = 41 To 50
                        source.WriteCellValue("Source Data", 1, 0, value)
                    Next
                End Sub,
                Sub()
                    For value = 1 To 10
                        target.WriteCellValue("Cases", 49, 0, value)
                    Next
                End Sub)

            Assert.That(ExternalValues(target), [Is].EqualTo({50.0, 116.0, 54.0}))
            Assert.That(target.LookupCellValue(Of Double)("External cases", 0, 7, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(60.0))
        End Using
    End Sub

    <Test>
    Public Sub KeepsPreviousFormulaCachesWhenAnAutomaticPassFails()
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, MakeOptions())
            workbook.WriteCellFormula("Sheet1", 0, 1, "IF(A1>0,B1,A1)", False)
            workbook.WriteCellValue("Sheet1", 0, 0, 0)
            Assert.That(Result(workbook, 1), [Is].EqualTo(0.0))
            Assert.Throws(Of CircularReferenceException)(Sub() workbook.WriteCellValue("Sheet1", 0, 0, 1))
            Assert.That(Result(workbook, 0), [Is].EqualTo(1.0))
            Assert.That(Result(workbook, 1), [Is].EqualTo(0.0))
            workbook.WriteCellValue("Sheet1", 0, 0, -2)
            Assert.That(Result(workbook, 1), [Is].EqualTo(-2.0))
        End Using
    End Sub

    Private Shared Function MakeOptions(Optional session As ExcelFormulaSession = Nothing) As ExcelDataOperationsOptions
        Return New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .CalculateOnCellEdit = True, .FormulaSession = session
        }
    End Function

    Private Shared Function Result(workbook As OpenXmlExcelDataOperations, column As Integer) As Double
        Return workbook.LookupCellValue(Of Double)("Sheet1", 0, column, FormulaCalculationMode.UseCachedValueOnly)
    End Function

    Private Shared Function ExternalValues(workbook As OpenXmlExcelDataOperations) As Double()
        Return Enumerable.Range(1, 3).Select(Function(row) workbook.LookupCellValue(Of Double)(
            "External cases", row, 2, FormulaCalculationMode.UseCachedValueOnly)).ToArray()
    End Function

End Class
