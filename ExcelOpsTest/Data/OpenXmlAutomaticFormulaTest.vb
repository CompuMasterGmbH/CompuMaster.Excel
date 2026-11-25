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
