Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlIterativeFormulaTest

    ' The cycle starts at the value saved by Excel. The second cell participates in its stopping condition.
    <TestCase("Steps", True)>
    <TestCase("Convergence", False)>
    Public Sub MatchesSuccessiveExcelPassesAndAnInputChange(caseName As String, automatic As Boolean)
        Using workbook = OpenReference(caseName)
            workbook.AutoCalculationEnabledWorkbookSetting = False
            AssertReferencePass(workbook, 0)
            For pass = 1 To 3
                workbook.RecalculateAll()
                AssertReferencePass(workbook, pass)
            Next
            Assert.That(Cached(workbook, "Cycle", 0, 2), [Is].EqualTo(1.0))

            workbook.AutoCalculationEnabledWorkbookSetting = automatic
            workbook.WriteCellValue("Cycle", 0, 3, 16.0)
            If Not automatic Then
                AssertReferencePass(workbook, 3)
                workbook.RecalculateAll()
            End If
            AssertReferencePass(workbook, 4)
        End Using
    End Sub

    <Test>
    Public Sub UsesZeroForAMissingCircularCacheAndPreservesTheNextSeedOnSave()
        Dim output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Try
            Using workbook = OpenReference("Steps")
                workbook.ResetCellValueFromFormulaCell("Cycle", 0, 0)
                workbook.ResetCellValueFromFormulaCell("Cycle", 0, 1)
                workbook.RecalculateCell("Cycle", 0, 1)
                AssertReferencePass(workbook, 0)
                workbook.RecalculateCell("Cycle", 0, 1)
                AssertReferencePass(workbook, 1)
                workbook.SaveCopyAs(output, recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using

            Using reopened As New OpenXmlExcelDataOperations(File.ReadAllBytes(output), CalculationOptions())
                AssertReferencePass(reopened, 1)
                reopened.RecalculateCell("Cycle", 0, 1)
                AssertReferencePass(reopened, 2)
            End Using
        Finally
            If File.Exists(output) Then File.Delete(output)
        End Try
    End Sub

    <Test>
    Public Sub AdvancesAnExcelCycleAcrossSheetsInNameOrder()
        Using workbook = OpenReference("Cycles")
            For pass = 0 To 5
                If pass > 0 Then workbook.RecalculateCell("Z cycle", 0, 1)
                Assert.That(Cached(workbook, "Z cycle", 0, 0),
                    [Is].EqualTo(Cached(workbook, "Reference", pass + 1, 4)).Within(1.0E-12), "Z cycle on pass " & pass)
                Assert.That(Cached(workbook, "A cycle", 0, 0),
                    [Is].EqualTo(Cached(workbook, "Reference", pass + 1, 5)).Within(1.0E-12), "A cycle on pass " & pass)
                Assert.That(Cached(workbook, "Z cycle", 0, 1),
                    [Is].EqualTo(Cached(workbook, "Reference", pass + 1, 6)).Within(1.0E-12), "Dependent on pass " & pass)
            Next
        End Using
    End Sub

    <TestCase(False)>
    <TestCase(True)>
    Public Sub AdvancesTheExcelTwoCellCycleFromItsSavedOrZeroSeeds(clearSeeds As Boolean)
        Using workbook = OpenReference("Cycles")
            If clearSeeds Then
                workbook.ResetCellValueFromFormulaCell("Cycle")
                workbook.RecalculateAll()
            End If

            For pass = 0 To 5
                If pass > 0 Then
                    ' Refreshing the dependent alone must use the same circular sweep as refreshing the full sheet.
                    If pass Mod 2 = 0 Then
                        workbook.RecalculateSheet("Cycle")
                    Else
                        workbook.RecalculateCell("Cycle", 0, 2)
                    End If
                End If
                For column = 0 To 2
                    Assert.That(Cached(workbook, "Cycle", 0, column),
                        [Is].EqualTo(Cached(workbook, "Reference", pass + 1, column + 1)).Within(1.0E-12),
                        "Cell " & ExcelCell.LocalCellAddress(0, column) & " on pass " & pass)
                Next
            Next
        End Using
    End Sub

    <TestCase(False)>
    <TestCase(True)>
    Public Sub MatchesExcelPassesForTwoOpenCircularWorkbooks(firstSortsBeforeSecond As Boolean)
        Dim session As New ExcelFormulaSession()
        Dim firstCase = If(firstSortsBeforeSecond, "CrossAFirst", "CrossZFirst")
        Dim secondCase = If(firstSortsBeforeSecond, "CrossZSecond", "CrossASecond")
        ' Measured on the Excel reference build, with one iteration per full calculation.
        Dim expectedFirst = {0.5, 1.125, 1.28125, 1.3203125, 1.330078125, 1.33251953125}
        Dim expectedSecond = If(firstSortsBeforeSecond,
            {1.25, 1.5625, 1.640625, 1.66015625, 1.6650390625, 1.666259765625},
            {1.25, 1.25, 1.5625, 1.640625, 1.66015625, 1.6650390625})
        Dim expectedDependent = {1.75, 2.375, 2.84375, 2.9609375, 2.990234375, 2.99755859375}

        Dim firstAlias = If(firstSortsBeforeSecond, "Z alias", "A alias")
        Dim secondAlias = If(firstSortsBeforeSecond, "A alias", "Z alias")
        Using first = OpenReference(firstCase, session, firstAlias), second = OpenReference(secondCase, session, secondAlias)
            For pass = 0 To expectedFirst.Length - 1
                If pass > 0 Then first.RecalculateCell("Data", 0, 1)
                Assert.That(Cached(first, "Data", 0, 0), [Is].EqualTo(expectedFirst(pass)), "First workbook pass " & pass)
                Assert.That(Cached(second, "Data", 0, 0), [Is].EqualTo(expectedSecond(pass)), "Second workbook pass " & pass)
                Assert.That(Cached(first, "Data", 0, 1), [Is].EqualTo(expectedDependent(pass)), "Dependent pass " & pass)
            Next
        End Using
    End Sub

    <Test>
    Public Sub ARecalculatedCopyUsesTheOriginalIterationOrderAndKeepsBothOriginalCaches()
        Dim session As New ExcelFormulaSession()
        Dim output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Try
            Using first = OpenReference("CrossZFirst", session), second = OpenReference("CrossASecond", session)
                first.RecalculateAll()
                first.SaveCopyAs(output, recalculateWorkbook:=True, SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(Cached(first, "Data", 0, 0), [Is].EqualTo(1.125))
                Assert.That(Cached(first, "Data", 0, 1), [Is].EqualTo(2.375))
                Assert.That(Cached(second, "Data", 0, 0), [Is].EqualTo(1.25))
            End Using
            Using saved As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, CalculationOptions())
                Assert.That(Cached(saved, "Data", 0, 0), [Is].EqualTo(1.28125))
                Assert.That(Cached(saved, "Data", 0, 1), [Is].EqualTo(2.84375))
            End Using
        Finally
            If File.Exists(output) Then File.Delete(output)
        End Try
    End Sub

    <Test>
    Public Sub RejectsConflictingWorkbookIterationSettingsWithoutChangingTheirCaches()
        Dim session As New ExcelFormulaSession()
        Using data As New MemoryStream()
            Using input = File.OpenRead(ReferencePath("CrossASecond"))
                input.CopyTo(data)
            End Using
            data.Position = 0
            Using package = SpreadsheetDocument.Open(data, True)
                package.WorkbookPart.Workbook.CalculationProperties.IterateCount = 2UI
            End Using

            Dim firstOptions = CalculationOptions(session)
            firstOptions.FormulaWorkbookName = Path.GetFileName(ReferencePath("CrossZFirst"))
            Dim secondOptions = CalculationOptions(session)
            secondOptions.FormulaWorkbookName = Path.GetFileName(ReferencePath("CrossASecond"))
            Using first As New OpenXmlExcelDataOperations(File.ReadAllBytes(ReferencePath("CrossZFirst")), firstOptions),
                  second As New OpenXmlExcelDataOperations(data.ToArray(), secondOptions)
                Dim failure = Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() first.RecalculateAll())
                Assert.That(failure.Message, Does.Contain("matching iteration settings"))
                Assert.That(Cached(first, "Data", 0, 0), [Is].EqualTo(0.5))
                Assert.That(Cached(first, "Data", 0, 1), [Is].EqualTo(1.75))
                Assert.That(Cached(second, "Data", 0, 0), [Is].EqualTo(1.25))
            End Using
        End Using
    End Sub

    <Test>
    Public Sub DisabledIterationReportsACircularValueButAllowsAnAddressQuery()
        Using bytes As New MemoryStream()
            Dim input = ReferencePath("Steps")
            Using inputStream = File.OpenRead(input)
                inputStream.CopyTo(bytes)
            End Using
            bytes.Position = 0
            Using package = SpreadsheetDocument.Open(bytes, True)
                package.WorkbookPart.Workbook.CalculationProperties.Iterate = False
            End Using
            Using workbook As New OpenXmlExcelDataOperations(bytes.ToArray(), CalculationOptions())
                Assert.Throws(Of CircularReferenceException)(Sub() workbook.RecalculateCell("Cycle", 0, 0))
                AssertReferencePass(workbook, 0)
                workbook.RecalculateCell("Cycle", 0, 2)
                Assert.That(Cached(workbook, "Cycle", 0, 2), [Is].EqualTo(1.0))
            End Using
        End Using
    End Sub

    Private Shared Sub AssertReferencePass(workbook As OpenXmlExcelDataOperations, pass As Integer)
        Dim circular = Cached(workbook, "Reference", pass + 1, 1)
        Dim dependent = Cached(workbook, "Reference", pass + 1, 2)
        Assert.That(Cached(workbook, "Cycle", 0, 0), [Is].EqualTo(circular).Within(1.0E-12), "Circular result on pass " & pass)
        Assert.That(Cached(workbook, "Cycle", 0, 1), [Is].EqualTo(dependent).Within(1.0E-12), "Dependent result on pass " & pass)
    End Sub

    Private Shared Function Cached(workbook As OpenXmlExcelDataOperations, sheet As String, row As Integer, column As Integer) As Double
        Return workbook.LookupCellValue(Of Double)(sheet, row, column, FormulaCalculationMode.UseCachedValueOnly)
    End Function

    Private Shared Function OpenReference(caseName As String, Optional session As ExcelFormulaSession = Nothing,
                                          Optional aliasName As String = Nothing) As OpenXmlExcelDataOperations
        Dim options = CalculationOptions(session)
        options.FormulaWorkbookName = aliasName
        Return New OpenXmlExcelDataOperations(ReferencePath(caseName), OpenMode.OpenExistingFile, options)
    End Function

    Private Shared Function ReferencePath(caseName As String) As String
        Return TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", "FormulaReferenceIteration" & caseName & ".xlsx")
    End Function

    Private Shared Function CalculationOptions(Optional session As ExcelFormulaSession = Nothing) As ExcelDataOperationsOptions
        Return New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .CalculateOnCellEdit = True, .FormulaSession = session
        }
    End Function

End Class
