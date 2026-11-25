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
    Public Sub ReportsAnIndirectCycleWithoutChangingThePreviousResults()
        Using workbook = OpenReference("Steps")
            workbook.WriteCellFormula("Cycle", 0, 0, "(B1+D1)/2", False)
            Dim failure = Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.RecalculateSheet("Cycle"))
            Assert.That(failure.Message, Does.Contain("multiple circular formula cells"))
            Assert.That(workbook.HasMissingCachedFormulaResult("Cycle", 0, 0), [Is].True)
            Assert.That(Cached(workbook, "Cycle", 0, 1), [Is].EqualTo(15.0))
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

    Private Shared Function OpenReference(caseName As String) As OpenXmlExcelDataOperations
        Return New OpenXmlExcelDataOperations(ReferencePath(caseName), OpenMode.OpenExistingFile, CalculationOptions())
    End Function

    Private Shared Function ReferencePath(caseName As String) As String
        Return TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", "FormulaReferenceIteration" & caseName & ".xlsx")
    End Function

    Private Shared Function CalculationOptions() As ExcelDataOperationsOptions
        Return New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .CalculateOnCellEdit = True
        }
    End Function

End Class
