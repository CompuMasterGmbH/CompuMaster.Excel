Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    ' The worksheet results were calculated by Excel against the embedded model.
    <Test>
    Public Sub PreservesEmbeddedDataModelAndCachedCubeResultsAfterAnEdit()
        Dim source = TestEnvironment.FullPathOfExistingTestFile("test_data", "OpenXmlCubeReference.xlsx")
        Dim output = Path.Combine(_directory, "CubeRoundTrip.xlsx")
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False,
            .DisableInitialCalculation = True
        }

        Using workbook As New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile, options)
            AssertCachedCubeResults(workbook)
            workbook.WriteCellValue("Results", 0, 4, "Round-trip edit")
            workbook.SaveCopyAs(output, recalculateWorkbook:=False, cachedCalculationsOption:=SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        For Each relationship In {"powerPivotData", "connections"}
            Assert.That(ReadWorkbookPart(output, relationship), [Is].EqualTo(ReadWorkbookPart(source, relationship)), relationship)
        Next

        Using reopened As New OpenXmlExcelDataOperations(output, OpenMode.OpenExistingFile, options)
            Assert.That(reopened.LookupCellValue(Of String)("Results", 0, 4, FormulaCalculationMode.UseCachedValueOnly),
                        [Is].EqualTo("Round-trip edit"))
            AssertCachedCubeResults(reopened)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() reopened.RecalculateCell("Results", 1, 1))
        End Using
    End Sub

    Private Shared Sub AssertCachedCubeResults(workbook As OpenXmlExcelDataOperations)
        Dim expected As Object() = {25.0, "Apple", 20.0, 5.0, "All products", 2.0}
        For index = 0 To expected.Length - 1
            Dim rowIndex = index + 1
            Dim value = workbook.LookupCellValueAsObject("Results", rowIndex, 1, FormulaCalculationMode.UseCachedValueOnly)
            Assert.That(value, [Is].TypeOf(expected(index).GetType()).And.EqualTo(expected(index)), $"Results!B{rowIndex + 1}")
        Next
    End Sub

    Private Shared Function ReadWorkbookPart(file As String, relationship As String) As Byte()
        Using document = SpreadsheetDocument.Open(file, False)
            Dim part = document.WorkbookPart.Parts.Select(Function(item) item.OpenXmlPart).
                Single(Function(item) item.RelationshipType.EndsWith("/" & relationship, StringComparison.Ordinal))
            Using content = part.GetStream(FileMode.Open, FileAccess.Read), buffer As New MemoryStream()
                content.CopyTo(buffer)
                Return buffer.ToArray()
            End Using
        End Using
    End Function

End Class
