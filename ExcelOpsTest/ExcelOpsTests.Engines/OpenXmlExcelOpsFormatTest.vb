Imports System.IO
Imports System.Threading.Tasks
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlExcelOpsFormatTest

    Private _directory As String

    <SetUp>
    Public Sub CreateDirectory()
        _directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_directory)
    End Sub

    <TearDown>
    Public Sub RemoveDirectory()
        Directory.Delete(_directory, recursive:=True)
    End Sub

    <TestCase(".xlsx", SpreadsheetDocumentType.Workbook, False)>
    <TestCase(".xlsm", SpreadsheetDocumentType.MacroEnabledWorkbook, True)>
    <TestCase(".xltx", SpreadsheetDocumentType.Template, False)>
    <TestCase(".xltm", SpreadsheetDocumentType.MacroEnabledTemplate, True)>
    Public Sub SaveCopyUsesDestinationFormat(extension As String, expectedType As SpreadsheetDocumentType, expectVba As Boolean)
        Dim destination = Path.Combine(_directory, "copy" & extension)
        Using workbook = OpenMacroWorkbook()
            Dim sheet = workbook.SheetNames().First()
            workbook.WriteCellValue(sheet, 0, 0, "Unsaved edit")
            workbook.SaveCopyAs(destination, SaveOptionsForDisabledCalculationEngines.NoReset)
            Assert.That(workbook.HasVbaProject, [Is].True)
        End Using

        Using document = SpreadsheetDocument.Open(destination, False)
            Assert.That(document.DocumentType, [Is].EqualTo(expectedType))
            Assert.That(document.WorkbookPart.VbaProjectPart IsNot Nothing, [Is].EqualTo(expectVba))
        End Using

        Using reopened As New OpenXmlExcelDataOperations(destination, OpenMode.OpenExistingFile,
                New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.That(reopened.LookupCellValue(Of String)(reopened.SheetNames().First(), 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("Unsaved edit"))
        End Using
    End Sub

    <Test>
    Public Sub RemoveVbaPersistsInMacroEnabledOutput()
        Dim destination = Path.Combine(_directory, "removed.xlsm")
        Using workbook = OpenMacroWorkbook()
            Assert.That(workbook.HasVbaProject, [Is].True)
            workbook.RemoveVbaProject()
            workbook.RemoveVbaProject()
            Assert.That(workbook.HasVbaProject, [Is].False)
            workbook.SaveAs(destination, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(destination, False)
            Assert.That(document.DocumentType, [Is].EqualTo(SpreadsheetDocumentType.MacroEnabledWorkbook))
            Assert.That(document.WorkbookPart.VbaProjectPart, [Is].Null)
        End Using
    End Sub

    <Test>
    Public Sub ConcurrentCopiesKeepVbaInMacroEnabledCopy()
        Dim macroFree = Path.Combine(_directory, "copy.xlsx")
        Dim macroEnabled = Path.Combine(_directory, "copy.xlsm")
        Using workbook = OpenMacroWorkbook()
            Parallel.Invoke(
                Sub() workbook.SaveCopyAs(macroFree, SaveOptionsForDisabledCalculationEngines.NoReset),
                Sub() workbook.SaveCopyAs(macroEnabled, SaveOptionsForDisabledCalculationEngines.NoReset))
            Assert.That(workbook.HasVbaProject, [Is].True)
        End Using
        Using document = SpreadsheetDocument.Open(macroFree, False)
            Assert.That(document.WorkbookPart.VbaProjectPart, [Is].Null)
        End Using
        Using document = SpreadsheetDocument.Open(macroEnabled, False)
            Assert.That(document.WorkbookPart.VbaProjectPart, [Is].Not.Null)
        End Using
    End Sub

    <Test>
    Public Sub SaveAsMacroFreeDoesNotRestoreVbaOnLaterSave()
        Dim macroFree = Path.Combine(_directory, "book.xlsx")
        Dim macroEnabled = Path.Combine(_directory, "book.xlsm")
        Using workbook = OpenMacroWorkbook()
            workbook.SaveAs(macroFree, SaveOptionsForDisabledCalculationEngines.NoReset)
            Assert.That(workbook.FilePath, [Is].EqualTo(macroFree))
            Assert.That(workbook.HasVbaProject, [Is].False)
            workbook.SaveCopyAs(macroEnabled, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using
        Using document = SpreadsheetDocument.Open(macroEnabled, False)
            Assert.That(document.WorkbookPart.VbaProjectPart, [Is].Null)
        End Using
    End Sub

    <Test>
    Public Sub RecalculatingCopyKeepsOriginalFormulaCacheMissing()
        Dim source = Path.Combine(_directory, "source.xlsx")
        Dim destination = Path.Combine(_directory, "copy.xlsx")
        Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite,
            Nothing, disableInitialCalculation:=True, disableAutoCalculation:=Nothing, disableCalculationEngine:=False)
        Using workbook As New OpenXmlExcelDataOperations(source, OpenMode.CreateFile, options)
            workbook.WriteCellFormula("Sheet1", 0, 0, "1+2", immediatelyCalculateCellValue:=False)
            workbook.SaveCopyAs(destination, True, SaveOptionsForDisabledCalculationEngines.NoReset)
            Assert.That(workbook.HasMissingCachedFormulaResult("Sheet1", 0, 0), [Is].True)
        End Using
        Using workbook As New OpenXmlExcelDataOperations(destination, OpenMode.OpenExistingFile, options)
            Assert.That(workbook.LookupCellValue(Of Integer)("Sheet1", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(3))
        End Using
    End Sub

    <Test>
    Public Sub ClonePreservesVbaAndUnsavedEditsIndependently()
        Using source = OpenMacroWorkbook()
            Dim sheet = source.SheetNames().First()
            source.WriteCellValue(sheet, 0, 0, "Unsaved edit")
            Using copy = DirectCast(DirectCast(source, ICloneable).Clone(), OpenXmlExcelDataOperations)
                Assert.That(copy.FilePath, [Is].Null.Or.Empty)
                Assert.That(copy.ReadOnly, [Is].False)
                Assert.That(copy.HasVbaProject, [Is].True)
                Assert.That(copy.LookupCellValue(Of String)(sheet, 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("Unsaved edit"))
                copy.RemoveVbaProject()
                copy.WriteCellValue(sheet, 0, 0, "Copy only")
                Assert.That(source.HasVbaProject, [Is].True)
                Assert.That(source.LookupCellValue(Of String)(sheet, 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("Unsaved edit"))
                source.Close()
                Assert.That(copy.LookupCellValue(Of String)(sheet, 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("Copy only"))
            End Using
            Assert.Throws(Of InvalidOperationException)(Sub() source.Clone())
        End Using
    End Sub

    <Test>
    Public Sub ClonePreservesPresentAndMissingFormulaCaches()
        Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite,
            Nothing, disableInitialCalculation:=True, disableAutoCalculation:=Nothing, disableCalculationEngine:=False)
        Using source As New OpenXmlExcelDataOperations(Path.Combine(_directory, "source.xlsx"), OpenMode.CreateFile, options)
            source.WriteCellFormula("Sheet1", 0, 0, "1+2", immediatelyCalculateCellValue:=True)
            source.WriteCellFormula("Sheet1", 0, 1, "4+5", immediatelyCalculateCellValue:=False)
            Using copy = source.Clone()
                Assert.That(copy.LookupCellValue(Of Integer)("Sheet1", 0, 0, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(3))
                Assert.That(copy.HasMissingCachedFormulaResult("Sheet1", 0, 1), [Is].True)
                Assert.That(copy.RecalculationRequired, [Is].EqualTo(source.RecalculationRequired))
            End Using
        End Using
    End Sub

    <Test>
    Public Sub ConcurrentEditsAndClonesProduceCoherentSnapshots()
        Using source = OpenMacroWorkbook()
            Dim sheet = source.SheetNames().First()
            source.WriteCellValue(sheet, 0, 0, "Before")
            Parallel.Invoke(
                Sub()
                    For index = 0 To 9
                        source.WriteCellValue(sheet, 0, 0, "After")
                    Next
                End Sub,
                Sub()
                    For index = 0 To 9
                        Using copy = source.Clone()
                            Dim value = copy.LookupCellValue(Of String)(sheet, 0, 0, FormulaCalculationMode.UseCachedValueOnly)
                            Assert.That(value, [Is].EqualTo("Before").Or.EqualTo("After"))
                            Assert.That(copy.HasVbaProject, [Is].True)
                        End Using
                    Next
                End Sub)
        End Using
    End Sub

    Private Shared Function OpenMacroWorkbook() As OpenXmlExcelDataOperations
        Dim source = Path.Combine(TestContext.CurrentContext.TestDirectory, "test_data", "VbaProject.xlsm")
        Return New OpenXmlExcelDataOperations(source, OpenMode.OpenExistingFile,
            New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
    End Function

End Class
