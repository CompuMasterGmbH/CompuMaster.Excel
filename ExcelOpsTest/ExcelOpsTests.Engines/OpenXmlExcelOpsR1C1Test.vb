Imports System.Globalization
Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.Test.ExcelOpsTests.Engines
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

Partial Public Class OpenXmlExcelOpsTest

    ' Excel serializes direct arguments from its R1C1 formula API as A1 operands.
    <Test>
    Public Sub RewritesFunctionArgumentsSavedFromR1C1AcrossStructuralEdits()
        Dim source = Path.Combine(_directory, "r1c1-function-arguments.xlsx")
        Dim output = Path.Combine(_directory, "r1c1-function-arguments-edited.xlsx")
        Const renamedSheet As String = "O'Brien Input"
        Const renamedQualifier As String = "'O''Brien Input'!"
        Dim originals = R1C1FunctionFormulas(DataQualifier, "C", 4, 6)
        Dim checks = originals.Select(Function(formula, row) FormulaCell(ExcelCell.LocalCellAddress(row, 0), formula)).ToArray()
        CreateStructuralWorkbook(source,
            {ValueCell("C4", 40), ValueCell("C5", 50), ValueCell("C6", 60), FormulaCell("D7", "SUM(C4:C6)")}, checks)

        Using document = SpreadsheetDocument.Open(source, True)
            document.WorkbookPart.Workbook.CalculationProperties.SetAttribute(New OpenXmlAttribute("refMode", "", "R1C1"))
        End Using

        Using workbook = OpenStructuralWorkbook(source)
            workbook.AddRow(DataSheet, 3, 1, True)
            AssertFunctionFormulas(workbook, R1C1FunctionFormulas(DataQualifier, "C", 5, 7))
            Assert.That(workbook.LookupCellFormula(DataSheet, 7, 3), [Is].EqualTo("SUM(C5:C7)"))

            workbook.AddColumn(DataSheet, 2, 1, True)
            AssertFunctionFormulas(workbook, R1C1FunctionFormulas(DataQualifier, "D", 5, 7))
            Assert.That(workbook.LookupCellFormula(DataSheet, 7, 4), [Is].EqualTo("SUM(D5:D7)"))

            workbook.RenameWorksheet(DataSheet, renamedSheet, True)
            AssertFunctionFormulas(workbook, R1C1FunctionFormulas(renamedQualifier, "D", 5, 7))
            Assert.That(workbook.HasMissingCachedFormulaResult("Check", 0, 0), [Is].True)
            workbook.SaveCopyAs(output, SaveOptionsForDisabledCalculationEngines.NoReset)
        End Using

        Using saved = OpenStructuralWorkbook(output)
            AssertFunctionFormulas(saved, R1C1FunctionFormulas(renamedQualifier, "D", 5, 7))
        End Using
        Using document = SpreadsheetDocument.Open(output, False)
            Assert.That(document.WorkbookPart.Workbook.CalculationProperties.GetAttribute("refMode", "").Value, [Is].EqualTo("R1C1"))
        End Using
    End Sub

    ' Excel retains the dynamic range operator when a direct endpoint is deleted.
    <TestCase(3, 1, "$C$4:$C$5", "#REF!", "$C$5")>
    <TestCase(5, 1, "$C$4:$C$5", "$C$4", "#REF!")>
    <TestCase(3, 3, "#REF!", "#REF!", "#REF!")>
    Public Sub KeepsDynamicRangesWhenDirectEndpointsAreDeleted(row As Integer, count As Integer, expectedArea As String,
                                                              expectedLeft As String, expectedRight As String)
        Dim source = Path.Combine(_directory, "dynamic-range-deletion.xlsx")
        Dim before = DynamicRangeFormulas(DataQualifier, "$C$4:$C$6", "$C$4", "$C$6")
        CreateStructuralWorkbook(source,
            {ValueCell("C4", 40), ValueCell("C5", 50), ValueCell("C6", 60)},
            before.Select(Function(formula, index) FormulaCell(ExcelCell.LocalCellAddress(index, 0), formula)).ToArray())

        Using workbook = OpenStructuralWorkbook(source)
            workbook.RemoveRows(DataSheet, row, count, True)
            AssertFunctionFormulas(workbook, DynamicRangeFormulas(DataQualifier, expectedArea, expectedLeft, expectedRight))

            ' The rewritten #REF! operands are valid input for another structural edit.
            workbook.RenameWorksheet(DataSheet, "NewData", True)
            AssertFunctionFormulas(workbook, DynamicRangeFormulas("NewData!", expectedArea, expectedLeft, expectedRight))
        End Using
    End Sub

    ' Excel distinguishes a space between ranges from a call that supplies a range endpoint.
    <Test>
    Public Sub KeepsIntersectionsAndRelativeR1C1Text()
        Dim source = Path.Combine(_directory, "intersection-and-relative-indirect.xlsx")
        Const relativeIndirect As String = "INDIRECT(""Z(-1)S(-1)"",FALSE)"
        CreateStructuralWorkbook(source,
            {ValueCell("C4", 40), ValueCell("C5", 50), ValueCell("C6", 60), FormulaCell("D7", relativeIndirect)},
            {FormulaCell("A1", "SUM(" & DataQualifier & "C4:C6 (" & DataQualifier & "C4:C6))")})

        Using workbook = OpenStructuralWorkbook(source)
            workbook.RemoveRows(DataSheet, 3, 1, True)
            Assert.That(workbook.LookupCellFormula("Check", 0, 0),
                        [Is].EqualTo("SUM(" & DataQualifier & "C4:C5 (" & DataQualifier & "C4:C5))"))
            Assert.That(workbook.LookupCellFormula(DataSheet, 5, 3), [Is].EqualTo(relativeIndirect))
            workbook.AddRow(DataSheet, 5, 1, True)
            Assert.That(workbook.LookupCellFormula(DataSheet, 6, 3), [Is].EqualTo(relativeIndirect))
        End Using
    End Sub

    <TestCase(False, False)>
    <TestCase(True, False)>
    <TestCase(False, True)>
    <TestCase(True, True)>
    Public Sub DatabaseReferenceArgumentsMatchExcel(insertColumn As Boolean, includeCachedResults As Boolean)
        Dim source = Path.Combine(_directory, "database-references.xlsx")
        Dim output = Path.Combine(_directory, "database-references-edited.xlsx")
        DatabaseReferenceCoverage.CreateFixture(source, includeCachedResults)
        Using workbook = OpenStructuralWorkbook(source)
            Dim differences = DatabaseReferenceCoverage.CompareWithExcel(workbook, output, insertColumn)
            Assert.That(differences, [Is].Empty, String.Join(Environment.NewLine, differences))
        End Using
    End Sub

    Private Shared Function R1C1FunctionFormulas(qualifier As String, column As String, firstRow As Integer, lastRow As Integer) As String()
        Dim first = firstRow.ToString(CultureInfo.InvariantCulture)
        Dim last = lastRow.ToString(CultureInfo.InvariantCulture)
        Dim area = "$" & column & "$" & first & ":$" & column & "$" & last
        Dim operand = qualifier & "$" & column & "$" & first
        Dim literalSheet = """" & DataSheet & """"
        Return {
            "SUM(" & qualifier & area & ")",
            "INDEX(" & qualifier & area & ",2)",
            "CELL(""address""," & operand & ")",
            "INDIRECT(""" & DataSheet & "!R4C3"",FALSE)",
            "INDIRECT(ADDRESS(ROW(" & operand & "),COLUMN(" & operand & "),1,FALSE," & literalSheet & "),FALSE)",
            "INDIRECT(ADDRESS(4,3,1,FALSE," & literalSheet & "),FALSE)",
            "SUM(" & qualifier & "$" & first & ":$" & last & ")",
            "SUM(" & qualifier & "$" & column & ":$" & column & ")",
            "SUM(" & qualifier & "$" & column & first & ":$" & column & last & ")",
            "INDIRECT(""" & DataSheet & "!Z4S3"",FALSE)",
            "N(""Z4S3 and R[1]C[2]"")+SUM(" & operand & ")"
        }.Concat(DynamicRangeFormulas(qualifier, area, "$" & column & "$" & first, "$" & column & "$" & last)).ToArray()
    End Function

    Private Shared Function DynamicRangeFormulas(qualifier As String, area As String, left As String, right As String) As String()
        Dim first = "INDEX(" & qualifier & area & ",1)"
        Dim second = "INDEX(" & qualifier & area & ",2)"
        Return {
            "SUM(" & first & ":" & second & ")",
            "SUM(" & qualifier & left & ":" & second & ")",
            "SUM(" & first & ":" & qualifier & right & ")"
        }
    End Function

    Private Shared Sub AssertFunctionFormulas(workbook As OpenXmlExcelDataOperations, expected As String())
        Dim actual = expected.Select(Function(formula, row) workbook.LookupCellFormula("Check", row, 0)).ToArray()
        Assert.That(actual, [Is].EqualTo(expected))
    End Sub

End Class
