Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlLetFormulaTest

    <TestCase("LET(rate,1/0,5)", 5.0)>
    <TestCase("LET(rate,NA(),ISNA(rate))", True)>
    <TestCase("LET(rate,2,LET(rate,rate+1,rate)+rate)", 5.0)>
    <TestCase("LET(rate,1,other,rate+2,LET(rate,other*2,rate+other)+rate)", 10.0)>
    <TestCase("LET(rate,9,Rate)", 9.0)>
    <TestCase("LET(rate,,rate)", 0.0)>
    <TestCase("LET(_rate,2,_rate)", 2.0)>
    <TestCase("LET(SUM,2,SUM(3)+SUM)", 5.0)>
    <TestCase("LET(amount,A1,ISBLANK(amount))", True)>
    <TestCase("LET(amount,A1,ROW(amount))", 1.0)>
    <TestCase("LET(amount,A1:A3,SUM(amount))", 2.0)>
    <TestCase("LET(amount,A1:A3,ISREF(amount))", True)>
    <TestCase("LET(myLocal,2,myLocal)+myLocal", "#NAME?")>
    <TestCase("LET(rate,6,rate)+LET(rate,8,rate)", 14.0)>
    <TestCase("LET(rate,SUM(SEQUENCE(1,3)),rate*rate)", 36.0)>
    <TestCase("_xlfn.LET(_xlpm.amount,A2,_xlpm.amount+_xlpm.amount)", 4.0)>
    Public Sub MatchesExcelLocalNameRules(formula As String, expected As Object)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False}
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, options)
            workbook.WriteCellValue("Sheet1", 1, 0, 2.0)
            workbook.WriteCellFormula("Sheet1", 0, 4, formula, False)
            Assert.That(workbook.LookupCellValueAsObject("Sheet1", 0, 4, FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo(expected), formula)
        End Using
    End Sub

    <Test>
    Public Sub EvaluatesABindingAtMostOnceInEachPass()
        Dim calls As Integer
        Dim session As New ExcelFormulaSession()
        session.RegisterFunction("APP.RATE",
            Function(context, arguments)
                calls += 1
                Return ExcelFormulaValue.FromNumber(4)
            End Function)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .FormulaSession = session, .FormulaWorkbookName = "Book"
        }
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, options)
            workbook.WriteCellFormula("Sheet1", 0, 4, "LET(rate,APP.RATE(),unused,APP.RATE(),rate+rate)", False)
            For pass = 1 To 2
                workbook.RecalculateCell("Sheet1", 0, 4)
                Assert.That(workbook.LookupCellValueAsObject("Sheet1", 0, 4, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(8.0))
                Assert.That(calls, [Is].EqualTo(pass))
            Next
        End Using
    End Sub

End Class
