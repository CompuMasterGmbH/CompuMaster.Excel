Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports NUnit.Framework

<TestFixture>
Public Class OpenXmlCustomFormulaTest

    <Test>
    Public Sub UsesFormulaIdentityAndResolvesCrossWorkbookPrecedentsAcrossFiveLevels()
        Dim session As New ExcelFormulaSession()
        Dim lastContext As ExcelFormulaContext = Nothing
        session.RegisterFunction("APP.RATE",
            Function(context, arguments)
                lastContext = context
                Assert.That(context.Workbook.Name, [Is].EqualTo("Target"))
                Assert.That(context.Worksheet, [Is].EqualTo("Sheet1"))
                Assert.That(context.FormulaCell.Address, [Is].EqualTo("A1"))
                Assert.That(context.ActiveWorkbookName, [Is].EqualTo("Rates"))
                Assert.That(context.ActiveCell.Address, [Is].EqualTo("C4"))
                Assert.That(context.NativeExcelApplication, [Is].Null)
                Assert.That(context.WorkbookSession.WorkbookNames, [Is].EquivalentTo({"Target", "Rates"}))
                Dim source = context.WorkbookSession.FindWorkbook("rates").ReadCell(Cell("B1"))
                Dim local = context.Workbook.ReadCell(Cell("B1"))
                Return ExcelFormulaValue.FromNumber(CDbl(source.Value) + CDbl(local.Value) + CDbl(arguments(0).Value))
            End Function)

        Using target = CreateWorkbook(session, "Target"), rates = CreateWorkbook(session, "Rates")
            session.SetActiveCell("Rates", Cell("C4"))
            rates.WriteCellValue(Cell("A1"), 7)
            rates.WriteCellFormula("Sheet1", 0, 1, "A1+1", False)
            target.WriteCellValue(Cell("B1"), 2)
            target.WriteCellFormula("Sheet1", 0, 0, "app.rate(1)", False)
            target.WriteCellFormula("Sheet1", 0, 2, "A1+1", False)
            target.WriteCellFormula("Sheet1", 0, 3, "C1+1", False)
            target.WriteCellFormula("Sheet1", 0, 4, "D1+1", False)
            target.WriteCellFormula("Sheet1", 0, 5, "E1+1", False)
            target.WriteCellFormula("Sheet1", 0, 6, "F1+1", False)

            target.RecalculateCell(Cell("G1"))
            Assert.That(Cached(target, "G1"), [Is].EqualTo(16.0))
            Assert.That(Cached(rates, "B1"), [Is].EqualTo(8.0))
            Assert.Throws(Of InvalidOperationException)(Sub() lastContext.Workbook.ReadCell(Cell("B1")))

            rates.WriteCellValue(Cell("A1"), 10)
            Assert.That(Cached(target, "G1"), [Is].EqualTo(16.0))
            target.RecalculateCell(Cell("G1"))
            Assert.That(Cached(target, "G1"), [Is].EqualTo(19.0))
            Assert.That(Cached(rates, "B1"), [Is].EqualTo(11.0))
        End Using
    End Sub

    <Test>
    Public Sub ReportsCrossWorkbookValueCycleWithoutCommittingResults()
        Dim session As New ExcelFormulaSession()
        session.RegisterFunction("OTHER",
            Function(context, arguments)
                Dim otherName = If(context.Workbook.Name = "First", "Second", "First")
                Return context.WorkbookSession.FindWorkbook(otherName).ReadCell(Cell("A1"))
            End Function)
        Using first = CreateWorkbook(session, "First"), second = CreateWorkbook(session, "Second")
            first.WriteCellFormula("Sheet1", 0, 0, "OTHER()", False)
            second.WriteCellFormula("Sheet1", 0, 0, "OTHER()", False)
            Assert.Throws(Of CircularReferenceException)(Sub() first.RecalculateCell(Cell("A1")))
            Assert.That(first.HasMissingCachedFormulaResult(Cell("A1")), [Is].True)
            Assert.That(second.HasMissingCachedFormulaResult(Cell("A1")), [Is].True)
        End Using
    End Sub

    <Test>
    Public Sub PreservesTypedArgumentsAndDoesNotSerializeCallbacks()
        Dim session As New ExcelFormulaSession()
        session.RegisterFunction("PROFILE",
            Function(context, arguments)
                Assert.That(context.ActiveCell, [Is].Null)
                Assert.That(arguments.Select(Function(value) value.Kind), [Is].EqualTo({
                    ExcelFormulaValueKind.Logical, ExcelFormulaValueKind.Text, ExcelFormulaValueKind.Error, ExcelFormulaValueKind.Blank
                }))
                Assert.That(arguments(0).Value, [Is].True)
                Assert.That(arguments(1).Value, [Is].EqualTo(""))
                Assert.That(arguments(2).Value, [Is].EqualTo("#N/A"))
                Assert.That(arguments(3).Value, [Is].Null)
                Return ExcelFormulaValue.FromError("#DIV/0!")
            End Function)
        Using first = CreateWorkbook(session, "Target"), unrelated = CreateWorkbook(New ExcelFormulaSession(), "Target")
            first.WriteCellFormula("Sheet1", 0, 0, "PROFILE(TRUE,"""",#N/A,)", False)
            unrelated.WriteCellFormula("Sheet1", 0, 0, "PROFILE(TRUE,"""",#N/A,)", False)
            Assert.That(first.LookupCellValueAsObject(Cell("A1"), FormulaCalculationMode.CalculateIfMissing), [Is].EqualTo("#DIV/0!"))
            Assert.Throws(Of FormulaComplexityNotSupportedException)(
                Sub() unrelated.LookupCellValueAsObject(Cell("A1"), FormulaCalculationMode.CalculateIfMissing))
        End Using
    End Sub

    <TestCase("_xlfn.TRANSLATE", "TRANSLATE")>
    <TestCase("_xlfn._xlws.APP.LOOKUP", "APP.LOOKUP")>
    <TestCase("_xll.APP.RATE", "APP.RATE")>
    Public Sub ResolvesAnExplicitRegistrationForAnExcelStoredFunctionName(storedName As String, registrationName As String)
        Dim session As New ExcelFormulaSession()
        Using workbook = CreateWorkbook(session, "Target")
            Dim formula = storedName & "(""test"")"
            workbook.WriteCellFormula("Sheet1", 0, 0, formula, False)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.RecalculateCell(Cell("A1")))
            session.RegisterFunction(registrationName, Function(context, arguments) ExcelFormulaValue.FromText(CStr(arguments(0).Value) & " result"))
            workbook.RecalculateCell(Cell("A1"))
            Assert.That(Cached(workbook, "A1"), [Is].EqualTo("test result"))
            Assert.That(workbook.LookupCellFormula("Sheet1", 0, 0), [Is].EqualTo(formula))
        End Using
    End Sub

    <Test>
    Public Sub RejectsDuplicateWorkbookNamesAndReleasesClosedWorkbooks()
        Dim session As New ExcelFormulaSession()
        Using first = CreateWorkbook(session, "Rates")
            Assert.Throws(Of ArgumentException)(Sub() CreateWorkbook(session, "rates"))
            session.SetActiveCell("Rates", Cell("A1"))
        End Using
        Using reopened = CreateWorkbook(session, "rates")
            session.RegisterFunction("APP.SELECTION", Function(context, arguments) ExcelFormulaValue.FromBoolean(context.ActiveCell Is Nothing))
            reopened.WriteCellFormula("Sheet1", 0, 0, "APP.SELECTION()", True)
            Assert.That(Cached(reopened, "A1"), [Is].True)
        End Using
    End Sub

    Private Shared Function CreateWorkbook(session As ExcelFormulaSession, name As String) As OpenXmlExcelDataOperations
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .FormulaSession = session, .FormulaWorkbookName = name
        }
        Return New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, options)
    End Function

    Private Shared Function Cell(address As String) As ExcelCell
        Return New ExcelCell("Sheet1", address, ExcelCell.ValueTypes.All)
    End Function

    Private Shared Function Cached(workbook As OpenXmlExcelDataOperations, address As String) As Object
        Return workbook.LookupCellValueAsObject(Cell(address), FormulaCalculationMode.UseCachedValueOnly)
    End Function

End Class
