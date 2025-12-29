Imports NUnit.Framework
Imports NUnit.Framework.Legacy
Imports CompuMaster.Excel.ExcelOps
Imports System.Data

Namespace ExcelOpsTests.MsExcelSpecials

    ''' <summary>
    ''' Test (component design) issue: MsExcelAutoCalc doesn't work but only manual calc works -> Bug!?
    ''' </summary>
    <NonParallelizable>
    <TestFixture>
    Public MustInherit Class MsExcelCalcTestBase

        <SetUp> Public Sub ResetConsoleForTestOutput()
            CompuMaster.Excel.Test.Console.ResetConsoleForTestOutput()
            If CompuMaster.ComInterop.ComTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled("Excel.Application") = False Then
                Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
            End If
        End Sub

        Protected MustOverride ReadOnly Property EngineName As String

        Protected MustOverride Function CreateEngineInstanceWithCreateFileMode(testFile As String) As ExcelOps.ExcelDataOperationsBase

        Protected MustOverride Sub EngineResetCellValueFromFormulaCell(wb As ExcelOps.ExcelDataOperationsBase, sheetName As String, rowIndex As Integer, columnIndex As Integer)

#Region "Test Sample 1"
        <Test>
        Public Sub FormulaComplexityLevel1_Solution()
            Dim Eppeo As CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase = CreateSheetWithReproducableBug_FormulaComplexityLevel1()

            'Solve buggy cells in Excel workbook with Epplus
            Eppeo.ReloadFromFile()
            Dim FirstSheetName As String = Eppeo.SheetNames(0)

            System.Console.WriteLine("Formula B2 BEFORE RESET=" & Eppeo.LookupCellFormula(FirstSheetName, 1, 1))
            EngineResetCellValueFromFormulaCell(Eppeo, FirstSheetName, 1, 1)
            ClassicAssert.IsNotNull(Eppeo.LookupCellFormula(FirstSheetName, 1, 1))
            ClassicAssert.IsNotEmpty(Eppeo.LookupCellFormula(FirstSheetName, 1, 1))
            System.Console.WriteLine("Formula B2 AFTER RESET=" & Eppeo.LookupCellFormula(FirstSheetName, 1, 1))
            System.Console.WriteLine()
            EngineResetCellValueFromFormulaCell(Eppeo, FirstSheetName, 2, 1)
            EngineResetCellValueFromFormulaCell(Eppeo, FirstSheetName, 4, 1)
            EngineResetCellValueFromFormulaCell(Eppeo, FirstSheetName, 5, 1)

            Dim TestFilePattern As String = "MsExcelNoCalcBug_" & EngineName.ToString & "_FormulaComplexityLevel1{0}.xlsx"
            Dim TestFile As String = TestEnvironment.FullPathOfDynTestFile(GetType(MsExcelCalcTestBase), String.Format(TestFilePattern, "_11_FixedInEpplus"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)

            'Create workbook copy, open and recalculate and save in MS Excel
            TestFile = TestEnvironment.FullPathOfDynTestFile(GetType(MsExcelCalcTestBase), String.Format(TestFilePattern, "_12_ReSavedByMsExcel"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()
            Try
                CompuMaster.Excel.ExcelOps.MsVsEpplusTools.OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(TestFile)
            Catch ex As System.PlatformNotSupportedException
                ClassicAssert.Ignore("Platform not supported or MS Excel app not installed: " & ex.Message)
            End Try

            'Compare expected values
            Dim ETable As DataTable = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(TestFile, FirstSheetName, False)
            Dim TTable As New ExcelOps.TextTable(ETable)
            System.Console.WriteLine(TTable.ToUIExcelTable)
            ClassicAssert.AreEqual(20, ETable.Rows(0)(1))
            ClassicAssert.AreEqual(20, ETable.Rows(1)(1))
            ClassicAssert.AreEqual(20, ETable.Rows(2)(1))
            ClassicAssert.AreEqual(20, ETable.Rows(4)(1))
            ClassicAssert.AreEqual(20, ETable.Rows(5)(1))
        End Sub

        <Test> Public Sub FormulaComplexityLevel1_BugReproduction()
            CreateSheetWithReproducableBug_FormulaComplexityLevel1()
        End Sub

        Private Function CreateSheetWithReproducableBug_FormulaComplexityLevel1() As ExcelOps.ExcelDataOperationsBase
            Dim TestFilePattern As String = "MsExcelNoCalcBug_" & EngineName.ToString & "_FormulaComplexityLevel1{0}.xlsx"
            Dim TestFile As String = TestEnvironment.FullPathOfDynTestFile(GetType(MsExcelCalcTestBase), String.Format(TestFilePattern, "_01_InitialEpplus"))
            System.Console.WriteLine("Output path of test files: " & System.IO.Path.GetDirectoryName(TestFile))
            System.Console.WriteLine()

            'Create new Excel workbook with Epplus and add some cells with values and formulas
            Dim Eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateEngineInstanceWithCreateFileMode(TestFile)
            Dim FirstSheetName As String = Eppeo.SheetNames(0)

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 0, 0, "Static value initially set")
            Eppeo.WriteCellValue(Of Integer)(FirstSheetName, 0, 1, 50)

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 1, 0, "Formula referencing B1 immediately calculated by Epplus")
            Eppeo.WriteCellFormula(FirstSheetName, 1, 1, "B1", True)

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 2, 0, "Formula referencing B1 NOT calculated by Epplus")
            Eppeo.WriteCellFormula(FirstSheetName, 2, 1, "B1", False)

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 4, 0, "Formula referencing B2 immediately calculated by Epplus")
            Eppeo.WriteCellFormula(FirstSheetName, 4, 1, "B2", True)

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 5, 0, "Formula referencing B3 NOT calculated by Epplus")
            Eppeo.WriteCellFormula(FirstSheetName, 5, 1, "B3", False)

            Select Case EngineName
                Case (New ExcelOps.EpplusFreeExcelDataOperations(ExcelDataOperationsBase.OpenMode.Uninitialized)).EngineName
                    ClassicAssert.AreEqual(True, Eppeo.CalculationModuleDisabled)
                    ClassicAssert.Throws(Of FeatureDisabledException)(Sub() Eppeo.Save(True, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset))
                    Eppeo.CalculationModuleDisabled = False
                    ClassicAssert.AreEqual(False, Eppeo.CalculationModuleDisabled)
                Case Else
                    ClassicAssert.AreEqual(False, Eppeo.CalculationModuleDisabled)
            End Select
            Eppeo.Save(ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)

            'Create workbook copy, open and recalculate and save in MS Excel
            TestFile = TestEnvironment.FullPathOfDynTestFile(GetType(MsExcelCalcTestBase), String.Format(TestFilePattern, "_02_ReSavedByMsExcel"))
            Eppeo.CalculationModuleDisabled = False
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()
            Try
                CompuMaster.Excel.ExcelOps.MsVsEpplusTools.OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(TestFile)
            Catch ex As System.PlatformNotSupportedException
                ClassicAssert.Ignore("Platform not supported or MS Excel app not installed: " & ex.Message)
            End Try

            'Update single cells in calculated workbook with Epplus
            Eppeo.ReloadFromFile()

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 0, 0, "Static value rewritten")
            Eppeo.WriteCellValue(Of Integer)(FirstSheetName, 0, 1, 20)

            TestFile = TestEnvironment.FullPathOfDynTestFile(GetType(MsExcelCalcTestBase), String.Format(TestFilePattern, "_03_UpdatedByEpplus"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()

            'Compare expected values
            Dim ETable As DataTable = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(TestFile, FirstSheetName, False)
            Dim TTable As New ExcelOps.TextTable(ETable)
            System.Console.WriteLine(TTable.ToUIExcelTable)
            ClassicAssert.AreEqual(20, ETable.Rows(0)(1))
            ClassicAssert.AreEqual(50, ETable.Rows(1)(1))
            ClassicAssert.AreEqual(50, ETable.Rows(2)(1))
            ClassicAssert.AreEqual(50, ETable.Rows(4)(1))
            ClassicAssert.AreEqual(50, ETable.Rows(5)(1))

            'System.Diagnostics.Process.Start(TestFile)
            Return Eppeo

        End Function
#End Region

#Region "Test Sample 2"
        <Test> Public Sub FormulaComplexityLevel2_Solution()
            Dim Eppeo As ExcelOps.ExcelDataOperationsBase = CreateSheetWithReproducableBug_FormulaComplexityLevel2()

            Eppeo.ReloadFromFile()
            Dim FirstSheetName As String = Eppeo.SheetNames(0)

            'Solve buggy cells in whole Excel workbook with Epplus by resetting all formula cells in all worksheets
            Dim TestFilePattern As String = "MsExcelNoCalcBug_" & EngineName.ToString & "_FormulaComplexityLevel2{0}.xlsx"
            Dim TestFile As String
            TestFile = TestEnvironment.FullPathOfDynTestFile(GetType(MsExcelCalcTestBase), String.Format(TestFilePattern, "_12_ReSavedByMsExcel"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation) 'solution: reset all cell values in cells with formulas
            Eppeo.Close()

            'Open and recalculate and save in MS Excel
            Try
                CompuMaster.Excel.ExcelOps.MsVsEpplusTools.OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(TestFile)
            Catch ex As System.PlatformNotSupportedException
                ClassicAssert.Ignore("Platform not supported or MS Excel app not installed: " & ex.Message)
            End Try

            'Compare expected values
            Dim ETable As DataTable = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(TestFile, FirstSheetName, False)
            Dim TTable As New ExcelOps.TextTable(ETable)
            System.Console.WriteLine(TTable.ToUIExcelTable)
            ClassicAssert.AreEqual("1", ETable.Rows(1)(2))
            ClassicAssert.AreEqual("1", ETable.Rows(2)(2))
        End Sub

        <Test> Public Sub FormulaComplexityLevel2_BugReproduction()
            CreateSheetWithReproducableBug_FormulaComplexityLevel2()
        End Sub

        Private Function CreateSheetWithReproducableBug_FormulaComplexityLevel2() As ExcelOps.ExcelDataOperationsBase
            Dim TestFilePattern As String = "MsExcelNoCalcBug_" & EngineName.ToString & "_FormulaComplexityLevel2_{0}.xlsx"
            Dim TestFile As String = TestEnvironment.FullPathOfDynTestFile(GetType(MsExcelCalcTestBase), String.Format(TestFilePattern, "_01_InitialEpplus"))
            System.Console.WriteLine("Output path of test files: " & System.IO.Path.GetDirectoryName(TestFile))
            System.Console.WriteLine()

            'Create new Excel workbook with Epplus and add some cells with values and formulas
            Dim Eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateEngineInstanceWithCreateFileMode(TestFile)
            Dim FirstSheetName As String = Eppeo.SheetNames(0)
            Eppeo.AddSheet("Sheet2")

            Dim SecondSheetName As String = Eppeo.SheetNames(1)
            Eppeo.WriteCellValue(Of String)(SecondSheetName, 0, 0, "Personal-Monate")
            Eppeo.WriteCellValue(Of Integer)(SecondSheetName, 1, 0, 12)

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 1, 0, "Sofort-Epplus-Calc")
            Eppeo.WriteCellValue(Of String)(FirstSheetName, 2, 0, "No-Epplus-Calc")

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 0, 1, "Personal-Monate")
            Eppeo.WriteCellFormula(FirstSheetName, 1, 1, "Sheet2!A2", True)
            Eppeo.WriteCellFormula(FirstSheetName, 2, 1, "Sheet2!A2", False)

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 0, 2, "Plausi")
            Eppeo.WriteCellFormula(FirstSheetName, 1, 2, "IF(B2=0,0,B2-SUM(D2:O2))", False)
            Eppeo.WriteCellFormula(FirstSheetName, 2, 2, "IF(B3=0,0,B3-SUM(D3:O3))", False)

            For MyCounter As Integer = 0 To 12 - 1
                Eppeo.WriteCellValue(Of String)(FirstSheetName, 0, 3 + MyCounter, "Monat " & MyCounter + 1)
                Eppeo.WriteCellValue(Of Integer)(FirstSheetName, 1, 3 + MyCounter, 1)
                Eppeo.WriteCellValue(Of Integer)(FirstSheetName, 2, 3 + MyCounter, 1)
            Next

            Eppeo.CalculationModuleDisabled = False
            Eppeo.Save(ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)

            'Create workbook copy, open and recalculate and save in MS Excel
            TestFile = TestEnvironment.FullPathOfDynTestFile(GetType(MsExcelCalcTestBase), String.Format(TestFilePattern, "_02_ReSavedByMsExcel"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()
            Try
                CompuMaster.Excel.ExcelOps.MsVsEpplusTools.OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(TestFile)
            Catch ex As System.PlatformNotSupportedException
                ClassicAssert.Ignore("Platform not supported or MS Excel app not installed: " & ex.Message)
            End Try

            'Update single cells in calculated workbook with Epplus
            Eppeo.ReloadFromFile()

            Eppeo.WriteCellValue(Of Integer)(FirstSheetName, 1, 2 + 1, 0)
            Eppeo.WriteCellValue(Of Integer)(FirstSheetName, 2, 2 + 1, 0)

            TestFile = TestEnvironment.FullPathOfDynTestFile(GetType(MsExcelCalcTestBase), String.Format(TestFilePattern, "_03_UpdatedByEpplus"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()

            'Compare expected values
            Dim ETable As DataTable = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(TestFile, FirstSheetName, False)
            Dim TTable As New ExcelOps.TextTable(ETable)
            System.Console.WriteLine(TTable.ToUIExcelTable)
            ClassicAssert.AreEqual("0", ETable.Rows(1)(2))
            ClassicAssert.AreEqual("0", ETable.Rows(2)(2))

            'System.Diagnostics.Process.Start(TestFile)
            Return Eppeo

        End Function
#End Region

    End Class

End Namespace