﻿Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps
Imports System.Data

Namespace ExcelOpsTests.MsExcelSpecials

    '<TestFixture(Explicit:=True, IgnoreReason:="MS Excel not supported on Non-Windows platforms")> Public Class MsExcelAutoCalcDoesntWorkButOnlyManualCalcWorksBug
    <NonParallelizable>
    <TestFixture>
    Public Class MsExcelAutoCalcDoesntWorkButOnlyManualCalcWorksBug

        Public Shared ReadOnly Property EngineTest() As TestEngines()
            Get
                Return TestTools.EnumValues(Of TestEngines).ToArray
            End Get
        End Property

        Public Enum TestEngines As Byte
            Epplus45LgplEdition
            EpplusPolyformLicenseEdition
            FreeSpireXls
        End Enum

        Private Function CreateEngineInstance(engineType As TestEngines, testFile As String) As ExcelOps.ExcelDataOperationsBase
            Select Case engineType
                Case TestEngines.Epplus45LgplEdition
                    Return New ExcelOps.EpplusFreeExcelDataOperations(testFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, String.Empty)
                Case TestEngines.EpplusPolyformLicenseEdition
                    ExcelOpsTests.Engines.EpplusPolyformEditionOpsTest.AssignLicenseContext()
                    Return New ExcelOps.EpplusPolyformExcelDataOperations(testFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, String.Empty)
                Case TestEngines.FreeSpireXls
                    Return New ExcelOps.FreeSpireXlsDataOperations(testFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, String.Empty)
                Case Else
                    Throw New NotImplementedException(engineType.ToString)
            End Select
        End Function

        <SetUp> Public Sub ResetConsoleForTestOutput()
            CompuMaster.Excel.Test.Console.ResetConsoleForTestOutput()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
        End Sub

        Private Sub EngineResetCellValueFromFormulaCell(engine As TestEngines, wb As ExcelOps.ExcelDataOperationsBase, sheetName As String, rowIndex As Integer, columnIndex As Integer)
            Select Case engine
                Case TestEngines.Epplus45LgplEdition
                    CType(wb, ExcelOps.EpplusFreeExcelDataOperations).ResetCellValueFromFormulaCell(sheetName, rowIndex, columnIndex)
                Case TestEngines.EpplusPolyformLicenseEdition
                    Assert.Ignore("Test not applicable for engine " & engine.ToString)
                    'CType(wb, ExcelOps.EpplusPolyformExcelDataOperations).ResetCellValueFromFormulaCell(sheetName, rowIndex, columnIndex)
                Case TestEngines.FreeSpireXls
                    Assert.Ignore("Test not applicable for engine " & engine.ToString)
                    'CType(wb, ExcelOps.FreeSpireXlsDataOperations).ResetCellValueFromFormulaCell(sheetName, rowIndex, columnIndex)
                Case Else
                    Throw New NotImplementedException(engine.ToString)
            End Select
        End Sub

#Region "Test Sample 1"
        <Test>
        Public Sub FormulaComplexityLevel1_Solution(<ValueSource(NameOf(EngineTest))> testEngine As TestEngines)
            Dim Eppeo As CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase = CreateSheetWithReproducableBug_FormulaComplexityLevel1(testEngine)

            'Solve buggy cells in Excel workbook with Epplus
            Eppeo.ReloadFromFile()
            Dim FirstSheetName As String = Eppeo.SheetNames(0)

            System.Console.WriteLine("Formula B2 BEFORE RESET=" & Eppeo.LookupCellFormula(FirstSheetName, 1, 1))
            EngineResetCellValueFromFormulaCell(testEngine, Eppeo, FirstSheetName, 1, 1)
            Assert.IsNotNull(Eppeo.LookupCellFormula(FirstSheetName, 1, 1))
            Assert.IsNotEmpty(Eppeo.LookupCellFormula(FirstSheetName, 1, 1))
            System.Console.WriteLine("Formula B2 AFTER RESET=" & Eppeo.LookupCellFormula(FirstSheetName, 1, 1))
            System.Console.WriteLine()
            EngineResetCellValueFromFormulaCell(testEngine, Eppeo, FirstSheetName, 2, 1)
            EngineResetCellValueFromFormulaCell(testEngine, Eppeo, FirstSheetName, 4, 1)
            EngineResetCellValueFromFormulaCell(testEngine, Eppeo, FirstSheetName, 5, 1)

            Dim TestFilePattern As String = "MsExcelNoCalcBug_" & testEngine.ToString & "_FormulaComplexityLevel1{0}.xlsx"
            Dim TestFile As String = TestEnvironment.FullPathOfDynTestFile(String.Format(TestFilePattern, "_11_FixedInEpplus"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)

            'Create workbook copy, open and recalculate and save in MS Excel
            TestFile = TestEnvironment.FullPathOfDynTestFile(String.Format(TestFilePattern, "_12_ReSavedByMsExcel"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()
            Try
                CompuMaster.Excel.ExcelOps.MsVsEpplusTools.OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(TestFile)
            Catch ex As System.PlatformNotSupportedException
                Assert.Ignore("Platform not supported or MS Excel app not installed: " & ex.Message)
            End Try

            'Compare expected values
            Dim ETable As DataTable = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(TestFile, FirstSheetName, False)
            Dim TTable As New ExcelOps.TextTable(ETable)
            System.Console.WriteLine(TTable.ToUIExcelTable)
            Assert.AreEqual(20, ETable.Rows(0)(1))
            Assert.AreEqual(20, ETable.Rows(1)(1))
            Assert.AreEqual(20, ETable.Rows(2)(1))
            Assert.AreEqual(20, ETable.Rows(4)(1))
            Assert.AreEqual(20, ETable.Rows(5)(1))
        End Sub

        <Test> Public Sub FormulaComplexityLevel1_BugReproduction(<ValueSource(NameOf(EngineTest))> testEngine As TestEngines)
            CreateSheetWithReproducableBug_FormulaComplexityLevel1(testEngine)
        End Sub

        Private Function CreateSheetWithReproducableBug_FormulaComplexityLevel1(testEngine As TestEngines) As ExcelOps.ExcelDataOperationsBase
            Dim TestFilePattern As String = "MsExcelNoCalcBug_" & testEngine.ToString & "_FormulaComplexityLevel1{0}.xlsx"
            Dim TestFile As String = TestEnvironment.FullPathOfDynTestFile(String.Format(TestFilePattern, "_01_InitialEpplus"))
            System.Console.WriteLine("Output path of test files: " & System.IO.Path.GetDirectoryName(TestFile))
            System.Console.WriteLine()

            'Create new Excel workbook with Epplus and add some cells with values and formulas
            Dim Eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateEngineInstance(testEngine, TestFile)
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

            Select Case testEngine
                Case TestEngines.Epplus45LgplEdition
                    Assert.AreEqual(True, Eppeo.CalculationModuleDisabled)
                    Assert.Throws(Of FeatureDisabledException)(Sub() Eppeo.Save(ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset))
                Case Else
                    Assert.AreEqual(False, Eppeo.CalculationModuleDisabled)
            End Select
            Eppeo.CalculationModuleDisabled = False
            Assert.AreEqual(False, Eppeo.CalculationModuleDisabled)
            Eppeo.Save(ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)

            'Create workbook copy, open and recalculate and save in MS Excel
            TestFile = TestEnvironment.FullPathOfDynTestFile(String.Format(TestFilePattern, "_02_ReSavedByMsExcel"))
            Eppeo.CalculationModuleDisabled = False
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()
            Try
                CompuMaster.Excel.ExcelOps.MsVsEpplusTools.OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(TestFile)
            Catch ex As System.PlatformNotSupportedException
                Assert.Ignore("Platform not supported or MS Excel app not installed: " & ex.Message)
            End Try

            'Update single cells in calculated workbook with Epplus
            Eppeo.ReloadFromFile()

            Eppeo.WriteCellValue(Of String)(FirstSheetName, 0, 0, "Static value rewritten")
            Eppeo.WriteCellValue(Of Integer)(FirstSheetName, 0, 1, 20)

            TestFile = TestEnvironment.FullPathOfDynTestFile(String.Format(TestFilePattern, "_03_UpdatedByEpplus"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()

            'Compare expected values
            Dim ETable As DataTable = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(TestFile, FirstSheetName, False)
            Dim TTable As New ExcelOps.TextTable(ETable)
            System.Console.WriteLine(TTable.ToUIExcelTable)
            Assert.AreEqual(20, ETable.Rows(0)(1))
            Assert.AreEqual(50, ETable.Rows(1)(1))
            Assert.AreEqual(50, ETable.Rows(2)(1))
            Assert.AreEqual(50, ETable.Rows(4)(1))
            Assert.AreEqual(50, ETable.Rows(5)(1))

            'System.Diagnostics.Process.Start(TestFile)
            Return Eppeo

        End Function
#End Region

#Region "Test Sample 2"
        <Test> Public Sub FormulaComplexityLevel2_Solution(<ValueSource(NameOf(EngineTest))> testEngine As TestEngines)
            Dim Eppeo As ExcelOps.ExcelDataOperationsBase = CreateSheetWithReproducableBug_FormulaComplexityLevel2(testEngine)

            Eppeo.ReloadFromFile()
            Dim FirstSheetName As String = Eppeo.SheetNames(0)

            'Solve buggy cells in whole Excel workbook with Epplus by resetting all formula cells in all worksheets
            Dim TestFilePattern As String = "MsExcelNoCalcBug_" & testEngine.ToString & "_FormulaComplexityLevel2{0}.xlsx"
            Dim TestFile As String
            TestFile = TestEnvironment.FullPathOfDynTestFile(String.Format(TestFilePattern, "_12_ReSavedByMsExcel"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation) 'solution: reset all cell values in cells with formulas
            Eppeo.Close()

            'Open and recalculate and save in MS Excel
            Try
                CompuMaster.Excel.ExcelOps.MsVsEpplusTools.OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(TestFile)
            Catch ex As System.PlatformNotSupportedException
                Assert.Ignore("Platform not supported or MS Excel app not installed: " & ex.Message)
            End Try

            'Compare expected values
            Dim ETable As DataTable = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(TestFile, FirstSheetName, False)
            Dim TTable As New ExcelOps.TextTable(ETable)
            System.Console.WriteLine(TTable.ToUIExcelTable)
            Assert.AreEqual("1", ETable.Rows(1)(2))
            Assert.AreEqual("1", ETable.Rows(2)(2))
        End Sub

        <Test> Public Sub FormulaComplexityLevel2_BugReproduction(<ValueSource(NameOf(EngineTest))> testEngine As TestEngines)
            CreateSheetWithReproducableBug_FormulaComplexityLevel2(testEngine)
        End Sub

        Private Function CreateSheetWithReproducableBug_FormulaComplexityLevel2(testEngine As TestEngines) As ExcelOps.ExcelDataOperationsBase
            Dim TestFilePattern As String = "MsExcelNoCalcBug_" & testEngine.ToString & "_FormulaComplexityLevel2_{0}.xlsx"
            Dim TestFile As String = TestEnvironment.FullPathOfDynTestFile(String.Format(TestFilePattern, "_01_InitialEpplus"))
            System.Console.WriteLine("Output path of test files: " & System.IO.Path.GetDirectoryName(TestFile))
            System.Console.WriteLine()

            'Create new Excel workbook with Epplus and add some cells with values and formulas
            Dim Eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateEngineInstance(testEngine, TestFile)
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
            TestFile = TestEnvironment.FullPathOfDynTestFile(String.Format(TestFilePattern, "_02_ReSavedByMsExcel"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()
            Try
                CompuMaster.Excel.ExcelOps.MsVsEpplusTools.OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(TestFile)
            Catch ex As System.PlatformNotSupportedException
                Assert.Ignore("Platform not supported or MS Excel app not installed: " & ex.Message)
            End Try

            'Update single cells in calculated workbook with Epplus
            Eppeo.ReloadFromFile()

            Eppeo.WriteCellValue(Of Integer)(FirstSheetName, 1, 2 + 1, 0)
            Eppeo.WriteCellValue(Of Integer)(FirstSheetName, 2, 2 + 1, 0)

            TestFile = TestEnvironment.FullPathOfDynTestFile(String.Format(TestFilePattern, "_03_UpdatedByEpplus"))
            Eppeo.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Eppeo.Close()

            'Compare expected values
            Dim ETable As DataTable = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(TestFile, FirstSheetName, False)
            Dim TTable As New ExcelOps.TextTable(ETable)
            System.Console.WriteLine(TTable.ToUIExcelTable)
            Assert.AreEqual("0", ETable.Rows(1)(2))
            Assert.AreEqual("0", ETable.Rows(2)(2))

            'System.Diagnostics.Process.Start(TestFile)
            Return Eppeo

        End Function
#End Region

    End Class

End Namespace