Imports NUnit.Framework

Namespace ExcelOpsEngineTests
    <TestFixture> Public Class ToolsTest

        <OneTimeSetUp()> Public Sub OneTimeInitConfig()
        End Sub

        <SetUp> Public Sub Setup()
            CompuMaster.Excel.Test.Console.ResetConsoleForTestOutput()
        End Sub

        <Test> Public Sub IsFormulaSimpleSumFunction()
            Assert.IsTrue(ExcelOps.Tools.IsFormulaSimpleSumFunction("SUM(E8:E46)"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaSimpleCellReference("14100-(104.1*12)"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaSimpleSumFunction("E46"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaSimpleSumFunction("Grunddaten!E46"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaSimpleSumFunction("'Renta Plan'!E46"))
        End Sub

        <Test> Public Sub IsFormulaSimpleCellReference()
            Assert.IsTrue(ExcelOps.Tools.IsFormulaSimpleCellReference("E46"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaSimpleCellReference("Grunddaten!E46"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaSimpleCellReference("'Renta Plan'!E46"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaSimpleCellReference("SUM(E8:E46)"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaSimpleCellReference("14100-(104.1*12)"))
        End Sub

        <Test> Public Sub IsFormulaWithoutCellReferences()
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("14100-(104.1*12)"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("20574-(267*12)"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("30500-(355.1*12)"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*2"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52+2"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52^2"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52-2"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52/2"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52.9712343241/2.4"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("A52*2"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*A2"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*TestSheet!A2"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*'TestSheet'!A2"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*XFD2"))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*XFF2")) 'invalid column part, so no a cell reference -> test successful!
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("=52*2")) 'invalid formula, but without a cell reference -> test successful!
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52\2"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*$XFD2"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*XFD$2"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*$XFD$2"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("SUM(B2,C3)"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("SUM(B2:C3)"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("$B2"))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("$B$2"))
            Assert.True(ExcelOps.Tools.IsFormulaWithoutCellReferences("$B$$2"))
            Assert.True(ExcelOps.Tools.IsFormulaWithoutCellReferences("B2$"))
            Assert.True(ExcelOps.Tools.IsFormulaWithoutCellReferences("$Bull$hit2"))
        End Sub

        <Test> Public Sub IsFormulaWithoutCellReferencesOrCellReferenceInSameRow()
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("14100-(104.1*12)", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("20574-(267*12)", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("30500-(355.1*12)", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*2", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52+2", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52^2", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52-2", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52/2", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52.9712343241/2.4", 1))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("A52*2", 1))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("TestSheet!A52*2", 1))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("'TestSheet'!A52*2", 1))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*TestSheet!A2", 1))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*'TestSheet'!A2", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*A2", 1)) 'cell in same row
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*XFD2", 1)) 'cell in same row
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*XFF2", 1)) 'invalid column part, so no a cell reference -> test successful!
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("=52*2", 1)) 'invalid formula, but without a cell reference -> test successful!
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52\2", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*$XFD2", 1)) 'cell in same row
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*XFD$2", 1)) 'cell in same row
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*$XFD$2", 1)) 'cell in same row
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("SUM(B2,C3)", 1))
            Assert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("SUM(B2:C3)", 1))
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("$B2", 1)) 'cell in same row
            Assert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("$B$2", 1)) 'cell in same row
            Assert.True(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("$B$$2", 1))
            Assert.True(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("B2$", 1))
            Assert.True(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("$Bull$hit2", 1))
            Assert.True(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("E15-8625", 14))
            Assert.False(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("E15-8625", 13))
        End Sub

        <Test> Public Sub TryParseToDoubleCultureSafe()
            Assert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("lakdfj0.000,00", Nothing)) 'invalid chars
            Assert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("+0.000,00", Nothing)) 'invalid char "+"
            Assert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("0.000,000,00", Nothing)) 'too many decimal separators

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1.000.000,00", Nothing))
            Assert.AreEqual(1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1.000.000,00"))

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,000,000.00", Nothing))
            Assert.AreEqual(1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,000,000.00"))

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe(" 1,000,000.00 ", Nothing))
            Assert.AreEqual(1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe(" 1,000,000.00 "))

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("-1.000.000,00", Nothing))
            Assert.AreEqual(-1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("-1.000.000,00"))

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("-1,000,000.00", Nothing))
            Assert.AreEqual(-1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("-1,000,000.00"))

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("-1,000,000.", Nothing))
            Assert.AreEqual(-1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("-1,000,000."))

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1", Nothing))
            Assert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1"))

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,", Nothing))
            Assert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,"))

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,0", Nothing))
            Assert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,0"))

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,00", Nothing))
            Assert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,00"))

            Assert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,000", Nothing)) 'not safe!

            Assert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,0000", Nothing))
            Assert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,0000"))

            Assert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,000,00", Nothing)) '"," must be a decimal separtor, but used as thousands separator, too
            Assert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,000,000,00", Nothing)) '"," must be a decimal separtor, but used as thousands separator, too
            Assert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1.000,000,000,00", Nothing)) '"," must be a decimal separtor, but used as thousands separator, too + another thousands separator
            Assert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,00,000,000", Nothing)) '"," must be a thousands separtor, but missing digits between thousands separators
            Assert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,00,000", Nothing)) '"," must be a thousands separtor, but missing digits between thousands separators
        End Sub

        <Test> Public Sub HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator()
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000,0", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("-1,000,0", New KeyValuePair(Of Char, Integer)(","c, 2)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1.000.0", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("-1.000.0", New KeyValuePair(Of Char, Integer)(","c, 2)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1.000,0", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("-1.000,0", New KeyValuePair(Of Char, Integer)(","c, 2)))

            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(",000", New KeyValuePair(Of Char, Integer)(","c, 0)))
            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("-,000", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000,0", New KeyValuePair(Of Char, Integer)(","c, 5)))

            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00,000,000", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00,000,000", New KeyValuePair(Of Char, Integer)(","c, 4)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00,000,000", New KeyValuePair(Of Char, Integer)(","c, 8)))

            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(",", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,0", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000,", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000.", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,0000", New KeyValuePair(Of Char, Integer)(","c, 1)))
            Assert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00000", New KeyValuePair(Of Char, Integer)(","c, 1)))
        End Sub

        <Test> Public Sub DetectXlsxDateSystem()
            'default file
            Assert.AreEqual(ExcelOps.Tools.XlsxDateSystem.Date1900, ExcelOps.Tools.DetectXlsxDateSystem(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelBaseDateSystem_windows_1900.xlsx")))

            'created manually by hardcoding the date1904 attribute in the zip's workbook.xml
            Assert.AreEqual(ExcelOps.Tools.XlsxDateSystem.Date1904, ExcelOps.Tools.DetectXlsxDateSystem(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelBaseDateSystem_mac_1904.xlsx")))

            'created by Excel 16.0 (MS Office 365, 2024-05)
            Assert.AreEqual(ExcelOps.Tools.XlsxDateSystem.Date1904, ExcelOps.Tools.DetectXlsxDateSystem(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelBaseDateSystem_mac_1904b.xlsx")))
        End Sub

        <Test> Public Sub ConvertExcelDateToDateTime()
            Assert.AreEqual(New DateTime(1900, 1, 1), ExcelOps.Tools.ConvertExcelDateToDateTime(2.0, ExcelOps.Tools.XlsxDateSystem.Date1900))
            Assert.AreEqual(New DateTime(1904, 1, 1), ExcelOps.Tools.ConvertExcelDateToDateTime(0.0, ExcelOps.Tools.XlsxDateSystem.Date1904))
        End Sub

    End Class

End Namespace