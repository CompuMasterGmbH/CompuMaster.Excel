Imports NUnit.Framework
Imports NUnit.Framework.Legacy

Namespace ExcelOpsEngineTests
    <TestFixture> Public Class ToolsTest

        <OneTimeSetUp()> Public Sub OneTimeInitConfig()
        End Sub

        <SetUp> Public Sub Setup()
            CompuMaster.Excel.Test.Console.ResetConsoleForTestOutput()
        End Sub

        <Test> Public Sub IsFormulaSimpleSumFunction()
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaSimpleSumFunction("SUM(E8:E46)"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaSimpleCellReference("14100-(104.1*12)"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaSimpleSumFunction("E46"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaSimpleSumFunction("Grunddaten!E46"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaSimpleSumFunction("'Renta Plan'!E46"))
        End Sub

        <Test> Public Sub IsFormulaSimpleCellReference()
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaSimpleCellReference("E46"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaSimpleCellReference("Grunddaten!E46"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaSimpleCellReference("'Renta Plan'!E46"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaSimpleCellReference("SUM(E8:E46)"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaSimpleCellReference("14100-(104.1*12)"))
        End Sub

        <Test> Public Sub IsFormulaWithoutCellReferences()
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("14100-(104.1*12)"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("20574-(267*12)"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("30500-(355.1*12)"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*2"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52+2"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52^2"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52-2"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52/2"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52.9712343241/2.4"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("A52*2"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*A2"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*TestSheet!A2"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*'TestSheet'!A2"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*XFD2"))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*XFF2")) 'invalid column part, so no a cell reference -> test successful!
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("=52*2")) 'invalid formula, but without a cell reference -> test successful!
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferences("52\2"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*$XFD2"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*XFD$2"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("52*$XFD$2"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("SUM(B2,C3)"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("SUM(B2:C3)"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("$B2"))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferences("$B$2"))
            ClassicAssert.True(ExcelOps.Tools.IsFormulaWithoutCellReferences("$B$$2"))
            ClassicAssert.True(ExcelOps.Tools.IsFormulaWithoutCellReferences("B2$"))
            ClassicAssert.True(ExcelOps.Tools.IsFormulaWithoutCellReferences("$Bull$hit2"))
        End Sub

        <Test> Public Sub IsFormulaWithoutCellReferencesOrCellReferenceInSameRow()
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("14100-(104.1*12)", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("20574-(267*12)", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("30500-(355.1*12)", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*2", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52+2", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52^2", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52-2", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52/2", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52.9712343241/2.4", 1))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("A52*2", 1))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("TestSheet!A52*2", 1))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("'TestSheet'!A52*2", 1))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*TestSheet!A2", 1))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*'TestSheet'!A2", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*A2", 1)) 'cell in same row
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*XFD2", 1)) 'cell in same row
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*XFF2", 1)) 'invalid column part, so no a cell reference -> test successful!
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("=52*2", 1)) 'invalid formula, but without a cell reference -> test successful!
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52\2", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*$XFD2", 1)) 'cell in same row
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*XFD$2", 1)) 'cell in same row
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("52*$XFD$2", 1)) 'cell in same row
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("SUM(B2,C3)", 1))
            ClassicAssert.IsFalse(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("SUM(B2:C3)", 1))
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("$B2", 1)) 'cell in same row
            ClassicAssert.IsTrue(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("$B$2", 1)) 'cell in same row
            ClassicAssert.True(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("$B$$2", 1))
            ClassicAssert.True(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("B2$", 1))
            ClassicAssert.True(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("$Bull$hit2", 1))
            ClassicAssert.True(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("E15-8625", 14))
            ClassicAssert.False(ExcelOps.Tools.IsFormulaWithoutCellReferencesOrCellReferenceInSameRow("E15-8625", 13))
        End Sub

        <Test> Public Sub TryParseToDoubleCultureSafe()
            ClassicAssert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("lakdfj0.000,00", Nothing)) 'invalid chars
            ClassicAssert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("+0.000,00", Nothing)) 'invalid char "+"
            ClassicAssert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("0.000,000,00", Nothing)) 'too many decimal separators

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1.000.000,00", Nothing))
            ClassicAssert.AreEqual(1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1.000.000,00"))

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,000,000.00", Nothing))
            ClassicAssert.AreEqual(1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,000,000.00"))

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe(" 1,000,000.00 ", Nothing))
            ClassicAssert.AreEqual(1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe(" 1,000,000.00 "))

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("-1.000.000,00", Nothing))
            ClassicAssert.AreEqual(-1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("-1.000.000,00"))

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("-1,000,000.00", Nothing))
            ClassicAssert.AreEqual(-1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("-1,000,000.00"))

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("-1,000,000.", Nothing))
            ClassicAssert.AreEqual(-1000000.0, ExcelOps.Tools.ParseToDoubleCultureSafe("-1,000,000."))

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1", Nothing))
            ClassicAssert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1"))

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,", Nothing))
            ClassicAssert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,"))

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,0", Nothing))
            ClassicAssert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,0"))

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,00", Nothing))
            ClassicAssert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,00"))

            ClassicAssert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,000", Nothing)) 'not safe!

            ClassicAssert.IsTrue(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,0000", Nothing))
            ClassicAssert.AreEqual(1.0, ExcelOps.Tools.ParseToDoubleCultureSafe("1,0000"))

            ClassicAssert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,000,00", Nothing)) '"," must be a decimal separtor, but used as thousands separator, too
            ClassicAssert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,000,000,00", Nothing)) '"," must be a decimal separtor, but used as thousands separator, too
            ClassicAssert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1.000,000,000,00", Nothing)) '"," must be a decimal separtor, but used as thousands separator, too + another thousands separator
            ClassicAssert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,00,000,000", Nothing)) '"," must be a thousands separtor, but missing digits between thousands separators
            ClassicAssert.IsFalse(ExcelOps.Tools.TryParseToDoubleCultureSafe("1,00,000", Nothing)) '"," must be a thousands separtor, but missing digits between thousands separators
        End Sub

        <Test> Public Sub HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator()
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000,0", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("-1,000,0", New KeyValuePair(Of Char, Integer)(","c, 2)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1.000.0", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("-1.000.0", New KeyValuePair(Of Char, Integer)(","c, 2)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1.000,0", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("-1.000,0", New KeyValuePair(Of Char, Integer)(","c, 2)))

            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(",000", New KeyValuePair(Of Char, Integer)(","c, 0)))
            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("-,000", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000,0", New KeyValuePair(Of Char, Integer)(","c, 5)))

            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00,000,000", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00,000,000", New KeyValuePair(Of Char, Integer)(","c, 4)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00,000,000", New KeyValuePair(Of Char, Integer)(","c, 8)))

            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(",", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,0", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000,", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsTrue(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,000.", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,0000", New KeyValuePair(Of Char, Integer)(","c, 1)))
            ClassicAssert.IsFalse(ExcelOps.Tools.HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator("1,00000", New KeyValuePair(Of Char, Integer)(","c, 1)))
        End Sub

        <Test> Public Sub DetectXlsxDateSystem()
            'default file
            ClassicAssert.AreEqual(ExcelOps.Tools.XlsxDateSystem.Date1900, ExcelOps.Tools.DetectXlsxDateSystem(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelBaseDateSystem_windows_1900.xlsx")))

            'created manually by hardcoding the date1904 attribute in the zip's workbook.xml
            ClassicAssert.AreEqual(ExcelOps.Tools.XlsxDateSystem.Date1904, ExcelOps.Tools.DetectXlsxDateSystem(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelBaseDateSystem_mac_1904.xlsx")))

            'created by Excel 16.0 (MS Office 365, 2024-05)
            ClassicAssert.AreEqual(ExcelOps.Tools.XlsxDateSystem.Date1904, ExcelOps.Tools.DetectXlsxDateSystem(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelBaseDateSystem_mac_1904b.xlsx")))
        End Sub

        <Test> Public Sub ConvertExcelDateToDateTime()
            ClassicAssert.AreEqual(New DateTime(1900, 1, 1), ExcelOps.Tools.ConvertExcelDateToDateTime(2.0, ExcelOps.Tools.XlsxDateSystem.Date1900))
            ClassicAssert.AreEqual(New DateTime(1904, 1, 1), ExcelOps.Tools.ConvertExcelDateToDateTime(0.0, ExcelOps.Tools.XlsxDateSystem.Date1904))
        End Sub

        <Test> Public Sub LookupCellAddresFromRange()
            Dim LookupMode As ExcelOps.Tools.LookupCellAddresFromRangeMode

            LookupMode = ExcelOps.Tools.LookupCellAddresFromRangeMode.FirstCell
            ClassicAssert.AreEqual("B2", ExcelOps.Tools.LookupCellAddresFromRange("B2:C4", LookupMode))
            ClassicAssert.AreEqual("B2", ExcelOps.Tools.LookupCellAddresFromRange("B2", LookupMode))
            Assert.Throws(Of ArgumentNullException)(Sub() ExcelOps.Tools.LookupCellAddresFromRange(Nothing, LookupMode))
            'ClassicAssert.AreEqual("", ExcelOps.Tools.LookupCellAddresFromRange("", LookupMode))
            Assert.Throws(Of ArgumentNullException)(Sub() ExcelOps.Tools.LookupCellAddresFromRange("", LookupMode))

            Assert.Throws(Of ArgumentException)(Sub() ExcelOps.Tools.LookupCellAddresFromRange("A1:B2:C3", LookupMode))
            Assert.Throws(Of ArgumentOutOfRangeException)(Sub() ExcelOps.Tools.LookupCellAddresFromRange("A1:B2", CType(LookupMode + 2, ExcelOps.Tools.LookupCellAddresFromRangeMode)))

            LookupMode = ExcelOps.Tools.LookupCellAddresFromRangeMode.LastCell
            ClassicAssert.AreEqual("C4", ExcelOps.Tools.LookupCellAddresFromRange("B2:C4", LookupMode))
            ClassicAssert.AreEqual("B2", ExcelOps.Tools.LookupCellAddresFromRange("B2", LookupMode))
            Assert.Throws(Of ArgumentNullException)(Sub() ExcelOps.Tools.LookupCellAddresFromRange(Nothing, LookupMode))
            'ClassicAssert.AreEqual("", ExcelOps.Tools.LookupCellAddresFromRange("", LookupMode))
            Assert.Throws(Of ArgumentNullException)(Sub() ExcelOps.Tools.LookupCellAddresFromRange("", LookupMode))

        End Sub

    End Class

End Namespace