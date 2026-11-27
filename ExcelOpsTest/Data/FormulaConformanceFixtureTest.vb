Imports System.Globalization
Imports System.IO
Imports System.Xml
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsOptions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework

<TestFixture>
Public Class FormulaConformanceFixtureTest

    Private Const ReferenceWorkbook As String = "FormulaReference.xlsx"
    Private Const CompanionWorkbook As String = "FormulaReferenceSource.xlsx"

    ' The checked-in caches were recorded by Microsoft Excel build 16.0.20625.20108.
    <TestCase(ReferenceWorkbook, "Cases", 26)>
    <TestCase(CompanionWorkbook, "Source Data", 1)>
    <TestCase("FormulaReference1900.xlsx", "Standard", 37)>
    <TestCase("FormulaReference1900.xlsx", "Lookup", 34)>
    <TestCase("FormulaReference1900.xlsx", "Dates", 15)>
    <TestCase("FormulaReference1904.xlsx", "Dates", 5)>
    <TestCase("FormulaReferenceCalendar1900.xlsx", "Calendar", 107)>
    <TestCase("FormulaReferenceCalendar1904.xlsx", "Calendar", 107)>
    <TestCase("FormulaReferenceWorkdays1900.xlsx", "Workdays", 204)>
    <TestCase("FormulaReferenceWorkdays1904.xlsx", "Workdays", 204)>
    <TestCase("FormulaReferenceDateText1900.xlsx", "Date text", 111)>
    <TestCase("FormulaReferenceDateText1904.xlsx", "Date text", 111)>
    <TestCase("FormulaReferenceDateDifference1900.xlsx", "Date differences", 399)>
    <TestCase("FormulaReferenceDateDifference1904.xlsx", "Date differences", 399)>
    <TestCase("FormulaReferenceYearFraction1900.xlsx", "Year fractions", 323)>
    <TestCase("FormulaReferenceYearFraction1904.xlsx", "Year fractions", 323)>
    <TestCase("FormulaReferenceLogical1900.xlsx", "Logical functions", 91)>
    <TestCase("FormulaReferenceLogical1900.xlsx", "Logical arrays", 13)>
    <TestCase("FormulaReferenceInformation1900.xlsx", "Information functions", 241)>
    <TestCase("FormulaReferenceInformation1900.xlsx", "Information arrays", 40)>
    <TestCase("FormulaReferenceStatistics1900.xlsx", "Order statistics", 388)>
    <TestCase("FormulaReferenceStatistics1900.xlsx", "Statistical arrays", 27)>
    <TestCase("FormulaReferenceDispersion1900.xlsx", "Summary statistics", 648)>
    <TestCase("FormulaReferenceDispersion1900.xlsx", "Mode arrays", 10)>
    <TestCase("FormulaReferenceMath1900.xlsx", "Math functions", 546)>
    <TestCase("FormulaReferenceMath1900.xlsx", "Math arrays", 55)>
    <TestCase("FormulaReferenceInteger1900.xlsx", "Integer functions", 801)>
    <TestCase("FormulaReferenceInteger1900.xlsx", "Integer arrays", 49)>
    <TestCase("FormulaReferenceCounting1900.xlsx", "Counting functions", 687)>
    <TestCase("FormulaReferenceCounting1900.xlsx", "Counting arrays", 13)>
    <TestCase("FormulaReferenceDelimiters1900.xlsx", "Delimiters functions", 211)>
    <TestCase("FormulaReferenceDelimiters1900.xlsx", "Delimiters arrays", 66)>
    <TestCase("FormulaReferenceConversion1900.xlsx", "Conversion functions", 448)>
    <TestCase("FormulaReferenceConversion1900.xlsx", "Conversion arrays", 16)>
    <TestCase("FormulaReferenceTextPositionLegacy.xlsx", "Text positions", 79)>
    <TestCase("FormulaReferenceTextPositionLegacy.xlsx", "Position arrays", 4)>
    <TestCase("FormulaReferenceTextPositionUnicode.xlsx", "Text positions", 79)>
    <TestCase("FormulaReferenceTextPositionUnicode.xlsx", "Position arrays", 4)>
    <TestCase("FormulaReferenceHorizontal1900.xlsx", "Horizontal functions", 93)>
    <TestCase("FormulaReferenceHorizontal1900.xlsx", "Horizontal arrays", 16)>
    <TestCase("FormulaReferenceVertical1900.xlsx", "Vertical functions", 25)>
    <TestCase("FormulaReferenceVertical1900.xlsx", "Vertical arrays", 5)>
    <TestCase("FormulaReferenceMatrix1900.xlsx", "Matrix functions", 95)>
    <TestCase("FormulaReferenceMatrix1900.xlsx", "Matrix arrays", 27)>
    <TestCase("FormulaReferenceIf1900.xlsx", "If functions", 33)>
    <TestCase("FormulaReferenceIf1900.xlsx", "If arrays", 30)>
    <TestCase("FormulaReferenceRadix1900.xlsx", "Radix functions", 769)>
    <TestCase("FormulaReferenceRadix1900.xlsx", "Radix arrays", 17)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "Roman functions", 105)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "Roman arrays", 4)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "All Roman forms", 1)>
    <TestCase("FormulaReferenceComplex1900.xlsx", "Complex functions", 501)>
    <TestCase("FormulaReferenceComplex1900.xlsx", "Complex arrays", 7)>
    <TestCase("FormulaReferenceComplexMath1900.xlsx", "ComplexMath functions", 376)>
    <TestCase("FormulaReferenceComplexMath1900.xlsx", "ComplexMath arrays", 3)>
    <TestCase("FormulaReferenceLoan1900.xlsx", "Loan functions", 392)>
    <TestCase("FormulaReferenceLoan1900.xlsx", "Loan arrays", 7)>
    <TestCase("FormulaReferenceAmortization1900.xlsx", "Amortization functions", 588)>
    <TestCase("FormulaReferenceAmortization1900.xlsx", "Amortization arrays", 8)>
    <TestCase("FormulaReferenceCashflow1900.xlsx", "Cashflow functions", 508)>
    <TestCase("FormulaReferenceCashflow1900.xlsx", "Cashflow arrays", 6)>
    <TestCase("FormulaReferenceDatedCashflow1900.xlsx", "DatedCashflow functions", 650)>
    <TestCase("FormulaReferenceDatedCashflow1900.xlsx", "DatedCashflow arrays", 2)>
    <TestCase("FormulaReferenceDatedCashflow1904.xlsx", "DatedCashflow functions", 650)>
    <TestCase("FormulaReferenceDatedCashflow1904.xlsx", "DatedCashflow arrays", 2)>
    <TestCase("FormulaReferenceGrowth1900.xlsx", "Growth functions", 581)>
    <TestCase("FormulaReferenceGrowth1900.xlsx", "Growth arrays", 7)>
    <TestCase("FormulaReferenceSeries1900.xlsx", "Series functions", 300)>
    <TestCase("FormulaReferenceSeries1900.xlsx", "Series arrays", 2)>
    <TestCase("FormulaReferenceDepreciation1900.xlsx", "Depreciation functions", 980)>
    <TestCase("FormulaReferenceDepreciation1900.xlsx", "Depreciation arrays", 10)>
    <TestCase("FormulaReferenceVariableDepreciation1900.xlsx", "Variable depreciation", 3096)>
    <TestCase("FormulaReferenceNormal1900.xlsx", "Normal functions", 919)>
    <TestCase("FormulaReferenceNormal1900.xlsx", "Normal arrays", 8)>
    <TestCase("FormulaReferenceGamma1900.xlsx", "Gamma functions", 1506)>
    <TestCase("FormulaReferenceGamma1900.xlsx", "Gamma arrays", 5)>
    <TestCase("FormulaReferenceBeta1900.xlsx", "Beta functions", 1907)>
    <TestCase("FormulaReferenceBeta1900.xlsx", "Beta arrays", 5)>
    <TestCase("FormulaReferenceDiscrete1900.xlsx", "Discrete functions", 1845)>
    <TestCase("FormulaReferenceDiscrete1900.xlsx", "Discrete arrays", 5)>
    <TestCase("FormulaReferenceReferences1900.xlsx", "References functions", 579)>
    <TestCase("FormulaReferenceReferences1900.xlsx", "References arrays", 33)>
    <TestCase("FormulaReferenceReshape1900.xlsx", "Reshape functions", 622)>
    <TestCase("FormulaReferenceReshape1900.xlsx", "Reshape arrays", 215)>
    <TestCase("FormulaReferenceFits1900.xlsx", "Fits functions", 713)>
    <TestCase("FormulaReferenceFits1900.xlsx", "Fits arrays", 23)>
    <TestCase("FormulaReferenceGermanText1900.xlsx", "German formula text", 34)>
    <TestCase("FormulaReferenceHypothesis1900.xlsx", "Hypothesis functions", 1657)>
    <TestCase("FormulaReferenceHypothesis1900.xlsx", "Hypothesis arrays", 159)>
    <TestCase("FormulaReferenceSecurities1900.xlsx", "Securities functions", 1367)>
    <TestCase("FormulaReferenceSecurities1900.xlsx", "Securities arrays", 2)>
    <TestCase("FormulaReferenceSecurities1904.xlsx", "Securities functions", 1306)>
    <TestCase("FormulaReferenceSecurities1904.xlsx", "Securities arrays", 2)>
    <TestCase("FormulaReferenceThreshold1900.xlsx", "Threshold and URL", 267)>
    <TestCase("FormulaReferenceThreshold1900.xlsx", "Threshold arrays", 59)>
    <TestCase("FormulaReferencePercentRank1900.xlsx", "Percent rank", 532)>
    <TestCase("FormulaReferencePercentRank1900.xlsx", "Percent rank arrays", 96)>
    <TestCase("FormulaReferenceHistograms1900.xlsx", "Distribution results", 579)>
    <TestCase("FormulaReferenceHistograms1900.xlsx", "Distribution arrays", 234)>
    <TestCase("FormulaReferenceDatabase1900.xlsx", "Database results", 626)>
    <TestCase("FormulaReferenceDatabase1900.xlsx", "Database arrays", 138)>
    <TestCase("FormulaReferenceLambda1900.xlsx", "Lambda results", 337)>
    <TestCase("FormulaReferenceLambda1900.xlsx", "Lambda arrays", 161)>
    <TestCase("FormulaReferenceRowAggregates1900.xlsx", "Row visibility", 84)>
    <TestCase("FormulaReferenceRowAggregates1900.xlsx", "Filter modes", 524)>
    <TestCase("FormulaReferenceFormats1900.xlsx", "Formats", 146)>
    <TestCase("FormulaReferenceFormats1900.xlsx", "Array display", 98)>
    <TestCase("FormulaReferenceFormats1900.xlsx", "Links and XML", 87)>
    <TestCase("FormulaReferenceFormats1900.xlsx", "Byte text", 95)>
    <TestCase("FormulaReferenceFormats1900.xlsx", "Random rules", 29)>
    <TestCase("FormulaReferenceEngineering1900.xlsx", "Bessel values", 228)>
    <TestCase("FormulaReferenceEngineering1900.xlsx", "Bessel J values", 76)>
    <TestCase("FormulaReferenceEngineering1900.xlsx", "Unit examples", 38)>
    <TestCase("FormulaReferenceEngineering1900.xlsx", "Baht text", 43)>
    <TestCase("FormulaReferenceEngineering1900.xlsx", "Unit catalog", 144)>
    <TestCase("FormulaReferenceFrench1900.xlsx", "Straight-line depreciation", 1451)>
    <TestCase("FormulaReferenceFrench1900.xlsx", "Declining depreciation", 1425)>
    <TestCase("FormulaReferenceFrench1904.xlsx", "Straight-line depreciation", 1451)>
    <TestCase("FormulaReferenceFrench1904.xlsx", "Declining depreciation", 1425)>
    <TestCase("FormulaReferenceUnary1900.xlsx", "Unary plus", 44)>
    <TestCase("FormulaReferenceOddLast1900.xlsx", "Odd last coupon price", 1217)>
    <TestCase("FormulaReferenceOddLast1900.xlsx", "Odd last coupon yield", 1217)>
    <TestCase("FormulaReferenceOddLast1904.xlsx", "Odd last coupon price", 1217)>
    <TestCase("FormulaReferenceOddLast1904.xlsx", "Odd last coupon yield", 1217)>
    <TestCase("FormulaReferenceAccrual1900.xlsx", "Accrued periodic interest", 3031)>
    <TestCase("FormulaReferenceAccrual1904.xlsx", "Accrued periodic interest", 3031)>
    <TestCase("FormulaReferenceGrouping1900.xlsx", "GroupBy probes", 54)>
    <TestCase("FormulaReferenceGrouping1900.xlsx", "GroupBy shapes", 51)>
    <TestCase("FormulaReferencePivotGroups1900.xlsx", "PivotBy probes", 75)>
    <TestCase("FormulaReferencePivotGroups1900.xlsx", "PivotBy options", 38)>
    <TestCase("FormulaReferencePivotTables1900.xlsx", "GETPIVOTDATA", 16)>
    <TestCase("FormulaReferenceOddFirst1900.xlsx", "Odd first coupon price", 756)>
    <TestCase("FormulaReferenceOddFirst1900.xlsx", "Odd first coupon yield", 756)>
    <TestCase("FormulaReferenceOddFirst1904.xlsx", "Odd first coupon price", 756)>
    <TestCase("FormulaReferenceOddFirst1904.xlsx", "Odd first coupon yield", 756)>
    <TestCase("FormulaReferenceMetadata1900.xlsx", "Cell metadata", 173)>
    <TestCase("FormulaReferenceMetadata1900.xlsx", "Info metadata", 20)>
    <TestCase("FormulaReferenceAxes1900.xlsx", "Whole axes", 56)>
    <TestCase("FormulaReferenceAxes1900.xlsx", "Other worksheet", 8)>
    <TestCase("FormulaReferenceRegression1900.xlsx", "Fit coefficients", 210)>
    <TestCase("FormulaReferenceRegression1900.xlsx", "Forecast values", 230)>
    <TestCase("FormulaReferenceBonds1900.xlsx", "Coupon periods", 1140)>
    <TestCase("FormulaReferenceBonds1900.xlsx", "Bond valuations", 429)>
    <TestCase("FormulaReferenceBonds1900.xlsx", "Interest at maturity", 540)>
    <TestCase("FormulaReferenceBonds1900.xlsx", "Bond yields", 143)>
    <TestCase("FormulaReferenceBonds1904.xlsx", "Coupon periods", 1140)>
    <TestCase("FormulaReferenceBonds1904.xlsx", "Bond valuations", 429)>
    <TestCase("FormulaReferenceBonds1904.xlsx", "Interest at maturity", 540)>
    <TestCase("FormulaReferenceBonds1904.xlsx", "Bond yields", 143)>
    <TestCase("FormulaReferenceLookupVariants.xlsx", "Lookups", 128)>
    <TestCase("FormulaReferenceLookupVariants.xlsx", "Regex", 1)>
    <TestCase("FormulaReferenceRegexVariants.xlsx", "Regex lookups", 33)>
    <TestCase("FormulaReferenceRegexFunctions.xlsx", "Regex functions", 91)>
    <TestCase("FormulaReferenceCriteria.xlsx", "Criteria", 89)>
    <TestCase("FormulaReferenceArrayVariants.xlsx", "Array variants", 65)>
    <TestCase("FormulaReferenceTextVariants.xlsx", "Text variants", 97)>
    <TestCase("FormulaReferenceSortVariants.xlsx", "Sort and unique", 89)>
    <TestCase("FormulaReferenceModern.xlsx", "Arrays", 25)>
    <TestCase("FormulaReferenceModern.xlsx", "Local names", 16)>
    Public Sub RecalculatesTheExcelReferenceFixture(fileName As String, sheetName As String, caseCount As Integer)
        Dim fixture = FixturePath(fileName)
        Dim expected = ReadFormulaCaches(fixture, sheetName)
        Assert.That(expected.Count, [Is].EqualTo(caseCount))

        Dim output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Try
            Using workbook = OpenForCalculation(fixture)
                workbook.ResetCellValueFromFormulaCell(sheetName)
                For Each address In expected.Keys
                    Assert.That(workbook.HasMissingCachedFormulaResult(New ExcelCell(sheetName, address, ExcelCell.ValueTypes.All)), [Is].True, address)
                Next
                workbook.RecalculateSheet(sheetName)
                workbook.SaveCopyAs(output, recalculateWorkbook:=False, SaveOptionsForDisabledCalculationEngines.NoReset)
            End Using

            Dim actual = ReadFormulaCaches(output, sheetName)
            Assert.That(actual.Keys, [Is].EquivalentTo(expected.Keys), sheetName)
            Using assertions = Assert.EnterMultipleScope()
                For Each item In expected
                    Dim address = sheetName & "!" & item.Key
                    Dim result = actual(item.Key)
                    Assert.That(result.Kind, [Is].EqualTo(item.Value.Kind), address)
                    If result.Kind <> item.Value.Kind Then Continue For
                    If result.Kind = CellValues.Number Then
                        Assert.That(result.Value, [Is].EqualTo(item.Value.Value).Within(ExcelOpsTests.Engines.FormulaReferenceCoverage.ReferenceNumberTolerance(DirectCast(item.Value.Value, Double), fileName, sheetName)), address)
                    Else
                        Assert.That(result.Value, [Is].EqualTo(item.Value.Value), address)
                    End If
                Next
            End Using
            If fileName = ReferenceWorkbook Then
                Assert.That(ReadFormulaCaches(output, "External cases"), [Is].EqualTo(ReadFormulaCaches(fixture, "External cases")))
            End If
            If fileName = "FormulaReferenceModern.xlsx" AndAlso sheetName = "Arrays" Then
                AssertSameCalcErrors(fixture, output, sheetName, {"I4", "I5", "I21", "I24"})
            ElseIf fileName = "FormulaReferenceArrayVariants.xlsx" Then
                AssertSameCalcErrors(fixture, output, sheetName, {"S22", "S23", "S31", "S32"})
            ElseIf fileName = "FormulaReferenceTextVariants.xlsx" Then
                AssertSameCalcErrors(fixture, output, sheetName, {"S9", "S37"})
            ElseIf fileName = "FormulaReferenceSortVariants.xlsx" Then
                AssertSameCalcErrors(fixture, output, sheetName, {"S63"})
            End If
        Finally
            If File.Exists(output) Then File.Delete(output)
        End Try
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelRegexResultsAndTheirTypes()
        AssertSpilledFormulaResults("FormulaReferenceRegexFunctions.xlsx", "Array results", "B2", "C2", "D2", "B3", "C3", "D3", "B4", "C4")
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelConditionalResults()
        AssertSpilledFormulaResults("FormulaReferenceLogical1900.xlsx", "Logical arrays",
            "E2", "E6", "D7", "E7", "E10", "E18", "D19", "E19", "E22", "E26", "E30",
            "D31", "E31", "E34", "D43", "E46", "F46", "D47", "E47", "F47")
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelInformationResults()
        AssertSpilledFormulaResults("FormulaReferenceInformation1900.xlsx", "Information arrays",
            "E2", "F2", "G2", "D9", "D10", "D11", "E14", "D15", "E15", "D27", "E32",
            "F32", "G32", "D39", "D40", "D41", "E44", "D45", "E45", "D57", "E62", "F62",
            "G62", "E74", "D75", "E75", "E92", "F92", "G92", "E104", "D105", "E105", "E122",
            "F122", "G122", "E134", "D135", "E135", "E152", "F152", "G152", "E164", "D165", "E165",
            "E212", "F212", "G212", "D219", "D220", "D221", "E224", "D225", "E225", "D237")
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelOrderStatistics()
        AssertSpilledFormulaResults("FormulaReferenceStatistics1900.xlsx", "Statistical arrays",
            "E2", "F2", "D8", "D9", "D13", "D14", "E17", "F17", "D23", "D24",
            "D28", "D29", "E32", "E37", "F37", "D43", "D44", "D48", "D49", "E52",
            "E57", "F57", "D63", "D64", "D68", "D69", "E72", "E77", "F77", "D83",
            "D84", "D88", "D89", "E92", "F92", "G92", "H92", "E97", "F97", "G97",
            "H97", "E102", "F102", "G102", "H102", "E107", "F107", "G107", "H107", "E112",
            "F112", "G112", "H112", "E117", "F117", "G117", "H117")
    End Sub

    <Test>
    Public Sub RecalculatesTheSpilledExcelModes()
        AssertSpilledFormulaResults("FormulaReferenceDispersion1900.xlsx", "Mode arrays",
            "D3", "D8", "D13", "D18", "D19", "D23", "D24", "D28", "D33", "D34", "D38", "D39", "D43")
    End Sub

    <TestCase("FormulaReferenceMath1900.xlsx", "Math arrays", 144)>
    <TestCase("FormulaReferenceInteger1900.xlsx", "Integer arrays", 121)>
    <TestCase("FormulaReferenceCounting1900.xlsx", "Counting arrays", 56)>
    <TestCase("FormulaReferenceDelimiters1900.xlsx", "Delimiters arrays", 125)>
    <TestCase("FormulaReferenceConversion1900.xlsx", "Conversion arrays", 31)>
    <TestCase("FormulaReferenceTextPositionLegacy.xlsx", "Position arrays", 10)>
    <TestCase("FormulaReferenceTextPositionUnicode.xlsx", "Position arrays", 10)>
    <TestCase("FormulaReferenceHorizontal1900.xlsx", "Horizontal arrays", 28)>
    <TestCase("FormulaReferenceVertical1900.xlsx", "Vertical arrays", 7)>
    <TestCase("FormulaReferenceMatrix1900.xlsx", "Matrix arrays", 136)>
    <TestCase("FormulaReferenceIf1900.xlsx", "If arrays", 84)>
    <TestCase("FormulaReferenceRadix1900.xlsx", "Radix arrays", 49)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "Roman arrays", 30)>
    <TestCase("FormulaReferenceRoman1900.xlsx", "All Roman forms", 19999)>
    <TestCase("FormulaReferenceComplex1900.xlsx", "Complex arrays", 17)>
    <TestCase("FormulaReferenceComplexMath1900.xlsx", "ComplexMath arrays", 14)>
    <TestCase("FormulaReferenceLoan1900.xlsx", "Loan arrays", 14)>
    <TestCase("FormulaReferenceAmortization1900.xlsx", "Amortization arrays", 16)>
    <TestCase("FormulaReferenceCashflow1900.xlsx", "Cashflow arrays", 10)>
    <TestCase("FormulaReferenceDatedCashflow1900.xlsx", "DatedCashflow arrays", 2)>
    <TestCase("FormulaReferenceDatedCashflow1904.xlsx", "DatedCashflow arrays", 2)>
    <TestCase("FormulaReferenceGrowth1900.xlsx", "Growth arrays", 14)>
    <TestCase("FormulaReferenceSeries1900.xlsx", "Series arrays", 5)>
    <TestCase("FormulaReferenceDepreciation1900.xlsx", "Depreciation arrays", 19)>
    <TestCase("FormulaReferenceVariableDepreciation1900.xlsx", "Variable depreciation", 15)>
    <TestCase("FormulaReferenceNormal1900.xlsx", "Normal arrays", 20)>
    <TestCase("FormulaReferenceGamma1900.xlsx", "Gamma arrays", 20)>
    <TestCase("FormulaReferenceBeta1900.xlsx", "Beta arrays", 20)>
    <TestCase("FormulaReferenceDiscrete1900.xlsx", "Discrete arrays", 24)>
    <TestCase("FormulaReferenceReferences1900.xlsx", "References arrays", 85)>
    <TestCase("FormulaReferenceReshape1900.xlsx", "Reshape arrays", 1120)>
    <TestCase("FormulaReferenceFits1900.xlsx", "Fits arrays", 46)>
    <TestCase("FormulaReferenceHypothesis1900.xlsx", "Hypothesis arrays", 325)>
    <TestCase("FormulaReferenceSecurities1900.xlsx", "Securities arrays", 4)>
    <TestCase("FormulaReferenceSecurities1904.xlsx", "Securities arrays", 4)>
    <TestCase("FormulaReferenceThreshold1900.xlsx", "Threshold arrays", 118)>
    <TestCase("FormulaReferencePercentRank1900.xlsx", "Percent rank arrays", 528)>
    <TestCase("FormulaReferenceHistograms1900.xlsx", "Distribution arrays", 526)>
    <TestCase("FormulaReferenceDatabase1900.xlsx", "Database arrays", 63)>
    <TestCase("FormulaReferenceLambda1900.xlsx", "Lambda arrays", 409)>
    <TestCase("FormulaReferenceRowAggregates1900.xlsx", "Filter modes", 25)>
    <TestCase("FormulaReferenceFormats1900.xlsx", "Formats", 2)>
    <TestCase("FormulaReferenceFormats1900.xlsx", "Array display", 52)>
    <TestCase("FormulaReferenceFormats1900.xlsx", "Links and XML", 24)>
    <TestCase("FormulaReferenceFormats1900.xlsx", "Random rules", 10)>
    <TestCase("FormulaReferenceEngineering1900.xlsx", "Bessel values", 15)>
    <TestCase("FormulaReferenceEngineering1900.xlsx", "Bessel J values", 5)>
    <TestCase("FormulaReferenceEngineering1900.xlsx", "Unit examples", 3)>
    <TestCase("FormulaReferenceEngineering1900.xlsx", "Baht text", 1)>
    <TestCase("FormulaReferenceFrench1900.xlsx", "Straight-line depreciation", 4)>
    <TestCase("FormulaReferenceFrench1900.xlsx", "Declining depreciation", 4)>
    <TestCase("FormulaReferenceFrench1904.xlsx", "Straight-line depreciation", 4)>
    <TestCase("FormulaReferenceFrench1904.xlsx", "Declining depreciation", 4)>
    <TestCase("FormulaReferenceUnary1900.xlsx", "Unary plus", 14)>
    <TestCase("FormulaReferenceOddLast1900.xlsx", "Odd last coupon price", 6)>
    <TestCase("FormulaReferenceOddLast1900.xlsx", "Odd last coupon yield", 6)>
    <TestCase("FormulaReferenceOddLast1904.xlsx", "Odd last coupon price", 6)>
    <TestCase("FormulaReferenceOddLast1904.xlsx", "Odd last coupon yield", 6)>
    <TestCase("FormulaReferenceAccrual1900.xlsx", "Accrued periodic interest", 6)>
    <TestCase("FormulaReferenceAccrual1904.xlsx", "Accrued periodic interest", 6)>
    <TestCase("FormulaReferenceGrouping1900.xlsx", "GroupBy probes", 396)>
    <TestCase("FormulaReferenceGrouping1900.xlsx", "GroupBy shapes", 873)>
    <TestCase("FormulaReferencePivotGroups1900.xlsx", "PivotBy probes", 1982)>
    <TestCase("FormulaReferencePivotGroups1900.xlsx", "PivotBy options", 4732)>
    <TestCase("FormulaReferenceOddFirst1900.xlsx", "Odd first coupon price", 16)>
    <TestCase("FormulaReferenceOddFirst1900.xlsx", "Odd first coupon yield", 16)>
    <TestCase("FormulaReferenceOddFirst1904.xlsx", "Odd first coupon price", 16)>
    <TestCase("FormulaReferenceOddFirst1904.xlsx", "Odd first coupon yield", 16)>
    <TestCase("FormulaReferenceMetadata1900.xlsx", "Cell metadata", 10)>
    <TestCase("FormulaReferenceMetadata1900.xlsx", "Info metadata", 1)>
    <TestCase("FormulaReferenceRegression1900.xlsx", "Fit coefficients", 570)>
    <TestCase("FormulaReferenceRegression1900.xlsx", "Forecast values", 248)>
    Public Sub RecalculatesTheSpilledExcelResults(fileName As String, sheetName As String, expectedChildren As Integer)
        Dim addresses As New List(Of String)
        Using package = SpreadsheetDocument.Open(FixturePath(fileName), False)
            Dim workbook = package.WorkbookPart
            Dim sheet = workbook.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = sheetName)
            Dim worksheet = DirectCast(workbook.GetPartById(sheet.Id.Value), WorksheetPart).Worksheet
            For Each cell In worksheet.Descendants(Of Cell)()
                If IsFormulaAnchor(cell) OrElse cell.CellValue Is Nothing Then Continue For
                Dim address As New ExcelCell(sheetName, cell.CellReference.Value, ExcelCell.ValueTypes.All)
                If address.RowIndex > 0 AndAlso address.ColumnIndex >= 3 Then addresses.Add(cell.CellReference.Value)
            Next
        End Using
        Assert.That(addresses, Has.Count.EqualTo(expectedChildren))
        AssertSpilledFormulaResults(fileName, sheetName, addresses.ToArray())
    End Sub

    Private Shared Sub AssertSpilledFormulaResults(fileName As String, sheetName As String, ParamArray addresses As String())
        Dim path = FixturePath(fileName)
        Using expected = OpenForCalculation(path), calculated = OpenForCalculation(path)
            calculated.ResetCellValueFromFormulaCell(sheetName)
            calculated.RecalculateSheet(sheetName)
            Using assertions = Assert.EnterMultipleScope()
                For Each address In addresses
                    Dim cell As New ExcelCell(sheetName, address, ExcelCell.ValueTypes.All)
                    Dim actual = calculated.LookupCellValueAsObject(cell, FormulaCalculationMode.UseCachedValueOnly)
                    Dim reference = expected.LookupCellValueAsObject(cell, FormulaCalculationMode.UseCachedValueOnly)
                    If TypeOf actual Is Double AndAlso TypeOf reference Is Double Then
                        Assert.That(actual, [Is].EqualTo(reference).Within(ExcelOpsTests.Engines.FormulaReferenceCoverage.ReferenceNumberTolerance(DirectCast(reference, Double), fileName, sheetName)), address)
                    Else
                        Assert.That(actual, [Is].EqualTo(reference), address)
                    End If
                Next
            End Using
        End Using
    End Sub

    <Test, Explicit("Reports Excel formula coverage for Open XML.")>
    Public Sub ReportsFormulaReferenceCoverageAgainstExcel()
        Dim result = ExcelOpsTests.Engines.FormulaReferenceCoverage.Report("Open XML",
            Function(path, options) New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, options),
            Function(workbook, address) workbook.LookupCellValueAsObject(address, FormulaCalculationMode.CalculateIfMissing))
        Assert.That(result.Matched, [Is].EqualTo(result.Total))
    End Sub

    ' External results retain their Excel caches, and the companion link is relative to the fixture.
    <Test>
    Public Sub KeepsTheExternalReferenceCasesPortable()
        Dim fixture = FixturePath(ReferenceWorkbook)
        Using package = SpreadsheetDocument.Open(fixture, False)
            Dim links = package.WorkbookPart.ExternalWorkbookParts.Single().ExternalRelationships.ToArray()
            Assert.That(links, [Is].Not.Empty)
            For Each link In links
                Assert.That(link.Uri.OriginalString, [Is].EqualTo(CompanionWorkbook))
                Assert.That(link.Uri.IsAbsoluteUri, [Is].False)
            Next
            Assert.That(File.Exists(Path.Combine(Path.GetDirectoryName(fixture), CompanionWorkbook)), [Is].True)
        End Using

        Using workbook = OpenForCalculation(fixture)
            Const sheetName As String = "External cases"
            Dim expected = ReadFormulaCaches(fixture, sheetName)
            Assert.That(expected.Count, [Is].EqualTo(3))
            For Each item In expected
                Dim address As New ExcelCell(sheetName, item.Key, ExcelCell.ValueTypes.All)
                Assert.That(workbook.LookupCellValueAsObject(address, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(item.Value.Value), item.Key)
            Next

            Dim firstLink As New ExcelCell(sheetName, "C2", ExcelCell.ValueTypes.All)
            workbook.ResetCellValueFromFormulaCell(firstLink)
            Assert.Throws(Of FormulaComplexityNotSupportedException)(
                Sub() workbook.LookupCellValueAsObject(firstLink, FormulaCalculationMode.CalculateIfMissing))
        End Using
    End Sub

    <TestCase(True)>
    <TestCase(False)>
    Public Sub RecalculatesExcelExternalReferencesFromOpenSource(usePaths As Boolean)
        Dim session As New ExcelFormulaSession()
        Dim fixture = FixturePath(ReferenceWorkbook)
        Dim sourcePath = FixturePath(CompanionWorkbook)
        Using workbook = OpenInSession(fixture, session, usePaths)
            Using source = OpenInSession(sourcePath, session, usePaths)
                workbook.ResetCellValueFromFormulaCell("External cases")
                workbook.RecalculateSheet("External cases")
                Assert.That(ExternalValues(workbook), [Is].EqualTo({42.0, 100.0, 46.0}))
                source.WriteCellValue("Source Data", 1, 0, 45.0)
                workbook.RecalculateSheet("External cases")
                Assert.That(ExternalValues(workbook), [Is].EqualTo({45.0, 106.0, 49.0}))
            End Using
            Assert.Throws(Of FormulaComplexityNotSupportedException)(Sub() workbook.RecalculateSheet("External cases"))
            Assert.That(ExternalValues(workbook), [Is].EqualTo({45.0, 106.0, 49.0}))
        End Using
    End Sub

    <Test>
    Public Sub RecalculatingASaveCopyUsesSessionSourcesAndKeepsTheirCaches()
        Dim fixture = FixturePath(ReferenceWorkbook)
        Dim sourcePath = FixturePath(CompanionWorkbook)
        Dim outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".xlsx")
        Dim session As New ExcelFormulaSession()
        session.RegisterFunction("APP.READ",
            Function(context, arguments)
                Assert.That(context.Workbook.Name, [Is].EqualTo(ReferenceWorkbook))
                Dim sourceCell As New ExcelCell("Source Data", "D1", ExcelCell.ValueTypes.All)
                Return context.WorkbookSession.FindWorkbook(CompanionWorkbook).ReadCell(sourceCell)
            End Function)

        Try
            Using target = OpenInSession(fixture, session, usePath:=True), source = OpenInSession(sourcePath, session, usePath:=True)
                source.WriteCellValue("Source Data", 1, 0, 45.0)
                source.WriteCellFormula("Source Data", 0, 3, "A2+3", False)
                target.WriteCellFormula("External cases", 0, 7, "APP.READ()", False)

                target.SaveCopyAs(outputPath, recalculateWorkbook:=True, SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.That(ExternalValues(target), [Is].EqualTo({42.0, 100.0, 46.0}))
                Assert.That(source.HasMissingCachedFormulaResult("Source Data", 0, 3), [Is].True)
                Assert.That(target.HasMissingCachedFormulaResult("External cases", 0, 7), [Is].True)

                Using output = OpenForCalculation(outputPath)
                    Assert.That(ExternalValues(output), [Is].EqualTo({45.0, 106.0, 49.0}))
                    Assert.That(output.LookupCellValue(Of Double)("External cases", 0, 7, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo(48.0))
                End Using
            End Using
        Finally
            If File.Exists(outputPath) Then File.Delete(outputPath)
        End Try
    End Sub

    Private Shared Sub AssertSameCalcErrors(fixture As String, output As String, sheetName As String, addresses As IEnumerable(Of String))
        Using original = OpenForCalculation(fixture), recalculated = OpenForCalculation(output)
            For Each address In addresses
                Dim cell As New ExcelCell(sheetName, address, ExcelCell.ValueTypes.All)
                Assert.That(original.LookupCellValueAsObject(cell, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("#CALC!"), address)
                Assert.That(recalculated.LookupCellValueAsObject(cell, FormulaCalculationMode.UseCachedValueOnly), [Is].EqualTo("#CALC!"), address)
            Next
        End Using
    End Sub

    Private Shared Function ExternalValues(workbook As OpenXmlExcelDataOperations) As Double()
        Return Enumerable.Range(1, 3).Select(
            Function(row) workbook.LookupCellValue(Of Double)("External cases", row, 2, FormulaCalculationMode.UseCachedValueOnly)).ToArray()
    End Function

    Private Shared Function OpenInSession(path As String, session As ExcelFormulaSession, usePath As Boolean) As OpenXmlExcelDataOperations
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .FormulaSession = session
        }
        If usePath Then Return New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, options)
        options.FormulaWorkbookName = IO.Path.GetFileName(path)
        Return New OpenXmlExcelDataOperations(File.ReadAllBytes(path), options)
    End Function

    Private Shared Function OpenForCalculation(path As String) As OpenXmlExcelDataOperations
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {
            .DisableCalculationEngine = False, .FormulaCulture = CultureInfo.GetCultureInfo("de-DE")
        }
        If IO.Path.GetFileName(path) = "FormulaReferenceMetadata1900.xlsx" Then
            Return New OpenXmlExcelDataOperations(path, OpenMode.OpenExistingFile, options)
        End If
        Return New OpenXmlExcelDataOperations(File.ReadAllBytes(path), options)
    End Function

    Private Shared Function FixturePath(fileName As String) As String
        Return TestEnvironment.FullPathOfExistingTestFile("test_data", "FormulaConformance", fileName)
    End Function

    <TestCase("EUROCONVERT(1.2,""DEM"",""EUR"")", 0.61)>
    <TestCase("EUROCONVERT(1,""FRF"",""EUR"",TRUE,3)", 0.152)>
    <TestCase("EUROCONVERT(1,""FRF"",""EUR"",FALSE,3)", 0.15)>
    <TestCase("EUROCONVERT(1,""FRF"",""DEM"",TRUE,3)", 0.29728616)>
    <TestCase("EUROCONVERT(123.456,""EUR"",""EUR"")", 123.456)>
    <TestCase("EUROCONVERT(1,""EUR"",""ATS"",TRUE)", 13.7603)>
    Public Sub CalculatesDocumentedEuroConversions(formula As String, expected As Double)
        Dim options As New ExcelDataOperationsOptions(WriteProtectionMode.ReadOnly) With {.DisableCalculationEngine = False}
        Using workbook As New OpenXmlExcelDataOperations(DirectCast(Nothing, String), OpenMode.CreateFile, options)
            Dim sheet = workbook.SheetNames()(0)
            workbook.WriteCellFormula(sheet, 0, 0, formula, immediatelyCalculateCellValue:=True)
            Assert.That(workbook.LookupCellValue(Of Double)(sheet, 0, 0, FormulaCalculationMode.UseCachedValueOnly),
                        [Is].EqualTo(expected).Within(1.0E-12))
        End Using
    End Sub

    Private Shared Function IsFormulaAnchor(cell As Cell) As Boolean
        Dim formula = cell.CellFormula
        Return formula IsNot Nothing AndAlso Not (String.IsNullOrEmpty(formula.Text) AndAlso
            If(formula.CalculateCell?.Value, False) AndAlso If(formula.FormulaType?.Value, CellFormulaValues.Normal) = CellFormulaValues.Normal)
    End Function

    Private Shared Function ReadFormulaCaches(fileName As String, sheetName As String) As Dictionary(Of String, (Kind As CellValues, Value As Object))
        Dim caches As New Dictionary(Of String, (Kind As CellValues, Value As Object))(StringComparer.Ordinal)
        Using package = SpreadsheetDocument.Open(fileName, False)
            Dim workbook = package.WorkbookPart
            Dim sheet = workbook.Workbook.Sheets.Elements(Of Sheet)().Single(Function(item) item.Name.Value = sheetName)
            Dim worksheet = DirectCast(workbook.GetPartById(sheet.Id.Value), WorksheetPart).Worksheet
            For Each cell In worksheet.Descendants(Of Cell)().Where(AddressOf IsFormulaAnchor)
                Dim address = cell.CellReference.Value
                Assert.That(cell.CellValue, [Is].Not.Null, sheetName & "!" & address)
                Dim text = If(cell.CellValue.Text, String.Empty)
                Dim kind = If(cell.DataType?.Value, CellValues.Number)
                Dim value As Object = text
                If kind = CellValues.Number Then
                    value = Double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture)
                ElseIf kind = CellValues.Boolean Then
                    value = XmlConvert.ToBoolean(text)
                End If
                caches.Add(address, (kind, value))
            Next
        End Using
        Return caches
    End Function

End Class
