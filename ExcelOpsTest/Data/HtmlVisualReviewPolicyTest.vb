Imports NUnit.Framework

Namespace Data

    Public Class HtmlVisualReviewPolicyTest

        Private Const SampleMethod As String = "HtmlExportWorksheetGrunddatenV19"
        Private Const SampleFullName As String = "CompuMaster.Excel.Test.ExcelOpsTests.Engines.EpplusPolyformEditionOpsTest." & SampleMethod

        <TestCase(Nothing, True, False, True)>
        <TestCase(Nothing, False, False, False)>
        <TestCase(" ", True, False, True)>
        <TestCase("off", True, False, False)>
        <TestCase(" off ", True, False, False)>
        <TestCase(SampleMethod, False, False, True)>
        <TestCase(SampleFullName, False, False, True)>
        <TestCase("AnotherTest, " & SampleMethod, False, False, True)>
        <TestCase("OtherEngine." & SampleMethod, True, False, False)>
        <TestCase("HtmlExportWorksheet", True, False, False)>
        <TestCase(Nothing, True, True, False)>
        <TestCase(SampleMethod, True, True, False)>
        Public Sub BrowserLaunchRespectsTestSelectionAndCi(selection As String, openByDefault As Boolean,
                                                           runningInCi As Boolean, expected As Boolean)
            Assert.That(HtmlVisualReview.ShouldOpen(selection, SampleMethod, SampleFullName, openByDefault, runningInCi),
                        [Is].EqualTo(expected))
        End Sub

    End Class

End Namespace
