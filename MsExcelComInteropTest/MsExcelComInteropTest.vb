Imports NUnit.Framework

Namespace MsExcelComInteropTest

    Public Class Tests

        <SetUp>
        Public Sub Setup()
        End Sub

        <OneTimeSetUp>
        Public Sub OneTimeSetUp()
            Try
                ExcelApp = New CompuMaster.Excel.MsExcelComInterop.ExcelApplication()
            Catch
                Assert.Ignore("Platform not supported or MS Excel application not installed")
            End Try
        End Sub

        Private ExcelApp As CompuMaster.Excel.MsExcelComInterop.ExcelApplication

        Protected Function OpenExcelAppAndWorkbook(path As String) As CompuMaster.Excel.MsExcelComInterop.ExcelWorkbook
            Return ExcelApp.Workbooks.Open(path)
        End Function

        <Test>
        Public Sub ExportAsFixedFormat()
            Dim TargetTestFile As String = TestEnvironment.FullPathOfDynTestFile(
                                       "pdf-export",
                                       System.IO.Path.GetFileNameWithoutExtension(TestEnvironment.TestFiles.TestFileExcelOpsTestCollection.FullName) & ".pdf"
                                       )
            If System.IO.File.Exists(TargetTestFile) Then System.IO.File.Delete(TargetTestFile)
            Dim Wb = OpenExcelAppAndWorkbook(TestEnvironment.TestFiles.TestFileExcelOpsTestCollection.FullName)
            Wb.ExportAsFixedFormat(MsExcelComInterop.Enumerations.XlFixedFormatType.XlTypePDF, TargetTestFile)
            Assert.True(System.IO.File.Exists(TargetTestFile))
            Wb.Close()
        End Sub

    End Class

End Namespace