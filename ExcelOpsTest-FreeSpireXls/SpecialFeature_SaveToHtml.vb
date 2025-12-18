Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps

Public Class SpecialFeature_SaveToHtml

    Private Const OPEN_OUTPUT_IN_BROWSER_AFTER_TEST As Boolean = False
    Private Const UNIQUE_TEST_OUTPUT_SUBDIR_NAME_FOR_PROVIDER = "FreeSpire"

    Private ReadOnly Property DisabledCalculationEngineOptions As ExcelDataOperationsOptions = New ExcelDataOperationsOptions("", True, True, True)

    <Test>
    Public Sub ExportWorkbook()
        Dim TestXlsxFile = TestFiles.TestFileGrund01()
        Dim TestHtmlOutputFile = TestEnvironment.FullPathOfDynTestFile(UNIQUE_TEST_OUTPUT_SUBDIR_NAME_FOR_PROVIDER, TestXlsxFile.Name & ".html")
        System.Console.WriteLine("TEST IN FILE: " & TestXlsxFile.FullName)
        System.Console.WriteLine("TEST OUT FILE: " & TestHtmlOutputFile)

        Try
            Dim Wb As New FreeSpireXlsDataOperations(TestXlsxFile.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, DisabledCalculationEngineOptions)
            Wb.SaveToHtml(TestHtmlOutputFile, False)
        Catch ex As TypeInitializationException
            Assert.Ignore("Not supported on this platform " & System.Environment.OSVersion.Platform.ToString)
        End Try

        If OPEN_OUTPUT_IN_BROWSER_AFTER_TEST Then
            Dim OpenFileProcess = System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                .UseShellExecute = True,
                .FileName = TestHtmlOutputFile
                })
        End If
    End Sub

    <Test>
    Public Sub ExportWorksheetGrunddatenV19()
        Dim TestXlsxFile = TestFiles.TestFileGrund01()
        System.Console.WriteLine("TEST IN FILE: " & TestXlsxFile.FullName)

        Dim Wb As FreeSpireXlsDataOperations = Nothing
        Try
            Wb = New FreeSpireXlsDataOperations(TestXlsxFile.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, DisabledCalculationEngineOptions)
        Catch ex As TypeInitializationException
            Assert.Ignore("Not supported on this platform " & System.Environment.OSVersion.Platform.ToString)
        End Try
        For Each WorkSheetName In Wb.WorkSheetNames
            System.Console.WriteLine("FOUND WORKSHEET: " & WorkSheetName)
        Next
        For Each WorkSheetName In Wb.WorkSheetNames
            Dim TestHtmlOutputFile = TestEnvironment.FullPathOfDynTestFile(UNIQUE_TEST_OUTPUT_SUBDIR_NAME_FOR_PROVIDER, TestXlsxFile.Name & "." & WorkSheetName & ".html")
            System.Console.WriteLine("TEST OUT FILE: " & TestHtmlOutputFile)
            Wb.SaveWorksheetToHtml(WorkSheetName, TestHtmlOutputFile)
            If OPEN_OUTPUT_IN_BROWSER_AFTER_TEST Then
                Dim OpenFileProcess = System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                    .UseShellExecute = True,
                    .FileName = TestHtmlOutputFile
                    })
            End If
        Next

    End Sub

End Class
