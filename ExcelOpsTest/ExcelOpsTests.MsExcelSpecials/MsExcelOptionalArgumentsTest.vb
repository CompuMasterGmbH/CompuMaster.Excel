Option Explicit On
Option Strict On

Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework

Namespace ExcelOpsTests.MsExcelSpecials

    <NonParallelizable>
    Public Class MsExcelOptionalArgumentsTest

        ''' <summary>
        ''' Verifies SaveAs with an omitted password and with an explicit password using installed Excel.
        ''' </summary>
        <TestCase(False)>
        <TestCase(True)>
        Public Sub SaveAsPreservesOptionalPassword(encrypted As Boolean)
            If Not CompuMaster.ComInterop.ComTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled("Excel.Application") Then
                Assert.Ignore("Microsoft Excel is not available on this test host.")
            End If
            Dim path = IO.Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") & ".xlsx")
            Dim password = If(encrypted, "password", Nothing)
            Try
                Using app As New CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper()
                    Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite, password)
                    Dim workbook As MsExcelDataOperations = Nothing
                    Try
                        workbook = New MsExcelDataOperations(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile, app, False, options)
                        Dim sheet = workbook.SheetNames()(0)
                        workbook.WriteCellValue(sheet, 0, 0, "test value")
                        workbook.WriteCellValue(sheet, 1, 0, 42)
                        workbook.WriteCellFormula(sheet, 2, 0, "A2+1", True)
                        workbook.SaveAs(path, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                    Finally
                        If workbook IsNot Nothing Then workbook.Close()
                    End Try
                    Dim reopened = app.Workbooks.Open(path, True, password)
                    reopened.CloseAndDispose()
                End Using
                Using reopened As New CompuMaster.Epplus4.ExcelPackage(New FileInfo(path), password)
                    Assert.That(reopened.Workbook.Worksheets(0).Cells(1, 1).Value, [Is].EqualTo("test value"))
                    Assert.That(reopened.Workbook.Worksheets(0).Cells(3, 1).Formula, [Is].EqualTo("A2+1"))
                End Using
            Finally
                If File.Exists(path) Then File.Delete(path)
                CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
            End Try
        End Sub

    End Class

End Namespace
