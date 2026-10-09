Imports System.IO
Imports CompuMaster.ComInterop
Imports CompuMaster.Excel.MsExcelComInterop
Imports NUnit.Framework

<NonParallelizable>
Public Class ManualPrinterTest

    ''' <summary>
    ''' Submits one synthetic page to the default printer; runs only when explicitly selected.
    ''' </summary>
    <Explicit("Submits a physical print job to the development computer's default printer.")>
    <TestCase(False)>
    <TestCase(True)>
    Public Sub ManualPrintOnePageToDefaultPrinter(worksheet As Boolean)
        If Not ComTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled("Excel.Application") Then
            Assert.Ignore("Microsoft Excel is not available on this test host.")
        End If
        Dim path = IO.Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") & ".xlsx")
        Dim target = If(worksheet, "Worksheet", "Workbook")
        Try
            Using app As New ExcelApplication()
                app.Visible = False
                app.DisplayAlerts = False
                app.Interactive = False
                Using draft As New ComRootObject(Of Object)(app.Workbooks.InvokeFunction(Of Object)("Add"), Sub(book) book.InvokeMethod("Close", False))
                    Using sheets As New ComRootObject(Of Object)(draft.InvokePropertyGet(Of Object)("Worksheets"), Nothing)
                        Using sheet As New ComRootObject(Of Object)(sheets.InvokePropertyGet(Of Object)("Item", 1), Nothing)
                            Using cell As New ComRootObject(Of Object)(sheet.InvokePropertyGet(Of Object)("Range", "A1"), Nothing)
                                cell.InvokePropertySet("Value2", "CompuMaster Excel Interop print test")
                                cell.InvokePropertySet("ColumnWidth", 48)
                            End Using
                            Using cell As New ComRootObject(Of Object)(sheet.InvokePropertyGet(Of Object)("Range", "A2"), Nothing)
                                cell.InvokePropertySet("Value2", target & ".PrintOut: omitted printer and file name")
                            End Using
                            Using cell As New ComRootObject(Of Object)(sheet.InvokePropertyGet(Of Object)("Range", "A3"), Nothing)
                                cell.InvokePropertySet("Value2", "Synthetic test data - one page, one copy")
                            End Using
                            Using layout As New ComRootObject(Of Object)(sheet.InvokePropertyGet(Of Object)("PageSetup"), Nothing)
                                layout.InvokePropertySet("PrintArea", "$A$1:$A$3")
                                layout.InvokePropertySet("Zoom", False)
                                layout.InvokePropertySet("FitToPagesWide", 1)
                                layout.InvokePropertySet("FitToPagesTall", 1)
                            End Using
                        End Using
                    End Using
                    draft.InvokeMethod("SaveAs", path, 51)
                End Using
                Dim workbook = app.Workbooks.Open(path)
                Try
                    If worksheet Then
                        Using sheet = workbook.Sheets.Item(0)
                            sheet.PrintOut(toPageIndex:=0)
                        End Using
                    Else
                        workbook.PrintOut(toPageIndex:=0)
                    End If
                    TestContext.Progress.WriteLine(target & ".PrintOut submitted one page to the default printer. Confirm the physical output manually.")
                Finally
                    workbook.Close()
                End Try
            End Using
        Finally
            If File.Exists(path) Then File.Delete(path)
            ComTools.GarbageCollectAndWaitForPendingFinalizers()
        End Try
    End Sub

End Class
