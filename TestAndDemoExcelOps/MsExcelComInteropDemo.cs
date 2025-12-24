class MsExcelComInteropDemo
{
    public static void Demo()
    {
        CompuMaster.Excel.MsExcelComInterop.ExcelApplication? excelApp = null;
        CompuMaster.Excel.MsExcelComInterop.ExcelWorkbook? wb = null;
        try
        {
            // Open excel engine
            try
            {
                excelApp = new CompuMaster.Excel.MsExcelComInterop.ExcelApplication();
            }
            catch
            {
                System.Console.WriteLine("Platform not supported or MS Excel application not installed");
                return;
            }

            // Workbook actions
            string sourceTestFile = "test.xlsx";
            string targetTestFile = "test.pdf";
            if (System.IO.File.Exists(targetTestFile))
                System.IO.File.Delete(targetTestFile);
            wb = excelApp!.Workbooks.Open(sourceTestFile);
            wb.ExportAsFixedFormat(CompuMaster.Excel.MsExcelComInterop.Enumerations.XlFixedFormatType.XlTypePDF, targetTestFile);
        }
        finally
        {
            // Close and dispose everything
            if (wb != null && wb.IsClosed == false)
                wb.Close();
            if (excelApp != null)
                excelApp.Dispose();
            CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers();
        }
    }

}