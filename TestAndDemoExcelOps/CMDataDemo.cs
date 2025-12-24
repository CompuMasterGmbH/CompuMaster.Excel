class CMDataDemo
{
    public static void Epplus()
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

    public static void WriteAndReadTableEpplusLgpl()
    {
        string filePath = "SampleTable.xlsx";

        var t1 = SampleTableDyn01();
        CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFileAndFirstSheet(filePath, t1);

        System.Data.DataTable t = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(filePath, true);
        CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFileAndFirstSheet(filePath, t);

        System.Data.DataSet ds = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataSetFromXlsFile(filePath, true);
        CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFileAndFirstSheet(filePath, ds.Tables[0]);
    }

    public static void WriteAndReadTableEpplusPolyform()
    {
        string filePath = "SampleTable.xlsx";

        var t1 = SampleTableDyn01();
        CompuMaster.Data.XlsEpplusPolyformEdition.WriteDataTableToXlsFileAndFirstSheet(filePath, t1);

        System.Data.DataTable t = CompuMaster.Data.XlsEpplusPolyformEdition.ReadDataTableFromXlsFile(filePath, true);
        CompuMaster.Data.XlsEpplusPolyformEdition.WriteDataTableToXlsFileAndFirstSheet(filePath, t);

        System.Data.DataSet ds = CompuMaster.Data.XlsEpplusPolyformEdition.ReadDataSetFromXlsFile(filePath, true);
        CompuMaster.Data.XlsEpplusPolyformEdition.WriteDataTableToXlsFileAndFirstSheet(filePath, ds.Tables[0]);
    }

    private static System.Data.DataTable SampleTableDyn01()
    {
        System.Data.DataTable t1 = new System.Data.DataTable("test");
        t1.Columns.Add();
        t1.Columns.Add();
        t1.Columns.Add();
        var r = t1.NewRow();
        r.ItemArray = new object[] { "1", "R1", "V1" };
        t1.Rows.Add(r);
        r = t1.NewRow();
        r.ItemArray = new object[] { "2", "R2", "V2" };
        t1.Rows.Add(r);
        r = t1.NewRow();
        r.ItemArray = new object[] { "3", "R3", "V3" };
        t1.Rows.Add(r);
        return t1;
    }

}