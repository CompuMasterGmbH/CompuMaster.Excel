# CompuMaster.Excel.MsExcelComInterop

Use Microsoft.Office.Interop.Excel v15 (MS Office 2013) or higher, for solutions targetting .NET Framework 4.8 or .NET 6 or higher

Allow COM interop at clients without using Microsoft.Office.Excel.Interop assemblies for light-weight deployments of applications to customers/clients, regardless of the installed version of Microsoft Office

Limitation: supports only a very tiny, but often-used feature set, e.g. print, export to PDF, run VBA code

## Licensing

  * Please see license file in project directory
  * Pay attention to required licensing of the 3rd party components (commercial vs. community licensing, user licensing, etc.)

## Examples

### Quick-Start: Create a workbook and put some values and formulas, then output the result to console

```C#
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
```
</details>

