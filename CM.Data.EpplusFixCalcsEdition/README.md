# CompuMaster.Data.EpplusFreeFixCalcsEdition

A library to write and read System.Data.DataTable or System.Data.DataSet

Based on EPPlus 4.5.3.3 with LGPL license for solutions targeting .NET Framework 4.8, .NET Standard 2.0, .NET 6, or .NET 8 and higher.

## Security and untrusted workbooks

`CompuMaster.EPPlus4` is based on the older EPPlus 4 codebase. Generating new workbooks from trusted data is a different risk from parsing workbooks supplied by unknown parties. A specially crafted XLSX file can attempt to exhaust memory or CPU through ZIP expansion, very large XML parts, or expensive password hashing. The CompuMaster fork has load limits to reduce these risks, but the limits are defense in depth, not a guarantee that every hostile file is safe. For externally supplied workbooks, prefer a currently supported reader or process files in an isolated, resource-constrained service. Do not disable limits for untrusted input.

The default limits for each loaded workbook are:

| Limit | Default |
| --- | ---: |
| Compressed or encrypted input size | 128 MiB |
| ZIP entries, including empty entries | 20,000 |
| Uncompressed size of one entry | 512 MiB |
| Total uncompressed size of all entries | 1 GiB |
| Size of one XML, relationships, or VML part | 512 MiB |
| Uncompressed-to-compressed ratio of one entry | 200:1 |
| Password-hash iterations in encryption metadata | 1,000,000 |
| Encryption metadata size | 1 MiB |

A valid XLSX larger than 50 MiB is covered by a regression test and loads with these defaults. Size limits are checked before and during ZIP extraction; an oversized or inconsistent package raises `InvalidDataException`. Loading can still require substantial memory, especially for many workbook objects.

For an unusually large **trusted** workbook, configure only the limits that the specific input requires. This example doubles the permitted input size and ZIP-entry count while leaving the other safety boundaries at their defaults:

```csharp
using CompuMaster.Data;
using CompuMaster.Epplus4;

var limits = new ExcelPackageLoadLimits(
    maxInputBytes: 256L * 1024 * 1024,
    maxZipEntries: 40_000);

var options = new XlsEpplusFixCalcsEdition.ReadOptions(
    firstRowContainsColumnNames: true,
    startReadingAtRowIndex: 0,
    packageLoadLimits: limits);
var table = XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFileWithOptions(
    "trusted-large-workbook.xlsx", options);
```

For direct `CompuMaster.EPPlus4` use, call `ExcelPackage.OpenWithLoadLimits(new FileInfo(path), limits)` or the stream overload. `ExcelPackage.CreateFromTemplateWithLoadLimits(...)` also applies limits to template input. Raising these limits for files from unknown senders weakens the DoS protection; instead, inspect the source, cap upload size, isolate processing, and set application-level time and memory budgets.

## Quick & dirty engine comparison / why you shouldn't use MS Excel for all situations

For a full engine overview and comparison chart, please see https://github.com/CompuMasterGmbH/CompuMaster.Excel/blob/main/README.md

## Licensing

  * Please see license file in project directory
  * Pay attention to required licensing of the 3rd party components (commercial vs. community licensing, user licensing, etc.)

## Examples

### Quick-Start: Write a table into a workbook and re-read table from single sheet and re-read dataset with all tables from all sheets

```C#
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
```
