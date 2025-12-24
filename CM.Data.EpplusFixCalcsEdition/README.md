# CompuMaster.Data.EpplusFreeFixCalcsEdition

A library to write and read System.Data.DataTable or System.Data.DataSet

Use Epplus 4.5 with LGPL license for solutions targetting .NET Framework 4.8 or .NET 6 or higher

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
