# CompuMaster.Excel.MicrosoftExcel

A library to access and edit Excel files with common interface CompuMaster.Excel.ExcelOps for several Excel engines

Use Microsoft.Office.Interop.Excel v15 (MS Office 2013) or higher, for solutions targetting .NET Framework 4.8 or .NET 6 or higher

## Quick & dirty engine comparison / why you shouldn't use MS Excel for all situations

For a full engine overview and comparison chart, please see https://github.com/CompuMasterGmbH/CompuMaster.Excel/blob/main/README.md

## Licensing

  * Please see license file in project directory
  * Pay attention to required licensing of the 3rd party components (commercial vs. community licensing, user licensing, etc.)

## Examples

### Excel diagnostic logs

`MsExcelDataOperations.DiagnosticLogFiles` and `MsExcelApplicationWrapper.DiagnosticLogFiles` return read-only snapshots of existing, positively associated Office diagnostic logs. Reading these properties does not start Excel or create logs. Surviving paths remain available after closing the application.

`PreserveDiagnosticLogFiles` defaults to `false` and is shared by workbooks using the same Excel process. After the existing shutdown sequence confirms that the exact process has exited, a successful quit with exit code zero removes its associated logs. An error or forced exit retains them and trims only trailing NUL padding, preserving Unicode terminal code units. Set `PreserveDiagnosticLogFiles = true` to leave every log byte unchanged, including padding.

BOMs and recognizable Office headers identify the text encoding. Ambiguous BOM-less logs may retain up to three trailing padding bytes to protect UTF-16/UTF-32 terminal code units.

Ownership requires the process ID and start time, valid log timestamps, and the Primary/Additional session GUID. A query-only Windows Restart Manager session can identify exclusively locked Primary logs. Unknown formats, uncertain ownership, inaccessible or replaced files, and reparse points are skipped. Housekeeping never changes the existing Quit/COM-release/wait/kill sequence or Office logging settings. Housekeeping makes at most three attempts, waiting 350 ms between attempts. A host crash cannot run in-process housekeeping.

### Quick-Start: Create a workbook and put some values and formulas, then output the result to console

```C#
using CompuMaster.Excel.ExcelOps;

string FirstSheetName;
TextTable formulasOrValues;
TextTable values;

//Create a workbook and put some values and formulas
workbook = new MsExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions());
FirstSheetName = workbook.SheetNames()[0];
workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);

//Output table with formulas or alternatively with formatted cell value
formulasOrValues = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText);
System.Console.WriteLine(formulasOrValues.ToUIExcelTable());

//Output table with calculated or static values
values = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues);
System.Console.WriteLine(values.ToUIExcelTable());
```

leads to following output

```text
# |A  |B      |C
--+---+-------+-----------
1 |123|456,123|=SUM(A1:B1)

# |A  |B      |C
--+---+-------+-------
1 |123|456,123|579,123
```
</details>
