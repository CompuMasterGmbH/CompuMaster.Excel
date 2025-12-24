# CompuMaster.Excel.MicrosoftExcel

A library to access and edit Excel files with common interface CompuMaster.Excel.ExcelOps for several Excel engines

Use Microsoft.Office.Interop.Excel v15 (MS Office 2013) or higher, for solutions targetting .NET Framework 4.8 or .NET 6 or higher

## Quick & dirty engine comparison / why you shouldn't use MS Excel for all situations

For a full engine overview and comparison chart, please see https://github.com/CompuMasterGmbH/CompuMaster.Excel/blob/main/README.md

## Licensing

  * Please see license file in project directory
  * Pay attention to required licensing of the 3rd party components (commercial vs. community licensing, user licensing, etc.)

## Examples

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

