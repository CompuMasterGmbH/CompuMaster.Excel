# CompuMaster.Excel

A bunch of libraries to access and edit Excel files with common interface for several Excel engines

## Why this project?

### One common API for several Excel engines

  * **CompuMaster.Excel.ExcelOps**
    * Provides the common API for all following Excel engines
    * Sometimes there is the need to switch the Excel engine under the hood, because the standard MS Excel engine (via COM interop) is too slow for many operations.
    * Depending on your custom project or developer license, you might want/require to use a specific alternative Excel engine
    * Allow parallel use of several Excel engines in your project
        * at least for some work items, you need Excel engine A, while for some other work item you need Excel engine B because of its feature set, bugs, etc.
        * allow step-by-step-migrations instead of one big migration task
  * **CompuMaster.Excel.MicrosoftExcel**
    * Use Microsoft.Office.Interop.Excel v15 (MS Office 2013) or higher, for solutions targetting .NET Framework 4.8 or .NET 6 or higher
  * **CompuMaster.Excel.EpplusFreeFixCalcsEdition** 
    * Use Epplus 4.5 with LGPL license for solutions targetting .NET Framework 4.8 or .NET 6 or higher
  * **CompuMaster.Excel.EpplusPolyformEdition**
    * Use latest Epplus with Epplus Software's polyform license for solutions targetting .NET Framework 4.8 or .NET 6 or higher
  * **CompuMaster.Excel.(Free)SpireXls**
    * Use (Free)Spire.Xls for solutions targetting .NET Framework 4.8 or .NET 6 or higher

### Common helper libraries

  * **CompuMaster.Excel.EpplusFreeFixCalcsEdition** 
    * a fork of Epplus 4.5 with LGPL license, providing features to clear calculation caches which might lead to wrong calculation cache reset behavioiur in MS Excel (MS Excel doesn't recalculate all required cells)
    * Supports .NET Framework 4.8 + .NET 6 or higher
    * Feature support for checkup if cell requires recalculation: https://github.com/EPPlusSoftware/EPPlus/issues/113
    * Correct refreshing of indirect cell references by resetting all calculation caches (enforces MS Excel to recalculate all cells)
  * **CompuMaster.Excel.Tools.MicrosoftAndEpplus**
    * Provide workarounds for some well-known issues in some of the alternative Excel engines
  * **CompuMaster.Excel.MsExcelComInterop**
    * Allow COM interop at clients without using Microsoft.Office.Excel.Interop assemblies for light-weight deployments of applications to customers/clients, regardless of the installed version of Microsoft Office
    * Limitation: supports only a very tiny, but often-used feature set, e.g. print, export to PDF, run VBA code

## Licensing

  * Please see license type for every sub project within its directory
  * Pay attention to required licensing of the 3rd party components (commercial vs. community licensing, user licensing, etc.)

## Examples

### Epplus 4 (LGPL) 

<details>
<summary>Quick-Start: Create a workbook and put some values and formulas, then output the result to console</summary>

```C#
using CompuMaster.Excel.ExcelOps;

string FirstSheetName;
TextTable formulasOrValues;
TextTable values;

//Create a workbook and put some values and formulas
ExcelDataOperationsBase workbook = new EpplusFreeExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, true, null);
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

### Epplus (Polyform license edition) 

<details>
<summary>Quick-Start: Create a workbook and put some values and formulas, then output the result to console</summary>

```C#
using CompuMaster.Excel.ExcelOps;

string FirstSheetName;
TextTable formulasOrValues;
TextTable values;

//Assign required license context for Epplus component
EpplusPolyformExcelDataOperations.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

//Create a workbook and put some values and formulas
workbook = new EpplusPolyformExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, true, null);
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

### FreeSpire.Xls

<details>
<summary>Quick-Start: Create a workbook and put some values and formulas, then output the result to console</summary>

```C#
using CompuMaster.Excel.ExcelOps;

string FirstSheetName;
TextTable formulasOrValues;
TextTable values;

//Create a workbook and put some values and formulas
workbook = new FreeSpireXlsDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, true, null);
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

### Microsoft Excel (via COM interop)

<details>
<summary>Quick-Start: Create a workbook and put some values and formulas, then output the result to console</summary>

```C#
using CompuMaster.Excel.ExcelOps;

string FirstSheetName;
TextTable formulasOrValues;
TextTable values;

//Create a workbook and put some values and formulas
workbook = new MsExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, true, null);
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

