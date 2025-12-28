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
  * **CompuMaster.Excel.FreeSpireXls**
    * Use FreeSpire.Xls for solutions targetting .NET Framework 4.8 or .NET 6 or higher

### Quick & dirty engine comparison / why you shouldn't use MS Excel for all situations

In following a simplified comparison without any warranties. Subjects might change over time, too. Please contact the responsible manufacturerer especially for licensing issues.

| Engine | Pros | Cons | Note on licensing or support | Manufacturer website |
|--------|------|------|------------------------------|----------------------|
| Microsoft Excel | + 100% compatibility to Micrsoft Excel ;-) | - Speed<br/>- Dependency to windows platform only (and maybe MacOS) | * licensing per user (!)<br/>* NOT recommended/supported for software deployment, for servers or for similar automation <sup>1)</sup> | www.microsoft.com |
| Epplus 5 or higher | + Speed | - no export of chart images<br/>- calculation issue when re-opening in MS Excel <sup>2)</sup><br/>- limited VBA/macro support  | * Polyform license<br/>* limited community licensing<br/>* commercial licensing available | www.epplussoftware.com |
| Epplus 4 | + Speed | - no export of chart images<br/>- some (seldom-used) calculation functions not implemented<br/>- calculation issue when re-opening in MS Excel <sup>2)</sup>, but workaround AVAILABLE (in this project's fork)<br/>- limited VBA/macro support<br/>- limited Linux/MacOS support (because of references to System.Drawing/libgdiplus) | * LGPL<br/>* "free license"<br/>* no manufacturer support (end of life) | www.github.com/JanKallman/EPPlus |
| Spire.Xls | + Speed<br/>+ Charts export (windows platform only) | - calculation issue when re-opening in MS Excel <sup>2)</sup><br/>- limited VBA/macro support | * commercial licensing available | www.e-iceblue.com |
| FreeSpire.Xls | + Speed<br/>+ Charts export (windows platform only) | - Limitations by manufacturer due to free edition<br/>- calculation issue when re-opening in MS Excel <sup>2)</sup><br/>- limited VBA/macro support | * "free license"<br/>* no official support by manufacturer | www.e-iceblue.com |

PLEASE NOTE:

<sup>1)</sup> A great article on Microsoft Excel for automation, inclusing licensing issues, is available at https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2

<sup>2)</sup> calculation issue when re-opening in MS Excel: after Excel file has been written to disk, cell formulas are usually calculated and their results are buffered in the saved Excel file. In certain cases, MS Excel is not able to refresh calculated cell values when they depend (indirectly) on other cells which have changed.
  * This issue applies for all 3rd-party engines (as far as I know)
  * Available workaround in MS Excel: enter each single cell manually (and in correct dependency order!) and confirm its formula (and sorry, F9 for full automatic recalculation doesn't work)
  * Available workaround in Epplus 4 special edition (provided within this project, see CompuMaster.Excel.EpplusFreeFixCalcsEdition): Clear all cached values from cells with formulas to enforce MS Excel to recalculate them (without depending of any caches)

### Common helper libraries

  * **CompuMaster.EPPlus4** 
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
ExcelDataOperationsBase workbook = new EpplusFreeExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions());
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
EpplusPolyformExcelDataOperations.LicenseContext = new EpplusPolyformExcelDataOperations.EpplusLicenseActivator(OfficeOpenXml.EPPlusLicenseType.NonCommercialPersonal, "Your Name");

//Create a workbook and put some values and formulas
workbook = new EpplusPolyformExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions());
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
workbook = new FreeSpireXlsDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions());
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

