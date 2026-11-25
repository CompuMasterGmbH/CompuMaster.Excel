# CompuMaster.Excel.OpenXml

Edit XLSX templates with the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) through the ExcelOps API. The engine runs on .NET Framework 4.8 on Windows and modern .NET on Windows and Linux. It copies the input package and edits the affected worksheet and workbook XML. Embedded images and drawing relationships are preserved without using Excel or a graphics library.

```csharp
using CompuMaster.Excel.ExcelOps;
using static CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase;

using var workbook = new OpenXmlExcelDataOperations(
    "template.xlsx", OpenMode.OpenExistingFile, new ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly));

workbook.WriteCellValue("Inputs", 0, 0, "Example");
workbook.WriteCellValue("Inputs", 0, 1, 12);
workbook.RequestRecalculationOnOpen();
workbook.SaveCopyAs("filled.xlsx", SaveOptionsForDisabledCalculationEngines.NoReset);
```

- Open an unencrypted XLSX from a path, byte array or stream, or create a workbook with one worksheet. The engine snapshots the input. The caller owns any supplied stream.
- Read and write strings, booleans, finite numbers and blanks with zero-based cell coordinates. Numeric reads return the stored number and do not interpret number formats as dates or currency. Writing a value replaces the cell's formula or old value and retains its formatting. Ordinary formulas can also be stored for later calculation.
- Inspect sheet visibility, hide and very-hide sheets, unhide sheets, and inspect or configure worksheet protection. Protection levels that preserve the existing cell-lock styles are supported. Removing password-protected sheet protection requires a different engine.
- Use `SaveCopyAs` with a different destination to retain the source path, write protection and original file bytes. `SaveAs` switches the current path to the new file. `Save` requires a writable file path.
- By default the evaluator is disabled and saving preserves workbook calculation mode and cached formula values. Caches can be stale after input edits. `RequestRecalculationOnOpen()` requests automatic full recalculation in Excel. The existing save cache options can clear cached formula results in the output. Reading a formula cell without a cached value raises `MissingCalculatedCellValueException`.
- Inspect cached formula results without calculating by calling `HasMissingCachedFormulaResult`. The `ResetCellValueFromFormulaCell` and `ResetCellValueFromFormulaCellInWholeWorkbook` methods clear stored results in the open workbook while retaining formula text.
- Export worksheets and workbooks as HTML tables with merged cells, common number formats, and font, fill, alignment and wrap styles. The export follows the EPPlus 4 CM table conventions. The default uses cached formula results and throws `MissingCalculatedCellValueException` if an exported formula result is missing. Set `MissingFormulaResultText` to explicitly render replacement text, including an empty string.
- Structural row, column, cell and sheet operations, edits inside shared, array or data-table formula ranges, and chart rendering raise `NotSupportedException`.

## Formula calculation

Enable the evaluator by setting `ExcelDataOperationsOptions.DisableCalculationEngine` to `false`. Opening a workbook still preserves its caches by default. Set `DisableInitialCalculation` to `false` as well to recalculate on open.

With the evaluator enabled, call `RecalculateCell`, `RecalculateSheet` or `RecalculateAll` to refresh results, or use `FormulaCalculationMode.CalculateIfMissing` for a cell lookup or HTML export. `CalculateIfMissing` uses an existing result for the requested cell. When it calculates a missing result, it also refreshes any evaluated formula precedents. An explicit calculation attempt with the evaluator disabled raises `FeatureDisabledException`. On save, pass `recalculateWorkbook: false` to preserve the current caches without requesting calculation after edits.

The evaluator supports ordinary scalar formulas with numbers, text, booleans, arithmetic, percentages, comparisons, concatenation, and A1 references on the same or another worksheet. `SUM`, `AVERAGE`, `MIN`, `MAX` and `COUNT` also accept rectangular A1 ranges. The `ABS` and `IF` functions are supported, and `IF` evaluates only the selected branch. `ROW` and `COLUMN` return the first row or column of a reference in scalar formulas, or the formula cell's position when called without arguments. `ROWS` and `COLUMNS` return the dimensions of a reference. These functions can reference the formula cell without reading its value. Formula results are stored as numeric, text, Boolean or Excel error values in the XLSX file.

Use English function names, periods in decimal literals and commas between arguments, for example `SUM(1.5,2)`. Arithmetic treats blanks as zero, converts booleans to 1 or 0, and parses numeric text using invariant culture. Text comparisons ignore case. Aggregates coerce Boolean and numeric text arguments passed as values. References contribute their numeric cells. `COUNT` skips errors, while the other aggregates propagate them.

An unsupported formula raises `FormulaComplexityNotSupportedException`. A direct or indirect dependency on a formula's own calculated value raises `CircularReferenceException`. In either case, the calculation preserves the original caches. Iterative calculation of circular references is outside the supported scope. Shared, array and data-table formulas, defined names, structured or external references, custom functions and VBA execution are outside the supported scope. ISO date cells, full-column references and implicit intersection are also outside this initial scope. The evaluator limits expressions to 512 tokens and dependency chains to 256 formulas. It also limits nested expression evaluation.

The engine is licensed under the MIT License.
