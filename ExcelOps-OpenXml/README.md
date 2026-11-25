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

Additional functions cover rounding and arithmetic (`ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `INT`, `MOD`, `SIGN`, `SQRT`, `POWER`, `PRODUCT`), text (`TRIM`, `UPPER`, `LOWER`, `LEN`, `LEFT`, `RIGHT`, `MID`, `CONCATENATE`), logical and type checks (`IFERROR`, `IFNA`, `NOT`, `AND`, `OR`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `ISERROR`, `ISNA`, `NA`, `COUNTA`, `COUNTBLANK`), and date and time components (`DATE`, `TIME`, `YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `SECOND`, `TODAY`, `NOW`). Date functions use the workbook's 1900 or 1904 date system. `FormulaLocalClock` optionally supplies local time for deterministic calculation. A pass samples it once when a time-dependent formula needs it.

An unsupported formula raises `FormulaComplexityNotSupportedException`. A direct or indirect dependency on a formula's own calculated value raises `CircularReferenceException`. In either case, the calculation preserves the original caches. Iterative calculation of circular references is outside the supported scope. Shared, array and data-table formulas, defined names, structured references and VBA execution are outside the supported scope. Full-column references and implicit intersection are outside the supported scope. The evaluator limits expressions to 512 tokens and dependency chains to 256 formulas. It also limits nested expression evaluation.

### Application-defined functions

Create an `ExcelFormulaSession`, explicitly register each .NET function, and set the session on the workbook's `ExcelDataOperationsOptions.FormulaSession`. Path-backed workbooks default to the file name as their session name. Set `FormulaWorkbookName` for a workbook opened from bytes or a stream, or to give a path-backed workbook an alias. Closing the workbook removes it from the session.

A registered function receives typed scalar arguments and an `ExcelFormulaContext`. The context identifies the workbook, worksheet and formula cell. The caller can independently supply a selection with `SetActiveCell`. Headless Open XML has no native Excel application object. Use the context's read-only workbook views during the synchronous invocation to read other cells, including cells in other open workbooks in the same session. Those reads resolve formula precedents before returning. A circular dependency can span workbooks. A failed evaluation preserves existing caches across the session.

Excel worksheet formulas can also read linked XLSX workbooks when the sources are explicitly opened in the same session. Path-backed sources are matched to the stored link target. A relative single-file link can use a matching session name when the source is opened from memory. Calculation reports an unavailable source and preserves existing caches. Opening or calculating a workbook never opens a source file or fetches a remote link.

Registered functions are application code. XLSX stores the formula and its cached result. The caller supplies the registration again when it wants to recalculate. An unregistered function is reported as unsupported. This API supports scalar inputs and outputs. Excel itself needs a separately installed function provider to recalculate the custom formula.

The engine is licensed under the MIT License.
