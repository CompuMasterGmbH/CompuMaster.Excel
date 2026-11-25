# CompuMaster.Excel.OpenXml

Edit Open XML workbooks and templates with the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) through the ExcelOps API. The engine runs on .NET Framework 4.8 on Windows and modern .NET on Windows and Linux. It copies the input package and edits the affected worksheet and workbook XML. Embedded images and drawing relationships are preserved without using Excel or a graphics library.

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

- Open an unencrypted XLSX, XLSM, XLTX or XLTM file from a path, byte array or stream, or create a file with one worksheet. The engine snapshots the input. The caller owns any supplied stream.
- Read and write strings, booleans, finite numbers and blanks with zero-based cell coordinates. Numeric reads return the stored number and do not interpret number formats as dates or currency. Writing a value replaces the cell's formula or old value and retains its formatting. Ordinary formulas can also be stored for later calculation.
- Inspect sheet visibility, hide and very-hide sheets, unhide sheets, and inspect or configure worksheet protection. Protection levels that preserve the existing cell-lock styles are supported. Removing password-protected sheet protection requires a different engine.
- The destination extension determines the saved format. XLSX and XLTX output omits the VBA project. XLSM and XLTM output preserves the project unless you call `RemoveVbaProject()`. Excel 4.0 macro sheets prevent export to a macro-free format.
- Use `SaveCopyAs` with a different destination to retain the source path, write protection and in-memory workbook state, including VBA and cached formula values. `SaveAs` switches the current path to the new file and discards the current VBA project when saving to a macro-free format. `Save` requires a writable file path.
- `Clone()` returns an independent, writable copy without a file path. It includes unsaved edits, VBA and cached formula results without recalculating. `ICloneable.Clone()` returns the same kind of copy.
- Cloning, saving and workbook mutations are serialized per workbook or across all workbooks in the same formula session. Use separate clones for independent work on different threads.
- By default the evaluator is disabled and saving preserves workbook calculation mode and cached formula values. Caches can be stale after input edits. `RequestRecalculationOnOpen()` requests automatic full recalculation in Excel. The existing save cache options can clear cached formula results in the output. Reading a formula cell without a cached value raises `MissingCalculatedCellValueException`.
- Inspect cached formula results without calculating by calling `HasMissingCachedFormulaResult`. The `ResetCellValueFromFormulaCell` and `ResetCellValueFromFormulaCellInWholeWorkbook` methods clear stored results in the open workbook while retaining formula text.
- Export worksheets and workbooks as HTML tables with merged cells, common number formats, and font, fill, alignment and wrap styles. The export follows the EPPlus 4 CM table conventions. The default uses cached formula results and throws `MissingCalculatedCellValueException` if an exported formula result is missing. Set `MissingFormulaResultText` to explicitly render replacement text, including an empty string.
- Adding or removing individual sheets, direct cell writes inside shared, array or data-table formula ranges, and chart rendering raise `NotSupportedException`.

## Formula calculation

Enable the evaluator by setting `ExcelDataOperationsOptions.DisableCalculationEngine` to `false`. Opening a workbook still preserves its caches by default. Set `DisableInitialCalculation` to `false` as well to recalculate on open.

With the evaluator enabled, call `RecalculateCell`, `RecalculateSheet` or `RecalculateAll` to refresh results, or use `FormulaCalculationMode.CalculateIfMissing` for a cell lookup or HTML export. `CalculateIfMissing` uses an existing result for the requested cell. When it calculates a missing result, it also refreshes any evaluated formula precedents. An explicit calculation attempt with the evaluator disabled raises `FeatureDisabledException`. On save, pass `recalculateWorkbook: false` to preserve the current caches without requesting calculation after edits.

The evaluator supports ordinary scalar formulas with numbers, text, booleans, arithmetic, percentages, comparisons, concatenation, and A1 references on the same or another worksheet. `SUM`, `AVERAGE`, `MIN`, `MAX` and `COUNT` also accept rectangular A1 ranges. The `ABS` and `IF` functions are supported, and `IF` evaluates only the selected branch. `ROW` and `COLUMN` return the first row or column of a reference in scalar formulas, or the formula cell's position when called without arguments. `ROWS` and `COLUMNS` return the dimensions of a reference. These functions can reference the formula cell without reading its value. Formula results are stored as numeric, text, Boolean or Excel error values in the XLSX file.

Use English function names, periods in decimal literals and commas between arguments, for example `SUM(1.5,2)`. Arithmetic treats blanks as zero and converts booleans to 1 or 0. Text conversion and numeric text use invariant culture by default. Set `FormulaCulture` to use a specific culture, including German decimal separators and Boolean text. Text comparisons ignore case. Aggregates coerce Boolean and numeric text arguments passed as values. References contribute their numeric cells. `COUNT` skips errors, while the other aggregates propagate them.

Additional functions cover rounding and arithmetic (`ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `INT`, `MOD`, `SIGN`, `SQRT`, `POWER`, `PRODUCT`), text (`TRIM`, `UPPER`, `LOWER`, `LEN`, `LEFT`, `RIGHT`, `MID`, `CONCATENATE`), logical and type checks (`IFERROR`, `IFNA`, `NOT`, `AND`, `OR`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `ISERROR`, `ISNA`, `NA`, `COUNTA`, `COUNTBLANK`), and date and time components (`DATE`, `TIME`, `YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `SECOND`, `TODAY`, `NOW`). Date functions use the workbook's 1900 or 1904 date system. `FormulaLocalClock` optionally supplies local time for deterministic calculation. A pass samples it once when a time-dependent formula needs it.

An unsupported formula raises `FormulaComplexityNotSupportedException`. A direct or indirect dependency on a formula's own calculated value raises `CircularReferenceException`. In either case, the calculation preserves the original caches. Iterative calculation of circular references is outside the supported scope. Shared, grouped array and data-table formulas, relative or external defined names, structured references and VBA execution are outside the supported scope. Full-column references and implicit intersection are outside the supported scope. The evaluator limits expressions to 512 tokens and dependency chains to 256 formulas. It also limits nested expression evaluation.

Workbook names and worksheet-local names support constants and formulas with absolute cell references. `LET` assigns formula-local names to scalars, references or intermediate arrays. Each binding is calculated once when used. `COUNTIF`, `COUNTIFS`, `SUMIF` and `SUMIFS` support comparison criteria and Excel wildcards. `INDEX`, `MATCH`, `VLOOKUP` and `XLOOKUP` cover lookup and reference selection. Use sorted inputs for legacy approximate lookups and modern binary searches. `XLOOKUP` and `XMATCH` support forward, reverse and binary search for exact or approximate matches. Wildcards use linear search. `XLOOKUP` can return a row or a column of results for use inside another formula. Regular expression matching is outside the current calculation scope.

`SEQUENCE`, `TRANSPOSE` and `FILTER` can produce intermediate arrays inside scalar formulas such as `SUM(FILTER(B2:B10,A2:A10>0))`. `INDEX`, `ROWS`, `COLUMNS` and numeric aggregates can consume these arrays. An array that would spill into multiple worksheet cells raises `FormulaComplexityNotSupportedException`. Excel's dynamic-array metadata for single-cell results and its rich value format for `#CALC!` caches are supported.

### Automatic calculation after value edits

Set `CalculateOnCellEdit = true` and enable the calculation engine to opt in to recalculating after `WriteCellValue`. The workbook's automatic calculation setting controls when this happens. In manual mode, value edits retain the previous formula caches until a recalculation is requested. The existing formula-write argument still controls whether a formula write calculates immediately.

The first automatic pass builds dependencies by evaluating the workbook. Later value edits recalculate the affected formulas and the dependents of volatile functions. This includes registered open workbooks in the same formula session. Workbooks can also participate in an `OpenXmlWorkbookSession` for structural reference updates. After a structural change, the next automatic pass rebuilds the calculation dependencies. The dependency graph records value reads through custom function contexts and linked ranges, including cells that were empty. If recalculation fails, the input write remains applied, the previous formula caches remain intact, and the next edit retries the pending work.

### Application-defined functions

Create an `ExcelFormulaSession`, explicitly register each .NET function, and set the session on the workbook's `ExcelDataOperationsOptions.FormulaSession`. Path-backed workbooks default to the file name as their session name. Set `FormulaWorkbookName` for a workbook opened from bytes or a stream, or to give a path-backed workbook an alias. Closing the workbook removes it from the session. A recalculating `SaveCopyAs` uses the session's registered functions and open workbook inputs. The new caches are written to the copy.

A registered function receives typed scalar arguments and an `ExcelFormulaContext`. The context identifies the workbook, worksheet and formula cell. The caller can independently supply a selection with `SetActiveCell`. Headless Open XML has no native Excel application object. Use the context's read-only workbook views during the synchronous invocation to read other cells, including cells in other open workbooks in the same session. Those reads resolve formula precedents before returning. A circular dependency can span workbooks. A failed evaluation preserves existing caches across the session.

Excel worksheet formulas can also read linked XLSX workbooks when the sources are explicitly opened in the same session. Path-backed sources are matched to the stored link target. A relative single-file link can use a matching session name when the source is opened from memory. Calculation reports an unavailable source and preserves existing caches. Opening or calculating a workbook never opens a source file or fetches a remote link.

Registered functions are application code. XLSX stores the formula and its cached result. The caller supplies the registration again when it wants to recalculate. An unregistered function is reported as unsupported. Use the registration overload with `isVolatile: true` if the result depends on state outside the tracked workbook reads. This API supports scalar inputs and outputs. Excel itself needs a separately installed function provider to recalculate the custom formula.

## Structural edits

- Insert or remove whole rows and columns, shift cell rectangles down, right, up or left, and rename worksheets with `updateFormulasAndReferences: true`. Structural operations use zero-based coordinates. A request for `false` raises `NotSupportedException`.
- The engine updates A1 cells and ranges, including absolute and mixed addresses, whole-row and whole-column references, and references from other worksheets. During partial cell shifts, Excel adjusts a range when its entire width or height lies in the shifted strip. A deleted reference with no surviving cells becomes `#REF!`. Formula strings retain their original text.
- Direct 3-D references to a span of worksheets retain their cell coordinates when one sheet in the span changes. When a sheet at either endpoint is renamed, the reference follows that name.
- Merged areas move when a shift starts before them, expand when cells are inserted inside them, and shrink when cells are removed. A partial shift that cuts through only part of a merged area removes the merge. Other merged areas retain their positions unless the shift reaches them.
- Insert rows above an Excel table, within its data, or before its totals row. Full-width downward shifts work as well. The table and its filter range follow the new bounds, and inserted data rows receive their calculated-column formulas. Structured references to those columns retain their text. Shared cell formulas are expanded to individual formulas before rewriting. Excel expands direct references and defined names that cover a complete table data range at its first row or before totals.
- Whole-column edits before a table move its bounds and filter. Whole-row deletions inside a table can remove data rows after the first, provided the deletion leaves the header and totals intact.
- The workbook's R1C1 display setting is preserved. Excel stores direct R1C1 references as A1 in XLSX, including function arguments. The engine follows direct references in dynamic range endpoints such as `INDEX(...):INDEX(...)` and `A1:INDEX(...)`. It preserves literal text passed to `INDIRECT` and updates direct references in expressions such as `INDIRECT(ADDRESS(ROW(A1),COLUMN(A1),1,FALSE),FALSE)`. Excel interprets R1C1 address text using its local reference notation.
- In lookup and database functions, explicit search, data, criteria and field-cell references follow structural edits. Literal field names and numeric column indexes keep their values.
- Defined names can use absolute A1 references with an explicit sheet or a worksheet-local scope. Renaming also updates explicit sheet qualifiers in relative defined names. Indirect references encoded as text retain their original text.
- Successful edits clear stored formula results across the workbook, including possible indirect dependents, and remove the calculation chain on save. The workbook requests full recalculation on load and retains its calculation mode.
- Direct references to another workbook, such as `'[1]Source Data'!$B$2`, retain their external coordinates during a local structural edit. When the linked workbook itself changes, a shared workbook session can update incoming references in attached workbooks.
- Direct sheet-qualified references in classic charts on worksheets and chart sheets follow structural changes. Supported chart roles include titles, series names, categories, values, and scatter X and Y values. The engine moves cached points with their cells and updates supported chart anchors during whole-row and whole-column edits.
- Renaming a worksheet also updates ChartEx sources stored in workbook-scoped `_xlchart` names. The ChartEx part and its caches are preserved.
- The engine validates the workbook before applying a structural edit. Unsupported cases include unregistered or unsupported external links, external 3-D references, spill and compound range forms with unsupported endpoints, array and data-table formulas, shifts in workbooks with ChartEx, unsupported chart sources or placement, pivots, hyperlinks, data validations, conditional formatting, and unknown reference-bearing extensions. Deleting an entire source range in a chart with multiple series is also unsupported. A shift requires the edited worksheet to have no standalone filters or active sheet protection. Inserting or deleting columns within a table, removing table headers, first data rows or totals, shifts that cover only part of a table, and changes to table ranges with saved sort state raise `NotSupportedException`. Inserting or removing whole columns requires default column formatting.

## Copying or moving sheets between workbooks

- Use `CopySheet`, `CopySheets`, `MoveSheet`, or `MoveSheets` with two open Open XML workbooks. An optional `targetSheetIndex` inserts before the tab at that zero-based position. The index counts worksheets and chart sheets together. Omitting it appends the group in source workbook order. Each `SheetTransferResult` reports the assigned `TargetSheetName` and actual `TargetSheetIndex` in the destination.
- `CopyWorksheetFormulaMode.ExcelCompatible` follows worksheet identity. References between sheets in the copied group point to their copies. Explicit references to other source worksheets become workbook links. These require a saved source path or a shared workbook session.
- `CopyWorksheetFormulaMode.PreserveFormulas` copies the text of supported cell and chart formulas exactly. Sheet qualifiers and names then resolve in the context of the destination. The copy clears cached formula results and formula-based chart caches so Excel can recalculate them.
- A move stages both workbooks before applying either change. Direct references in the source to moved sheets become links to the destination. Existing destination links to the moved sheets become local references, including in classic chart sheets. Unused destination links to the source are removed and later link indexes are adjusted. Save both workbooks to persist the move.
- Compatible groups can contain worksheets with embedded classic charts and classic chart sheets. ChartEx, Excel tables, sheet-scoped defined names on the transferred sheets, array and data-table formulas, and copying formulas from a workbook with external links raise `NotSupportedException`. Excel-compatible adaptation of transferred formulas also requires a source without defined names or 3-D references in the formulas being adapted.

## Working with several open workbooks

Create an `OpenXmlWorkbookSession` and call `Attach` for each participating workbook. Structural edits also update supported direct links and their cached source coordinates in attached dependents, including links used by supported tables and classic charts. Moves between two attached workbooks also redirect references in other attached workbooks. Closed or detached workbooks stay outside the operation. Incoming structural updates of cached external defined names remain unsupported.

Copying or moving sheets within a session can create links to unsaved workbooks. The generated temporary name identifies the linked workbook until it is saved. Keep the dependents attached and save them after the linked workbook has a path to persist the resolved link. A previously saved dependent file changes only on its next save. Save and close the workbook instances as usual.

See [Excel chart reference behavior](chart-reference-excel-behavior.md) for the reference cases for embedded charts and chart sheets.

See [Excel worksheet transfer behavior](worksheet-transfer-excel-behavior.md) for the native Excel reference tests and session limitations.

The engine is licensed under the MIT License.
