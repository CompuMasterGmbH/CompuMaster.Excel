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
- After ordinary input edits, the default save preserves workbook calculation mode and cached formula values. These caches may be stale. `RequestRecalculationOnOpen()` requests automatic full recalculation in Excel. The existing save cache options can clear cached formula results in the output. Reading a formula cell without a cached value raises `MissingCalculatedCellValueException`.
- Inspect cached formula results without calculating by calling `HasMissingCachedFormulaResult`. The `ResetCellValueFromFormulaCell` and `ResetCellValueFromFormulaCellInWholeWorkbook` methods clear stored results in the open workbook while retaining formula text.
- Export worksheets and workbooks as HTML tables with merged cells, common number formats, and font, fill, alignment and wrap styles. The export follows the EPPlus 4 CM table conventions. The default uses cached formula results and throws `MissingCalculatedCellValueException` if an exported formula result is missing. Set `MissingFormulaResultText` to explicitly render replacement text, including an empty string.
- The explicit `CalculateIfMissing` lookup raises `NotSupportedException` when a result must be calculated. Existing recalculation methods continue to raise `FeatureDisabledException`. Adding or removing worksheets, editing cells within shared, array or data-table formula ranges, and chart rendering raise `NotSupportedException`.

## Structural edits

- Insert or remove whole rows and columns, shift cell rectangles down, right, up or left, and rename worksheets with `updateFormulasAndReferences: true`. Structural operations use zero-based coordinates. A request for `false` raises `NotSupportedException`.
- The engine updates A1 cells and ranges, including absolute and mixed addresses, whole-row and whole-column references, and references from other worksheets. During partial cell shifts, Excel adjusts a range when its entire width or height lies in the shifted strip. A deleted reference with no surviving cells becomes `#REF!`. Formula strings retain their original text.
- Merged areas move when a shift starts before them, expand when cells are inserted inside them, and shrink when cells are removed. A partial shift that cuts through only part of a merged area removes the merge. Other merged areas retain their positions unless the shift reaches them.
- Insert rows above an Excel table, within its data, or before its totals row. Full-width downward shifts work as well. The table and its filter range follow the new bounds, and inserted data rows receive their calculated-column formulas. Structured references to those columns retain their text. Shared cell formulas are expanded to individual formulas before rewriting. Excel expands direct references and defined names that cover a complete table data range at its first row or before totals.
- The workbook's R1C1 display setting is preserved. Excel stores direct R1C1 references as A1 in XLSX, including function arguments. The engine follows direct references in dynamic range endpoints such as `INDEX(...):INDEX(...)` and `A1:INDEX(...)`. It preserves literal text passed to `INDIRECT` and updates direct references in expressions such as `INDIRECT(ADDRESS(ROW(A1),COLUMN(A1),1,FALSE),FALSE)`. Excel interprets R1C1 address text using its local reference notation.
- In lookup and database functions, explicit search, data, criteria and field-cell references follow structural edits. Literal field names and numeric column indexes keep their values.
- Defined names can use absolute A1 references with an explicit sheet or a worksheet-local scope. Renaming also updates explicit sheet qualifiers in relative defined names. Indirect references encoded as text retain their original text.
- Successful edits clear stored formula results across the workbook, including possible indirect dependents, and remove the calculation chain on save. The workbook requests full recalculation on load and retains its calculation mode.
- The engine validates the workbook before applying a structural edit. Unsupported cases include external, 3-D, spill and compound range forms with unsupported endpoints, array and data-table formulas, charts, pivots, hyperlinks, data validations, conditional formatting, and unknown reference-bearing extensions. A shift also requires the edited worksheet to have no standalone filters, cell-anchored drawings, or active sheet protection. Table column shifts, deletions through a table, shifts that cover only part of a table, and changes to table ranges with saved sort state raise `NotSupportedException`. Inserting or removing whole columns requires default column formatting. Existing cell styles move with their cells and inserted cells use the workbook's existing row and column defaults.

The engine is licensed under the MIT License.
