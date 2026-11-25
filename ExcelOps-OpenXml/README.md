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
- The default save preserves workbook calculation mode and cached formula values. Caches can be stale after input edits. `RequestRecalculationOnOpen()` requests automatic full recalculation in Excel. The existing save cache options can clear cached formula results in the output. Reading a formula cell without a cached value raises `MissingCalculatedCellValueException`.
- Inspect cached formula results without calculating by calling `HasMissingCachedFormulaResult`. The `ResetCellValueFromFormulaCell` and `ResetCellValueFromFormulaCellInWholeWorkbook` methods clear stored results in the open workbook while retaining formula text.
- Export worksheets and workbooks as HTML tables with merged cells, common number formats, and font, fill, alignment and wrap styles. The export follows the EPPlus 4 CM table conventions. The default uses cached formula results and throws `MissingCalculatedCellValueException` if an exported formula result is missing. Set `MissingFormulaResultText` to explicitly render replacement text, including an empty string.
- The explicit `CalculateIfMissing` lookup raises `NotSupportedException` when a result must be calculated. Existing recalculation methods continue to raise `FeatureDisabledException`. Structural row, column, cell and sheet operations, edits inside shared, array or data-table formula ranges, and chart rendering raise `NotSupportedException`.

The engine is licensed under the MIT License.
