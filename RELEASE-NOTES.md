# Release notes

## Unreleased

### Breaking Changes

- BREAKING CHANGE: `CompuMaster.Excel.ExcelOps` adds a strict default for known structural formula and reference update differences in `CompuMaster.Excel.EpplusFreeFixCalcsEdition`, `CompuMaster.Excel.EpplusPolyformEdition`, `CompuMaster.Excel.FreeSpireXls`, and `CompuMaster.Excel.SpireXls`. Affected edits now throw `StructuralReferenceCompatibilityException` before changing the workbook. To retain the previous behavior, set `ExcelDataOperationsOptions.AllowKnownStructuralReferenceDifferences = True` before constructing the engine.

### Changes

- `CompuMaster.Excel.OpenXml` preserves registered external workbook references when rewriting local references during structural edits.
- `CompuMaster.Excel.ExcelOps` adds single-sheet and group sheet transfer methods with positional overloads, destination names and indexes in the results, and a per-call copy formula mode. `CompuMaster.Excel.MicrosoftExcel` supports native copy and move, including chart sheets. `CompuMaster.Excel.OpenXml` supports both copy modes and moves for compatible groups of worksheets and classic chart sheets. A move updates supported direct references in the source and destination workbooks.
- `CompuMaster.Excel.OpenXml` updates supported classic chart references, cached points, and drawing anchors during structural edits. Transfers can include embedded classic charts. Worksheet renames also update supported ChartEx source names. `CompuMaster.Excel.MicrosoftExcel` implements the common worksheet rename method using Excel's native reference updates.
- `CompuMaster.Excel.OpenXml` adds workbook sessions for incoming reference rewrites, move updates in additional open workbooks, and links between workbooks saved after the transfer.
- `CompuMaster.Excel.OpenXml` supports deleting selected rows in an Excel table's data body and moving a table through whole-column edits before it.
