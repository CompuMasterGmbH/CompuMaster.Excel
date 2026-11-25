# Release notes

## Unreleased

### Breaking Changes

- BREAKING CHANGE: `CompuMaster.Excel.ExcelOps` adds a strict default for known structural formula and reference update differences in `CompuMaster.Excel.EpplusFreeFixCalcsEdition`, `CompuMaster.Excel.EpplusPolyformEdition`, `CompuMaster.Excel.FreeSpireXls`, and `CompuMaster.Excel.SpireXls`. Affected edits now throw `StructuralReferenceCompatibilityException` before changing the workbook. To retain the previous behavior, set `ExcelDataOperationsOptions.AllowKnownStructuralReferenceDifferences = True` before constructing the engine.

### Changes

- `CompuMaster.Excel.OpenXml` preserves registered external workbook references when rewriting local references during structural edits.
- `CompuMaster.Excel.ExcelOps` adds single-sheet and group sheet transfer methods with positional overloads, destination names and indexes in the results, and a per-call copy formula mode. `CompuMaster.Excel.MicrosoftExcel` supports native copy and move, including chart sheets. `CompuMaster.Excel.OpenXml` supports both copy modes and moves for compatible groups of worksheets and classic chart sheets. A move updates supported direct references in the source and destination workbooks.
