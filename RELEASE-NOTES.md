# Release notes

## Unreleased

### Breaking Changes

- BREAKING CHANGE: `CompuMaster.Excel.ExcelOps` adds a strict default for known structural formula and reference update differences in `CompuMaster.Excel.EpplusFreeFixCalcsEdition`, `CompuMaster.Excel.EpplusPolyformEdition`, `CompuMaster.Excel.FreeSpireXls`, and `CompuMaster.Excel.SpireXls`. Affected edits now throw `StructuralReferenceCompatibilityException` before changing the workbook. To retain the previous behavior, set `ExcelDataOperationsOptions.AllowKnownStructuralReferenceDifferences = True` before constructing the engine.

### Changes

- `CompuMaster.Excel.OpenXml` preserves registered external workbook references when rewriting local references during structural edits.
- `CompuMaster.Excel.ExcelOps` adds single-sheet and group sheet transfer methods with positional overloads, destination names and indexes in the results, and a per-call copy formula mode. `CompuMaster.Excel.MicrosoftExcel` supports native copy and move, including chart sheets. `CompuMaster.Excel.OpenXml` supports both copy modes and moves for compatible groups of worksheets and classic chart sheets. A move updates supported direct references in the source and destination workbooks.
- `CompuMaster.Excel.OpenXml` updates supported classic chart references, cached points, and drawing anchors during structural edits, including complete source deletions in supported multiple-series column, line and scatter charts. Transfers can include embedded classic charts. Worksheet renames also update supported ChartEx source names. `CompuMaster.Excel.MicrosoftExcel` implements the common worksheet rename method using Excel's native reference updates.
- `CompuMaster.Excel.OpenXml` adds workbook sessions for incoming reference rewrites, move updates in additional open workbooks, and links between workbooks saved after the transfer.
- `CompuMaster.Excel.OpenXml` supports deleting selected rows in an Excel table's data body, including the first data row, deleting the totals row, and moving a table through whole-column edits before it.
- `CompuMaster.Excel.OpenXml` supports clearing an entire supported table data body with Excel's retained insertion row. Whole-column edits update table columns, structured references and supported table filters, including tables with hidden headers. After deleting a filtered column, the engine reapplies supported numeric filters to the remaining cells.
- `CompuMaster.Excel.OpenXml` transfers supported workbook and worksheet-local defined names, including constants, absolute ranges and calculated names. It preserves their bindings during name collisions, including dependencies between names.
- `CompuMaster.Excel.OpenXml` rewrites supported dynamic spill references and uncached ChartEx source ranges and drawing anchors during structural edits. Deleting an entire supported ChartEx value source removes the affected series and retains the chart.
- `CompuMaster.Excel.OpenXml` preserves direct 3-D coordinates during single-sheet shifts and updates the endpoint names when a worksheet is renamed.
