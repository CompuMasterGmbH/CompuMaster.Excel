# Excel worksheet transfer behavior

The reference run used Windows desktop Excel 16.0, build 20625. The tests call the native worksheet and worksheet-group `Copy` and `Move` operations. They open the source and target in one Excel instance. Each case builds fresh workbooks and uses disposable local files when it saves them. The assertions check formula text, calculated values, and the primary external link in the XLSX package.

## References during a transfer

| Scenario | Excel behavior |
| --- | --- |
| A transferred sheet refers to itself | The reference stays local. An explicit sheet qualifier follows the name assigned in the target. |
| A transferred sheet refers to a sheet that stays in the source | The reference becomes external and continues to point to the source. A sheet with the same name in the target does not capture it. |
| The transferred sheet's name already exists in the target | Excel assigns a new name. In the tested first collision, `Transfer` became `Transfer (2)`. A pre-existing local formula in the target still refers to the original `Transfer` there. |
| A formula elsewhere points to the transferred sheet before a copy | It keeps pointing to the original, including an external reference already present in the target. |
| A formula elsewhere points to the transferred sheet before a move | In the tested open workbooks it follows the moved sheet. A formula left in the source becomes external. An existing external formula in the target becomes local. Both follow the name assigned in the target. |
| Two mutually referencing sheets are copied or moved as a group | Both transferred formulas refer to the sheets in the target. When both names collide, Excel uses both names assigned in the target. |
| The same two sheets are copied one at a time | The two copied formulas continue to refer to the original sheets in the source. Nothing automatically redirects the first copy's dependency to the second copy. |
| The same two sheets are moved one at a time | The first move creates an external dependency. Moving the second sheet into the same open target makes that dependency local again. |
| The operation creates a new workbook | Excel opens an unsaved destination. The same workbook link rules apply. |

## Workbook identity and paths

- With both workbooks open, a formula shows the other workbook's current name. An unsaved workbook has an Excel-assigned name, which depends on the session and its locale.
- When the linked workbook is saved for the first time, Excel updates references in the other open workbook immediately. If that dependent workbook was already saved, its file retains the previous link target until it is saved again. This applies in either direction after a move. The tests cover both save orders.
- The saved formula uses a numeric workbook index, for example `'[1]Data Set'!$A$1`. The link relationship holds the path. In the test with the source in an `Inputs` subdirectory, Excel kept a relative path from the target to the source. The reverse link resolved to the workbook in the parent directory. The local Excel build used a drive-rooted primary path for that reverse link and also wrote alternate path metadata. An implementation must preserve the identity of the linked workbook across storage representations.
- In the existing linked-workbook insertion test, an incoming reference follows a structural edit on the source only when the dependent workbook is open in the same Excel instance. Excel does not rewrite the closed dependent file.

## Formula modes on the shared API

`ExcelDataOperationsBase.CopySheet` and `CopySheets` accept a per-call `CopyWorksheetFormulaMode`. The default `ExcelCompatible` preserves the identity of the referenced sheets. `PreserveFormulas` keeps the original cell and supported chart formula text, which can make a same-named worksheet in the target supply the value. Excel's native provider implements `ExcelCompatible` and rejects `PreserveFormulas` before transferring sheets. Open XML implements both modes for supported worksheets and classic chart sheets, including mixed groups.

Excel and Open XML implement `MoveSheet` and `MoveSheets` for supported groups containing worksheets and chart sheets. Transfer overloads accept a zero-based index in the combined destination tab order. Each `SheetTransferResult` contains the assigned name and actual index in that combined order.

## Open XML workbook sessions

`OpenXmlWorkbookSession` connects participating workbook instances. It stages supported incoming reference changes before applying a structural edit and redirects references in additional attached workbooks after a move. It uses saved paths and temporary names to establish workbook identity. When the linked workbook receives a path, the session resolves the temporary name to that path. An attached dependent writes the resolved link on its next save. Microsoft Excel reopened the output and calculated the linked formulas in a move that was saved in both possible orders.

Excel-compatible transfer copies sheet-local definitions and imports the workbook names used by transferred cell formulas. Supported definitions include constants, absolute ranges and calculated expressions that use workbook-scoped names. Relative and mixed references to the transferred sheet using the name preserve their offsets, including when the sheet is renamed during transfer. Workbook-scoped definitions must explicitly qualify these references. When a dependency name already exists at workbook scope in the destination, the incoming definition uses worksheet scope. Excel qualifies references to it from other defined names so they still resolve to the transferred definition when the workbook is reopened. A workbook-scoped name still used in the source refers to the destination after a move, preserving its offsets. Relative references to other sheets in transferred names and references from retained sheets to a moved worksheet's local name remain outside the supported scope.

Excel stores ChartEx sources as workbook-scoped chart names. A copied or moved chart follows its transferred data sheet or an external link to a sheet left in the source workbook. Open XML preserves the destination's existing chart names and assigns distinct names to transferred sources when necessary. In `PreserveFormulas` mode the chart source formulas keep their text, and a different destination definition for the same source name prevents the copy.

Excel cleared the series when a ChartEx chart sheet preceded its separate data sheet in the source tab order and both were moved together. The same occurred when an embedded ChartEx chart's host preceded its separate data sheet and a destination collision renamed the data sheet. In the reverse tab order, Excel retained the tested sources. Open XML detects the divergent cases before changing either workbook. With `AllowKnownStructuralReferenceDifferences` enabled on the source engine, Open XML moves the sheets and preserves the chart series.

Excel-compatible transfer of 3-D formulas and table transfers remain outside the supported scope. Incoming structural updates cover direct links from cells, supported tables, classic charts and uncached ChartEx sources. The session writes resolved external link paths as absolute file URIs.

The native reference tests are in [MsExcelOpsWorksheetTransferTest.vb](../ExcelOpsTest/ExcelOpsTests.Engines/MsExcelOpsWorksheetTransferTest.vb). Microsoft describes the distinction between workbook-link display and storage in its [workbook link documentation](https://support.microsoft.com/en-us/excel/description-of-workbook-link-management-and-storage-in-excel).
