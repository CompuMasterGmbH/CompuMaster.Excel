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

Excel and Open XML implement `MoveSheet` and `MoveSheets` for supported groups containing worksheets and chart sheets. Open XML updates direct references in the two participating workbooks, stages both changes together, and requires a saved path when it creates a workbook link. Transfer overloads accept a zero-based index in the combined destination tab order. Each `SheetTransferResult` contains the assigned name and actual index in that combined order. A shared session is needed to extend this behavior to additional open workbooks and to rebind links when an unsaved workbook acquires a filename.

## Requirements for an Open XML workbook session

1. Maintain an identity for each open workbook, including an unsaved workbook and one that acquires a path on its first save.
2. Determine destination sheet names and the full set of transferred sheets before rewriting references. Preserve the identity of any existing sheets with the same name.
3. On a copy, redirect references between simultaneously copied sheets to their copies and keep other references bound to their original sheets. On a move, update registered dependents on both sides and convert qualifiers between local and external as necessary.
4. Update references in open dependents when an unsaved workbook receives a path. Persist those changes when the dependent workbook is saved.
5. Use the existing cross-workbook insertion and worksheet transfer tests as the Excel reference. Defined-name collisions, 3-D formulas, and worksheet objects such as charts and tables require separate behavior tests before claiming compatibility.

The native reference tests are in [MsExcelOpsWorksheetTransferTest.vb](../ExcelOpsTest/ExcelOpsTests.Engines/MsExcelOpsWorksheetTransferTest.vb). Microsoft describes the distinction between workbook-link display and storage in its [workbook link documentation](https://support.microsoft.com/en-us/excel/description-of-workbook-link-management-and-storage-in-excel).
