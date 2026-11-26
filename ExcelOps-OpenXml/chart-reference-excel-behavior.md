# Excel chart reference behavior

The reference tests use desktop Excel on Windows to edit a data worksheet and save the result as XLSX. The same data feeds a column chart embedded on the data sheet, an XY scatter chart on a dashboard worksheet, and a column chart on a chart sheet. A chart sheet is a sheet dedicated to displaying a chart. [Microsoft documentation](https://support.microsoft.com/de-de/excel/move-or-resize-a-chart)

## Reference cases

| Edit | Observed Excel behavior |
| --- | --- |
| Insert before the source range or at its first data row | The range moves and retains its length. The series name and linked title follow their own cells. |
| Insert inside a source range | The ranges expand to include the inserted row. A blank point has no numeric cache entry, and following points retain their new indices. |
| Insert immediately after a source range | The chart data ranges keep their original bounds. |
| Delete a row inside a source range | The ranges shrink and the saved cache reflects the remaining points. |
| Insert a column before the data | The series name, category or X range, value range, and title move with the corresponding cells. |
| Shift only cells in the values column down | The values range expands. The category or X range keeps its bounds. The saved values cache records the inserted gap. |
| Delete the entire categories or X range | In the saved column chart, Excel drops the category reference. In the saved scatter chart, Excel keeps an X reference with `#REF!`. The values range still refers to the original surviving cells. |
| Delete the entire values range | The XLSX chart keeps an invalid values reference and a cached fallback point with value 1 in the tested charts. If the series name cell is also deleted, Excel drops the single-series chart's name reference. In the tested multiple-series column, line and scatter charts, the name reference remains and its cache contains `#REF!`. |
| Clear the entire data body of a table used by a chart | The table retains a blank insertion row. The tested column and funnel chart series keep references to that row. The classic chart stores a point count of one with no cached point for the blank cells. Sources below the table shift up with their cells. |
| Rename the source worksheet | All explicit sheet qualifiers follow the new name. Excel doubles an apostrophe in the quoted qualifier. |

The embedded chart on the edited worksheet was configured to move and resize with cells. Its anchor followed whole-row and whole-column edits before it, including insertion immediately after the source range. The chart on the dashboard stayed anchored on that sheet. In the tested partial shift, the source chart's anchor also stayed in place.

## Embedded chart placement

The placement tests put chart data on an unaffected worksheet and edit only the worksheet hosting the chart. They compare Excel's chart presence, position and dimensions in points for whole rows and whole columns. The cases include default dimensions and explicit column widths and row heights. The standard three placement modes are defined by Excel's chart object API. [Excel placement modes](https://learn.microsoft.com/en-us/office/vba/api/excel.xlplacement)

| Edit relative to the chart | Move and size with cells | Move without sizing | Free floating |
| --- | --- | --- | --- |
| Insert or delete before the chart | Moves with the cells and retains its size. | Moves with the cells and retains its size. | Retains its position and size. |
| Insert or delete after the chart | Retains its position and size. | Retains its position and size. | Retains its position and size. |
| Delete some of the covered rows or columns | Shrinks along the edited axis. If the first anchor cell is removed, its edge moves to the deletion boundary. | Retains its size. If its first anchor cell is removed, its edge moves to the deletion boundary. | Retains its position and size. |
| Delete all covered rows or all covered columns | Removes the chart or retains it with zero height or width, depending on the original used range. | Retains the chart and its size, with its starting edge at the deletion boundary. | Retains the chart, its position and its size. |
| Insert within the chart's extent | Expands along the edited axis. | Retains its position and size. | Retains its position and size. |

Insertion at the first anchor cell moves the chart with that cell when the placement is cell-dependent. Insertion before the last covered cell counts as insertion inside the chart. If the chart ends exactly on a cell boundary, an edit starting at that boundary is outside its extent.

When the chart collapses, Excel compares the original used range with the start of the chart on the other axis. For a chart starting at B3, deleting its entire row span retains a zero-height chart if the original used range includes column C or later. Deleting its entire column span retains a zero-width chart if the original used range includes row 4 or later. A formatted cell counts. A cell removed by the same edit still counts.

Excel removes all eight tested types of ChartEx chart when deleting their full column span, even when the used range would retain a zero-width classic chart. The same ChartEx types can remain with zero height after deleting their full row span.

Open XML applies the placement rules to supported embedded classic and ChartEx charts. It removes detached chart relationships along with a deleted drawing. The cross-engine fixture in `test_data/EmbeddedChartPlacement.xlsx` uses all three placement modes on one dashboard. The Excel Interop tests also verify the rendered positions and sizes after reopening the Open XML output.

Both EPPlus adapters and both Spire adapters produced placement differences on the cross-engine fixture. They reject whole-row and whole-column edits on a worksheet with embedded charts by default. `AllowKnownStructuralReferenceDifferences` permits the adapter's existing behavior for callers that accept those differences.

## Survey by chart type

The survey enumerates all 103 values in Microsoft's `XlChartType` list and records the chart parts before editing, after inserting a row inside the data, and after deleting a source column. It uses embedded charts and inputs appropriate for stock, scatter, bubble, and map charts. The dedicated reference tests above also cover chart sheets and drawing placement. [XlChartType enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype) · [Survey CSV](chart-type-reference-survey.csv) · [Survey script](../tools/compare-excel-chart-references.ps1)

On desktop Excel 16.0 build 20625, the survey obtained reference snapshots for 85 chart requests. Of these, 77 used classic chart parts and eight used ChartEx. Excel reported a generic combination type for the custom combination and for the three built-in combination presets created directly through `AddChart2`. The saved parts retain the corresponding bar and line or area and bar groups. The built-in combinations expand their source ranges on row insertion and retain invalid references for the deleted series on column removal. Two enum values are chart selectors. Excel rejected the 16 `Ex` values when supplied directly to `AddChart2`. It also rejected them when assigned to an existing chart. The CSV records these outcomes separately.

The eight ChartEx cases include treemap, histogram, waterfall, sunburst, box and whisker, Pareto, funnel, and map. Their chart formulas used hidden `_xlchart` defined names in the tested packages. The survey resolves those names to the sheet ranges to show the effect of each edit. The histogram, Pareto, and map examples lost their source references after the input column was deleted.

## Open XML support

The engine rewrites supported direct source references in classic chart parts. It maps cached points onto the new coordinates and updates supported cell anchors on the edited worksheet. The tests compare the saved references and points for each scenario above and reopen the Open XML output in Excel. The transfer tests also reopen copied embedded charts and chart sheets in both formula modes.

Worksheet renames and supported shifts update the workbook-scoped `_xlchart` source names used by supported ChartEx charts. The row insertion and full value-column deletion tests cover all eight ChartEx types produced by the survey. A deleted category range or series name continues to resolve to `#REF!` in the tested charts. Deleting an entire value range removes its series. If no values remain, Excel saves the empty chart without plot axes and retains its drawing anchor. The tests compare resolved references, series, axes and anchors, then reopen the files in Excel.

Complete source deletions in the tested multiple-series column, line and scatter charts keep the series and update their point caches. Classic multiple-series chart roles outside those supported series remain guarded. Supported ChartEx transfers cover embedded charts and chart sheets with direct source ranges. Changing or transferring cached ChartEx dimensions, deleting a value dimension from a ChartEx series with several numeric dimensions, and named or extended classic chart formulas remain unsupported. [Reference tests](../ExcelOpsTest/ExcelOpsTests.Engines/MsExcelOpsChartReferenceTest.vb) · [Sheet transfer behavior](worksheet-transfer-excel-behavior.md) · [Chart source references in Open XML](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.formula)
