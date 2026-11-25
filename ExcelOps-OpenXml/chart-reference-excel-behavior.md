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
| Rename the source worksheet | All explicit sheet qualifiers follow the new name. Excel doubles an apostrophe in the quoted qualifier. |

The embedded chart on the edited worksheet was configured to move and resize with cells. Its anchor followed whole-row and whole-column edits before it, including insertion immediately after the source range. The chart on the dashboard stayed anchored on that sheet. In the tested partial shift, the source chart's anchor also stayed in place.

## Survey by chart type

The survey enumerates all 103 values in Microsoft's `XlChartType` list and records the chart parts before editing, after inserting a row inside the data, and after deleting a source column. It uses embedded charts and inputs appropriate for stock, scatter, bubble, and map charts. The dedicated reference tests above also cover chart sheets and drawing placement. [XlChartType enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype) · [Survey CSV](chart-type-reference-survey.csv) · [Survey script](../tools/compare-excel-chart-references.ps1)

On desktop Excel 16.0 build 20625, the survey obtained reference snapshots for 85 chart requests. Of these, 77 used classic chart parts and eight used ChartEx. Excel reported a generic combination type for the custom combination and for the three built-in combination presets created directly through `AddChart2`. The saved parts retain the corresponding bar and line or area and bar groups. The built-in combinations expand their source ranges on row insertion and retain invalid references for the deleted series on column removal. Two enum values are chart selectors. Excel rejected the 16 `Ex` values when supplied directly to `AddChart2`. It also rejected them when assigned to an existing chart. The CSV records these outcomes separately.

The eight ChartEx cases include treemap, histogram, waterfall, sunburst, box and whisker, Pareto, funnel, and map. Their chart formulas used hidden `_xlchart` defined names in the tested packages. The survey resolves those names to the sheet ranges to show the effect of each edit. The histogram, Pareto, and map examples lost their source references after the input column was deleted.

## Open XML support

The engine rewrites supported direct source references in classic chart parts. It maps cached points onto the new coordinates and updates supported cell anchors on the edited worksheet. The tests compare the saved references and points for each scenario above and reopen the Open XML output in Excel. The transfer tests also reopen copied embedded charts and chart sheets in both formula modes.

Worksheet renames and supported shifts update the workbook-scoped `_xlchart` source names used by supported ChartEx charts. The row insertion and full value-column deletion tests cover all eight ChartEx types produced by the survey. A deleted category range or series name continues to resolve to `#REF!` in the tested charts. Deleting an entire value range removes its series. If no values remain, Excel saves the empty chart without plot axes and retains its drawing anchor. The tests compare resolved references, series, axes and anchors, then reopen the files in Excel.

Complete source deletions in the tested multiple-series column, line and scatter charts keep the series and update their point caches. Classic multiple-series chart roles outside those supported series remain guarded. Supported ChartEx transfers cover embedded charts and chart sheets with direct source ranges. Changing or transferring cached ChartEx dimensions, deleting a value dimension from a ChartEx series with several numeric dimensions, and named or extended classic chart formulas remain unsupported. [Reference tests](../ExcelOpsTest/ExcelOpsTests.Engines/MsExcelOpsChartReferenceTest.vb) · [Sheet transfer behavior](worksheet-transfer-excel-behavior.md) · [Chart source references in Open XML](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.formula)
