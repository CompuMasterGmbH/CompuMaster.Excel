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
| Delete the entire values range | The XLSX chart keeps an invalid values reference and a cached fallback point with value 1 in the tested charts. If the series name cell is also deleted, Excel drops that name reference. |
| Rename the source worksheet | All explicit sheet qualifiers follow the new name. Excel doubles an apostrophe in the quoted qualifier. |

The embedded chart on the edited worksheet was configured to move and resize with cells. Its anchor followed whole-row and whole-column edits before it, including insertion immediately after the source range. The chart on the dashboard stayed anchored on that sheet. In the tested partial shift, the source chart's anchor also stayed in place.

## Survey by chart type

The survey enumerates all 103 values in Microsoft's `XlChartType` list and records the chart parts before editing, after inserting a row inside the data, and after deleting a source column. It uses embedded charts and inputs appropriate for stock, scatter, bubble, and map charts. The dedicated reference tests above also cover chart sheets and drawing placement. [XlChartType enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype) · [Survey CSV](chart-type-reference-survey.csv) · [Survey script](../tools/compare-excel-chart-references.ps1)

On desktop Excel 16.0 build 20625, the survey obtained reference snapshots for 82 chart requests. Of these, 74 used classic chart parts and eight used ChartEx. Excel normalized the custom combination request to a mixed chart with bar and line groups. Two enum values are chart selectors. The 19 remaining values were rejected during creation with this setup. These were three predefined combination requests and the 16 `Ex` variants. The CSV records these outcomes separately.

The eight ChartEx cases include treemap, histogram, waterfall, sunburst, box and whisker, Pareto, funnel, and map. Their chart formulas used hidden `_xlchart` defined names in the tested packages. The survey resolves those names to the sheet ranges to show the effect of each edit. The histogram, Pareto, and map examples lost their source references after the input column was deleted.

## Open XML milestone

- Resolve the chart parts through the drawing relationships of both worksheet parts and chart sheet parts.
- Rewrite the chart's source references using the edited worksheet and the reference's role. Preserve the Excel distinction between a missing category reference and an invalid X or values reference.
- Keep saved chart caches consistent with the new references, including gaps. Handle the drawing anchor according to its host worksheet and placement. Stage chart, drawing, and worksheet changes together before applying them.
- Verify round trips with the reference cases and Excel reopening. Defined names, external workbook chart sources, linked axis titles and labels, and specialized chart types are candidates for additional reference tests.

The current Open XML guard tests verify that unsupported chart edits throw before changing either kind of workbook. The tests cover both native behavior and the guard. [Reference tests](../ExcelOpsTest/ExcelOpsTests.Engines/MsExcelOpsChartReferenceTest.vb) · [Chart source references in Open XML](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.formula)
