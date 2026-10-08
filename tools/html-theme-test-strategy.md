# HTML export theme color references (issue #9)

## Reference generation

Run `./tools/new-html-theme-reference-fixtures.ps1` on Windows with Microsoft Excel
installed. It uses a private Excel instance and creates the test matrix directly
in Excel. The legacy palette is read from the existing Excel-generated legacy
reference workbook. The older synthetic fixtures have been removed.

The new files cover the legacy Office palette, the installed Office theme, Office
2013 - 2022, Ion, and a custom Office palette with a pure red Accent 2. The installed
themes are applied by Excel. Excel also sets the custom accent, writes the XLSX
files, and captures colors after closing and reopening each saved workbook.

Each workbook has 39 reference cells: all six accents as font and solid fill,
each at tint 0, -0.25 and +0.60, two fixed red RGB controls, and one cell without
a fill. Visible sample texts describe the formatting and include resolved RGB
values. Adjacent columns show the expected font color of column B and fill color
of column C. The custom red case includes "Roter Text auf weissem Untergrund".

`HtmlExportThemeExcelReferences.csv` captures sheet/address, case ID, sample text,
displayed font/fill RGB, Excel version/build, and the SHA-256 of the saved workbook.
An empty fill value means no fill, not a white fill. Values come from
`Range.DisplayFormat`, with Excel's OLE_COLOR byte order converted to CSS hex.
Expected RGB values are never computed using an export engine or its tint helper.

The output guard rejects existing files unless `-Overwrite` is explicitly supplied.
Reference regeneration is an intentional maintenance operation. Changing the
theme in Excel after capture requires a new capture and visual review.

## Review and regression tests

1. Use the captured Excel colors as the independent automated reference. The manual
   review of the saved workbooks and HTML exports follows successful automated tests.
2. Verify each workbook's hash against the CSV, then export it through each engine.
3. Locate the HTML table cell by its unique sample text (or an explicit address
   mapping) and resolve its effective CSS font and background colors. Do not merely
   search the whole HTML document for a color, since controls may contain it too.
4. Compare colors case by case, with engine, workbook, sheet, address, case ID,
   expected color, and actual color in failure messages. Test inherited styles and
   explicit styles consistently. No fill must remain distinct from an explicit
   white fill. Add border references when border-color export is implemented.
5. Run durable tests from checked-in workbooks and CSV without requiring Excel in
   CI. Keep unsupported-engine outcomes explicit until their export is implemented.

This first set isolates theme resolution from conditional formatting and table
styles. Those need separate fixtures if their rendered appearance is included in
the HTML export contract. Built-in Office themes can change between Office releases;
checked-in workbook themes and captured colors are the stable test inputs.

## Implemented and verified

- Both EPPlus HTML exporters resolve the embedded workbook palette for each export.
- EPPlus 4 exposes a detached `ThemeXml` snapshot resolved through the workbook's
  package relationship. EPPlus 8 uses its existing `ThemeManager` API.
- Tint/shade uses Excel's integer HLS conversion and separate luminance-term
  truncation, including neutral gray rounding. Additional direct Excel measurements
  cover unusual tint values and black/white endpoints.
- Shared engine regression tests compare the exact font and fill RGB of every
  captured cell in sheet and workbook HTML, including byte-array/stream inputs.
  A separate case changes the red theme colors to equivalent direct RGB colors
  with their captured tints, to verify tint-aware color caching.
- Unit tests cover detached theme snapshots, non-default theme part names, missing
  themes, tint boundaries, and invalid tint inputs. No Excel installation is needed
  for these color assertions in CI.
- HTML export XML documentation has been reviewed against the current implementation.
  The API checker already includes both option classes and `HtmlDocumentExportParts`
  without temporary HTML documentation exclusions. Documentation now states the
  actual section/document output, defaults, engine limitations, and legacy async behavior.

## Remaining ticket work

- Final manual visual review of the generated workbooks and HTML exports.
- Complete and verify HTML export support for the remaining engines.
- Revisit the HTML export API shape while adding the remaining engines. In particular,
  review ignored row/column range options, one-based EPPlus header row numbers,
  empty-sheet placeholder handling, and the non-awaitable legacy Async Sub methods.
  The asynchronous worksheet export currently emits engine content without the
  document/section wrappers used by the synchronous overloads. Preserve compatibility
  when improving these behaviors.
- Delete `9-add-export-feature-to-html` locally and remotely only at final cleanup,
  after confirming no valuable unmerged work remains and all required pipelines
  have passed. Keep the branch available until then.

## Microsoft references

- [Workbook.ApplyTheme](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.applytheme)
- [Range.DisplayFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.range.displayformat)
- [Interior.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.tintandshade)
