# Excel formula conformance

[Microsoft's worksheet function index](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical) supplies function names, categories, version markers and per-function documentation. The [XLSX function specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf) supplies additional storage tokens. Tokens saved by Excel in the repository's reference workbooks take precedence.

Run `pwsh -File tools/update-excel-function-catalog.ps1` to refresh [the catalog](excel-function-catalog.csv). The script also accepts local HTML snapshots via `-SupportIndexPath` and `-XlsxFunctionsPath`, and an alternate destination via `-OutputPath`.

- `OpenXmlEvaluator` records names routed to the formula evaluator or handled by the formula parser. A function can still have unsupported arguments or result shapes.
- `ReferenceFormulaCells` counts saved worksheet formulas containing the named function, including nested calls. A cell counts once per function. The catalog records the published index and sample inventory. The Windows Excel reference build determines availability and values for each tested variant.

Use Microsoft's argument definitions to design separate cases for defaults, optional modes, input types, errors and result shapes. When a detail is unspecified, keep an Excel-calculated case that records it explicitly. Save cases in `ExcelOpsTest/test_data/FormulaConformance` and synchronize the workbook copies in the Spire test projects. `FormulaConformanceFixtureTest` clears the results and compares the recalculated XLSX cells against saved Excel values. Cases that require unsupported behavior have their own tests for the saved value and the reported limitation.
