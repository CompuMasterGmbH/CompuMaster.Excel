# Excel formula conformance

[Microsoft's worksheet function index](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical) supplies function names, categories, version markers and per-function documentation. The [XLSX function specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf) supplies additional storage tokens. Tokens saved by Excel in the repository's reference workbooks take precedence.

Run `pwsh -File tools/update-excel-function-catalog.ps1` to refresh [the catalog](excel-function-catalog.csv). The script also accepts local HTML snapshots via `-SupportIndexPath` and `-XlsxFunctionsPath`, and an alternate destination via `-OutputPath`.

- `OpenXmlEvaluator` records names routed to the formula evaluator or handled by the formula parser. A function can still have unsupported arguments or result shapes.
- `ReferenceFormulaCells` counts saved worksheet formulas containing the named function, including nested calls. A cell counts once per function. The catalog records the published index and sample inventory. The Windows Excel reference build determines availability and values for each tested variant.

Run `pwsh -File tools/update-excel-formula-variants.ps1` to regenerate [the observed variants](excel-function-variants.csv) from the reference workbooks and the function catalog. Each row groups calls by argument count, argument kinds, calendar system and, for direct calls, saved result shape and anchor type or error. Small integer literals are recorded so common option values can be located. The example columns identify a workbook, worksheet, cell and, where saved, the array result range. The count records call occurrences, including repeated calls inside a formula.

A nested call has argument information and an example location. Its type and shape depend on the surrounding expression, so the report leaves those columns empty. The report describes the saved fixtures. Engine tests determine which of those formulas and array results recalculate successfully. Use the function catalog to locate functions with no example yet.
Use Microsoft's argument definitions to design separate cases for defaults, optional modes, input types, errors and result shapes. When a detail is unspecified, keep an Excel-calculated case that records it explicitly. Save cases in `ExcelOpsTest/test_data/FormulaConformance` and synchronize the workbook copies in the Spire test projects. `FormulaConformanceFixtureTest` clears the results and compares the recalculated XLSX cells against saved Excel values. Cases that require unsupported behavior have their own tests for the saved value and the reported limitation.

Run the shared formula coverage report with an exact test filter. Use a test project that contains the engines of interest.

```powershell
dotnet test ExcelOpsTest/TestExcelOps.vbproj -f net8.0 --filter 'Name=ReportsFormulaReferenceCoverageAgainstExcel' --logger 'console;verbosity=detailed'
```

Run the same filter on the FreeSpire and Spire test projects to include those engines. Run the Microsoft Excel comparison on Windows with Microsoft Excel installed. The test uses the German reference culture, clears the caches in temporary workbook copies, and reports differences and totals per engine. It covers the worksheet formula cases in the reference workbooks. External workbook links and the iterative calculation sequences have dedicated tests. The Open XML conformance tests separately assert typed cell results.
