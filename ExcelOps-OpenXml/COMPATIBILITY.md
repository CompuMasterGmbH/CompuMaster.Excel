# Open XML Excel compatibility

Use the [worksheet function catalog](../tools/excel-function-catalog.csv) to find the official Microsoft documentation, evaluator coverage, and Excel reference cases for each function. The [engine guide](README.md) describes supported workbook operations. This page gives examples for behavior that differs from Excel and for extensions supplied by the Open XML engine.

## Formula syntax and calculation

- **Formula syntax in the API.** Write invariant formula text, for example `SUM(1.5,2)`. A German Excel user can enter `=SUMME(1,5;2)` in the Excel interface. `FormulaCulture` controls text and number conversions during calculation.
- **Missing and existing results.** Suppose `A1` contains `B1+1`, its saved result is 11, and the application writes 20 to `B1`. In manual mode, `CalculateIfMissing` still reads 11. Use `RecalculateCell("Sheet1", 0, 0)` before reading `A1` to get 21. To calculate after value edits, enable the evaluator and set `CalculateOnCellEdit = True` when opening the workbook in automatic mode.
- **Known calculation limitations.** `AMORDEGRC(1200,DATE(2008,7,1),DATE(2008,12,31),150,0.25,0.15)` raises `FormulaComplexityNotSupportedException` in Open XML. Use a stored Excel result or perform this calculation in Excel. In another known numerical difference, Excel returns `#NUM!` for `CHISQ.INV(0.95,10000000000)` while Open XML returns an estimated quantile. Use the Excel engine when that error behavior is part of the workbook contract. [AMORDEGRC](https://support.microsoft.com/en-us/excel/functions/amordegrc-function) · [CHISQ.INV](https://support.microsoft.com/en-us/excel/functions/chisq-inv-function)
- **Pivot tables.** `GETPIVOTDATA("Sales",Report!A3,"Region","East")` reads the saved visible aggregate when `Report!A3` identifies a supported worksheet-based pivot table. If the source data changes, refresh the pivot in Excel before handing it back to Open XML. OLAP pivots, grouped-field lookups, and report-filter fields require a compatible pivot engine. [GETPIVOTDATA](https://support.microsoft.com/en-us/excel/functions/getpivotdata-function)
- **Linked workbooks.** For a formula such as `'[Rates.xlsx]Sheet1'!A1+1`, open the source and the dependent workbook in the same `ExcelFormulaSession` before recalculation. Open XML resolves the link from the workbooks in that session. A workbook with an unavailable source can still be used with its saved results.

## Workbook operations and output

- **Last worksheet.** Removing the last worksheet from a workbook can leave Open XML with an empty worksheet collection. Add a worksheet before saving if the application needs a specific name or contents. Saving an empty workbook creates a blank worksheet so the file opens again in Excel.
- **Structural edits.** If a deletion selection consists entirely of hidden columns and a table filter is active, Open XML raises `StructuralReferenceCompatibilityException` before changing the workbook. Clear the filter or set `AllowKnownStructuralReferenceDifferences = True` in the options before opening a workbook that requires the Open XML behavior. See the [structural edit guide](README.md#structural-edits).
- **HTML export.** When a formula result is missing, the default export raises `MissingCalculatedCellValueException`. Set `HtmlSheetExportOptions.MissingFormulaResultText = "#Value?"` to show a visible marker, or enable calculation and select `FormulaCalculationMode.CalculateIfMissing` to calculate missing values. See the [HTML export examples](../README.md#missing-formula-results-in-html).

## VBA and Excel 4.0 macros

Open XML edits the workbook package. To run existing VBA or Excel 4.0 (XLM) code, open the workbook in desktop Excel with the applicable macro settings. For a headless application, port the logic to .NET and choose how to invoke it based on what the original code does. [VBA worksheet functions](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel) · [VBA macro triggers](https://support.microsoft.com/en-us/excel/run-a-macro-in-excel) · [Excel 4.0 macros](https://support.microsoft.com/en-au/excel/working-with-excel-4-0-macros)

### Preserve or remove macro content

| Input and destination | Open XML behavior |
| --- | --- |
| A VBA workbook saved as `report.xlsm` or `report.xltm` | The macro-enabled format preserves the VBA project. Call `RemoveVbaProject()` to omit it on the next save. |
| A VBA workbook saved as `report.xlsx` or `report.xltx` | The macro-free format omits the VBA project. Worksheet formulas that call the old functions can still need a .NET implementation for recalculation. |
| An XLM workbook saved as `report.xlsm` or `report.xltm` | Open XML can carry the macro sheets through when saving the package. Execution remains an Excel operation. |
| An XLM workbook saved as `report.xlsx` or `report.xltx` | Open XML rejects the save with `NotSupportedException`. Use a macro-enabled destination or prepare a new macro-free template after migrating the XLM workflow. |

For example, this makes a macro-enabled copy and preserves the existing formula results for use in Excel. Passing `recalculateWorkbook:=False` lets the copy retain previously saved results, which may be stale after input edits.

```vbnet
workbook.SaveCopyAs("report.xlsm",
    recalculateWorkbook:=False,
    cachedCalculationsOption:=ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
```

`HasVbaProject` reports the presence of a VBA project. Use the XLM save rule above for workbooks with macro sheets, which are separate package parts.

### Port actions and event handlers

An event handler such as `Workbook_Open`, a button macro, or an XLM macro-sheet routine can perform actions. Port the actions to an application method and invoke it at the equivalent point in the application workflow. For example, if a workbook-open macro writes a status into `Inputs!A1`, call this method after opening the workbook.

```vbnet
Private Sub InitializeReport(workbook As OpenXmlExcelDataOperations)
    workbook.WriteCellValue("Inputs", 0, 0, "Ready")
End Sub
```

Keep that method in the application's open or command workflow. A worksheet can use the function registration described below when it needs a calculated result.

### Replace legacy native calls

A legacy macro sheet can use `CALL(A1,B2)` to invoke a native function with the registration ID in `A1`. `REGISTER.ID` can obtain that ID. Open XML reports native `CALL` and `REGISTER.ID` evaluation as unsupported. For a scalar calculation such as a rebate, migrate the old call to a worksheet formula such as `=REBATE(B2)` and register its .NET implementation. For a call that performs an action, port it to the application workflow. Use Excel with the original add-in when the workbook must run the legacy call as written. [CALL](https://support.microsoft.com/en-us/excel/functions/call-function) · [REGISTER.ID](https://support.microsoft.com/en-us/excel/functions/register-id-function)

## Application-defined worksheet functions

A cell formula such as `=REBATE(B2)` calls a function that returns a value. Port the original VBA function to .NET, register it under the same name, and pass the session when opening the workbook. The example calculates 25 when `B2` is 250.

```vbnet
Imports CompuMaster.Excel.ExcelOps
```

```vbnet
Dim session As New ExcelFormulaSession()
session.RegisterFunction("REBATE",
    Function(context, arguments)
        If arguments.Count <> 1 OrElse arguments(0).Kind <> ExcelFormulaValueKind.Number Then
            Return ExcelFormulaValue.FromError("#VALUE!")
        End If
        Return ExcelFormulaValue.FromNumber(CDbl(arguments(0).Value) * 0.1)
    End Function)

Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly) With {
    .DisableCalculationEngine = False,
    .FormulaSession = session
}

Using workbook As New OpenXmlExcelDataOperations(
    "report.xlsm", ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
    workbook.RecalculateAll()
End Using
```

The callback receives the typed scalar arguments and a context containing the workbook, worksheet, and formula cell. It can read cells in workbooks in the same session during the call. The callback returns a scalar value or an Excel error. Supply the registration whenever a workbook is opened for fresh calculation. The saved file stores the formula and the calculated result, so another reader can use the saved result. For fresh calculation in Excel, supply an equivalent Excel-side function, such as the original VBA function in a macro-enabled workbook. [Excel worksheet functions written in VBA](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel)

## Real-time data (RTD)

Both Open XML routes use the Excel formula syntax. The example uses a fictional provider identifier. The two topics select the temperature for Hall 1.

```excel
=RTD("Example.SensorRtd",,"Hall1","Temperature")
```

The second argument is the server name. Leave it empty for a local Windows server. When Excel and Open XML use the same installed server, the workbook can use the same formula in both. [RTD formula](https://support.microsoft.com/en-us/excel/functions/rtd-function) · [RTD COM server interface](https://learn.microsoft.com/en-us/office/vba/api/excel.irtdserver)

### Use a local Windows COM server

Install and register the RTD server using the provider's installer. Allow its `ProgID` in the application before opening the workbook. This controls which locally installed COM servers a workbook can activate.

```vbnet
Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly) With {
    .DisableCalculationEngine = False,
    .CalculateOnCellEdit = True,
    .AllowedNativeRtdProgIds = New String() {"Example.SensorRtd"}
}

Using workbook As New OpenXmlExcelDataOperations(
    "sensors.xlsx", ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
    workbook.RecalculateAll()
    ' Read cached formula results or export the workbook while it is open.
End Using
```

The first calculation subscribes to the topics. The adapter receives server notifications and retrieves the changed values. With `CalculateOnCellEdit = True` and the workbook set to automatic calculation, changed data schedules a calculation pass that updates dependent formulas. With automatic calculation disabled, call `RecalculateAll()` to use the latest received values. Topic subscriptions are shared within the workbook. If a formula changes from Hall 1 to Hall 2, the Hall 1 subscription is released at workbook close. If a background calculation raises an exception, Open XML reports it through `System.Diagnostics.Trace` and retains the previous results.

The native adapter supports local Windows COM servers. A formula such as `RTD("Example.SensorRtd","other-host","Hall1","Temperature")` requires an application provider. Open XML raises `FormulaComplexityNotSupportedException` when native activation is unavailable on the platform, the local `ProgID` is outside the allowed list, or the formula requests a remote COM server. An allowed server that is unavailable returns `#N/A`. Use an application provider if the COM implementation requires services from a running Excel application.

### Use an application data provider

Register only the identifier that the application owns. Other `RTD` formulas in the workbook can use other registered providers, or allowed local COM servers on Windows.

```vbnet
Imports System.Collections.Concurrent
Imports CompuMaster.Excel.ExcelOps
```

```vbnet
Dim latest As New ConcurrentDictionary(Of (String, String), Double)()
latest(("Hall1", "Temperature")) = 21.4

Dim session As New ExcelFormulaSession()
session.RegisterRtdProvider("Example.SensorRtd",
    Function(context, arguments)
        If arguments.Count <> 4 OrElse
           arguments(2).Kind <> ExcelFormulaValueKind.Text OrElse
           arguments(3).Kind <> ExcelFormulaValueKind.Text Then
            Return ExcelFormulaValue.FromError("#VALUE!")
        End If
        Dim topic = (CStr(arguments(2).Value), CStr(arguments(3).Value))
        Dim value As Double
        If latest.TryGetValue(topic, value) Then Return ExcelFormulaValue.FromNumber(value)
        Return ExcelFormulaValue.FromError("#N/A")
    End Function)

Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly) With {
    .DisableCalculationEngine = False,
    .FormulaSession = session
}

Using workbook As New OpenXmlExcelDataOperations(
    "sensors.xlsx", ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
    workbook.RecalculateAll()
    latest(("Hall1", "Temperature")) = 22.0
    workbook.RecalculateAll()
End Using
```

The callback receives the `ProgID`, server argument and all topics in formula order. The same code works on Windows, Linux and macOS. For a live source, replace the sample assignments with updates from the application. The registration is evaluated on automatic calculation passes. An update to the application's data source needs to request recalculation when the workbook must reflect the change immediately.

Open XML checks `RegisterRtdProvider` for the specific `ProgID` first, the existing session-wide `RegisterFunction("RTD", …)` second, and the allowed native Windows server third. If an existing session-wide function should yield to native providers, migrate it to one registration per `ProgID`. Use `RemoveRtdProvider` to remove a specific application registration.

## Functions that require other engines or data sources

The [function catalog](../tools/excel-function-catalog.csv) tracks evaluator support. The examples below show how to handle formulas that fall outside the current Open XML evaluator. When a workbook already contains the Excel result, a cached-only read or export can use it. A missing cached result requires a calculation engine or an application-supplied result. Open XML reports an unsupported fresh calculation through `FormulaComplexityNotSupportedException`. The EPPlus 4 formula compiler can return the worksheet error `#NAME?` for an unknown function.

When Open XML needs to calculate one of the four `FORECAST.ETS` functions, `TRANSLATE`, `DETECTLANGUAGE`, `STOCKHISTORY`, `CALL`, or `REGISTER.ID`, the exception names the function and states that it is not yet implemented. A registered .NET implementation can still handle that function. See [the implementation follow-up](https://github.com/CompuMasterGmbH/CompuMaster.Excel/issues/66).

| Formula or feature | Current behavior and application alternatives |
| --- | --- |
| `=TRANSLATE("Hola","es","en")` and `=DETECTLANGUAGE("Hola")` | Fresh calculation without a registered implementation raises `FormulaComplexityNotSupportedException`. Register a scalar .NET function that uses a translation or language detection provider, or calculate in Excel. A different provider can return different translations or language codes. [TRANSLATE](https://support.microsoft.com/en-us/excel/functions/translate-function) · [DETECTLANGUAGE](https://support.microsoft.com/en-us/excel/functions/detectlanguage-function) |
| `=STOCKHISTORY("MSFT",DATE(2025,1,2),DATE(2025,1,10))` | Use Excel to populate the historical series. Excel obtains the data from LSEG, whose direct historical-pricing API requires authenticated access. An application with authorized market data can write dates and prices into worksheet cells. A registered scalar function can return one value, while this Excel formula produces a spill range. [STOCKHISTORY](https://support.microsoft.com/en-gb/excel/functions/stockhistory-function) · [Excel's data provider](https://support.microsoft.com/en-us/excel/about-the-stocks-financial-data-sources) · [LSEG API access](https://developers.lseg.com/en/api-catalog/refinitiv-data-platform/refinitiv-data-platform-apis/quick-start/quick-start-guide-for-refinitiv-data-platform) |
| `=CUBEVALUE("ThisWorkbookDataModel","[Measures].[TotalSales]")` and `=CUBESETCOUNT(A1)` | Use Excel or a compatible OLAP engine to evaluate the cube or workbook data model. When `A1` contains a `CUBESET` formula, the count uses its evaluated set. The worksheet caption alone cannot determine its size. For a simple report, the application can provide a table of data or precomputed aggregates and use ordinary worksheet formulas over that table. [CUBEVALUE](https://support.microsoft.com/en-us/excel/functions/cubevalue-function) · [CUBESETCOUNT](https://support.microsoft.com/en-us/excel/functions/cubesetcount-function) |
| `=IMAGE("https://example.org/logo.png","Logo")` | Fresh evaluation of this image formula is outside the current scope. For a static logo, prepare the template with an embedded worksheet picture. See the [engine guide](README.md) for supported drawing preservation. [IMAGE](https://support.microsoft.com/en-us/excel/functions/image-function) |
| `=WEBSERVICE("https://example.org/data")` | Use application code to obtain the response and write it to a worksheet cell. A registered scalar .NET function can also supply the response text. Excel's own implementation uses Windows components. [WEBSERVICE](https://support.microsoft.com/en-us/excel/functions/webservice-function) |
| `=FORECAST.ETS(A13,B1:B12,A1:A12)`, `FORECAST.ETS.CONFINT`, `FORECAST.ETS.SEASONALITY`, and `FORECAST.ETS.STAT` | Use Excel for fresh calculation. The formulas operate on the supplied series. Matching Excel's smoothing, seasonality selection, and statistics is a calculation implementation task. [Excel forecasting functions](https://support.microsoft.com/en-us/excel/functions/forecasting-functions-reference) |

The [RTD section](#real-time-data-rtd) covers the implemented Windows COM and application-provider paths. It includes setup, update behavior, and alternatives for remote servers and other platforms.
