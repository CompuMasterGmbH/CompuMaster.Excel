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

## Application-defined worksheet functions

A workbook can use a .NET implementation for a cell formula such as `=REBATE(B2)`. Port the business logic from the original VBA worksheet function and register it under the formula name.

```vbnet
Dim session As New ExcelFormulaSession()
session.RegisterFunction("REBATE",
    Function(context, arguments)
        If arguments.Count <> 1 OrElse arguments(0).Kind <> ExcelFormulaValueKind.Number Then
            Return ExcelFormulaValue.FromError("#VALUE!")
        End If
        Return ExcelFormulaValue.FromNumber(CDbl(arguments(0).Value) * 0.1)
    End Function)
```

Pass this session in `ExcelDataOperationsOptions.FormulaSession`. The context identifies the formula cell and provides read-only access to workbooks in the session. The XLSX stores the formula and its result. Supply the registration again when reopening it for calculation. For a VBA macro driven by a button or by an event such as `Workbook_Open`, port the operation to application code and connect it to the equivalent application action. [Excel worksheet functions written in VBA](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel)

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
