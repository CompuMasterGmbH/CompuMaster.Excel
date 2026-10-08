# CompuMaster.EPPlus 4.5.3.3

This package is based on upstream EPPlus 4.5.3.3; CompuMaster NuGet releases use date-based package versions.

## Special CompuMaster Edition

### Added Features

* Resets internal calculation caches so Microsoft Excel recalculates dependent formulas when reopening a workbook.
* Exposes embedded workbook theme XML so consumers can resolve design-dependent accent colors without a fixed Office palette.
* Fixes internal races in cell storage, style updates, and calculation/lifetime coordination; general multithreaded workbook access remains unsupported.
* Uses zero-based worksheet indexing by default on every target framework; the existing compatibility setting can switch to one-based indexing.

### Security Features

* Uses `System.IO.Compression.ZipArchive` for XLSX packaging, with compatibility handling for older encrypted workbooks.
* Limits ZIP/XML expansion and encrypted-input work when loading XLSX files; callers can adjust the defaults for trusted files.

These limits do not make arbitrary untrusted workbooks safe to parse. See the [security guidance](https://github.com/CompuMasterGmbH/CompuMaster.Excel/blob/main/CM.Data.EpplusFixCalcsEdition/README.md#security-and-untrusted-workbooks).

## Known Issues

**Warning: Epplus4 (including its fork CompuMaster.EPPlus4) is not thread-safe. Multithreaded use of the library is not supported. Concurrent access can cause exceptions and incorrect or lost cell values without an exception.**

Use each `ExcelPackage` in one thread throughout its lifetime, including access to its workbook, worksheets and ranges, formula calculation, saving, and disposal. This follows [Jan Källman's recommendation to avoid multiple threads on a single workbook](https://github.com/EPPlusSoftware/EPPlus/issues/894#issuecomment-1578082491). Different worksheets in one package are not independent concurrency boundaries. If a package must be shared, serialize all access with one stable external lock per package/workbook, including complete select/read/modify/write sequences; locks per worksheet or cell are insufficient.

The CompuMaster edition fixes the reproduced internal races in cell-store growth and access, workbook-wide style updates, and concurrent formula calculations. Calculation, parser-manager operations, saving, and disposal now share a package-local monitor. Regression tests cover these cases and independently owned packages. These internal fixes do not establish general thread safety or replace caller coordination. See [issue #26](https://github.com/CompuMasterGmbH/CompuMaster.Excel/issues/26) for the original failures, implementation scope, and remaining investigation items.

The following known limitations remain:

1. **A shared, cached `ExcelRange` can target the wrong cell.** Its indexers change the range object's address and return the same object. Another thread can change that address between selection and use. This behavior is retained for compatibility; do not share cached range objects between threads.
2. **Compound operations and arbitrary concurrent workbook access remain unsupported.** Internal cell-store locks do not make selection followed by reading/writing atomic, provide snapshot cell enumeration, or coordinate every cell/structure change with calculation, saving, loading, or disposal. Concurrent edits can still invalidate a traversal or produce inconsistent output. Coordinate complete operations externally, including reads; direct XML/collection mutation and copying styles between workbooks also require caller coordination.
3. **Callbacks must not reenter workbook services during calculation.** Custom functions and loggers run while the package monitor is held. Recursively calculating/parsing, saving, or disposing the same package can interfere with active parser state or workbook lifetime. Waiting for another thread that needs the same package can deadlock. The monitor permits internal nested calls but does not make such callback behavior supported.
4. **Some stream copies are globally serialized.** The inherited `CopyStream` lock can limit throughput even for independent packages. It does not make sharing a stream safe; each package must own its stream exclusively.

Use separately owned packages and streams for parallel jobs, with one thread accessing each package at a time. The independent-package regression tests cover selected operations, not every feature. Other lazy initialization paths, the shared `RAND()` seed, and publicly mutable operator/validation objects still require investigation; these code observations are not all demonstrated runtime defects.

## Announcement: This is the last version of EPPlus under the LGPL License
EPPlus will from version 5 be licensed under the [Polyform Noncommercial 1.0.0]( https://polyformproject.org/licenses/noncommercial/1.0.0/) license.  
With the new license EPPlus is still free to use in some cases, but will require a commercial license to be used in a commercial business.  
More information on the license change on [our website]( https://www.epplussoftware.com/Home/LgplToPolyform)

## New features in version 4.5:
* .NET Core support
* Sparklines
* Sort method added to ExcelRange
* Bug fixes and minor changes, see below and visit https://github.com/JanKallman/EPPlus for tutorials, samples and the latest information

## Important Notes:
The CompuMaster edition initializes the Worksheets collection as zero-based on every target framework.
This can be altered programmatically by setting ExcelPackage.Compatibility.IsWorksheets1Based to true after constructing the package.

.NET Core uses a preview of System.Drawing.Common, so be aware of that. We will update it as Microsoft releases newer versions.
System.Drawing.Common requires libgdiplus to be installed on non-Windows operating systems.

## Use your favorite package manager to install it.
For example:

### Homebrew on MacOS:
brew install mono-libgdiplus

### apt-get:
apt-get install libgdiplus

## EPPlus-A .NET Spreadsheet API

Changes
4.5.3.3
* Support for .NET Standard 2.1.

4.5.3.2
* Added a target build for .NET Core 2.1 (netcoreapp2.1) with System.Drawing.Common 4.6.0-preview6.19303.8 
* Fixed Text property with short date format
* Fixed problem with defined names containing backslash 
* More bugfixes, see https://github.com/JanKallman/EPPlus/commits/master

4.5.3.1
* Fixed Lookup function ignoring result vector.
* Fixed address validation.

4.5.3
* Upgraded System.Drawing.Common for .NET Core to 4.5.1
* Enabled worksheetcharts to use a pivottable as source by adding a pivotTableSource parameter to the AddChart method of the Worksheets collection
* Pmt function
* And lots of bugfixes, see https://github.com/JanKallman/EPPlus/commits/master
      
4.5.2.1
* Upgraded System.Drawing.Common for .NET Core to 4.5.0
* Fixed problem with Apostrophe in worksheet name

4.5.2
* Upgraded System.Drawing.Common to 4.5.0-rc1
* Optimized image handling
* External Streams are not closed when disposing the package
* Fixed issue with Floor and Celing functions
* And more bugfixes, see https://github.com/JanKallman/EPPlus/commits/master

4.5.1
* Added web sample for .NET Core from Vahid Nasiri
* Added sample sparkline sample to sample project
* Fixed a few problems related to .NET Core on Mac

4.5.0.3
* Fix for compound documents (VBA and Encryption).
* Fix for Excel 2010 sha1 hashed agile encryption.
* Upgraded System.Drawing.Common to 4.5.0-preview1-26216-02
* Also see https://github.com/JanKallman/EPPlus/commits/master

4.5.0.2 rc
* Merge in e few pull requests and fixed a few issues. See https://github.com/JanKallman/EPPlus/commits/master


4.5.0.1 Beta 2
* Added sparkline support.
* Switched targetframework from netcoreapp2.0 to netstandardapp2.0
* Replaced CoreCompat.System.Drawing.v2 with System.Drawing.Common
* Fixed a few issues. See https://github.com/JanKallman/EPPlus/commits/master

4.5.0.0 Beta 1
* .Net Core support.
* Added ExcelPackage.Compatibility.IsWorksheets1Based to remove inconsistent 1 base collection on the worksheets collection.
Note: .Net Core will have this property set to false, and .Net 3.5 and .Net 4 version will have this property set to true for backward compatibility reasons.
This property can be set via the appsettings.json file in .Net Core or the app.config file. See sample project for examples.
* RoundedCorners property Add to ExcelChart
* DataTable propery Added  to ExcelPlotArea for charts
* Sort method added on ExcelRange
* Added functions NETWORKDAYS.INTL and NETWORKDAYS.
* And a lot of bug fixes. See https://github.com/JanKallman/EPPlus/commits/master

4.1.1
* Fix VBA bug in Excel 2016 - 1708 and later

4.1
* Added functions Rank, Rank.eq, Rank.avg and Search
* Applied a whole bunch of pull requests...
* Performance and memory usage tweeks
* Ability to set and retrieve 'custom' extended application propeties.
* Added style QuotePrefix
* Added support for MajorTimeUnit and MinorTimeUnit to chart axes
* Added GapWidth Property to BarChart and Gapwidth.
* Added Fill and Border properties to ChartSerie.
* Added support for MajorTimeUnit and MinorTimeUnit to chart axes
* Insert/delete row/column now shifts named ranges, comments, tables and pivottables.
* And fixed a lot of issues. See http://epplus.codeplex.com/SourceControl/list/changesets for more details

4.0.5 Fixes
* Switched to Visual Studio 2015 for code and sample projects.
* Added LineColor, MarkerSize, LineWidth and MarkerLineColor properties to line charts
* Added LineEnd properties to shapes
* Added functions Value, DateValue, TimeValue
* Removed WPF depedency.
* And fixed a lot of issues. See http://epplus.codeplex.com/SourceControl/list/changesets for more details

4.0.4 Fixes
* Added functions Daverage, Dvar Dvarp, DMax, DMin DSum,  DGet, DCount and DCountA 
* Exposed the formula parser logging functionality via FormulaParserManager.
* And fixed a lot of issues. See http://epplus.codeplex.com/SourceControl/list/changesets for more details

4.0.3 Fixes
* Added compilation directive for MONO (Thanks Danny)
* Added functions IfError, Char, Error.Type, Degrees, Fixed, IsNonText, IfNa and SumIfs
* And fixed a lot of issues. See http://epplus.codeplex.com/SourceControl/list/changesets for more details

4.0.2 Fixes
* Fixes a whole bunch of bugs related to the cell store (Worksheet.InsertColumn, Worksheet.InsertRow, Worksheet.DeleteColumn, Worksheet.DeleteRow, Range.Copy, Range.Clear)
* Added functions Acos, Acosh, Asinh, Atanh, Atan, CountBlank, CountIfs, Mina, Offset, Median, Hyperlink, Rept
* Fix for reading Excel comment content from the t-element.
* Fix to make Range.LoadFromCollection work better with inheritence
* And alot of other small fixes

4.0.1 Fixes
* VBA unreadable content
* Fixed a few issues with InsertRow and DeleteRow
* Fixed bug in Average and AverageA 
* Handling of Div/0 in functions
* Fixed VBA CodeModule error when copying a worksheet.
* Value decoding when reading str element for cell value.
* Better exception when accessing a worksheet out of range in the Excelworksheets indexer.
* Added Small and Large function to formula parser. Performance fix when encountering an unknown function.
* Fixed handling strings in formulas
* Calculate hangs if formula start with a parenthes.
* Worksheet.Dimension returned an invalid range in some cases.
* Rowheight was wrong in some cases.
* ExcelSeries.Header had an incorrect validation check.

New features 4.0

Replaced Packaging API with DotNetZip
* This will remove any problems with Isolated Storage and enable multithreading
* This historical upstream note does not establish thread safety; see Known Issues above.


New Cell store
* Less memory consumption
* Insert columns (not on the range level)
* Faster row inserts,

Formula Parser
* Calculates all formulas in a workbook, a worksheet or in a specified range
* 100+ functions implemented
* Access via Calculate methods on Workbook, Worksheet and Range objects.
* Add custom/missing Excel functions via Workbook. FormulaParserManager.
* Samples added to the EPPlusSamples project.

The formula parser does not support Array Formulas
* Intersect operator (Space)
* References to external workbooks
* And probably a whole lot of other stuff as well :)

Performance
*Of course the performance of the formula parser is nowhere near Excels. Our focus has been functionality.

Agile Encryption (Office 2012-)
* Support for newer type of encryption.

Minor new features
* Chart worksheets
* New Chart Types Bubblecharts
* Radar Charts
* Area Charts
* And lots of bug fixes...

Beta 2 Changes
* Fixed bug when using RepeatColumns & RepeatRows at the same time.
* VBA project will be left untouched if it�s not accessed.
* Fixed problem with strings on save.
* Added locks to the cell store for access by multiple threads.
* Implemented Indirect function
* Used DisplayNameAttribute to generate column headers from LoadFromCollection
* Rewrote ExcelRangeBase.Copy function. 
* Added caching to Save ZipStream for Cells and shared strings to speed up the Save method.
* Added Missing InsertColumn and DeleteColumn
* Added pull request to support Date1904 
* Added pull request ExcelWorksheet. LoadFromDataReader

Release Candidate changes
* Fixed some problems with Range.Copy Function
* InsertColumn and Delete column didn't work in some cases
* Chart.DisplayBlankAs had the wrong default type in Excel 2010+
* Datavalidation list overflow caused corruption of the package
* Fixed a few Calculation when referring ranges (for example If function)
* Added ChartAxis.DisplayUnit
* Fixed a bug related to shared formulas
* Named styles failed in some cases.
* Style.Indent got an invalid value in some cases.
* Fixed a problem with AutofitColumns method.
* Performance fix.
* A whole lot of other small fixes.
