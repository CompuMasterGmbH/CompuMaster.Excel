# EPPlus

## Special CompuMaster Edition

This folder contains the maintained `CompuMaster.EPPlus4` fork of EPPlus 4.5.3.3. See the [package README](EPPlus/README.md) for the CompuMaster feature list and notes, which are also included in the NuGet package.

### Added Features

* Resets internal calculation caches so Microsoft Excel recalculates dependent formulas when reopening a workbook.
* Fixes internal races in cell storage, style updates, and calculation/lifetime coordination; general multithreaded workbook access remains unsupported.
* Uses zero-based worksheet indexing by default on every target framework; the existing compatibility setting can switch to one-based indexing.

### Security Features

* Uses `System.IO.Compression.ZipArchive` for XLSX packaging, with compatibility handling for older encrypted workbooks.
* Limits ZIP/XML expansion and encrypted-input work when loading XLSX files; callers can adjust the defaults for trusted files.

These limits do not make arbitrary untrusted workbooks safe to parse. See the [security guidance](../CM.Data.EpplusFixCalcsEdition/README.md#security-and-untrusted-workbooks).

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

## Upstream EPPlus Archive

**This repository has moved to https://github.com/EPPlusSoftware/EPPlus.** 

**The code in this archive represents the final version of EPPlus under LGPL. There will be no more activity here.**  

EPPlus will from version 5 switch license from **LGPL** to [Polyform Noncommercial 1.0.0]( https://polyformproject.org/licenses/noncommercial/1.0.0/) license.  
With the new license EPPlus is still free to use in some cases, but will require a commercial license to be used in a commercial business.

More information on the license change on [our website]( https://www.epplussoftware.com)
***
Create advanced Excel spreadsheets using .NET, without the need of interop.

EPPlus is a .NET library that reads and writes Excel files using the Office Open XML format (xlsx). 
EPPlus has no dependencies other than .NET.
 
## EPPlus supports:
* Cell Ranges 
* Cell styling (Border, Color, Fill, Font, Number, Alignments) 
* Data validation 
* Conditional formatting 
* Charts 
* Pictures 
* Shapes 
* Comments 
* Tables 
* Pivot tables 
* Protection 
* Encryption 
* VBA 
* Formula calculation 
* Many more... 

## Overview
This project started with the source from ExcelPackage. It was a great project to start from.
It had the basic functionality needed to read and write a spreadsheet.
Advantages over other:
EPPlus uses dictionaries to access cell data, making performance a lot better.
Complete integration with .NET 

## Support
All support is currently referred to [Stack overflow](https://stackoverflow.com/questions/tagged/epplus). 
A tutorial is available in the wiki and the sample project can be downloaded with each version. 
The old site at [Codeplex](http://epplus.codeplex.com) also contains material that can be helpful. 
Bugs and new feature requests can be added to the issues tracker. 

## License
The project is licensed under the GNU Library General Public License (LGPL). 
