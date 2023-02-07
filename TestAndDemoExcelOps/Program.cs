using CompuMaster.Excel.ExcelOps;

ExcelDataOperationsBase workbook;
string FirstSheetName;
TextTable formulasOrValues;
TextTable values;

//Create a workbook and put some values and formulas
workbook = new EpplusFreeExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, true, null);
System.Console.WriteLine("Engine=" + workbook.EngineName);
FirstSheetName = workbook.SheetNames()[0];
workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);

//Output table with formulas or alternatively with formatted cell value
formulasOrValues = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText);
System.Console.WriteLine(formulasOrValues.ToUIExcelTable());

//Output table with calculated or static values
values = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues);
System.Console.WriteLine(values.ToUIExcelTable());



//Assign required license context for Epplus component
EpplusPolyformExcelDataOperations.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

//Create a workbook and put some values and formulas
workbook = new EpplusPolyformExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, true, null);
System.Console.WriteLine("Engine=" + workbook.EngineName);
FirstSheetName = workbook.SheetNames()[0];
workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);

//Output table with formulas or alternatively with formatted cell value
formulasOrValues = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText);
System.Console.WriteLine(formulasOrValues.ToUIExcelTable());

//Output table with calculated or static values
values = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues);
System.Console.WriteLine(values.ToUIExcelTable());



/*
//Create a workbook and put some values and formulas
workbook = new FreeSpireXlsDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, true, null);
System.Console.WriteLine("Engine=" + workbook.EngineName);
FirstSheetName = workbook.SheetNames()[0];
workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);

//Output table with formulas or alternatively with formatted cell value
formulasOrValues = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText);
System.Console.WriteLine(formulasOrValues.ToUIExcelTable());

//Output table with calculated or static values
values = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues);
System.Console.WriteLine(values.ToUIExcelTable());
*/



//Create a workbook and put some values and formulas
workbook = new MsExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, true, null);
System.Console.WriteLine("Engine=" + workbook.EngineName);
FirstSheetName = workbook.SheetNames()[0];
workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);

//Output table with formulas or alternatively with formatted cell value
formulasOrValues = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText);
System.Console.WriteLine(formulasOrValues.ToUIExcelTable());

//Output table with calculated or static values
values = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues);
System.Console.WriteLine(values.ToUIExcelTable());
