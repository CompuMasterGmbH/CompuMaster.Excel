using CompuMaster.Excel.ExcelOps;

public class Samples
{
    public static void DemoAddColumnsAndCheckForReferenceUpdates()
    {
        ExcelDataOperationsBase workbook;
        string FirstSheetName;

        if (false)
        {
            //Create a workbook and put some values and formulas
            workbook = new MsExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile));
            workbook.CalculationModuleDisabled = true;
            System.Console.WriteLine("Engine=" + workbook.EngineName);
            FirstSheetName = workbook.SheetNames()[0];
            workbook.ClearSheet(FirstSheetName);
            workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
            workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
            workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);
            workbook.WriteCellFormula(FirstSheetName, 0, 3, @"SUM(A1:B1)", false);
            workbook.WriteCellFormula(FirstSheetName, 0, 4, @"SUM(A2:D2)", false);
            workbook.AddColumn(FirstSheetName, 3, 3, true);
            System.IO.FileInfo file = new System.IO.FileInfo("Demo_AddCols_ModCalcEnabled_" + workbook.EngineName + ".xlsx");
            System.Console.WriteLine("Saving to " + file.FullName);
            workbook.SaveAs(file.FullName, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset);
            workbook.CloseExcelAppInstance();
        }

        if (false)
        {
            //Create a workbook and put some values and formulas
            workbook = new EpplusFreeExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile));
            workbook.CalculationModuleDisabled = true;
            System.Console.WriteLine("Engine=" + workbook.EngineName);
            FirstSheetName = workbook.SheetNames()[0];
            workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
            workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
            workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);
            workbook.WriteCellFormula(FirstSheetName, 0, 3, @"SUM(A1:B1)", false);
            workbook.WriteCellFormula(FirstSheetName, 0, 4, @"SUM(A2:D2)", false);
            workbook.AddColumn(FirstSheetName, 3, 3, true);
            System.IO.FileInfo file = new System.IO.FileInfo("Demo_AddCols_ModCalcDisabled_" + workbook.EngineName + ".xlsx");
            System.Console.WriteLine("Saving to " + file.FullName);
            workbook.SaveAs(file.FullName, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation);
        }

        if (false)
        {
            //Create a workbook and put some values and formulas
            workbook = new EpplusFreeExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile));
            workbook.CalculationModuleDisabled = true;
            System.Console.WriteLine("Engine=" + workbook.EngineName);
            FirstSheetName = workbook.SheetNames()[0];
            workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
            workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
            workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);
            workbook.WriteCellFormula(FirstSheetName, 0, 3, @"SUM(A1:B1)", false);
            workbook.WriteCellFormula(FirstSheetName, 0, 4, @"SUM(A2:D2)", false);
            workbook.AddColumn(FirstSheetName, 3, 3, false);
            System.IO.FileInfo file = new System.IO.FileInfo("Demo_AddCols_ModCalcDisabled_NoRefUpdates_" + workbook.EngineName + ".xlsx");
            System.Console.WriteLine("Saving to " + file.FullName);
            workbook.SaveAs(file.FullName, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation);
        }

        if (false)
        {
            //Create a workbook and put some values and formulas
            workbook = new EpplusFreeExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile));
            workbook.CalculationModuleDisabled = false;
            System.Console.WriteLine("Engine=" + workbook.EngineName);
            FirstSheetName = workbook.SheetNames()[0];
            workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
            workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
            workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);
            workbook.WriteCellFormula(FirstSheetName, 0, 3, @"SUM(A1:B1)", false);
            workbook.WriteCellFormula(FirstSheetName, 0, 4, @"SUM(A2:D2)", false);
            workbook.AddColumn(FirstSheetName, 3, 3, true);
            System.IO.FileInfo file = new System.IO.FileInfo("Demo_AddCols_ModCalcEnabled_" + workbook.EngineName + ".xlsx");
            System.Console.WriteLine("Saving to " + file.FullName);
            workbook.SaveAs(file.FullName, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation);
        }

        if (false)
        {
            //Create a workbook and put some values and formulas
            workbook = new EpplusFreeExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile));
            workbook.CalculationModuleDisabled = false;
            System.Console.WriteLine("Engine=" + workbook.EngineName);
            FirstSheetName = workbook.SheetNames()[0];
            workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
            workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
            workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);
            workbook.WriteCellFormula(FirstSheetName, 0, 3, @"SUM(A1:B1)", false);
            workbook.WriteCellFormula(FirstSheetName, 0, 4, @"SUM(A2:D2)", false);
            workbook.AddColumn(FirstSheetName, 3, 3, false);
            System.IO.FileInfo file = new System.IO.FileInfo("Demo_AddCols_ModCalcEnabled_NoRefUpdates_" + workbook.EngineName + ".xlsx");
            System.Console.WriteLine("Saving to " + file.FullName);
            workbook.SaveAs(file.FullName, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation);
        }
    }

    public static void SeveralDemos()
    {
        ExcelDataOperationsBase workbook;
        string FirstSheetName;
        TextTable formulasOrValues;
        TextTable values;

        //Create a workbook and put some values and formulas
        workbook = new EpplusFreeExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile));
        System.Console.WriteLine("Engine=" + workbook.EngineName);
        FirstSheetName = workbook.SheetNames()[0];
        workbook.WriteCellValue<int>(FirstSheetName, 0, 0, 123);
        workbook.WriteCellValue<double>(new ExcelCell(FirstSheetName, "B1", ExcelCell.ValueTypes.All), 456.123);
        workbook.WriteCellFormula(FirstSheetName, 0, 2, @"SUM(A1:B1)", true);
        workbook.WriteCellFormula(FirstSheetName, 0, 3, @"SUM(A1:E1)", false);
        workbook.AddColumn(FirstSheetName, 0, 3, true);


        //Output table with formulas or alternatively with formatted cell value
        formulasOrValues = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText);
        System.Console.WriteLine(formulasOrValues.ToUIExcelTable());

        //Output table with calculated or static values
        values = workbook.SheetContentMatrix(FirstSheetName, ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues);
        System.Console.WriteLine(values.ToUIExcelTable());



        //Assign required license context for Epplus component
        EpplusPolyformExcelDataOperations.LicenseContext = new EpplusPolyformExcelDataOperations.EpplusLicenseActivator(OfficeOpenXml.EPPlusLicenseType.NonCommercialPersonal, "Unit Testing");

        //Create a workbook and put some values and formulas
        workbook = new EpplusPolyformExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile));
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
        workbook = new FreeSpireXlsDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions());
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
        workbook = new MsExcelDataOperations(null, ExcelDataOperationsBase.OpenMode.CreateFile, new ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile));
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

    }

}
