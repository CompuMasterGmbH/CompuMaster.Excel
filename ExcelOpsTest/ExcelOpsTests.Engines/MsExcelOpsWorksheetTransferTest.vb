Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework
Imports NativeExcel = Microsoft.Office.Interop.Excel

Namespace ExcelOpsTests.Engines

    Partial Public Class MsExcelOpsTest

        Private Const TransferSourceFile As String = "Source model.xlsx"
        Private Const TransferTargetFile As String = "Target model.xlsx"

        ' A source-only dependency remains external even when the destination has a sheet with that name.
        <TestCase(False, False)>
        <TestCase(False, True)>
        <TestCase(True, False)>
        <TestCase(True, True)>
        Public Sub WorksheetTransfer_SavedWorkbookReferencesAndNameCollisions(move As Boolean, collision As Boolean)
            Dim caseDirectory = CreateTransferTestDirectory()
            Dim inputs = Directory.CreateDirectory(Path.Combine(caseDirectory, "Inputs")).FullName
            Dim sourcePath = Path.Combine(inputs, TransferSourceFile)
            Dim targetPath = Path.Combine(caseDirectory, TransferTargetFile)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            Try
                source = CreateTransferSource()
                target = CreateTransferTarget(collision)
                Dim targetCheck = TransferSheet(target.Workbook, "Target Check")
                targetCheck.Range("A1").Formula = "=" & OpenWorkbookQualifier(source.Workbook, "Transfer") & "$A$1"
                source.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                target.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)

                TransferWorksheet(TransferSheet(source.Workbook, "Transfer"), targetCheck, move)
                Dim receivedName = If(collision, "Transfer (2)", "Transfer")
                Dim received = TransferSheet(target.Workbook, receivedName)
                Dim remaining = TransferSheet(source.Workbook, "Remain")
                Assert.That(source.Workbook.Worksheets.Count, [Is].EqualTo(If(move, 2, 3)))
                AssertTransferFormula(received, "A2", "=A1+" & receivedName & "!$A$1", 20)
                AssertTransferFormula(received, "A3", "=" & OpenWorkbookQualifier(source.Workbook, "Data Set") & "$A$1", 20)
                AssertTransferFormula(received, "A4", "=SUM(" & OpenWorkbookQualifier(source.Workbook, "Data Set") & "$A$1:$A$2)", 50)
                AssertTransferFormula(remaining, "A1", "=" & If(move, OpenWorkbookQualifier(target.Workbook, receivedName), "Transfer!") & "$A$1", 10)
                AssertTransferFormula(targetCheck, "A1", "=" & If(move, receivedName & "!", OpenWorkbookQualifier(source.Workbook, "Transfer")) & "$A$1", 10)
                If collision Then AssertTransferFormula(targetCheck, "A2", "=Transfer!$A$1", 900)
                TestContext.Out.WriteLine("Excel " & source.Workbook.Application.Version & ", build " & source.Workbook.Application.Build)

                source.Save()
                target.Save()
                target.Close()
                source.Close()

                Assert.That(StoredWorkbookLinkTargets(targetPath), [Is].EqualTo({"Inputs/" & TransferSourceFile}))
                If move Then
                    Dim sourceLink = StoredWorkbookLinkTargets(sourcePath).Single()
                    Assert.That(New Uri(New Uri(sourcePath), sourceLink).LocalPath, [Is].EqualTo(targetPath))
                Else
                    Assert.That(StoredWorkbookLinkTargets(sourcePath), [Is].Empty)
                End If
                Using document = SpreadsheetDocument.Open(targetPath, False)
                    Dim sheet = StructuralReferenceCoverage.WorksheetFor(document.WorkbookPart, receivedName)
                    Assert.That(StructuralReferenceCoverage.FormulaAt(sheet, "A3"), [Is].EqualTo("'[1]Data Set'!$A$1"))
                End Using
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        ' When the linked book gets a filename, the in-memory formula changes before the dependent file is written again.
        <TestCase(False, False)>
        <TestCase(False, True)>
        <TestCase(True, False)>
        <TestCase(True, True)>
        Public Sub WorksheetTransfer_FirstSaveUpdatesOpenLinks(move As Boolean, targetFirst As Boolean)
            Dim caseDirectory = CreateTransferTestDirectory()
            Dim sourcePath = Path.Combine(caseDirectory, TransferSourceFile)
            Dim targetPath = Path.Combine(caseDirectory, TransferTargetFile)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            Try
                source = CreateTransferSource()
                target = CreateTransferTarget(False)
                Dim sourceNameBeforeSave = source.Workbook.Name
                Dim targetNameBeforeSave = target.Workbook.Name
                Dim remaining = TransferSheet(source.Workbook, "Remain")
                Dim targetCheck = TransferSheet(target.Workbook, "Target Check")
                TransferWorksheet(TransferSheet(source.Workbook, "Transfer"), targetCheck, move)
                Dim received = TransferSheet(target.Workbook, "Transfer")
                AssertTransferFormula(received, "A3", "=" & OpenWorkbookQualifier(source.Workbook, "Data Set") & "$A$1", 20)
                AssertTransferFormula(remaining, "A1", "=" & If(move, OpenWorkbookQualifier(target.Workbook, "Transfer"), "Transfer!") & "$A$1", 10)

                If targetFirst Then
                    target.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                    Assert.That(StoredWorkbookLinkTargets(targetPath), [Is].EqualTo({sourceNameBeforeSave}))
                    source.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                    Assert.That(StoredWorkbookLinkTargets(targetPath), [Is].EqualTo({sourceNameBeforeSave}))
                Else
                    source.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                    If move Then Assert.That(StoredWorkbookLinkTargets(sourcePath), [Is].EqualTo({targetNameBeforeSave}))
                    target.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                    If move Then Assert.That(StoredWorkbookLinkTargets(sourcePath), [Is].EqualTo({targetNameBeforeSave}))
                End If

                AssertTransferFormula(received, "A3", "=[" & TransferSourceFile & "]Data Set!$A$1", 20)
                AssertTransferFormula(remaining, "A1", "=" & If(move, "[" & TransferTargetFile & "]Transfer!", "Transfer!") & "$A$1", 10)
                source.Save()
                target.Save()
                Assert.That(StoredWorkbookLinkTargets(targetPath), [Is].EqualTo({TransferSourceFile}))
                If move Then Assert.That(StoredWorkbookLinkTargets(sourcePath), [Is].EqualTo({TransferTargetFile}))
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        ' Group transfers bind both sheets together, even when both names collide in the destination.
        <TestCase(False, False)>
        <TestCase(False, True)>
        <TestCase(True, False)>
        <TestCase(True, True)>
        Public Sub WorksheetTransfer_GroupAndSequentialTransfers(move As Boolean, asGroup As Boolean)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            Try
                source = CreateTransferSource()
                target = CreateTransferTarget(True)
                Dim targetCheck = TransferSheet(target.Workbook, "Target Check")
                If asGroup Then
                    Dim names = {"Data Set", "Transfer"}
                    Dim assigned = If(move, source.MoveSheets(names, target), source.CopySheets(names, target))
                    Assert.That(assigned.Select(Function(item) item.TargetSheetName), [Is].EqualTo({"Transfer (2)", "Data Set (2)"}))
                Else
                    TransferWorksheet(TransferSheet(source.Workbook, "Transfer"), targetCheck, move)
                    Dim receivedFirst = TransferSheet(target.Workbook, "Transfer (2)")
                    AssertTransferFormula(receivedFirst, "A3", "=" & OpenWorkbookQualifier(source.Workbook, "Data Set") & "$A$1", 20)
                    TransferWorksheet(TransferSheet(source.Workbook, "Data Set"), receivedFirst, move)
                End If

                Dim received = TransferSheet(target.Workbook, "Transfer (2)")
                Dim receivedData = TransferSheet(target.Workbook, "Data Set (2)")
                Dim dataQualifier = If(asGroup OrElse move, "Data Set (2)!", OpenWorkbookQualifier(source.Workbook, "Data Set"))
                Dim transferQualifier = If(asGroup OrElse move, "Transfer (2)!", OpenWorkbookQualifier(source.Workbook, "Transfer"))
                AssertTransferFormula(received, "A3", "=" & dataQualifier & "$A$1", 20)
                AssertTransferFormula(receivedData, "A3", "=" & transferQualifier & "$A$1", 10)
                AssertTransferFormula(targetCheck, "A2", "=Transfer!$A$1", 900)
                Assert.That(source.Workbook.Worksheets.Count, [Is].EqualTo(If(move, 1, 3)))
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
            End Try
        End Sub

        <TestCase(False, 0)>
        <TestCase(False, 1)>
        <TestCase(True, 1)>
        <TestCase(True, 2)>
        Public Sub WorksheetTransfer_TransfersMixedWorksheetAndChartSheetGroups(move As Boolean, targetIndex As Integer)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            Try
                source = CreateTransferSource()
                target = CreateTransferTarget(False)
                Dim sourceChart = DirectCast(source.Workbook.Charts.Add(After:=TransferSheet(source.Workbook, "Remain")), NativeExcel.Chart)
                sourceChart.Name = "Chart View"
                Dim targetChart = DirectCast(target.Workbook.Charts.Add(After:=TransferSheet(target.Workbook, "Target Check")), NativeExcel.Chart)
                targetChart.Name = "Target Chart"
                Dim names = {"Chart View", "Transfer"}
                Dim expectedTabs = target.SheetNames()
                expectedTabs.InsertRange(targetIndex, {"Transfer", "Chart View"})
                Dim destinationCount = target.SheetNames().Count
                Assert.Throws(Of NotSupportedException)(
                    Sub() source.CopySheets(names, target, targetIndex, ExcelDataOperationsBase.CopyWorksheetFormulaMode.PreserveFormulas))
                Assert.That(target.SheetNames().Count, [Is].EqualTo(destinationCount))

                Dim assigned = If(move, source.MoveSheets(names, target, targetIndex), source.CopySheets(names, target, targetIndex))
                Assert.That(assigned.Select(Function(item) item.TargetSheetName), [Is].EqualTo({"Transfer", "Chart View"}))
                Assert.That(assigned.Select(Function(item) item.TargetSheetIndex), [Is].EqualTo({targetIndex, targetIndex + 1}))
                Assert.That(target.SheetNames(), [Is].EqualTo(expectedTabs))
                Assert.That(target.ChartSheetNames(), [Is].EquivalentTo({"Chart View", "Target Chart"}))
                Assert.That(source.SheetNames().Contains("Chart View"), [Is].EqualTo(Not move))
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
            End Try
        End Sub

        <TestCase(False)>
        <TestCase(True)>
        Public Sub WorksheetTransfer_ExcelCreatesAnUnsavedDestination(move As Boolean)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As NativeExcel.Workbook = Nothing
            Try
                source = CreateTransferSource()
                Dim transfer = TransferSheet(source.Workbook, "Transfer")
                If move Then
                    transfer.Move()
                Else
                    transfer.Copy()
                End If
                target = source.Workbook.Application.ActiveWorkbook
                Assert.That(target.Path, [Is].Empty)
                Assert.That(target.Worksheets.Count, [Is].EqualTo(1))
                AssertTransferFormula(TransferSheet(target, "Transfer"), "A3",
                                      "=" & OpenWorkbookQualifier(source.Workbook, "Data Set") & "$A$1", 20)
                AssertTransferFormula(TransferSheet(source.Workbook, "Remain"), "A1",
                                      "=" & If(move, OpenWorkbookQualifier(target, "Transfer"), "Transfer!") & "$A$1", 10)
            Finally
                If target IsNot Nothing Then target.Close(False)
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
            End Try
        End Sub

        <TestCase(ExcelDataOperationsBase.CopyWorksheetFormulaMode.ExcelCompatible, 20, 20)>
        <TestCase(ExcelDataOperationsBase.CopyWorksheetFormulaMode.PreserveFormulas, 910, 700)>
        Public Sub WorksheetTransfer_ExcelCalculatesBothOpenXmlCopyModes(mode As ExcelDataOperationsBase.CopyWorksheetFormulaMode,
                                                                        expectedSelf As Integer, expectedReference As Integer)
            Dim caseDirectory = CreateTransferTestDirectory()
            Dim sourcePath = Path.Combine(caseDirectory, TransferSourceFile)
            Dim targetPath = Path.Combine(caseDirectory, TransferTargetFile)
            Dim output = Path.Combine(caseDirectory, "Edited target.xlsx")
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            Try
                source = CreateTransferSource()
                target = CreateTransferTarget(True)
                source.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                target.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                source.Close()
                target.Close()
                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
                Using origin As New OpenXmlExcelDataOperations(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options),
                      destination As New OpenXmlExcelDataOperations(targetPath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    Assert.That(origin.CopySheet("Transfer", destination, mode).TargetSheetName, [Is].EqualTo("Transfer (2)"))
                    destination.SaveCopyAs(output, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using

                source = Me.CreateInstance(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target = Me.CreateInstance(output, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target.RecalculateAll()
                Dim copied = TransferSheet(target.Workbook, "Transfer (2)")
                Assert.That(CDbl(copied.Range("A2").Value2), [Is].EqualTo(expectedSelf))
                Assert.That(CDbl(copied.Range("A3").Value2), [Is].EqualTo(expectedReference))
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        <TestCase(False)>
        <TestCase(True)>
        Public Sub WorksheetTransfer_ExcelReopensOpenXmlLinksCreatedBetweenUnsavedBooks(targetFirst As Boolean)
            Dim caseDirectory = CreateTransferTestDirectory()
            Dim sourcePath = Path.Combine(caseDirectory, TransferSourceFile)
            Dim targetPath = Path.Combine(caseDirectory, TransferTargetFile)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            Try
                source = CreateTransferSource()
                target = CreateTransferTarget(False)
                source.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                target.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                source.Close()
                target.Close()

                Dim originTemplate = File.ReadAllBytes(sourcePath)
                Dim targetTemplate = File.ReadAllBytes(targetPath)
                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite)
                Using session As New OpenXmlWorkbookSession(),
                      origin As New OpenXmlExcelDataOperations(originTemplate, options),
                      destination As New OpenXmlExcelDataOperations(targetTemplate, options)
                    session.Attach(origin)
                    session.Attach(destination)
                    origin.MoveSheet("Transfer", destination)
                    If targetFirst Then
                        destination.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                        origin.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                    Else
                        origin.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                        destination.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                    End If
                    origin.Save()
                    destination.Save()
                End Using

                source = Me.CreateInstance(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target = Me.CreateInstance(targetPath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target.RecalculateAll()
                AssertTransferFormula(TransferSheet(source.Workbook, "Remain"), "A1",
                    "=" & OpenWorkbookQualifier(target.Workbook, "Transfer") & "$A$1", 10)
                AssertTransferFormula(TransferSheet(target.Workbook, "Transfer"), "A3",
                    "=" & OpenWorkbookQualifier(source.Workbook, "Data Set") & "$A$1", 20)
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        <Test>
        Public Sub WorksheetTransfer_OpenXmlMovesMixedGroupAndExcelRecalculatesBothBooks()
            Dim caseDirectory = CreateTransferTestDirectory()
            Dim sourcePath = Path.Combine(Directory.CreateDirectory(Path.Combine(caseDirectory, "Inputs")).FullName, TransferSourceFile)
            Dim targetPath = Path.Combine(caseDirectory, TransferTargetFile)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            Try
                source = CreateTransferSource()
                target = CreateTransferTarget(True)
                Dim movedChart = AddReferenceChartSheet(source.Workbook, TransferSheet(source.Workbook, "Remain"))
                movedChart.Name = "Move Chart"
                Dim retainedChart = AddReferenceChartSheet(source.Workbook, TransferSheet(source.Workbook, "Remain"))
                retainedChart.Name = "Stay Chart"
                Dim targetCheck = TransferSheet(target.Workbook, "Target Check")
                Dim targetChart = AddReferenceChartSheet(target.Workbook, targetCheck)
                targetChart.Name = "Target Chart"
                SetTransferChartSeries(movedChart, "Transfer!")
                SetTransferChartSeries(retainedChart, "Transfer!")
                SetTransferChartSeries(targetChart, OpenWorkbookQualifier(source.Workbook, "Transfer"))
                targetCheck.Range("A1").Formula = "=" & OpenWorkbookQualifier(source.Workbook, "Transfer") & "$A$1"
                TransferSheet(source.Workbook, "Transfer").Range("A5").Formula = "=" & OpenWorkbookQualifier(target.Workbook, "Transfer") & "$A$1"

                source.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                target.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                source.Save()
                target.Close()
                source.Close()

                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite)
                Using origin As New OpenXmlExcelDataOperations(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options),
                      destination As New OpenXmlExcelDataOperations(targetPath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    Dim moved = origin.MoveSheets({"Move Chart", "Transfer"}, destination, 1)
                    Assert.That(moved.Select(Function(item) item.TargetSheetName), [Is].EqualTo({"Transfer (2)", "Move Chart"}))
                    Assert.That(moved.Select(Function(item) item.TargetSheetIndex), [Is].EqualTo({1, 2}))
                    Assert.That(origin.SheetNames(), [Is].EqualTo({"Data Set", "Stay Chart", "Remain"}))
                    Assert.That(destination.LookupCellFormula("Target Check", 0, 0), [Is].EqualTo("'Transfer (2)'!$A$1"))
                    Assert.That(destination.LookupCellFormula("Transfer (2)", 4, 0), [Is].EqualTo("Transfer!$A$1"))
                    origin.Save()
                    destination.Save()
                End Using

                source = Me.CreateInstance(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target = Me.CreateInstance(targetPath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target.RecalculateAll()
                AssertTransferFormula(TransferSheet(source.Workbook, "Remain"), "A1",
                    "=" & OpenWorkbookQualifier(target.Workbook, "Transfer (2)") & "$A$1", 10)
                AssertTransferFormula(TransferSheet(target.Workbook, "Transfer (2)"), "A3",
                    "=" & OpenWorkbookQualifier(source.Workbook, "Data Set") & "$A$1", 20)
                AssertTransferFormula(TransferSheet(target.Workbook, "Transfer (2)"), "A5", "=Transfer!$A$1", 900)
                AssertTransferFormula(TransferSheet(target.Workbook, "Target Check"), "A1", "=Transfer (2)!$A$1", 10)
                AssertTransferFormula(TransferSheet(target.Workbook, "Target Check"), "A2", "=Transfer!$A$1", 900)
                AssertChartTransferSeries(DirectCast(source.Workbook.Charts.Item("Stay Chart"), NativeExcel.Chart),
                    OpenWorkbookQualifier(target.Workbook, "Transfer (2)"), 10)
                AssertChartTransferSeries(DirectCast(target.Workbook.Charts.Item("Move Chart"), NativeExcel.Chart), "'Transfer (2)'!", 10)
                AssertChartTransferSeries(DirectCast(target.Workbook.Charts.Item("Target Chart"), NativeExcel.Chart), "'Transfer (2)'!", 10)
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        <Test>
        Public Sub ChartReferences_TrackEditsInTheLinkedWorkbook()
            Dim caseDirectory = CreateTransferTestDirectory()
            Dim sourcePath = Path.Combine(caseDirectory, TransferSourceFile)
            Dim targetPath = Path.Combine(caseDirectory, TransferTargetFile)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            Try
                source = CreateTransferSource()
                target = CreateTransferTarget(False)
                Dim targetCheck = TransferSheet(target.Workbook, "Target Check")
                Dim embedded = AddEmbeddedReferenceChart(targetCheck).Chart
                Dim chartSheet = AddReferenceChartSheet(target.Workbook, targetCheck)
                Dim qualifier = OpenWorkbookQualifier(source.Workbook, "Transfer")
                targetCheck.Range("A1").Formula = "=" & qualifier & "$A$1"
                SetTransferChartSeries(embedded, qualifier)
                SetTransferChartSeries(chartSheet, qualifier)
                source.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                target.SaveAs(targetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                source.AddRow("Transfer", 0, 1, True)
                qualifier = OpenWorkbookQualifier(source.Workbook, "Transfer")
                AssertChartTransferSeries(embedded, qualifier, 10, "$A$2")
                AssertChartTransferSeries(chartSheet, qualifier, 10, "$A$2")
                source.Close()
                target.Close()

                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite)
                Using session As New OpenXmlWorkbookSession(),
                      origin As New OpenXmlExcelDataOperations(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options),
                      dependent As New OpenXmlExcelDataOperations(targetPath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    session.Attach(origin)
                    session.Attach(dependent)
                    origin.AddRow("Transfer", 0, 1, True)
                    origin.Save()
                    dependent.Save()
                End Using

                source = Me.CreateInstance(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target = Me.CreateInstance(targetPath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                qualifier = OpenWorkbookQualifier(source.Workbook, "Transfer")
                targetCheck = TransferSheet(target.Workbook, "Target Check")
                embedded = DirectCast(targetCheck.ChartObjects(1), NativeExcel.ChartObject).Chart
                chartSheet = DirectCast(target.Workbook.Charts.Item(ChartViewSheet), NativeExcel.Chart)
                AssertChartTransferSeries(embedded, qualifier, 10, "$A$2")
                AssertChartTransferSeries(chartSheet, qualifier, 10, "$A$2")
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
                Directory.Delete(caseDirectory, True)
            End Try
        End Sub

        Private Shared Sub SetTransferChartSeries(chart As NativeExcel.Chart, qualifier As String)
            chart.ChartType = NativeExcel.XlChartType.xlColumnClustered
            Dim collection = DirectCast(chart.SeriesCollection(), NativeExcel.SeriesCollection)
            While collection.Count > 0
                collection.Item(1).Delete()
            End While
            Dim series = collection.NewSeries()
            series.Formula = "=SERIES(,," & qualifier & "$A$1,1)"
        End Sub

        Private Shared Sub AssertChartTransferSeries(chart As NativeExcel.Chart, qualifier As String, expected As Integer,
                                                       Optional address As String = "$A$1")
            chart.Refresh()
            Dim series = DirectCast(chart.SeriesCollection(1), NativeExcel.Series)
            Assert.That(series.Formula.Replace("'", ""), Does.Contain(qualifier.Replace("'", "") & address))
            Dim points = DirectCast(series.Values, Array)
            Assert.That(CDbl(points.GetValue(points.GetLowerBound(0))), [Is].EqualTo(expected))
        End Sub

        Private Function CreateTransferSource() As MsExcelDataOperations
            AlwaysCloseAllWorkbooksInNewEngineInstances = False
            Dim source = Me.CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile,
                                          New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile))
            Dim transfer = TransferSheet(source.Workbook, "Sheet1")
            transfer.Name = "Transfer"
            source.AddSheet("Data Set", Nothing)
            source.AddSheet("Remain", Nothing)
            Dim data = TransferSheet(source.Workbook, "Data Set")
            transfer.Range("A1").Value2 = 10
            transfer.Range("A2").Formula = "=A1+Transfer!$A$1"
            transfer.Range("A3").Formula = "='Data Set'!$A$1"
            transfer.Range("A4").Formula = "=SUM('Data Set'!$A$1:$A$2)"
            data.Range("A1").Value2 = 20
            data.Range("A2").Value2 = 30
            data.Range("A3").Formula = "=Transfer!$A$1"
            TransferSheet(source.Workbook, "Remain").Range("A1").Formula = "=Transfer!$A$1"
            Return source
        End Function

        Private Function CreateTransferTarget(collision As Boolean) As MsExcelDataOperations
            Dim target = Me.CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile,
                                          New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile))
            Dim check = TransferSheet(target.Workbook, "Sheet1")
            check.Name = "Target Check"
            If collision Then
                target.AddSheet("Transfer", Nothing)
                target.AddSheet("Data Set", Nothing)
                TransferSheet(target.Workbook, "Transfer").Range("A1").Value2 = 900
                TransferSheet(target.Workbook, "Data Set").Range("A1").Value2 = 700
                check.Range("A2").Formula = "=Transfer!$A$1"
            End If
            Return target
        End Function

        Private Shared Function CreateTransferTestDirectory() As String
            Return Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ExcelOps-SheetTransfer-" & Guid.NewGuid().ToString("N"))).FullName
        End Function

        Private Shared Function TransferSheet(workbook As NativeExcel.Workbook, name As String) As NativeExcel.Worksheet
            Return DirectCast(workbook.Worksheets.Item(name), NativeExcel.Worksheet)
        End Function

        Private Shared Sub TransferWorksheet(sheet As NativeExcel.Worksheet, after As NativeExcel.Worksheet, move As Boolean)
            If move Then
                sheet.Move(After:=after)
            Else
                sheet.Copy(After:=after)
            End If
        End Sub

        Private Shared Function OpenWorkbookQualifier(workbook As NativeExcel.Workbook, sheet As String) As String
            Return "'[" & workbook.Name.Replace("'", "''") & "]" & sheet & "'!"
        End Function

        Private Shared Sub AssertTransferFormula(sheet As NativeExcel.Worksheet, address As String, expected As String, value As Double)
            Dim cell = sheet.Range(address)
            Dim context = sheet.Name & "!" & address
            Assert.That(CStr(cell.Formula).Replace("'", ""), [Is].EqualTo(expected.Replace("'", "")), context)
            Assert.That(CDbl(cell.Value2), [Is].EqualTo(value), context)
        End Sub

        Private Shared Function StoredWorkbookLinkTargets(path As String) As String()
            Using stream As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using document = SpreadsheetDocument.Open(stream, False)
                    Return document.WorkbookPart.ExternalWorkbookParts.Select(
                        Function(part)
                            Dim id = part.ExternalLink.GetFirstChild(Of ExternalBook)().Id.Value
                            Dim link = part.ExternalRelationships.Single(Function(item) item.Id = id)
                            Return Uri.UnescapeDataString(link.Uri.OriginalString).Replace("\"c, "/"c)
                        End Function).ToArray()
                End Using
            End Using
        End Function

    End Class

End Namespace
