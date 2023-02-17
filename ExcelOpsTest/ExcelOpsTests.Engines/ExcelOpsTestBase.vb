Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework.Interfaces

Namespace ExcelOpsTests.Engines

    <NonParallelizable>
    Public MustInherit Class ExcelOpsTestBase(Of T As ExcelOps.ExcelDataOperationsBase)

        Protected MustOverride Function _CreateInstance() As T

        Protected MustOverride Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String) As T

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <returns></returns>
        Protected Function CreateInstance() As T
            Try
                Return _CreateInstance()
            Catch ex As Exception
                If ex.GetType() Is GetType(PlatformNotSupportedException) Then
                    Throw
                ElseIf ex.GetType() Is GetType(CompuMaster.ComInterop.ComApplicationNotAvailableException) Then
                    Throw
                Else
                    Dim InnerEx As Exception = ex.InnerException
                    Do While InnerEx IsNot Nothing
                        If InnerEx.GetType() Is GetType(PlatformNotSupportedException) Then
                            Throw InnerEx
                        ElseIf InnerEx.GetType() Is GetType(CompuMaster.ComInterop.ComApplicationNotAvailableException) Then
                            Throw InnerEx
                        Else
                            InnerEx = InnerEx.InnerException
                        End If
                    Loop
                End If
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <param name="file"></param>
        ''' <param name="mode"></param>
        ''' <param name="[readOnly]"></param>
        ''' <param name="passwordForOpening"></param>
        ''' <returns></returns>
        Protected Function CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String) As T
            Try
                Return _CreateInstance(file, mode, [readOnly], passwordForOpening)
            Catch ex As Exception
                If ex.GetType() Is GetType(PlatformNotSupportedException) Then
                    Throw
                ElseIf ex.GetType() Is GetType(CompuMaster.ComInterop.ComApplicationNotAvailableException) Then
                    Throw
                Else
                    Dim InnerEx As Exception = ex.InnerException
                    Do While InnerEx IsNot Nothing
                        If InnerEx.GetType() Is GetType(PlatformNotSupportedException) Then
                            Throw InnerEx
                        ElseIf InnerEx.GetType() Is GetType(CompuMaster.ComInterop.ComApplicationNotAvailableException) Then
                            Throw InnerEx
                        Else
                            InnerEx = InnerEx.InnerException
                        End If
                    Loop
                End If
                Throw
            End Try
        End Function

        <OneTimeSetUp>
        Public Sub CommonOneTimeSetup()
            Try
                Assert.NotNull(Me.CreateInstance(Nothing, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, True, Nothing))
            Catch ex As PlatformNotSupportedException
                Assert.Ignore("Platform not supported: " & ex.Message)
            Catch ex As CompuMaster.ComInterop.ComApplicationNotAvailableException
                Assert.Ignore("Platform supports COM, but requested COM application not installed: " & ex.Message)
            Catch ex As System.Runtime.InteropServices.COMException
                Assert.Ignore("Platform not supported or requested COM application not installed: " & ex.Message)
            End Try
        End Sub

        <SetUp>
        Public Sub CommonSetup()
            Test.Console.ResetConsoleForTestOutput()
        End Sub

        <TearDown>
        Public Sub CommonTearDown()
            CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
            If TestContext.CurrentContext.Result.Outcome = NUnit.Framework.Interfaces.ResultState.Failure Then
                Dim Log As String = Test.Console.GetConsoleLog
                System.Console.WriteLine(Log)
            End If
        End Sub

        Public MustOverride ReadOnly Property ExpectedEngineName As String

        <Test> Public Sub EngineName()
            Assert.AreEqual(ExpectedEngineName, Me.CreateInstance().EngineName)
        End Sub

        <Test> Public Sub HasVbaProject()
            Dim VbaTestFile = TestEnvironment.FullPathOfExistingTestFile("test_data", "VbaProject.xlsm")
            Assert.IsTrue(Me.CreateInstance(VbaTestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "").HasVbaProject)

            Dim NonVbaTestFile = TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund01.xlsx")
            Assert.IsFalse(Me.CreateInstance(NonVbaTestFile, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "").HasVbaProject)
        End Sub

        <Test> Public Sub SaveXlsxWithVbaProjectMustFail()
            Dim Wb As T

            'VBA test file must fail to save until VBA project has been removed
            Dim VbaTestFile = TestEnvironment.FullPathOfExistingTestFile("test_data", "VbaProject.xlsm")
            Dim VbaTestFileClone = TestEnvironment.FullPathOfDynTestFile(NameOf(SaveXlsxWithVbaProjectMustFail), "VbaProject.xlsm")
            Dim NewXlsxTargetPath As String = TestEnvironment.FullPathOfDynTestFile(NameOf(SaveXlsxWithVbaProjectMustFail), "VbaProject.xlsx")
            System.IO.File.Copy(VbaTestFile, VbaTestFileClone)

            Wb = Me.CreateInstance(VbaTestFileClone, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "")
            Assert.True(Wb.HasVbaProject)
            Assert.Throws(Of NotSupportedException)(Sub() Wb.SaveAs(NewXlsxTargetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour))
            Dim FilePathInEngineBefore As String = Wb.WorkbookFilePath
            Wb.RemoveVbaProject()
            If GetType(T) Is GetType(MsExcelDataOperations) Then
                'MS Excel engine: feature RemoveVbaProject not supported + workaround only partially possible
                Assert.IsNotEmpty(Wb.WorkbookFilePath)
            Else
                Assert.AreEqual(FilePathInEngineBefore, Wb.WorkbookFilePath)
            End If
            Assert.False(Wb.HasVbaProject)
            Wb.SaveAs(NewXlsxTargetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Wb.Close()

            Wb = Me.CreateInstance(VbaTestFileClone, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, "")
            Assert.True(Wb.HasVbaProject)
            Wb.Save()
            Assert.True(Wb.HasVbaProject, "VBA project hasn't been removed automatically")
            Wb.RemoveVbaProject()
            Assert.False(Wb.HasVbaProject)
            Wb.Save()
            Wb.Close()

            'But new created file saves with success
            Dim NonVbaTestFile = TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund01.xlsx")
            Dim NonVbaTestFileClone = TestEnvironment.FullPathOfDynTestFile(NameOf(SaveXlsxWithVbaProjectMustFail), "ExcelOpsGrund01.xlsx")
            NewXlsxTargetPath = TestEnvironment.FullPathOfDynTestFile(NameOf(SaveXlsxWithVbaProjectMustFail), "NonVbaProject.xlsx")
            System.IO.File.Copy(NonVbaTestFile, NonVbaTestFileClone)

            Wb = Me.CreateInstance(NonVbaTestFileClone, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "")
            Assert.False(Wb.HasVbaProject)
            Wb.SaveAs(NewXlsxTargetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Wb.Close()

            Wb = Me.CreateInstance(NonVbaTestFileClone, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, "")
            Assert.False(Wb.HasVbaProject)
            Wb.Save()
            Wb.Close()

            'Loading a workbook with VBA project + removing VBA project + saving workbook as XLSM + reloading workbook = must still HasVbaProject = False
            VbaTestFile = TestEnvironment.FullPathOfExistingTestFile("test_data", "VbaProject.xlsm")
            NewXlsxTargetPath = TestEnvironment.FullPathOfDynTestFile(NameOf(SaveXlsxWithVbaProjectMustFail), "VbaProject.xlsm")
            Wb = Me.CreateInstance(VbaTestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "")
            Assert.True(Wb.HasVbaProject)
            Wb.RemoveVbaProject()
            Assert.False(Wb.HasVbaProject)
            Wb.SaveAs(NewXlsxTargetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Wb.Close()
            Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "")
            Assert.False(Wb.HasVbaProject)

        End Sub

        <Test> Public Sub PasswordForOpening()
            Dim Wb As T
            'Testfile without password
            Dim TestFile As String = TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund01.xlsx")
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "")
            Assert.AreEqual("Grunddaten", Wb.SheetNames(0))

            'Now, save it with password
            Wb.PasswordForOpening = "dummy"
            Dim NewXlsxTargetPath As String = TestEnvironment.FullPathOfDynTestFile("PasswordProtectedFile.xlsx")
            Wb.SaveAs(NewXlsxTargetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            'Console.WriteLine("Saved password protected file to: " & NewXlsxTargetPath)
            Wb.Close()

            'Try to reload it without password -> it must fail
            Assert.Catch(Of Exception)(Sub() Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "something else"))
            Assert.Catch(Of Exception)(Sub() Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, ""))
            Assert.Catch(Of Exception)(Sub() Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing))

            'Reload it with password -> now it must succeed
            Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "dummy")
            Assert.AreEqual("Grunddaten", Wb.SheetNames(0))
        End Sub

        <Test> Public Sub CreateWorkbookWithoutFilePath()
            Dim Wb As T
            Dim TestFile As String

            TestFile = Nothing
            Wb = Me.CreateInstance()
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)
            Wb.Close()

            TestFile = Nothing
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, "")
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)
            Wb.Close()

            TestFile = ""
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, "")
            Assert.AreEqual(Nothing, Wb.FilePath)
            Assert.AreEqual(Nothing, Wb.WorkbookFilePath)
            Wb.Close()
        End Sub

        <Test> Public Sub CreateAndSaveAsAndFilePath()
            Dim Wb As T
            Dim TestFile As String = TestEnvironment.FullPathOfDynTestFile("created-workbook.xlsx")
            Dim TestFile2 As String = TestEnvironment.FullPathOfDynTestFile("created-workbook2.xlsx")

            'Creating a new workbook without pre-defined file name must fail on Save(), but successful on SaveAs()
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, "")
            Assert.AreEqual(TestFile = Nothing, Wb.ReadOnly, "Newly created files must be ReadOnly if file path hasn't been set up")
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(Nothing, Wb.WorkbookFilePath)
            Wb.Save()
            Wb.ReadOnly = True
            Assert.Throws(Of FileReadOnlyException)(Sub() Wb.Save())
            Assert.Throws(Of FileReadOnlyException)(Sub() Wb.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour))
            Wb.SaveAs(TestFile2, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Wb.Close()

            'Creating a new workbook must fail with a pre-defined file name if there is already a file
            Assert.Throws(Of FileAlreadyExistsException)(Sub()
                                                             Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, "")
                                                         End Sub)
            System.IO.File.Delete(TestFile) 'Delete the file for next test block

            'Creating a new workbook must always be ReadOnly and saving it without a name must be forbidden
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, "")
            Assert.AreEqual(TestFile = Nothing, Wb.ReadOnly, "Newly created files must always be ReadOnly")
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(Nothing, Wb.WorkbookFilePath)
            Wb.Save()
            Wb.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Assert.AreEqual(False, Wb.ReadOnly, "Newly saved files must always be ReadWrite")
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)
            Wb.Close()
            Assert.AreEqual(False, Wb.ReadOnly)
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(Nothing, Wb.WorkbookFilePath)
            Wb.ReloadFromFile()
            Assert.AreEqual(False, Wb.ReadOnly)
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)

            'Saving a ReadWrite file must be forbidden
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "")
            Assert.AreEqual(True, Wb.ReadOnly)
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)
            Assert.Throws(Of FileReadOnlyException)(Sub() Wb.Save())
            Assert.AreEqual(True, Wb.ReadOnly)
            Assert.Throws(Of FileReadOnlyException)(Sub() Wb.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour))
            Assert.AreEqual(True, Wb.ReadOnly)
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)
            Wb.Close()

            'Saving a ReadWrite file must be allowed
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, "")
            Assert.AreEqual(False, Wb.ReadOnly)
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)
            Wb.Save()
            Assert.AreEqual(False, Wb.ReadOnly)
            Wb.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Assert.AreEqual(False, Wb.ReadOnly)
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)
            Wb.Close()
        End Sub

        <Test>
        Public Sub CreateInstanceWithOrWithoutCreationOfWorkbook()
            Dim workbook As ExcelDataOperationsBase

            workbook = Me.CreateInstance()
            Select Case workbook.GetType
                Case GetType(MsExcelDataOperations)
                    'Accept fact that a new workbook is opened automatically
                    Assert.NotZero(workbook.SheetNames.Count)
                Case Else
                    'No workbook opened - must be done in 2ndary step
                    Assert.Throws(Of InvalidOperationException)(Function() workbook.SheetNames.Count)
            End Select
            workbook.Close()

            workbook = Me.CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile, True, Nothing)
            Assert.AreEqual(1, workbook.SheetNames.Count, "Sheets Count")
            Assert.AreEqual("Sheet1", workbook.SheetNames(0))
            workbook.Close()

        End Sub

        <Test> Public Overridable Sub CopySheetContent()
            Dim eppeoIn As ExcelOps.ExcelDataOperationsBase
            Dim eppeoOut As ExcelOps.ExcelDataOperationsBase = Nothing
            Dim TestControllingToolFileNameIn As String
            Dim TestControllingToolFileNameOutTemplate As String
            Dim TestControllingToolFileNameOut As String

            TestControllingToolFileNameIn = TestFiles.TestFileGrund01.FullName
            TestControllingToolFileNameOutTemplate = TestFiles.TestFileGrund02.FullName
            TestControllingToolFileNameOut = TestEnvironment.FullPathOfDynTestFile("CopySheetContent_" & GetType(T).Name & ".xlsx")
            Try
                Console.WriteLine("Test file in: " & TestControllingToolFileNameIn)
                Console.WriteLine("Test file output template: " & TestControllingToolFileNameOutTemplate)
                Console.WriteLine("Test file output: " & TestControllingToolFileNameOut)

                eppeoIn = Me.CreateInstance(TestControllingToolFileNameIn, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)
                eppeoOut = Me.CreateInstance(TestControllingToolFileNameOutTemplate, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)

                Const SheetToCopy As String = "Grunddaten"
                eppeoIn.CopySheetContent(SheetToCopy, eppeoOut, ExcelOps.ExcelDataOperationsBase.CopySheetOption.TargetSheetMightExist)
                eppeoOut.SelectSheet(SheetToCopy)
                eppeoOut.SaveAs(TestControllingToolFileNameOut, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.AreEqual(eppeoIn.SheetContentMatrix(SheetToCopy, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText), eppeoOut.SheetContentMatrix(SheetToCopy, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText))
                Assert.Pass("Required manual, optical review for comparison to check for formattings")
            Finally
                If eppeoOut IsNot Nothing Then eppeoOut.Close()
            End Try
        End Sub

        <Test> Public Sub ExcelOpsTestCollection_ZahlenUndProzentwerte()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestFiles.TestFileExcelOpsTestCollection.FullName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)
            Dim SheetName As String
            SheetName = "ZahlenUndProzentwerte"
            Assert.AreEqual("0.00", eppeo.LookupCellFormat(SheetName, 0, 1))
            Assert.AreEqual("0.00%", eppeo.LookupCellFormat(SheetName, 1, 1))
            Assert.AreEqual(10.0, System.Math.Round(eppeo.LookupCellValue(Of Double)(SheetName, 0, 1), 2))
            Assert.AreEqual(0.1, System.Math.Round(eppeo.LookupCellValue(Of Double)(SheetName, 1, 1), 2))
            eppeo.WriteCellValue(Of Double)(SheetName, 0, 1, 20.0)
            eppeo.WriteCellValue(Of Double)(SheetName, 1, 1, 0.2)
            Assert.AreEqual(20.0, System.Math.Round(eppeo.LookupCellValue(Of Double)(SheetName, 0, 1), 2))
            Assert.AreEqual(0.2, System.Math.Round(eppeo.LookupCellValue(Of Double)(SheetName, 1, 1), 2))
        End Sub

        <Test> Public Sub SheetNames()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)

            Dim EppeoSheetNamesList As List(Of String)
            EppeoSheetNamesList = eppeo.SheetNames
            System.Console.WriteLine(Strings.Join(EppeoSheetNamesList.ToArray, ","))
            Assert.AreEqual("Grunddaten,Kostenplanung", Strings.Join(EppeoSheetNamesList.ToArray, ","))
        End Sub

        <Test> Public Sub AddSheet()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            Dim BeforeSheet As String = "Grunddaten"
            Dim SheetNameTopPosition As String = "SheetOnTop"
            Dim SheetNameBottomPosition As String = "SheetOnBottom"
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)
            Dim ExpectedSheetNamesList, NewSheetNamesList As List(Of String)
            ExpectedSheetNamesList = eppeo.SheetNames
            ExpectedSheetNamesList.Add(SheetNameBottomPosition)
            ExpectedSheetNamesList.Insert(0, SheetNameTopPosition)

            eppeo.AddSheet(SheetNameBottomPosition)
            eppeo.AddSheet(SheetNameTopPosition, BeforeSheet)
            NewSheetNamesList = eppeo.SheetNames
            Assert.AreEqual(ExpectedSheetNamesList.ToArray, NewSheetNamesList.ToArray)
        End Sub

        <Test> Public Sub AllFormulasOfWorkbook()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String
            Dim AllFormulas As List(Of ExcelOps.TextTableCell)

            TestControllingToolFileName = TestFiles.TestFileGrund01.FullName
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)
            AllFormulas = eppeo.AllFormulasOfWorkbook
            Console.WriteLine("Test file: " & TestControllingToolFileName)
            Assert.NotZero(AllFormulas.Count)

            Dim ReferencesFromTestSheet As List(Of TextTableCell) = ExcelOps.Tools.FormulasWithSheetReferencesFromSheet("Kostenplanung", AllFormulas, eppeo.SheetNames.ToArray)
            System.Console.WriteLine("## Formulas of Kostenplanung targetting other sheets in workbook")
            For MyFormulaCounter As Integer = 0 To ReferencesFromTestSheet.Count - 1
                System.Console.WriteLine(ReferencesFromTestSheet(MyFormulaCounter).ToString)
            Next
            Assert.NotZero(ReferencesFromTestSheet.Count)
            Assert.IsTrue(ExcelOps.Tools.ContainsFormulasWithSheetReferencesFromSheet("Kostenplanung", AllFormulas, eppeo.SheetNames.ToArray))
            Assert.IsFalse(ExcelOps.Tools.ContainsFormulasWithSheetReferencesFromSheet("Grunddaten", AllFormulas, eppeo.SheetNames.ToArray))

            System.Console.WriteLine()
            System.Console.WriteLine("## Formulas of sheets in workbook targetting Grunddaten")
            Dim ReferencesToTestSheet As List(Of TextTableCell) = ExcelOps.Tools.FormulasWithSheetReferencesToSheet(AllFormulas, "Grunddaten", Nothing)
            For MyFormulaCounter As Integer = 0 To ReferencesToTestSheet.Count - 1
                System.Console.WriteLine(ReferencesToTestSheet(MyFormulaCounter).ToString)
            Next
            Assert.NotZero(ReferencesToTestSheet.Count)
            Assert.IsTrue(ExcelOps.Tools.ContainsFormulasWithSheetReferencesToSheet(AllFormulas, "Grunddaten", Nothing))
        End Sub

        <Test> Public Sub CellWithError()
            Dim wb As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestFiles.TestFileGrund02.FullName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)
            Dim SheetName As String = wb.SheetNames(0)

            wb.WriteCellFormula(SheetName, 0, 0, "B2", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual(Nothing, wb.LookupCellErrorValue(SheetName, 0, 0))

            wb.WriteCellFormula(SheetName, 0, 0, "1/1", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual(Nothing, wb.LookupCellErrorValue(SheetName, 0, 0))
            Assert.AreEqual(1, wb.LookupCellValue(Of Integer)(SheetName, 0, 0))

            wb.WriteCellFormula(SheetName, 0, 0, "INVALIDFUNCTION(B2)", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual("#NAME?", wb.LookupCellErrorValue(SheetName, 0, 0))

            wb.WriteCellFormula(SheetName, 0, 0, "1/0", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual("#DIV/0!", wb.LookupCellErrorValue(SheetName, 0, 0))

            wb.WriteCellFormula(SheetName, 0, 0, "B2/0", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual("#DIV/0!", wb.LookupCellErrorValue(SheetName, 0, 0))

            Select Case GetType(T)
                Case GetType(EpplusFreeExcelDataOperations), GetType(EpplusPolyformExcelDataOperations)
                    'bug in Epplus engine
                    Assert.Ignore("Bugs in Epplus formula manager engine")
                Case Else
                    wb.WriteCellFormula(SheetName, 0, 0, "A0", False)
                    wb.RecalculateCell(SheetName, 0, 0, False)
                    Assert.AreEqual("#NAME?", wb.LookupCellErrorValue(SheetName, 0, 0))
            End Select
        End Sub

        <Test> Public Sub LookupLastCellAddress()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)

            Dim LastCellFound As ExcelOps.ExcelCell
            LastCellFound = eppeo.LookupLastContentCell(TestSheet)
            Assert.AreEqual("E40", LastCellFound.Address)
            Assert.AreEqual(eppeo.LookupLastContentRowIndex(TestSheet), LastCellFound.RowIndex)
            Assert.AreEqual(eppeo.LookupLastContentColumnIndex(TestSheet), LastCellFound.ColumnIndex)
        End Sub

        Protected Delegate Sub TestInCultureContextAction()

        Protected Sub TestInCultureContext(cultureName As String, testMethod As TestInCultureContextAction)
            Dim OriginCulture = System.Threading.Thread.CurrentThread.CurrentCulture
            Try
                Select Case cultureName
                    Case "", "invariant"
                        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture
                    Case Else
                        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo(cultureName)
                End Select
                TestInCultureContext_AssignCurrentThreadCulture()
                testMethod()
            Finally
                System.Threading.Thread.CurrentThread.CurrentCulture = OriginCulture
                TestInCultureContext_AssignCurrentThreadCulture()
            End Try
        End Sub

        ''' <summary>
        ''' Assign current thread's culture to excel engine, if it requires additional steps after culture change (e.g. MS Excel)
        ''' </summary>
        Protected Overridable Sub TestInCultureContext_AssignCurrentThreadCulture()
        End Sub

        Protected Const PlaceHolderDecimalSeparator As String = "▲"c
        Protected Const PlaceHolderGroupSeparator As String = "▪"c

        Private Function ExpectedResultInCultureContext(expectedRawMatrix As String) As String
            Return expectedRawMatrix.
                Replace(PlaceHolderDecimalSeparator, System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator).
                Replace(PlaceHolderGroupSeparator, System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator)
        End Function

        <Test> Public Overridable Sub SheetContentMatrix_StaticOrCalculatedValues(<Values("invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)

                    ExpectedMatrix =
                         "# |A                           |B        |C  |D     |E    " & ControlChars.CrLf &
                         "--+----------------------------+---------+---+------+-----" & ControlChars.CrLf &
                         "1 |Jahr                        |2019     |   |      |False" & ControlChars.CrLf &
                         "2 |Geschäftsjahr von           |         |bis|      |     " & ControlChars.CrLf &
                         "3 |Aktueller Monat             |1        |   |Januar|     " & ControlChars.CrLf &
                         "4 |                            |         |   |      |     " & ControlChars.CrLf &
                         "5 |Name Betrieb                |Test     |   |      |     " & ControlChars.CrLf &
                         "6 |                            |         |   |      |     " & ControlChars.CrLf &
                         "7 |Arbeitgeberanteile in %     |         |   |      |     " & ControlChars.CrLf &
                         "8 |Chef: 14▲09                 |         |   |      |     " & ControlChars.CrLf &
                         "9 |Büroangestellte: 20▲00      |         |   |      |     " & ControlChars.CrLf &
                         "10|Produktivkraft: 25▲00       |         |   |      |     " & ControlChars.CrLf &
                         "11|Azubi / Aushilfen: 33▲00    |         |   |      |     " & ControlChars.CrLf &
                         "12|                            |         |   |      |     " & ControlChars.CrLf &
                         "13|Berechnung Jahresarbeitszeit|         |   |      |     " & ControlChars.CrLf &
                         "14|Tage / Jahr:                |365      |   |      |     " & ControlChars.CrLf &
                         "15|Wochenendtage               |104      |   |      |     " & ControlChars.CrLf &
                         "16|=Zahltage:                  |261      |   |      |     " & ControlChars.CrLf &
                         "17|Wochenarbeitszeit           |40       |   |      |     " & ControlChars.CrLf &
                         "18|Tagesarbeitszeit:           |8        |   |      |     " & ControlChars.CrLf &
                         "19|Normallohnstunden / Jahr:   |2▪088▲00 |   |      |     " & ControlChars.CrLf &
                         "20|                            |         |   |      |     " & ControlChars.CrLf &
                         "21|                            |         |   |      |     " & ControlChars.CrLf &
                         "22|                            |         |   |      |     " & ControlChars.CrLf &
                         "23|1                           |Januar   |   |      |     " & ControlChars.CrLf &
                         "24|2                           |Februar  |   |      |     " & ControlChars.CrLf &
                         "25|3                           |März     |   |      |     " & ControlChars.CrLf &
                         "26|4                           |April    |   |      |     " & ControlChars.CrLf &
                         "27|5                           |Mai      |   |      |     " & ControlChars.CrLf &
                         "28|6                           |Juni     |   |      |     " & ControlChars.CrLf &
                         "29|7                           |Juli     |   |      |     " & ControlChars.CrLf &
                         "30|8                           |August   |   |      |     " & ControlChars.CrLf &
                         "31|9                           |September|   |      |     " & ControlChars.CrLf &
                         "32|10                          |Oktober  |   |      |     " & ControlChars.CrLf &
                         "33|11                          |November |   |      |     " & ControlChars.CrLf &
                         "34|12                          |Dezember |   |      |     " & ControlChars.CrLf &
                         "35|Zusammensetzung AG Anteile  |         |   |      |     " & ControlChars.CrLf &
                         "36|Krankenkasse                |2▲8      |   |      |     " & ControlChars.CrLf &
                         "37|Rentenkasse                 |8        |   |      |     " & ControlChars.CrLf &
                         "38|Pflegekasse                 |1▲4      |   |      |     " & ControlChars.CrLf &
                         "39|Krankengeld                 |0▲25     |   |      |     " & ControlChars.CrLf &
                         "40|                            |12▲45    |   |      |     " & ControlChars.CrLf
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues, ExpectedResultInCultureContext(ExpectedMatrix))
                End Sub)
        End Sub

        <Test> Public Overridable Sub SheetContentMatrix_StaticValues(<Values("invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)

                    ExpectedMatrix =
                                 "# |A                           |B        |C  |D |E    " & ControlChars.CrLf &
                                 "--+----------------------------+---------+---+--+-----" & ControlChars.CrLf &
                                 "1 |Jahr                        |2019     |   |  |False" & ControlChars.CrLf &
                                 "2 |Geschäftsjahr von           |         |bis|  |     " & ControlChars.CrLf &
                                 "3 |Aktueller Monat             |1        |   |  |     " & ControlChars.CrLf &
                                 "4 |                            |         |   |  |     " & ControlChars.CrLf &
                                 "5 |Name Betrieb                |Test     |   |  |     " & ControlChars.CrLf &
                                 "6 |                            |         |   |  |     " & ControlChars.CrLf &
                                 "7 |Arbeitgeberanteile in %     |         |   |  |     " & ControlChars.CrLf &
                                 "8 |Chef: 14▲09                 |         |   |  |     " & ControlChars.CrLf &
                                 "9 |Büroangestellte: 20▲00      |         |   |  |     " & ControlChars.CrLf &
                                 "10|Produktivkraft: 25▲00       |         |   |  |     " & ControlChars.CrLf &
                                 "11|Azubi / Aushilfen: 33▲00    |         |   |  |     " & ControlChars.CrLf &
                                 "12|                            |         |   |  |     " & ControlChars.CrLf &
                                 "13|Berechnung Jahresarbeitszeit|         |   |  |     " & ControlChars.CrLf &
                                 "14|Tage / Jahr:                |365      |   |  |     " & ControlChars.CrLf &
                                 "15|Wochenendtage               |         |   |  |     " & ControlChars.CrLf &
                                 "16|=Zahltage:                  |         |   |  |     " & ControlChars.CrLf &
                                 "17|Wochenarbeitszeit           |40       |   |  |     " & ControlChars.CrLf &
                                 "18|Tagesarbeitszeit:           |         |   |  |     " & ControlChars.CrLf &
                                 "19|Normallohnstunden / Jahr:   |         |   |  |     " & ControlChars.CrLf &
                                 "20|                            |         |   |  |     " & ControlChars.CrLf &
                                 "21|                            |         |   |  |     " & ControlChars.CrLf &
                                 "22|                            |         |   |  |     " & ControlChars.CrLf &
                                 "23|1                           |Januar   |   |  |     " & ControlChars.CrLf &
                                 "24|2                           |Februar  |   |  |     " & ControlChars.CrLf &
                                 "25|3                           |März     |   |  |     " & ControlChars.CrLf &
                                 "26|4                           |April    |   |  |     " & ControlChars.CrLf &
                                 "27|5                           |Mai      |   |  |     " & ControlChars.CrLf &
                                 "28|6                           |Juni     |   |  |     " & ControlChars.CrLf &
                                 "29|7                           |Juli     |   |  |     " & ControlChars.CrLf &
                                 "30|8                           |August   |   |  |     " & ControlChars.CrLf &
                                 "31|9                           |September|   |  |     " & ControlChars.CrLf &
                                 "32|10                          |Oktober  |   |  |     " & ControlChars.CrLf &
                                 "33|11                          |November |   |  |     " & ControlChars.CrLf &
                                 "34|12                          |Dezember |   |  |     " & ControlChars.CrLf &
                                 "35|Zusammensetzung AG Anteile  |         |   |  |     " & ControlChars.CrLf &
                                 "36|Krankenkasse                |2▲8      |   |  |     " & ControlChars.CrLf &
                                 "37|Rentenkasse                 |8        |   |  |     " & ControlChars.CrLf &
                                 "38|Pflegekasse                 |1▲4      |   |  |     " & ControlChars.CrLf &
                                 "39|Krankengeld                 |0▲25     |   |  |     " & ControlChars.CrLf
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticValues, ExpectedResultInCultureContext(ExpectedMatrix))
                End Sub)
        End Sub

        <Test> Public Overridable Sub SheetContentMatrix_Formulas(<Values("invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)

                    ExpectedMatrix =
                                 "# |A |B            |C |D                                  " & ControlChars.CrLf &
                                 "--+--+-------------+--+-----------------------------------" & ControlChars.CrLf &
                                 "1 |  |             |  |                                   " & ControlChars.CrLf &
                                 "2 |  |             |  |                                   " & ControlChars.CrLf &
                                 "3 |  |             |  |=INDEX(B23:B34,MATCH(B3,A23:A34,0))" & ControlChars.CrLf &
                                 "4 |  |             |  |                                   " & ControlChars.CrLf &
                                 "5 |  |             |  |                                   " & ControlChars.CrLf &
                                 "6 |  |             |  |                                   " & ControlChars.CrLf &
                                 "7 |  |             |  |                                   " & ControlChars.CrLf &
                                 "8 |  |             |  |                                   " & ControlChars.CrLf &
                                 "9 |  |             |  |                                   " & ControlChars.CrLf &
                                 "10|  |             |  |                                   " & ControlChars.CrLf &
                                 "11|  |             |  |                                   " & ControlChars.CrLf &
                                 "12|  |             |  |                                   " & ControlChars.CrLf &
                                 "13|  |             |  |                                   " & ControlChars.CrLf &
                                 "14|  |             |  |                                   " & ControlChars.CrLf &
                                 "15|  |=2*52        |  |                                   " & ControlChars.CrLf &
                                 "16|  |=B14-B15     |  |                                   " & ControlChars.CrLf &
                                 "17|  |             |  |                                   " & ControlChars.CrLf &
                                 "18|  |=B17/5       |  |                                   " & ControlChars.CrLf &
                                 "19|  |=B18*B16     |  |                                   " & ControlChars.CrLf &
                                 "20|  |             |  |                                   " & ControlChars.CrLf &
                                 "21|  |             |  |                                   " & ControlChars.CrLf &
                                 "22|  |             |  |                                   " & ControlChars.CrLf &
                                 "23|  |             |  |                                   " & ControlChars.CrLf &
                                 "24|  |             |  |                                   " & ControlChars.CrLf &
                                 "25|  |             |  |                                   " & ControlChars.CrLf &
                                 "26|  |             |  |                                   " & ControlChars.CrLf &
                                 "27|  |             |  |                                   " & ControlChars.CrLf &
                                 "28|  |             |  |                                   " & ControlChars.CrLf &
                                 "29|  |             |  |                                   " & ControlChars.CrLf &
                                 "30|  |             |  |                                   " & ControlChars.CrLf &
                                 "31|  |             |  |                                   " & ControlChars.CrLf &
                                 "32|  |             |  |                                   " & ControlChars.CrLf &
                                 "33|  |             |  |                                   " & ControlChars.CrLf &
                                 "34|  |             |  |                                   " & ControlChars.CrLf &
                                 "35|  |             |  |                                   " & ControlChars.CrLf &
                                 "36|  |             |  |                                   " & ControlChars.CrLf &
                                 "37|  |             |  |                                   " & ControlChars.CrLf &
                                 "38|  |             |  |                                   " & ControlChars.CrLf &
                                 "39|  |             |  |                                   " & ControlChars.CrLf &
                                 "40|  |=SUM(B36:B39)|  |                                   " & ControlChars.CrLf
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.Formulas, ExpectedResultInCultureContext(ExpectedMatrix))
                End Sub)
        End Sub

        <Test> Public Overridable Sub SheetContentMatrix_FormattedText(<Values("en-US", "invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)

                    ExpectedMatrix =
                         "# |A                           |B        |C  |D     |E    " & ControlChars.CrLf &
                         "--+----------------------------+---------+---+------+-----" & ControlChars.CrLf &
                         "1 |Jahr                        |2019     |   |      |False" & ControlChars.CrLf &
                         "2 |Geschäftsjahr von           |         |bis|      |     " & ControlChars.CrLf &
                         "3 |Aktueller Monat             |1        |   |Januar|     " & ControlChars.CrLf &
                         "4 |                            |         |   |      |     " & ControlChars.CrLf &
                         "5 |Name Betrieb                |Test     |   |      |     " & ControlChars.CrLf &
                         "6 |                            |         |   |      |     " & ControlChars.CrLf &
                         "7 |Arbeitgeberanteile in %     |         |   |      |     " & ControlChars.CrLf &
                         "8 |Chef: 14▲09                 |         |   |      |     " & ControlChars.CrLf &
                         "9 |Büroangestellte: 20▲00      |         |   |      |     " & ControlChars.CrLf &
                         "10|Produktivkraft: 25▲00       |         |   |      |     " & ControlChars.CrLf &
                         "11|Azubi / Aushilfen: 33▲00    |         |   |      |     " & ControlChars.CrLf &
                         "12|                            |         |   |      |     " & ControlChars.CrLf &
                         "13|Berechnung Jahresarbeitszeit|         |   |      |     " & ControlChars.CrLf &
                         "14|Tage / Jahr:                |365      |   |      |     " & ControlChars.CrLf &
                         "15|Wochenendtage               |104      |   |      |     " & ControlChars.CrLf &
                         "16|=Zahltage:                  |261      |   |      |     " & ControlChars.CrLf &
                         "17|Wochenarbeitszeit           |40       |   |      |     " & ControlChars.CrLf &
                         "18|Tagesarbeitszeit:           |8        |   |      |     " & ControlChars.CrLf &
                         "19|Normallohnstunden / Jahr:   |2▪088▲00 |   |      |     " & ControlChars.CrLf &
                         "20|                            |         |   |      |     " & ControlChars.CrLf &
                         "21|                            |         |   |      |     " & ControlChars.CrLf &
                         "22|                            |         |   |      |     " & ControlChars.CrLf &
                         "23|1                           |Januar   |   |      |     " & ControlChars.CrLf &
                         "24|2                           |Februar  |   |      |     " & ControlChars.CrLf &
                         "25|3                           |März     |   |      |     " & ControlChars.CrLf &
                         "26|4                           |April    |   |      |     " & ControlChars.CrLf &
                         "27|5                           |Mai      |   |      |     " & ControlChars.CrLf &
                         "28|6                           |Juni     |   |      |     " & ControlChars.CrLf &
                         "29|7                           |Juli     |   |      |     " & ControlChars.CrLf &
                         "30|8                           |August   |   |      |     " & ControlChars.CrLf &
                         "31|9                           |September|   |      |     " & ControlChars.CrLf &
                         "32|10                          |Oktober  |   |      |     " & ControlChars.CrLf &
                         "33|11                          |November |   |      |     " & ControlChars.CrLf &
                         "34|12                          |Dezember |   |      |     " & ControlChars.CrLf &
                         "35|Zusammensetzung AG Anteile  |         |   |      |     " & ControlChars.CrLf &
                         "36|Krankenkasse                |2▲8      |   |      |     " & ControlChars.CrLf &
                         "37|Rentenkasse                 |8        |   |      |     " & ControlChars.CrLf &
                         "38|Pflegekasse                 |1▲4      |   |      |     " & ControlChars.CrLf &
                         "39|Krankengeld                 |0▲25     |   |      |     " & ControlChars.CrLf &
                         "40|                            |12▲45    |   |      |     " & ControlChars.CrLf
                    Assert.AreEqual(12.45.ToString, eppeo.LookupCellFormattedText(TestSheet, 40 - 1, 2 - 1))
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormattedText, ExpectedResultInCultureContext(ExpectedMatrix))
                End Sub)
            System.Console.WriteLine(Console.GetConsoleLog)
        End Sub

        <Test> Public Overridable Sub SheetContentMatrix_FormulaOrFormattedText(<Values("invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)

                    ExpectedMatrix =
                         "# |A                           |B            |C  |D                                  |E    " & ControlChars.CrLf &
                         "--+----------------------------+-------------+---+-----------------------------------+-----" & ControlChars.CrLf &
                         "1 |Jahr                        |2019         |   |                                   |False" & ControlChars.CrLf &
                         "2 |Geschäftsjahr von           |             |bis|                                   |     " & ControlChars.CrLf &
                         "3 |Aktueller Monat             |1            |   |=INDEX(B23:B34,MATCH(B3,A23:A34,0))|     " & ControlChars.CrLf &
                         "4 |                            |             |   |                                   |     " & ControlChars.CrLf &
                         "5 |Name Betrieb                |Test         |   |                                   |     " & ControlChars.CrLf &
                         "6 |                            |             |   |                                   |     " & ControlChars.CrLf &
                         "7 |Arbeitgeberanteile in %     |             |   |                                   |     " & ControlChars.CrLf &
                         "8 |Chef: 14▲09                 |             |   |                                   |     " & ControlChars.CrLf &
                         "9 |Büroangestellte: 20▲00      |             |   |                                   |     " & ControlChars.CrLf &
                         "10|Produktivkraft: 25▲00       |             |   |                                   |     " & ControlChars.CrLf &
                         "11|Azubi / Aushilfen: 33▲00    |             |   |                                   |     " & ControlChars.CrLf &
                         "12|                            |             |   |                                   |     " & ControlChars.CrLf &
                         "13|Berechnung Jahresarbeitszeit|             |   |                                   |     " & ControlChars.CrLf &
                         "14|Tage / Jahr:                |365          |   |                                   |     " & ControlChars.CrLf &
                         "15|Wochenendtage               |=2*52        |   |                                   |     " & ControlChars.CrLf &
                         "16|=Zahltage:                  |=B14-B15     |   |                                   |     " & ControlChars.CrLf &
                         "17|Wochenarbeitszeit           |40           |   |                                   |     " & ControlChars.CrLf &
                         "18|Tagesarbeitszeit:           |=B17/5       |   |                                   |     " & ControlChars.CrLf &
                         "19|Normallohnstunden / Jahr:   |=B18*B16     |   |                                   |     " & ControlChars.CrLf &
                         "20|                            |             |   |                                   |     " & ControlChars.CrLf &
                         "21|                            |             |   |                                   |     " & ControlChars.CrLf &
                         "22|                            |             |   |                                   |     " & ControlChars.CrLf &
                         "23|1                           |Januar       |   |                                   |     " & ControlChars.CrLf &
                         "24|2                           |Februar      |   |                                   |     " & ControlChars.CrLf &
                         "25|3                           |März         |   |                                   |     " & ControlChars.CrLf &
                         "26|4                           |April        |   |                                   |     " & ControlChars.CrLf &
                         "27|5                           |Mai          |   |                                   |     " & ControlChars.CrLf &
                         "28|6                           |Juni         |   |                                   |     " & ControlChars.CrLf &
                         "29|7                           |Juli         |   |                                   |     " & ControlChars.CrLf &
                         "30|8                           |August       |   |                                   |     " & ControlChars.CrLf &
                         "31|9                           |September    |   |                                   |     " & ControlChars.CrLf &
                         "32|10                          |Oktober      |   |                                   |     " & ControlChars.CrLf &
                         "33|11                          |November     |   |                                   |     " & ControlChars.CrLf &
                         "34|12                          |Dezember     |   |                                   |     " & ControlChars.CrLf &
                         "35|Zusammensetzung AG Anteile  |             |   |                                   |     " & ControlChars.CrLf &
                         "36|Krankenkasse                |2▲8          |   |                                   |     " & ControlChars.CrLf &
                         "37|Rentenkasse                 |8            |   |                                   |     " & ControlChars.CrLf &
                         "38|Pflegekasse                 |1▲4          |   |                                   |     " & ControlChars.CrLf &
                         "39|Krankengeld                 |0▲25         |   |                                   |     " & ControlChars.CrLf &
                         "40|                            |=SUM(B36:B39)|   |                                   |     " & ControlChars.CrLf
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText, ExpectedResultInCultureContext(ExpectedMatrix))
                End Sub)
        End Sub

        Private Sub AssertSheetContentMatrix(eo As ExcelOps.ExcelDataOperationsBase, sheetName As String, matrixContentType As ExcelOps.ExcelDataOperationsBase.MatrixContent, expectedMatrix As String)
            Dim MatrixContentName As String = matrixContentType.ToString
            Dim SheetData As TextTable = eo.SheetContentMatrix(sheetName, matrixContentType)
            SheetData.AutoTrim()
            Console.WriteLine("## Table " & eo.EngineName & " - " & MatrixContentName & " - " & System.Threading.Thread.CurrentThread.CurrentCulture.Name)
            Console.WriteLine(SheetData.ToUIExcelTable)
            Console.WriteLine("## /Table")
            Assert.AreEqual(expectedMatrix, SheetData.ToUIExcelTable)
        End Sub

        <Test> Public Sub LookupCellValue(<Values("invariant", "de-DE")> cultureName As String)
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing)

                    '## Expected matrix like following
                    '"# |A                           |B              |C  |D                                  |E     
                    '"--+----------------------------+---------------+---+-----------------------------------+----- 
                    '"1 |Jahr                        |2019           |   |                                   |False 
                    '"2 |Geschäftsjahr von           |               |bis|                                   |      
                    '"3 |Aktueller Monat             |1              |   |=INDEX(B23:B34,MATCH(B3,A23:A34,0))|      
                    '"4 |                            |               |   |                                   |      
                    '"5 |Name Betrieb                |Test           |   |                                   |      
                    '"6 |                            |               |   |                                   |      
                    '"7 |Arbeitgeberanteile In %     |               |   |                                   |      
                    '"8 |Chef: 14▲09                 |               |   |                                   |      
                    '"9 |Büroangestellte: 20▲00      |               |   |                                   |      
                    '"10|Produktivkraft: 25▲00       |               |   |                                   |      
                    '"11|Azubi / Aushilfen: 33▲00    |               |   |                                   |      
                    '"12|                            |               |   |                                   |      
                    '"13|Berechnung Jahresarbeitszeit|               |   |                                   |      
                    '"14|Tage / Jahr:                |365            |   |                                   |      
                    '"15|Wochenendtage               |=2*52          |   |                                   |      
                    '"16|=Zahltage:                  |=B14-B15       |   |                                   |      
                    '"17|Wochenarbeitszeit           |40             |   |                                   |      
                    '"18|Tagesarbeitszeit:           |=B17/5         |   |                                   |      
                    '"19|Normallohnstunden / Jahr:   |=B18*B16       |   |                                   |      
                    '"20|                            |               |   |                                   |      
                    '"36|Krankenkasse                |2▲8            |   |                                   |      
                    '"37|Rentenkasse                 |8              |   |                                   |      
                    '"38|Pflegekasse                 |1▲4            |   |                                   |      
                    '"39|Krankengeld                 |0▲25           |   |                                   |      
                    '"40|                            |=SUMME(B36:B39)|   |                                   |      

                    'D3
                    Assert.AreEqual("Januar", eppeo.LookupCellValue(Of String)(New ExcelOps.ExcelCell(TestSheet, "D3", ExcelOps.ExcelCell.ValueTypes.All)))
                    Assert.AreEqual("Januar", eppeo.LookupCellFormattedText(New ExcelOps.ExcelCell(TestSheet, "D3", ExcelOps.ExcelCell.ValueTypes.All)))
                    Assert.AreEqual("INDEX(B23:B34,MATCH(B3,A23:A34,0))", eppeo.LookupCellFormula(New ExcelOps.ExcelCell(TestSheet, "D3", ExcelOps.ExcelCell.ValueTypes.All)))

                    'A8
                    Assert.AreEqual(14.09D, eppeo.LookupCellValue(Of Double)(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))
                    Assert.AreEqual(ExpectedResultInCultureContext("Chef: 14▲09"), eppeo.LookupCellFormattedText(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))
                    Assert.AreEqual(Nothing, eppeo.LookupCellFormula(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))

                    'E1
                    Assert.AreEqual(False, eppeo.LookupCellValue(Of Boolean)(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                    Assert.AreEqual("False", eppeo.LookupCellValue(Of String)(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                    Assert.AreEqual("False", eppeo.LookupCellFormattedText(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                End Sub)
        End Sub

    End Class

End Namespace