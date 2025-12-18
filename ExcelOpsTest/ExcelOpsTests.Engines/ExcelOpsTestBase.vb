Option Strict On
Option Explicit On

Imports CompuMaster.Epplus4.FormulaParsing.Excel.Functions.Text
Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework
Imports NUnit.Framework.Interfaces

Namespace ExcelOpsTests.Engines

    <Apartment(Threading.ApartmentState.STA)>
    <NonParallelizable>
    Public MustInherit Class ExcelOpsTestBase(Of T As ExcelOps.ExcelDataOperationsBase)

#If CI_CD Then
        Private Const OPEN_HTML_OUTPUT_IN_BROWSER_AFTER_TEST As Boolean = False
#Else
        Private Const OPEN_HTML_OUTPUT_IN_BROWSER_AFTER_TEST As Boolean = True
#End If

        Protected MustOverride Function _CreateInstanceUninitialized() As T

#Disable Warning CA1716 ' Bezeichner dürfen nicht mit Schlüsselwörtern übereinstimmen
        Protected MustOverride Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, options As ExcelOps.ExcelDataOperationsOptions) As T
        Protected MustOverride Function _CreateInstance(data As Byte(), options As ExcelOps.ExcelDataOperationsOptions) As T
        Protected MustOverride Function _CreateInstance(data As System.IO.Stream, options As ExcelOps.ExcelDataOperationsOptions) As T
#Enable Warning CA1716 ' Bezeichner dürfen nicht mit Schlüsselwörtern übereinstimmen

        ''' <summary>
        ''' Create a new excel engine instance (reminder: set System.Threading.Thread.CurrentThread.CurrentCulture as required BEFORE creating the instance to ensure the engine uses the correct culture later on)
        ''' </summary>
        ''' <returns></returns>
        Protected Function CreateInstanceUninitialized() As T
            Try
                Return _CreateInstanceUninitialized()
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
        ''' <param name="data"></param>
        ''' <param name="passwordForOpening"></param>
        ''' <param name="disableCalculationEngine"></param>
        ''' <returns></returns>
        Protected Function CreateInstance(data As Byte(), options As ExcelOps.ExcelDataOperationsOptions) As T
            Try
                Return _CreateInstance(data, options)
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
        ''' <param name="data"></param>
        ''' <param name="passwordForOpening"></param>
        ''' <param name="disableCalculationEngine"></param>
        ''' <returns></returns>
        Protected Function CreateInstance(data As System.IO.Stream, options As ExcelOps.ExcelDataOperationsOptions) As T
            Try
                Return _CreateInstance(data, options)
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
        ''' <param name="disableCalculationEngine">True to disable calculation engine, e.g. sometimes required with some excel engines to load excel workbooks with circular reference errors successfully</param>
        ''' <returns></returns>
        Protected Function CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, options As ExcelOps.ExcelDataOperationsOptions) As T
            Try
                Return _CreateInstance(file, mode, options)
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
                Assert.NotNull(Me.CreateInstance(Nothing, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)))
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
            'Dispose/Finalize/Release COM objects
            CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
            'Show test details on failure (only)
            If TestContext.CurrentContext.Result.Outcome = NUnit.Framework.Interfaces.ResultState.Failure Then
                Dim Log As String = Test.Console.GetConsoleLog
                System.Console.WriteLine(Log)
            End If
        End Sub

        Public MustOverride ReadOnly Property ExpectedEngineName As String

        <Test> Public Sub EngineName()
            Assert.AreEqual(ExpectedEngineName, Me.CreateInstanceUninitialized().EngineName)
        End Sub

        <Test> Public Sub HasVbaProject()
            Dim VbaTestFile = TestEnvironment.FullPathOfExistingTestFile("test_data", "VbaProject.xlsm")
            Assert.IsTrue(Me.CreateInstance(VbaTestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)).HasVbaProject)

            Dim NonVbaTestFile = TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund01.xlsx")
            Assert.IsFalse(Me.CreateInstance(NonVbaTestFile, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)).HasVbaProject)
        End Sub

        <Test> Public Sub SaveXlsxWithVbaProjectMustFail()
            Dim Wb As T

            'VBA test file must fail to save until VBA project has been removed
            Dim VbaTestFile = TestEnvironment.FullPathOfExistingTestFile("test_data", "VbaProject.xlsm")
            Dim VbaTestFileClone = TestEnvironment.FullPathOfDynTestFile(NameOf(SaveXlsxWithVbaProjectMustFail), "VbaProject.xlsm")
            Dim NewXlsxTargetPath As String = TestEnvironment.FullPathOfDynTestFile(NameOf(SaveXlsxWithVbaProjectMustFail), "VbaProject.xlsx")
            System.IO.File.Copy(VbaTestFile, VbaTestFileClone)

            Wb = Me.CreateInstance(VbaTestFileClone, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
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

            Wb = Me.CreateInstance(VbaTestFileClone, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite))
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

            Wb = Me.CreateInstance(NonVbaTestFileClone, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.False(Wb.HasVbaProject)
            Wb.SaveAs(NewXlsxTargetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Wb.Close()

            Wb = Me.CreateInstance(NonVbaTestFileClone, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite))
            Assert.False(Wb.HasVbaProject)
            Wb.Save()
            Wb.Close()

            'Loading a workbook with VBA project + removing VBA project + saving workbook as XLSM + reloading workbook = must still HasVbaProject = False
            VbaTestFile = TestEnvironment.FullPathOfExistingTestFile("test_data", "VbaProject.xlsm")
            NewXlsxTargetPath = TestEnvironment.FullPathOfDynTestFile(NameOf(SaveXlsxWithVbaProjectMustFail), "VbaProject.xlsm")
            Wb = Me.CreateInstance(VbaTestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.True(Wb.HasVbaProject)
            Wb.RemoveVbaProject()
            Assert.False(Wb.HasVbaProject)
            Wb.SaveAs(NewXlsxTargetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Wb.Close()
            Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.False(Wb.HasVbaProject)

        End Sub

        <Test> Public Sub LoadFileFromByteArray()
            Dim Wb As T
            'Testfile without password
            Dim TestFile As String = TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund01.xlsx")
            Dim Data As Byte() = System.IO.File.ReadAllBytes(TestFile)
            If GetType(T) Is GetType(MsExcelDataOperations) Then
                'known to fail because no support for reading files from in-memory
                Assert.Throws(Of NotSupportedException)(Sub()
                                                            Me.CreateInstance(Data, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                                                        End Sub)
            Else
                Wb = Me.CreateInstance(Data, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Assert.AreEqual("Grunddaten", Wb.SheetNames(0))
                Assert.That(Wb.ReadOnly, [Is].True)
            End If
        End Sub

        <Test> Public Sub LoadFileFromStream()
            Dim Wb As T
            'Testfile without password
            Dim TestFile As String = TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund01.xlsx")
            Dim Data As System.IO.Stream = New System.IO.FileStream(TestFile, System.IO.FileMode.Open)
            If GetType(T) Is GetType(MsExcelDataOperations) Then
                'known to fail because no support for reading files from in-memory
                Assert.Throws(Of NotSupportedException)(Sub()
                                                            Me.CreateInstance(Data, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                                                        End Sub)
            Else
                Wb = Me.CreateInstance(Data, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Assert.AreEqual("Grunddaten", Wb.SheetNames(0))
                Assert.That(Wb.ReadOnly, [Is].True)
            End If
        End Sub

        <Test> Public Sub PasswordForOpening()
            Dim Wb As T
            'Testfile without password
            Dim TestFile As String = TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund01.xlsx")
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.AreEqual("Grunddaten", Wb.SheetNames(0))

            'Now, save it with password
            Wb.PasswordForOpening = "dummy"
            Dim NewXlsxTargetPath As String = TestEnvironment.FullPathOfDynTestFile(Wb, "PasswordProtectedFile.xlsx")
            Wb.SaveAs(NewXlsxTargetPath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            'Console.WriteLine("Saved password protected file to: " & NewXlsxTargetPath)
            Wb.Close()

            'Try to reload it without password -> it must fail
            Assert.Catch(Of Exception)(Sub() Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly, "something else")))
            Assert.Catch(Of Exception)(Sub() Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly, "")))
            Assert.Catch(Of Exception)(Sub() Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly, Nothing)))

            'Reload it with password -> now it must succeed
            Wb = Me.CreateInstance(NewXlsxTargetPath, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly, "dummy"))
            Assert.AreEqual("Grunddaten", Wb.SheetNames(0))
        End Sub

        <Test> Public Sub CreateWorkbookWithoutFilePath()
            Dim Wb As T
            Dim TestFile As String

            TestFile = Nothing
            Wb = Me.CreateInstanceUninitialized()
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)
            Wb.Close()

            TestFile = Nothing
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite))
            Assert.AreEqual(TestFile, Wb.FilePath)
            Assert.AreEqual(TestFile, Wb.WorkbookFilePath)
            Wb.Close()

            TestFile = ""
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite))
            Assert.AreEqual(Nothing, Wb.FilePath)
            Assert.AreEqual(Nothing, Wb.WorkbookFilePath)
            Wb.Close()
        End Sub

        <Test> Public Sub CreateAndSaveAsAndFilePath()
            Dim Wb As T
            Dim TestFile As String = TestEnvironment.FullPathOfDynTestFile(Me.CreateInstanceUninitialized, "created-workbook.xlsx")
            Dim TestFile2 As String = TestEnvironment.FullPathOfDynTestFile(Me.CreateInstanceUninitialized, "created-workbook2.xlsx")

            'Creating a new workbook without pre-defined file name must fail on Save(), but successful on SaveAs()
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.AreEqual(True, Wb.ReadOnly, "Newly created files must be ReadOnly if file path hasn't been set up, but ReadWrite if file path has been set up")
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite))
            Assert.AreEqual(False, Wb.ReadOnly, "Newly created files must be ReadOnly if file path hasn't been set up, but ReadWrite if file path has been set up")
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile))
            Assert.AreEqual(TestFile = Nothing, Wb.ReadOnly, "Newly created files must be ReadOnly if file path hasn't been set up, but ReadWrite if file path has been set up")
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
                                                             Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite))
                                                         End Sub)
            System.IO.File.Delete(TestFile) 'Delete the file for next test block

            'Creating a new workbook must always be ReadOnly and saving it without a name must be forbidden
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite))
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
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
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
            Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite))
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

            workbook = Me.CreateInstanceUninitialized()
            Select Case workbook.GetType
                Case GetType(MsExcelDataOperations)
                    'Accept fact that a new workbook is opened automatically
                    'Assert.NotZero(workbook.SheetNames.Count)
                    Assert.Throws(Of System.NullReferenceException)(Function() workbook.SheetNames.Count)
                Case Else
                    'No workbook opened - must be done in 2ndary step
                    Assert.Throws(Of InvalidOperationException)(Function() workbook.SheetNames.Count)
            End Select
            workbook.Close()

            workbook = Me.CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.AreEqual(1, workbook.SheetNames.Count, "Sheets Count")
            Assert.AreEqual("Sheet1", workbook.SheetNames(0))
            workbook.Close()

        End Sub

        <Test> Public Overridable Sub CopySheetContent()
            Dim eppeoIn As ExcelOps.ExcelDataOperationsBase = Nothing
            Dim eppeoOut As ExcelOps.ExcelDataOperationsBase = Nothing
            Dim TestControllingToolFileNameIn As String
            Dim TestControllingToolFileNameOutTemplate As String
            Dim TestControllingToolFileNameOut As String

            TestControllingToolFileNameIn = TestFiles.TestFileGrund01.FullName
            TestControllingToolFileNameOutTemplate = TestFiles.TestFileGrund02.FullName
            TestControllingToolFileNameOut = TestEnvironment.FullPathOfDynTestFile(Me.CreateInstanceUninitialized, "CopySheetContent_" & GetType(T).Name & ".xlsx")
            Try
                Console.WriteLine("Test file in: " & TestControllingToolFileNameIn)
                Console.WriteLine("Test file output template: " & TestControllingToolFileNameOutTemplate)
                Console.WriteLine("Test file output: " & TestControllingToolFileNameOut)

                eppeoIn = Me.CreateInstance(TestControllingToolFileNameIn, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                eppeoOut = Me.CreateInstance(TestControllingToolFileNameOutTemplate, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

                Const SheetToCopy As String = "Grunddaten"
                eppeoIn.CopySheetContent(SheetToCopy, eppeoOut, ExcelOps.ExcelDataOperationsBase.CopySheetOption.TargetSheetMightExist)
                eppeoOut.SelectSheet(SheetToCopy)
                eppeoOut.SaveAs(TestControllingToolFileNameOut, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.AreEqual(eppeoIn.SheetContentMatrix(SheetToCopy, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText), eppeoOut.SheetContentMatrix(SheetToCopy, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText))
                Assert.Pass("Required manual, optical review for comparison to check for formattings")
            Finally
                If eppeoIn IsNot Nothing Then eppeoIn.Close()
                If eppeoOut IsNot Nothing Then eppeoOut.Close()
            End Try
        End Sub

        <Test> Public Sub ExcelOpsTestCollection_ZahlenUndProzentwerte()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestFiles.TestFileExcelOpsTestCollection.FullName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
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

        <Test> Public Sub IsMergedCell()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileMergedCells.FullName
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Const SheetName As String = "MergedCellsTest"
            Assert.True(eppeo.IsMergedCell(SheetName, 0, 0))
            Assert.True(eppeo.IsMergedCell(SheetName, 1, 0))
            Assert.True(eppeo.IsMergedCell(SheetName, 0, 1))
            Assert.True(eppeo.IsMergedCell(SheetName, 2, 0))
            Assert.True(eppeo.IsMergedCell(SheetName, 0, 2))
            Assert.True(eppeo.IsMergedCell(SheetName, 2, 2))
            Assert.False(eppeo.IsMergedCell(SheetName, 3, 0))
            Assert.False(eppeo.IsMergedCell(SheetName, 0, 3))
        End Sub

        <Test> Public Sub MergeAndUnMergedCell()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileMergedCells.FullName
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Const SheetName As String = "MergedCellsTest"
            Assert.True(eppeo.IsMergedCell(SheetName, 0, 0))
            Assert.True(eppeo.IsMergedCell(SheetName, 2, 2))
            eppeo.UnMergeCells(SheetName, 1, 1)
            Assert.False(eppeo.IsMergedCell(SheetName, 0, 0))
            Assert.False(eppeo.IsMergedCell(SheetName, 2, 2))
            eppeo.MergeCells(SheetName, 0, 0, 2, 2)
            Assert.True(eppeo.IsMergedCell(SheetName, 0, 0))
            Assert.True(eppeo.IsMergedCell(SheetName, 2, 2))
        End Sub

        <Test> Public Sub MergedCells()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileMergedCells.FullName
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Dim SheetName As String = "MergedCellsTest"
            Assert.AreEqual("A1:C3", String.Join(";"c, eppeo.MergedCells(SheetName).Select(Of String)(Function(x) x.LocalAddress)))

            TestControllingToolFileName = TestFiles.TestFileGrund01.FullName
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            SheetName = "Grunddaten"
            Assert.AreEqual("", String.Join(";"c, eppeo.MergedCells(SheetName).Select(Of String)(Function(x) x.LocalAddress)))
        End Sub

        <Test> Public Sub AutoFitColumns()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileMergedCells.FullName
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Const SheetName As String = "MergedCellsTest"
            Try
                Assert.AreEqual(2, eppeo.LookupLastCell(SheetName).ColumnIndex)
                eppeo.AutoFitColumns(SheetName, 0, 250.0)
                eppeo.AutoFitColumns(SheetName, 0)
                eppeo.AutoFitColumns(SheetName, 80.0)
                eppeo.AutoFitColumns(SheetName)
            Catch ex As PlatformNotSupportedException
                'System.Drawing.Common is not supported on platform
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As System.TypeInitializationException
                'The type initializer for 'Gdip' threw an exception.
                '---> System.PlatformNotSupportedException System.Drawing.Common Is Not supported on non-Windows platforms. See https://aka.ms/systemdrawingnonwindows for more information.
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            End Try
        End Sub

        <Test> Public Sub SheetNames()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim EppeoSheetNamesList As List(Of String)

            Dim TestFileName As String
            TestFileName = TestFiles.TestFileGrund01.FullName
            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            EppeoSheetNamesList = eppeo.SheetNames
            System.Console.WriteLine("## " & System.IO.Path.GetFileName(TestFileName))
            System.Console.WriteLine(Strings.Join(EppeoSheetNamesList.ToArray, ","))
            Assert.AreEqual("Grunddaten,Kostenplanung", Strings.Join(EppeoSheetNamesList.ToArray, ","))

            TestFileName = TestFiles.TestFileChartSheet01.FullName
            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            EppeoSheetNamesList = eppeo.SheetNames
            System.Console.WriteLine()
            System.Console.WriteLine("## " & System.IO.Path.GetFileName(TestFileName))
            System.Console.WriteLine(Strings.Join(EppeoSheetNamesList.ToArray, ","))
            Assert.AreEqual("data,chart", Strings.Join(EppeoSheetNamesList.ToArray, ","))
        End Sub

        <Test> Public Sub SheetIndex()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestFileName As String
            TestFileName = TestFiles.TestFileChartSheet01.FullName
            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.AreEqual(0, eppeo.SheetIndex("data"))
            Assert.AreEqual(1, eppeo.SheetIndex("chart"))
            Assert.AreEqual(-1, eppeo.SheetIndex("doesntexist"))
        End Sub

        <Test> Public Sub WorkSheetNames()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim EppeoSheetNamesList As List(Of String)

            Dim TestFileName As String
            TestFileName = TestFiles.TestFileChartSheet01.FullName
            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            EppeoSheetNamesList = eppeo.WorkSheetNames
            System.Console.WriteLine("## " & System.IO.Path.GetFileName(TestFileName))
            System.Console.WriteLine(Strings.Join(EppeoSheetNamesList.ToArray, ","))
            Assert.AreEqual("data", Strings.Join(EppeoSheetNamesList.ToArray, ","))
        End Sub

        <Test> Public Sub WorkSheetIndex()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestFileName As String
            TestFileName = TestFiles.TestFileChartSheet01.FullName
            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.AreEqual(0, eppeo.WorkSheetIndex("data"))
            Assert.AreEqual(-1, eppeo.WorkSheetIndex("chart"))
        End Sub

        <Test> Public Sub SelectedSheetName()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase

            Dim TestFileName As String
            TestFileName = TestFiles.TestFileGrund03.FullName
            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

            Assert.That(eppeo.SelectedSheetName, [Is].EqualTo("Ausgewählt"))

            If eppeo.EngineName = "Spire.Xls" Then
                'known to insert sheet "Evaluation Warning" --> do not test
                Assert.Ignore("Testing must fail because of added sheet ""Evaluation Warning""")
            End If

            eppeo.SelectSheet(0)
            eppeo.SaveAs(TestEnvironment.FullPathOfDynTestFile(eppeo, "SelectedSheet.Grunddaten.0.xlsx"), ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            System.Console.WriteLine("OUT: " & eppeo.FilePath)
            Assert.That(eppeo.SelectedSheetName, [Is].EqualTo("Grunddaten"))

            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            eppeo.SelectSheet(1)
            eppeo.SaveAs(TestEnvironment.FullPathOfDynTestFile(eppeo, "SelectedSheet.Ausgewählt.1.xlsx"), ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            System.Console.WriteLine("OUT: " & eppeo.FilePath)
            Assert.That(eppeo.SelectedSheetName, [Is].EqualTo("Ausgewählt"))

            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            eppeo.SelectSheet(2)
            eppeo.SaveAs(TestEnvironment.FullPathOfDynTestFile(eppeo, "SelectedSheet.Kostenplanung.2.xlsx"), ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            System.Console.WriteLine("OUT: " & eppeo.FilePath)
            Assert.That(eppeo.SelectedSheetName, [Is].EqualTo("Kostenplanung"))

            eppeo.SelectSheet("Grunddaten")
            eppeo.SaveAs(TestEnvironment.FullPathOfDynTestFile(eppeo, "SelectedSheet.Grunddaten.xlsx"), ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            System.Console.WriteLine("OUT: " & eppeo.FilePath)
            Assert.That(eppeo.SelectedSheetName, [Is].EqualTo("Grunddaten"))

            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            eppeo.SelectSheet("Kostenplanung")
            eppeo.SaveAs(TestEnvironment.FullPathOfDynTestFile(eppeo, "SelectedSheet.Kostenplanung.xlsx"), ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            System.Console.WriteLine("OUT: " & eppeo.FilePath)
            Assert.That(eppeo.SelectedSheetName, [Is].EqualTo("Kostenplanung"))

        End Sub

        <Test> Public Sub SelectedSheetName2()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase

            Dim TestFileName As String
            TestFileName = TestFiles.TestFileHtmlExport01.FullName
            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

            Assert.That(eppeo.SelectedSheetName, [Is].EqualTo("Onboarding offen"))
            eppeo.SelectSheet("Upload erledigt")
            Assert.That(eppeo.SelectedSheetName, [Is].EqualTo("Upload erledigt"))
        End Sub

        <Test> Public Sub ChartSheetNames()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim EppeoSheetNamesList As List(Of String)

            Dim TestFileName As String
            TestFileName = TestFiles.TestFileChartSheet01.FullName
            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            EppeoSheetNamesList = eppeo.ChartSheetNames
            System.Console.WriteLine("## " & System.IO.Path.GetFileName(TestFileName))
            System.Console.WriteLine(Strings.Join(EppeoSheetNamesList.ToArray, ","))
            Assert.AreEqual("chart", Strings.Join(EppeoSheetNamesList.ToArray, ","))
        End Sub

        <Test> Public Sub ChartSheetIndex()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestFileName As String
            TestFileName = TestFiles.TestFileChartSheet01.FullName
            eppeo = Me.CreateInstance(TestFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.AreEqual(0, eppeo.ChartSheetIndex("chart"))
            Assert.AreEqual(-1, eppeo.ChartSheetIndex("data"))
        End Sub

        <Test> Public Sub AddSheet()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            Dim BeforeSheet As String = "Grunddaten"
            Dim SheetNameTopPosition As String = "SheetOnTop"
            Dim SheetNameBottomPosition As String = "SheetOnBottom"
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
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
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
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
            Dim wb As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestFiles.TestFileGrund02.FullName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
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

            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.AreEqual("E40", eppeo.LookupLastCell(TestSheet).Address)
            Assert.AreEqual(39, eppeo.LookupLastRowIndex(TestSheet))
            Assert.AreEqual(4, eppeo.LookupLastColumnIndex(TestSheet))

            eppeo.WriteCellValue(Of String)(New ExcelCell(TestSheet, "J50", ExcelCell.ValueTypes.All), "Content! :-)")
            Assert.AreEqual("J50", eppeo.LookupLastCell(TestSheet).Address)
            Assert.AreEqual(49, eppeo.LookupLastRowIndex(TestSheet))
            Assert.AreEqual(9, eppeo.LookupLastColumnIndex(TestSheet))

            TestControllingToolFileName = TestFiles.TestFileMergedCells.FullName
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            TestSheet = "MergedCellsTest"
            Assert.AreEqual("C3", eppeo.LookupLastCell(TestSheet).Address)
            Assert.AreEqual(2, eppeo.LookupLastRowIndex(TestSheet))
            Assert.AreEqual(2, eppeo.LookupLastColumnIndex(TestSheet))
        End Sub

        <Test> Public Sub LookupLastContentCellAddress()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.AreEqual("E40", eppeo.LookupLastContentCell(TestSheet).Address)
            Assert.AreEqual(39, eppeo.LookupLastContentRowIndex(TestSheet))
            Assert.AreEqual(4, eppeo.LookupLastContentColumnIndex(TestSheet))

            eppeo.WriteCellValue(Of String)(New ExcelCell(TestSheet, "I49", ExcelCell.ValueTypes.All), "content :-)")
            eppeo.WriteCellValue(Of String)(New ExcelCell(TestSheet, "J50", ExcelCell.ValueTypes.All), "")
            Assert.AreEqual("I49", eppeo.LookupLastContentCell(TestSheet).Address)
            Assert.AreEqual(48, eppeo.LookupLastContentRowIndex(TestSheet))
            Assert.AreEqual(8, eppeo.LookupLastContentColumnIndex(TestSheet))

            TestControllingToolFileName = TestFiles.TestFileMergedCells.FullName
            eppeo = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            TestSheet = "MergedCellsTest"
            Assert.AreEqual("C3", eppeo.LookupLastContentCell(TestSheet).Address)
            Assert.AreEqual(2, eppeo.LookupLastContentRowIndex(TestSheet))
            Assert.AreEqual(2, eppeo.LookupLastContentColumnIndex(TestSheet))
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

        Private Function ExpectedResultInCultureContextAndPlateformLineBreakEncoding(expectedRawMatrix As String) As String
            Return expectedRawMatrix.
                Replace(PlaceHolderDecimalSeparator, System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator).
                Replace(PlaceHolderGroupSeparator, System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator).
                Replace(ControlChars.CrLf, System.Environment.NewLine)
        End Function

        <Test> Public Overridable Sub SheetContentMatrix_StaticOrCalculatedValues(<Values("invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

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
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues, ExpectedResultInCultureContextAndPlateformLineBreakEncoding(ExpectedMatrix))
                End Sub)
        End Sub

        <Test> Public Overridable Sub SheetContentMatrix_StaticValues(<Values("invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

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
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticValues, ExpectedResultInCultureContextAndPlateformLineBreakEncoding(ExpectedMatrix))
                End Sub)
        End Sub

        <Test> Public Overridable Sub SheetContentMatrix_Formulas(<Values("invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

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
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.Formulas, ExpectedResultInCultureContextAndPlateformLineBreakEncoding(ExpectedMatrix))
                End Sub)
        End Sub

        <Test> Public Overridable Sub SheetContentMatrix_FormattedText(<Values("en-US", "invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

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
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormattedText, ExpectedResultInCultureContextAndPlateformLineBreakEncoding(ExpectedMatrix))
                End Sub)
        End Sub

        <Test> Public Overridable Sub SheetContentMatrix_FormulaOrFormattedText(<Values("invariant", "de-DE")> cultureName As String)
            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

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
                    Me.AssertSheetContentMatrix(eppeo, TestSheet, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText, ExpectedResultInCultureContextAndPlateformLineBreakEncoding(ExpectedMatrix))
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
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

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
                    Assert.AreEqual(ExpectedResultInCultureContextAndPlateformLineBreakEncoding("Chef: 14▲09"), eppeo.LookupCellFormattedText(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))
                    Assert.AreEqual(Nothing, eppeo.LookupCellFormula(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))

                    'E1
                    Assert.AreEqual(False, eppeo.LookupCellValue(Of Boolean)(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                    Assert.AreEqual("False", eppeo.LookupCellValue(Of String)(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                    Assert.AreEqual("False", eppeo.LookupCellFormattedText(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                End Sub)
        End Sub

#Region "ExcelCharting"
        Private Function PrepareAndFillExcelFileWithChart(variantOfImage As Byte) As ExcelOps.ExcelDataOperationsBase
            Dim ExcelFile As String = TestEnvironment.FullPathOfExistingTestFile(TestFiles.TestFileChartSheet01.FullName)
            Dim Workbook = Me.CreateInstance(ExcelFile, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Select Case variantOfImage
                Case 0 'master
                    Workbook.WriteCellValue(Of String)(New ExcelCell("data", "B1", ExcelCell.ValueTypes.All), "Sample Chart")
                    Workbook.WriteCellValue(Of String)(New ExcelCell("data", "B2", ExcelCell.ValueTypes.All), "Sub title")
                    Workbook.WriteCellValue(Of Integer)(New ExcelCell("data", "B3", ExcelCell.ValueTypes.All), 2022)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "B6", ExcelCell.ValueTypes.All), 1234.56D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "C6", ExcelCell.ValueTypes.All), 2345.67D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "D6", ExcelCell.ValueTypes.All), 3456.78D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "F6", ExcelCell.ValueTypes.All), 4000)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "G6", ExcelCell.ValueTypes.All), 4444.44D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "I6", ExcelCell.ValueTypes.All), 5000D)
                Case 1 'difference at just 1 letter (year number)
                    Workbook.WriteCellValue(Of String)(New ExcelCell("data", "B1", ExcelCell.ValueTypes.All), "Sample Chart")
                    Workbook.WriteCellValue(Of String)(New ExcelCell("data", "B2", ExcelCell.ValueTypes.All), "Sub title")
                    Workbook.WriteCellValue(Of Integer)(New ExcelCell("data", "B3", ExcelCell.ValueTypes.All), 2023) '<-- here is the difference!
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "B6", ExcelCell.ValueTypes.All), 1234.56D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "C6", ExcelCell.ValueTypes.All), 2345.67D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "D6", ExcelCell.ValueTypes.All), 3456.78D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "F6", ExcelCell.ValueTypes.All), 4000)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "G6", ExcelCell.ValueTypes.All), 4444.44D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "I6", ExcelCell.ValueTypes.All), 5000D)
                Case 2 'stronger difference
                    Workbook.WriteCellValue(Of String)(New ExcelCell("data", "B1", ExcelCell.ValueTypes.All), "Sample Chart")
                    Workbook.WriteCellValue(Of String)(New ExcelCell("data", "B2", ExcelCell.ValueTypes.All), "Sub title")
                    Workbook.WriteCellValue(Of Integer)(New ExcelCell("data", "B3", ExcelCell.ValueTypes.All), 2023) '<-- here is the difference!
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "B6", ExcelCell.ValueTypes.All), 934.56D) '<-- here is the difference!
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "C6", ExcelCell.ValueTypes.All), 2345.67D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "D6", ExcelCell.ValueTypes.All), 3456.78D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "F6", ExcelCell.ValueTypes.All), 4300) '<-- here is the difference!
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "G6", ExcelCell.ValueTypes.All), 4444.44D)
                    Workbook.WriteCellValue(Of Decimal)(New ExcelCell("data", "I6", ExcelCell.ValueTypes.All), 5000D)
                Case Else
                    Throw New NotImplementedException
            End Select
#Disable Warning CA1862 ' "StringComparison"-Methodenüberladungen verwenden, um Zeichenfolgenvergleiche ohne Beachtung der Groß-/Kleinschreibung durchzuführen
            If Workbook.CalculationModuleDisabled AndAlso Workbook.EngineName.ToLowerInvariant.Contains("epplus") Then
#Enable Warning CA1862 ' "StringComparison"-Methodenüberladungen verwenden, um Zeichenfolgenvergleiche ohne Beachtung der Groß-/Kleinschreibung durchzuführen
                Workbook.CalculationModuleDisabled = False 'for this test, calculcation module can be re-enabled
            End If
            Workbook.RecalculateAll()
            Return Workbook
        End Function

        <Test>
        Public Sub ExportChartSheetImage()
            Try
                Dim TempFilePng As String = TestEnvironment.FullPathOfDynTestFile(Me.CreateInstanceUninitialized, "excel_test_chart.png")
                Dim MasterImg As System.Drawing.Image
#Disable Warning CA1416
                MasterImg = System.Drawing.Image.FromFile(TestEnvironment.FullPathOfExistingTestFile("test_comparison_masters", "excel_test_chart.png"))

                Dim Workbook = PrepareAndFillExcelFileWithChart(0)
                Assert.AreEqual(2, Workbook.SheetNames.Count) '"data","chart"
                Dim ChartImage = Workbook.ExportChartSheetImage("chart")

                ChartImage.Save(TempFilePng, System.Drawing.Imaging.ImageFormat.Png)
                System.Console.WriteLine("PNG saved to " & TempFilePng)
                'If Debugger.IsAttached Then IOTools.LaunchFileWithAssociatedApp(TempFilePng)
#Enable Warning CA1416

                If Workbook.EngineName = "Spire.Xls" Then
                    TestImageComparison.AssertImagesAreEqual(MasterImg, ChartImage, 0.023) 'accept difference if it's just the evaluation note in image
                ElseIf Workbook.EngineName = "FreeSpire.Xls" Then
                    TestImageComparison.AssertImagesAreEqual(MasterImg, ChartImage, 0.02) 'accept difference of slightly re-located Y axis descriptions
                Else
                    TestImageComparison.AssertImagesAreEqual(MasterImg, ChartImage)
                End If

                'Now, do the negative test: image must be detected as "is different"
                Dim DifferentChartImage As System.Drawing.Image
                DifferentChartImage = PrepareAndFillExcelFileWithChart(2).ExportChartSheetImage("chart") 'stronger difference
                Assert.Throws(Of NUnit.Framework.AssertionException)(Sub() TestImageComparison.AssertImagesAreEqual(MasterImg, DifferentChartImage))

                DifferentChartImage = PrepareAndFillExcelFileWithChart(1).ExportChartSheetImage("chart") 'minor difference of just 1 letter 
                Assert.Throws(Of NUnit.Framework.AssertionException)(Sub() TestImageComparison.AssertImagesAreEqual(MasterImg, DifferentChartImage))
            Catch ex As PlatformNotSupportedException
                'System.Drawing.Common is not supported on platform
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As System.TypeInitializationException
                'The type initializer for 'Gdip' threw an exception.
                '---> System.PlatformNotSupportedException System.Drawing.Common Is Not supported on non-Windows platforms. See https://aka.ms/systemdrawingnonwindows for more information.
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As NotSupportedException
                Assert.Ignore("EngineNotSupported: " & ex.Message)
            Catch ex As NotImplementedException
                Assert.Ignore("EngineImplementation missing: " & ex.Message)
            End Try
        End Sub
#End Region

#Region "Excel error values in cells"
        <Test> Public Sub LookupErrorCellValue(<Values("invariant", "de-DE", "en-US")> cultureName As String)
            Dim TestControllingToolFileName As String = TestFiles.TestFileExcelOpsErrorValues.FullName
            Dim TestSheet As String = "Tabelle1"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

                    Console.WriteLine(eppeo.SheetContentMatrix(TestSheet, ExcelDataOperationsBase.MatrixContent.Errors).ToUIExcelTable)
                    '## Expected matrix like following
                    '# |A      |B     |C    |D    |E    |F      
                    '--+-------+------+-----+-----+-----+-------
                    '1 |#DIV/0!|#NAME?|#REF!|#REF!|#NUM!|#VALUE!

                    Assert.AreEqual("#DIV/0!", eppeo.LookupCellErrorValue(New ExcelOps.ExcelCell(TestSheet, "A1", ExcelOps.ExcelCell.ValueTypes.All)))
                    Assert.AreEqual("#NAME?", eppeo.LookupCellErrorValue(New ExcelOps.ExcelCell(TestSheet, "B1", ExcelOps.ExcelCell.ValueTypes.All)))
                    Assert.AreEqual("#REF!", eppeo.LookupCellErrorValue(New ExcelOps.ExcelCell(TestSheet, "C1", ExcelOps.ExcelCell.ValueTypes.All)))
                    Assert.AreEqual("#VALUE!", eppeo.LookupCellErrorValue(New ExcelOps.ExcelCell(TestSheet, "F1", ExcelOps.ExcelCell.ValueTypes.All)))

                    'NOTE: Known expected behaviour for #NUM! error value is differently between the several engines
                    Select Case eppeo.EngineName
                        Case "Spire.Xls", "FreeSpire.Xls", "Epplus (Polyform license edition)"
                            Assert.AreEqual(Nothing, eppeo.LookupCellErrorValue(New ExcelOps.ExcelCell(TestSheet, "D1", ExcelOps.ExcelCell.ValueTypes.All)))
                            Assert.Ignore(eppeo.EngineName & " is not fully compatible and doesn't show up with #REF! error value in cell D1 (cell reference to a removed sheet (by EPPlus lib, which doesn't update formulas with #REF! on sheet removal)")
                            'Assert.AreEqual(Nothing, eppeo.LookupCellErrorValue(New ExcelOps.ExcelCell(TestSheet, "E1", ExcelOps.ExcelCell.ValueTypes.All)))
                            'Assert.Ignore(eppeo.EngineName & " is not fully compatible and doesn't show up with #NUM! error value in cell E1 (=SQRT(-1) alias =WURZEL(-1))")
                        Case Else
                            Assert.AreEqual("#REF!", eppeo.LookupCellErrorValue(New ExcelOps.ExcelCell(TestSheet, "D1", ExcelOps.ExcelCell.ValueTypes.All)))
                            Assert.AreEqual("#NUM!", eppeo.LookupCellErrorValue(New ExcelOps.ExcelCell(TestSheet, "E1", ExcelOps.ExcelCell.ValueTypes.All)))
                    End Select

                End Sub)
        End Sub

        <Test> Public Sub FindErrorCellsInWorkbook_LookupLastCell()
            Dim TestControllingToolFileName As String = TestFiles.TestFileExcelOpsErrorValues.FullName
            Dim TestSheet As String = "Tabelle1"
            System.Console.WriteLine("Testing XLSX: " & TestControllingToolFileName)
            Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Assert.AreEqual("F1", eppeo.LookupLastCell(TestSheet).Address)
        End Sub

        <Test> Public Sub FindErrorCellsInWorkbook(<Values("invariant", "de-DE", "en-US")> cultureName As String)
            Dim TestControllingToolFileName As String = TestFiles.TestFileExcelOpsErrorValues.FullName
            Dim TestSheet As String = "Tabelle1"

            TestInCultureContext(
                cultureName,
                Sub()
                    Dim eppeo As ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestControllingToolFileName, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))

                    Assert.AreEqual("F1", eppeo.LookupLastCell(TestSheet).Address)
                    Console.WriteLine(eppeo.SheetContentMatrix(TestSheet, ExcelDataOperationsBase.MatrixContent.Errors).ToUIExcelTable)
                    '## Expected matrix like following
                    '# |A      |B     |C    |D    |E    |F      
                    '--+-------+------+-----+-----+-----+-------
                    '1 |#DIV/0!|#NAME?|#REF!|#REF!|#NUM!|#VALUE!

                    Assert.AreEqual(1, eppeo.FindErrorCellsInWorkbook("#DIV/0!").Count)
                    Assert.AreEqual(1, eppeo.FindErrorCellsInWorkbook("#NAME?").Count)
                    Assert.AreEqual(1, eppeo.FindErrorCellsInWorkbook("#VALUE!").Count)

                    'NOTE: Known expected behaviour for #NUM! error value is differently between the several engines
                    Select Case eppeo.EngineName
                        Case "Spire.Xls", "FreeSpire.Xls"
                            Assert.AreEqual(Nothing, eppeo.LookupCellErrorValue(TestSheet, 0, 3), "D1 >> #REF") 'Error detection not working (correctly) in engine at D1
                            Assert.AreEqual("#NUM!", eppeo.LookupCellErrorValue(TestSheet, 0, 4), "E1 >> #NUM") 'Error detection not working (correctly) in engine at E1
                            Assert.AreEqual(1, eppeo.FindErrorCellsInWorkbook("#NUM!").Count)
                            Assert.AreEqual(1, eppeo.FindErrorCellsInWorkbook("#REF!").Count) 'Error detection not working (correctly) in engine at D1
                            Assert.AreEqual(5, eppeo.FindErrorCellsInWorkbook().Count)
                            Assert.Ignore(eppeo.EngineName & " is not fully compatible and doesn't show up with #NUM! error value in cell E1 (=SQRT(-1) alias =WURZEL(-1))")
                        Case "Epplus (Polyform license edition)"
                            Assert.AreEqual(Nothing, eppeo.LookupCellErrorValue(TestSheet, 0, 3), "D1 >> #REF") 'Error detection not working (correctly) in engine at D1
                            Assert.AreEqual(1, eppeo.FindErrorCellsInWorkbook("#NUM!").Count)
                            Assert.AreEqual(1, eppeo.FindErrorCellsInWorkbook("#REF!").Count)
                            Assert.AreEqual(5, eppeo.FindErrorCellsInWorkbook().Count)
                            Assert.Ignore(eppeo.EngineName & " is not fully compatible and doesn't show up with #REF! error value in cell D1 (cell reference to a removed sheet (by EPPlus lib, which doesn't update formulas with #REF! on sheet removal)")
                        Case Else
                            Assert.AreEqual(1, eppeo.FindErrorCellsInWorkbook("#NUM!").Count)
                            Assert.AreEqual(2, eppeo.FindErrorCellsInWorkbook("#REF!").Count)
                            Assert.AreEqual(6, eppeo.FindErrorCellsInWorkbook().Count)
                    End Select

                    If eppeo.CalculationModuleDisabled = False Then
                        eppeo.RecalculateAll()
                        eppeo.RecalculateCell(TestSheet, 0, 3)
                        Console.WriteLine("Values after explicit recalculation")
                        Console.WriteLine(eppeo.SheetContentMatrix(TestSheet, ExcelDataOperationsBase.MatrixContent.Errors).ToUIExcelTable)
                    End If
                End Sub)
        End Sub


#End Region

#Region "Excel workbooks with embedded pictures might fail to load/save on different platforms (due to System.Drawing not being accessible, e.g. on Linux with .Net 7"
        ''' <summary>
        ''' Embedded "picture": a chart in diagram sheet
        ''' </summary>
        <Test>
        Public Sub TestFileWithEmbeddedPicture01()
            Try
                Dim ExcelFile As String = TestEnvironment.FullPathOfExistingTestFile(TestFiles.TestFileEmbeddedPicture01.FullName)
                Dim Workbook = Me.CreateInstance(ExcelFile, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Dim OutputFile As String = TestEnvironment.FullPathOfDynTestFile(Workbook.GetType, "test_embeddedpicture01.xlsx")
                Workbook.SaveAs(OutputFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Catch ex As PlatformNotSupportedException
                'System.Drawing.Common is not supported on platform
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As System.TypeInitializationException
                'The type initializer for 'Gdip' threw an exception.
                '---> System.PlatformNotSupportedException System.Drawing.Common Is Not supported on non-Windows platforms. See https://aka.ms/systemdrawingnonwindows for more information.
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As NotSupportedException
                Assert.Ignore("EngineNotSupported: " & ex.Message)
            Catch ex As NotImplementedException
                Assert.Ignore("EngineImplementation missing: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Embedded "picture": a chart in worksheet
        ''' </summary>
        <Test>
        Public Sub TestFileWithEmbeddedPicture02()
            Try
                Dim ExcelFile As String = TestEnvironment.FullPathOfExistingTestFile(TestFiles.TestFileEmbeddedPicture02.FullName)
                Dim Workbook = Me.CreateInstance(ExcelFile, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Dim OutputFile As String = TestEnvironment.FullPathOfDynTestFile(Workbook.GetType, "test_embeddedpicture02.xlsx")
                Workbook.SaveAs(OutputFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Catch ex As PlatformNotSupportedException
                'System.Drawing.Common is not supported on platform
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As System.TypeInitializationException
                'The type initializer for 'Gdip' threw an exception.
                '---> System.PlatformNotSupportedException System.Drawing.Common Is Not supported on non-Windows platforms. See https://aka.ms/systemdrawingnonwindows for more information.
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As NotSupportedException
                Assert.Ignore("EngineNotSupported: " & ex.Message)
            Catch ex As NotImplementedException
                Assert.Ignore("EngineImplementation missing: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Embedded static picture
        ''' </summary>
        <Test>
        Public Sub TestFileWithEmbeddedPicture03()
            Try
                Dim ExcelFile As String = TestEnvironment.FullPathOfExistingTestFile(TestFiles.TestFileEmbeddedPicture03.FullName)
                Dim Workbook = Me.CreateInstance(ExcelFile, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Dim OutputFile As String = TestEnvironment.FullPathOfDynTestFile(Workbook.GetType, "test_embeddedpicture02.xlsx")
                Workbook.SaveAs(OutputFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
            Catch ex As PlatformNotSupportedException
                'System.Drawing.Common is not supported on platform
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As System.TypeInitializationException
                'The type initializer for 'Gdip' threw an exception.
                '---> System.PlatformNotSupportedException System.Drawing.Common Is Not supported on non-Windows platforms. See https://aka.ms/systemdrawingnonwindows for more information.
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As NotSupportedException
                Assert.Ignore("EngineNotSupported: " & ex.Message)
            Catch ex As NotImplementedException
                Assert.Ignore("EngineImplementation missing: " & ex.Message)
            End Try
        End Sub
#End Region

#Region "Circular references behaviour"
        <Test>
        Public Sub TestFileWithCircularReference01_LoadWithCalculationEngineDefault()
            Try
                Dim ExcelFile As String = TestEnvironment.FullPathOfExistingTestFile(TestFiles.TestFileCircularReference01.FullName)
                Dim Workbook As T

                Dim CatchedEx As Exception = Nothing
                Try
                    Workbook = Me.CreateInstance(ExcelFile, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Catch ex As Exception
                    CatchedEx = ex
                    Workbook = Me.CreateInstanceUninitialized()
                End Try
                Select Case Workbook.EngineName
                    Case "Epplus 4 (LGPL)"
                        Assert.Multiple(
                            Sub()
                                Assert.AreEqual(True, Workbook.DefaultCalculationOptions.DisableCalculationEngine, "DefaultCalculationOptions.DisableCalculationEngine")
                                Assert.AreEqual(False, Workbook.DefaultCalculationOptions.DisableInitialCalculation, "DefaultCalculationOptions.DisableInitialCalculation")
                                Assert.AreEqual(True, Workbook.CalculationModuleDisabled, "CalculationModuleDisabled")
                                Assert.AreEqual(True, Workbook.AutoCalculationEnabledWorkbookSetting, "AutoCalculationEnabledWorkbookSetting")
                                Assert.AreEqual(True, Workbook.AutoCalculationOnLoad, "AutoCalculationOnLoad")
                                Assert.AreEqual(False, Workbook.AutoCalculationOnLoadEffectively, "AutoCalculationOnLoadEffectively")
                                Assert.Null(CatchedEx)
                                Assert.NotNull(Workbook.FilePath)
                            End Sub)
                    Case Else
                        Console.WriteLine("Workbook.EngineName=" & Workbook.EngineName)
                        Assert.Multiple(
                            Sub()
                                Assert.AreEqual(False, Workbook.DefaultCalculationOptions.DisableCalculationEngine, "DefaultCalculationOptions.DisableCalculationEngine")
                                Assert.AreEqual(False, Workbook.DefaultCalculationOptions.DisableInitialCalculation, "DefaultCalculationOptions.DisableInitialCalculation")
                                Assert.AreEqual(False, Workbook.CalculationModuleDisabled, "CalculationModuleDisabled")
                                Assert.AreEqual(True, Workbook.AutoCalculationEnabledWorkbookSetting, "AutoCalculationEnabledWorkbookSetting")
                                Assert.AreEqual(True, Workbook.AutoCalculationOnLoad, "AutoCalculationOnLoad")
                                Assert.AreEqual(True, Workbook.AutoCalculationOnLoadEffectively, "AutoCalculationOnLoadEffectively")
                                Assert.Null(CatchedEx)
                                Assert.NotNull(Workbook.FilePath)
                            End Sub)
                End Select
            Catch ex As PlatformNotSupportedException
                'System.Drawing.Common is not supported on platform
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As System.TypeInitializationException
                'The type initializer for 'Gdip' threw an exception.
                '---> System.PlatformNotSupportedException System.Drawing.Common Is Not supported on non-Windows platforms. See https://aka.ms/systemdrawingnonwindows for more information.
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As NotSupportedException
                Assert.Ignore("EngineNotSupported: " & ex.Message)
            Catch ex As NotImplementedException
                Assert.Ignore("EngineImplementation missing: " & ex.Message)
            End Try
        End Sub

        <Test>
        Public Sub TestFileWithCircularReference01_LoadAndResaveWithDisabledCalculationEngine()
            Try
                Dim ExcelFile As String = TestEnvironment.FullPathOfExistingTestFile(TestFiles.TestFileCircularReference01.FullName)
                Dim Workbook As T = Me.CreateInstance(ExcelFile, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions("", True, True, True))
                Dim OutputFile As String = TestEnvironment.FullPathOfDynTestFile(Workbook.GetType, "test_circularref01_rewritten.xlsx")
                Workbook.SaveAs(OutputFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            Catch ex As PlatformNotSupportedException
                'System.Drawing.Common is not supported on platform
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As System.TypeInitializationException
                'The type initializer for 'Gdip' threw an exception.
                '---> System.PlatformNotSupportedException System.Drawing.Common Is Not supported on non-Windows platforms. See https://aka.ms/systemdrawingnonwindows for more information.
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As NotSupportedException
                Assert.Ignore("EngineNotSupported: " & ex.Message)
            Catch ex As NotImplementedException
                Assert.Ignore("EngineImplementation missing: " & ex.Message)
            End Try
        End Sub

        <Test>
        Public Sub TestFileWithCircularReference01_LoadAndRecalculateOnEnabledCalculationEngine()
            Try
                Dim ExcelFile As String = TestEnvironment.FullPathOfExistingTestFile(TestFiles.TestFileCircularReference01.FullName)
                Dim Workbook As T = Me.CreateInstance(ExcelFile, ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions("", True, True, False))
                Dim CatchedEx As Exception = Nothing
                Try
                    If Workbook.CalculationModuleDisabled = False Then
                        Workbook.RecalculateAll()
                    End If
                Catch ex As Exception
                    CatchedEx = ex
                End Try
                Select Case Workbook.EngineName
                    Case Else
                        Console.WriteLine("Workbook.EngineName=" & Workbook.EngineName)
                        Assert.Null(CatchedEx)
                End Select

            Catch ex As PlatformNotSupportedException
                'System.Drawing.Common is not supported on platform
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As System.TypeInitializationException
                'The type initializer for 'Gdip' threw an exception.
                '---> System.PlatformNotSupportedException System.Drawing.Common Is Not supported on non-Windows platforms. See https://aka.ms/systemdrawingnonwindows for more information.
                'just ignore AutoFit feature
                Assert.Ignore("PlatformNotSupported: " & ex.Message)
            Catch ex As NotSupportedException
                Assert.Ignore("EngineNotSupported: " & ex.Message)
            Catch ex As NotImplementedException
                Assert.Ignore("EngineImplementation missing: " & ex.Message)
            End Try
        End Sub
#End Region

#Region "HTML exports"
        <Test>
        Public Sub HtmlExportWorkbookGrunddaten01()
            Dim TestXlsxFile = TestFiles.TestFileGrund01()
            System.Console.WriteLine("TEST IN FILE: " & TestXlsxFile.FullName)

            Try
                Dim Wb As CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestXlsxFile.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Dim TestHtmlOutputFile = TestEnvironment.FullPathOfDynTestFile(Wb, TestXlsxFile.Name & ".html")
                System.Console.WriteLine("TEST OUT FILE: " & TestHtmlOutputFile)
                Wb.ExportWorkbookToHtml(TestHtmlOutputFile, New HtmlWorkbookExportOptions() With {.SheetNavigationActionStyle = HtmlWorkbookExportOptions.SheetNavigationActionStyles.JumpToAnchor})
                If OPEN_HTML_OUTPUT_IN_BROWSER_AFTER_TEST Then
                    Dim OpenFileProcess = System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                .UseShellExecute = True,
                .FileName = TestHtmlOutputFile
                })
                End If
            Catch ex As NotImplementedException
                Assert.Ignore("Not implemented for this Excel engine")
            Catch ex As TypeInitializationException
                Assert.Ignore("Not supported on this platform " & System.Environment.OSVersion.Platform.ToString)
            End Try

        End Sub

        <Test>
        Public Sub HtmlExportWorkbookHtmlExport01()
            Dim TestXlsxFile = TestFiles.TestFileHtmlExport01()
            System.Console.WriteLine("TEST IN FILE: " & TestXlsxFile.FullName)

            Try
                With Nothing
                    Dim Wb As CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestXlsxFile.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                    Dim TestHtmlOutputFile = TestEnvironment.FullPathOfDynTestFile(Wb, TestXlsxFile.Name & ".nav-anchor.html")
                    System.Console.WriteLine("TEST OUT FILE: " & TestHtmlOutputFile)
                    Wb.ExportWorkbookToHtml(TestHtmlOutputFile, New HtmlWorkbookExportOptions() With {.SheetNavigationActionStyle = HtmlWorkbookExportOptions.SheetNavigationActionStyles.JumpToAnchor})
                    If OPEN_HTML_OUTPUT_IN_BROWSER_AFTER_TEST Then
                        Dim OpenFileProcess = System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                            .UseShellExecute = True,
                            .FileName = TestHtmlOutputFile
                            })
                    End If
                End With
                With Nothing
                    Dim Wb As CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestXlsxFile.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                    Dim TestHtmlOutputFile = TestEnvironment.FullPathOfDynTestFile(Wb, TestXlsxFile.Name & ".nav-anchor-fixed-on-top.html")
                    System.Console.WriteLine("TEST OUT FILE: " & TestHtmlOutputFile)
                    Wb.ExportWorkbookToHtml(TestHtmlOutputFile, New HtmlWorkbookExportOptions() With {.SheetNavigationActionStyle = HtmlWorkbookExportOptions.SheetNavigationActionStyles.JumpToAnchor, .SheetNavigationAlwaysVisible = True})
                    If OPEN_HTML_OUTPUT_IN_BROWSER_AFTER_TEST Then
                        Dim OpenFileProcess = System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                            .UseShellExecute = True,
                            .FileName = TestHtmlOutputFile
                            })
                    End If
                End With
                With Nothing
                    Dim Wb As CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase = Me.CreateInstance(TestXlsxFile.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                    Dim TestHtmlOutputFile = TestEnvironment.FullPathOfDynTestFile(Wb, TestXlsxFile.Name & ".nav-switch.html")
                    System.Console.WriteLine("TEST OUT FILE: " & TestHtmlOutputFile)
                    Wb.ExportWorkbookToHtml(TestHtmlOutputFile, New HtmlWorkbookExportOptions() With {.SheetNavigationActionStyle = HtmlWorkbookExportOptions.SheetNavigationActionStyles.SwitchVisibleSheet})
                    If OPEN_HTML_OUTPUT_IN_BROWSER_AFTER_TEST Then
                        Dim OpenFileProcess = System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                            .UseShellExecute = True,
                            .FileName = TestHtmlOutputFile
                            })
                    End If
                End With
            Catch ex As NotImplementedException
                Assert.Ignore("Not implemented for this Excel engine")
            Catch ex As TypeInitializationException
                Assert.Ignore("Not supported on this platform " & System.Environment.OSVersion.Platform.ToString)
            End Try

        End Sub

        <Test>
        Public Sub HtmlExportWorksheetGrunddatenV19()
            Dim TestXlsxFile = TestFiles.TestFileGrund01()
            System.Console.WriteLine("TEST IN FILE: " & TestXlsxFile.FullName)

            Dim TempFilePng As String = TestEnvironment.FullPathOfDynTestFile(Me.CreateInstanceUninitialized, "excel_test_chart.png")
            Dim Workbook = PrepareAndFillExcelFileWithChart(0)

            Dim Wb As CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase = Nothing
            Try
                Wb = Me.CreateInstance(TestXlsxFile.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
            Catch ex As TypeInitializationException
                Assert.Ignore("Not supported on this platform " & System.Environment.OSVersion.Platform.ToString)
            End Try
            For Each WorkSheetName In Wb.WorkSheetNames
                System.Console.WriteLine("FOUND WORKSHEET: " & WorkSheetName)
            Next
            Try
                For Each WorkSheetName In Wb.WorkSheetNames
                    With Nothing
                        Dim TestHtmlOutputFile = TestEnvironment.FullPathOfDynTestFile(Me.CreateInstanceUninitialized, TestXlsxFile.Name & "." & WorkSheetName & ".no-title.html")
                        System.Console.WriteLine("TEST OUT FILE: " & TestHtmlOutputFile)
                        Wb.ExportSheetToHtml(WorkSheetName, TestHtmlOutputFile, New HtmlSheetExportOptions)
                        If OPEN_HTML_OUTPUT_IN_BROWSER_AFTER_TEST Then
                            Dim OpenFileProcess = System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                        .UseShellExecute = True,
                        .FileName = TestHtmlOutputFile
                        })
                        End If
                    End With
                    With Nothing
                        Dim TestHtmlOutputFile = TestEnvironment.FullPathOfDynTestFile(Me.CreateInstanceUninitialized, TestXlsxFile.Name & "." & WorkSheetName & ".h2.html")
                        Wb.ExportSheetToHtml(WorkSheetName, TestHtmlOutputFile, New HtmlSheetExportOptions() With {.ExportSheetNameAsTitle = HtmlSheetExportOptions.SheetTitleStyles.H2})
                        If OPEN_HTML_OUTPUT_IN_BROWSER_AFTER_TEST Then
                            Dim OpenFileProcess = System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                        .UseShellExecute = True,
                        .FileName = TestHtmlOutputFile
                        })
                        End If
                    End With
                Next
            Catch ex As NotImplementedException
                Assert.Ignore("Not implemented for this Excel engine")
            End Try

        End Sub
#End Region

    End Class

End Namespace