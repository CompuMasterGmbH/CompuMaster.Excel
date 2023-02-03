Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps

Public MustInherit Class ExcelOpsTestBase(Of T As ExcelOps.ExcelDataOperationsBase)

    Protected MustOverride Function CreateInstance() As T

    Protected MustOverride Function CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String) As T

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
        Assert.AreEqual(FilePathInEngineBefore, Wb.WorkbookFilePath)
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
        Wb.Close()

        'Try to reload it without password -> it must fail
        Assert.Catch(Of Exception)(Sub() Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "something else"))
        Assert.Catch(Of Exception)(Sub() Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, ""))
        Assert.Catch(Of Exception)(Sub() Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, Nothing))

        'Reload it with password -> now it must succeed
        Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, "dummy")
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

        'Creating a new workbook without pre-defined file name must fail on Save(), but successful on SaveAs()
        Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, "")
        Assert.AreEqual(True, Wb.ReadOnly, "Newly created files must always be ReadOnly")
        Assert.AreEqual(TestFile, Wb.FilePath)
        Assert.AreEqual(Nothing, Wb.WorkbookFilePath)
        Assert.Throws(Of FileReadOnlyException)(Sub() Wb.Save())
        Wb.SaveAs(TestFile, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.DefaultBehaviour)
        Wb.Close()

        'Creating a new workbook must fail with a pre-defined file name if there is already a file
        Assert.Throws(Of FileAlreadyExistsException)(Sub()
                                                         Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, "")
                                                     End Sub)
        System.IO.File.Delete(TestFile) 'Delete the file for next test block

        'Creating a new workbook must always be ReadOnly and saving it without a name must be forbidden
        Wb = Me.CreateInstance(TestFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, "")
        Assert.AreEqual(True, Wb.ReadOnly, "Newly created files must always be ReadOnly")
        Assert.AreEqual(TestFile, Wb.FilePath)
        Assert.AreEqual(Nothing, Wb.WorkbookFilePath)
        Assert.Throws(Of FileReadOnlyException)(Sub() Wb.Save())
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
End Class
