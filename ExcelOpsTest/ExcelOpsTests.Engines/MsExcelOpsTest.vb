Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework
Imports NUnit.Framework.Legacy

Namespace ExcelOpsTests.Engines

    <NonParallelizable>
    Partial Public Class MsExcelOpsTest
        Inherits ExcelOpsTestBase(Of ExcelOps.MsExcelDataOperations)

        Public Overrides ReadOnly Property ExpectedEngineName As String = "Microsoft Excel (2013 or higher)"

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.MsExcelDataOperations
            If MsExcelInstance Is Nothing OrElse MsExcelInstance.IsClosed OrElse MsExcelInstance.IsDisposed Then
                'recreate excel instance
                MsExcelInstance = New CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper
            ElseIf AlwaysCloseAllWorkbooksInNewEngineInstances Then
                MsExcelInstance.Workbooks.CloseAllWorkbooks()
            End If
            Return New ExcelOps.MsExcelDataOperations(file, mode, MsExcelInstance, False, options)
        End Function

        Protected Overrides Function _CreateInstanceUninitialized() As ExcelOps.MsExcelDataOperations
            If MsExcelInstance Is Nothing OrElse MsExcelInstance.IsClosed OrElse MsExcelInstance.IsDisposed Then
                'recreate excel instance
                MsExcelInstance = New CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper
            ElseIf AlwaysCloseAllWorkbooksInNewEngineInstances Then
                MsExcelInstance.Workbooks.CloseAllWorkbooks()
            End If
            Return New ExcelOps.MsExcelDataOperations(ExcelDataOperationsBase.OpenMode.Uninitialized)
        End Function

        <OneTimeSetUp>
        Public Sub OneTimeSetup()
        End Sub

        <SetUp>
        Public Sub Setup()
            AlwaysCloseAllWorkbooksInNewEngineInstances = True
        End Sub

        <TearDown>
        Public Sub TearDown()
            Dim WbCount As Integer = MsExcelInstance.Workbooks.Count
            Dim AssertionMessageWbCount As String = Nothing
            For MyCounter As Integer = 0 To WbCount - 1
                If AssertionMessageWbCount <> Nothing Then AssertionMessageWbCount &= ","
                AssertionMessageWbCount &= MsExcelInstance.Workbooks.Workbook(MyCounter + 1).Name
            Next
            If AlwaysCloseAllWorkbooksInNewEngineInstances = False Then
                For MyCounter As Integer = WbCount To 2 Step -1
                    MsExcelInstance.Workbooks.Workbook(MyCounter).CloseAndDispose()
                Next
            End If
            ClassicAssert.LessOrEqual(WbCount, 1, AssertionMessageWbCount)
            If WbCount = 1 Then
                MsExcelInstance.Workbooks.Workbook(1).CloseAndDispose()
            End If
        End Sub

        <OneTimeTearDown>
        Public Sub OneTimeTearDown()
            MsExcelInstance.Workbooks.CloseAllWorkbooks()
            MsExcelInstance.Close()
            If MsExcelInstance IsNot Nothing Then MsExcelInstance.Dispose()
            CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
        End Sub

        ''' <summary>
        ''' Allows multiple worksheets since 2nd CreateInstance method call doesn't close previously opened workbook(s)
        ''' </summary>
        ''' <returns></returns>
        Protected Property AlwaysCloseAllWorkbooksInNewEngineInstances As Boolean = True

        ''' <summary>
        ''' Configures the environment to allow multiple workbooks in MS Excel (otherwise all workbook would get closed in test setup)
        ''' </summary>
        Protected Sub ClosedAllWorkbooksAndAllowMultipleWorkbooksForThisTestRun()
            MsExcelInstance.Workbooks.CloseAllWorkbooks()
            AlwaysCloseAllWorkbooksInNewEngineInstances = False
        End Sub

        Private MsExcelInstance As CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper

        <TestCase(True)>
        <TestCase(False)>
        Public Sub WorkbookReferences_ExcelTracksSourceEditsOnlyInAnOpenDestination(destinationOpen As Boolean)
            Dim directory = IO.Path.Combine(IO.Path.GetTempPath(), "ExcelOps-WorkbookLinks-" & Guid.NewGuid().ToString("N"))
            IO.Directory.CreateDirectory(directory)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            Try
                WorkbookReferenceCoverage.CopyPair(directory)
                AlwaysCloseAllWorkbooksInNewEngineInstances = False
                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite)
                source = Me.CreateInstance(IO.Path.Combine(directory, WorkbookReferenceCoverage.SourceName), ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                Dim targetPath = IO.Path.Combine(directory, WorkbookReferenceCoverage.TargetName)
                If destinationOpen Then target = Me.CreateInstance(targetPath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)

                source.AddRow(WorkbookReferenceCoverage.RemoteSheet, 1, 1, True)
                source.Save()

                If Not destinationOpen Then target = Me.CreateInstance(targetPath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                Dim expectedAddress = If(destinationOpen, "$B$3", "$B$2")
                Dim expectedSuffix = WorkbookReferenceCoverage.RemoteSheet & "'!" & expectedAddress
                Assert.That(target.LookupCellFormula(WorkbookReferenceCoverage.ChecksSheet, 0, 0), Does.EndWith(expectedSuffix))
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
                IO.Directory.Delete(directory, True)
            End Try
        End Sub

        <Test>
        Public Sub WorkbookReferences_ExcelReopensAndCalculatesOpenXmlOutput()
            Dim directory = IO.Path.Combine(IO.Path.GetTempPath(), "ExcelOps-WorkbookLinks-" & Guid.NewGuid().ToString("N"))
            IO.Directory.CreateDirectory(directory)
            Dim target As MsExcelDataOperations = Nothing
            Try
                WorkbookReferenceCoverage.CopyPair(directory)
                Dim original = IO.Path.Combine(directory, WorkbookReferenceCoverage.TargetName)
                Dim edited = IO.Path.Combine(directory, "edited.xlsx")
                Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
                Using editor As New OpenXmlExcelDataOperations(original, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    WorkbookReferenceCoverage.EditTarget(editor)
                    editor.SaveCopyAs(edited, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using

                target = Me.CreateInstance(edited, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target.RecalculateAll()
                Assert.That(target.LookupCellValue(Of Integer)(WorkbookReferenceCoverage.ChecksSheet, 2, 0, FormulaCalculationMode.EngineDefault),
                            [Is].EqualTo(36))
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                IO.Directory.Delete(directory, True)
            End Try
        End Sub

        <Test, Explicit("Known2Fail But Less Important")> Public Sub ManualRunOnly_KillAllMsExcelAppProcesses()
            Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Console.WriteLine("Found " & MsExcelProcessesBefore.Length & " EXCEL processes")
            ExcelOps.MsExcelDataOperations.CheckForRunningMsExcelInstancesAndAskUserToKill()
            ClassicAssert.Pass()
        End Sub

        <Test> Public Overrides Sub CopySheetContent()
            Me.ClosedAllWorkbooksAndAllowMultipleWorkbooksForThisTestRun()
            MyBase.CopySheetContent()
        End Sub

        Protected Overrides Sub TestInCultureContext_AssignCurrentThreadCulture()
            MsExcelInstance.SetCultureContext(System.Threading.Thread.CurrentThread.CurrentCulture)
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Function ReadFormulaReferenceResult(workbook As ExcelDataOperationsBase, address As ExcelCell) As Object
            ' Read native values after worksheet calculation, including empty text and typed Excel errors.
            Dim native = DirectCast(workbook, MsExcelDataOperations).Workbook
            Dim sheet = DirectCast(native.Worksheets.Item(address.SheetName), Microsoft.Office.Interop.Excel.Worksheet)
            Dim cell = sheet.Range(address.Address)
            Try
                Dim value As Object = cell.Value2
                If value Is Nothing OrElse (TypeOf value Is String AndAlso DirectCast(value, String).Length = 0) Then
                    ' Excel may leave an unloaded dynamic array anchor clean after the worksheet pass.
                    cell.Dirty()
                    cell.Calculate()
                    value = cell.Value2
                End If
                If Not TypeOf value Is Integer Then Return value
                Select Case DirectCast(value, Integer) And &HFFFF
                    Case 2045 : Return "#SPILL!"
                    Case 2050 : Return "#CALC!"
                End Select
                Return workbook.LookupCellErrorValue(address.SheetName, address.RowIndex, address.ColumnIndex)
            Finally
                Runtime.InteropServices.Marshal.ReleaseComObject(cell)
                Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            End Try
        End Function

        Protected Overrides Function _CreateInstance(data() As Byte, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.MsExcelDataOperations
            Throw New NotSupportedException
        End Function

        Protected Overrides Function _CreateInstance(data As IO.Stream, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.MsExcelDataOperations
            Throw New NotSupportedException
        End Function
    End Class

End Namespace
