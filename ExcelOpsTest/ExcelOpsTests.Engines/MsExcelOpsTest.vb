Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework

Namespace ExcelOpsTests.Engines

    <NonParallelizable>
    Public Class MsExcelOpsTest
        Inherits ExcelOpsTestBase(Of ExcelOps.MsExcelDataOperations)

        Public Overrides ReadOnly Property ExpectedEngineName As String = "Microsoft Excel (2013 or higher)"

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.MsExcelDataOperations
            If MsExcelInstance Is Nothing OrElse MsExcelInstance.IsDisposed Then
                'recreate excel instance
                MsExcelInstance = New CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper
            ElseIf AlwaysCloseAllWorkbooksInNewEngineInstances Then
                MsExcelInstance.Workbooks.CloseAllWorkbooks()
            End If
            Return New ExcelOps.MsExcelDataOperations(file, mode, MsExcelInstance, False, options)
        End Function

        Protected Overrides Function _CreateInstanceUninitialized() As ExcelOps.MsExcelDataOperations
            If MsExcelInstance Is Nothing OrElse MsExcelInstance.IsDisposed Then
                'recreate excel instance
                MsExcelInstance = New CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper
            ElseIf AlwaysCloseAllWorkbooksInNewEngineInstances Then
                MsExcelInstance.Workbooks.CloseAllWorkbooks()
            End If
            Return New ExcelOps.MsExcelDataOperations(ExcelDataOperationsBase.OpenMode.Uninitialized)
        End Function

        <OneTimeSetUp>
        Public Sub OneTimeSetup()
            If Global.CompuMaster.Excel.MsExcelCom.MsExcelTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled = False Then
                AssertNoExcelProcessesAvailable()
            End If
        End Sub

        <SetUp>
        Public Sub Setup()
            AssertExactlyOur1ExcelProcessAvailableInProcessList()
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
            Assert.LessOrEqual(WbCount, 1, AssertionMessageWbCount)
            If WbCount = 1 Then
                MsExcelInstance.Workbooks.Workbook(1).CloseAndDispose()
            End If
        End Sub

        <OneTimeTearDown>
        Public Sub OneTimeTearDown()
            If MsExcelInstance IsNot Nothing Then MsExcelInstance.Dispose()
            CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
            AssertNoExcelProcessesAvailable()
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

        Private Sub AssertExactlyOur1ExcelProcessAvailableInProcessList()
            If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
                Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                If MsExcelProcesses.Length <> 1 Then
                    Assert.Fail("Found " & MsExcelProcesses.Length & " EXCEL processes, but 1 excel processes allowed/required")
                Else
                    Assert.AreEqual(MsExcelInstance.ExcelProcessId, MsExcelProcesses(0).Id)
                End If
            End If
        End Sub

        Private Sub AssertNoExcelProcessesAvailable()
            If NUnit.Framework.TestContext.CurrentContext.Test.Name <> NameOf(ManualRunOnly_KillAllMsExcelAppProcesses) Then
                Dim MsExcelProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                If MsExcelProcesses.Length <> 0 Then
                    Assert.Fail("Found " & MsExcelProcesses.Length & " EXCEL processes, but no excel processes allowed")
                End If
            End If
        End Sub

        Private MsExcelInstance As CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper

        <Test, Explicit("Known2Fail But Less Important")> Public Sub ManualRunOnly_KillAllMsExcelAppProcesses()
            Dim MsExcelProcessesBefore As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Console.WriteLine("Found " & MsExcelProcessesBefore.Length & " EXCEL processes")
            ExcelOps.MsExcelDataOperations.CheckForRunningMsExcelInstancesAndAskUserToKill()
            Assert.Pass()
        End Sub

        <Test> Public Overrides Sub CopySheetContent()
            Me.ClosedAllWorkbooksAndAllowMultipleWorkbooksForThisTestRun()
            MyBase.CopySheetContent()
        End Sub

        Protected Overrides Sub TestInCultureContext_AssignCurrentThreadCulture()
            MsExcelInstance.SetCultureContext(System.Threading.Thread.CurrentThread.CurrentCulture)
        End Sub

        Protected Overrides Function _CreateInstance(data() As Byte, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.MsExcelDataOperations
            Throw New NotSupportedException
        End Function

        Protected Overrides Function _CreateInstance(data As IO.Stream, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.MsExcelDataOperations
            Throw New NotSupportedException
        End Function
    End Class

End Namespace