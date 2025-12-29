Imports NUnit.Framework
Imports NUnit.Framework.Legacy
Imports CompuMaster.Excel.ExcelOps
Imports System.Data

Namespace ExcelOpsTests.MsExcelSpecials

    ''' <summary>
    ''' Test (component design) issue: MsExcelAutoCalc doesn't work but only manual calc works -> Bug!?
    ''' </summary>
    <NonParallelizable>
    <TestFixture>
    Public Class MsExcelCalcTestSpireXls
        Inherits CompuMaster.Excel.Test.ExcelOpsTests.MsExcelSpecials.MsExcelCalcTestBase

        Protected Overrides ReadOnly Property EngineName As String
            Get
                Static Result As String
                If Result Is Nothing Then
                    ExcelOps.SpireXlsDataOperations.AllowInstancingForNonLicencedContextForTestingPurposesOnly = True
                    Result = (New ExcelOps.SpireXlsDataOperations(ExcelDataOperationsBase.OpenMode.Uninitialized)).EngineName
                End If
                Return Result
            End Get
        End Property

        Protected Overrides Function CreateEngineInstanceWithCreateFileMode(testFile As String) As ExcelOps.ExcelDataOperationsBase
            Return New ExcelOps.SpireXlsDataOperations(testFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile))
        End Function

        Protected Overrides Sub EngineResetCellValueFromFormulaCell(wb As ExcelOps.ExcelDataOperationsBase, sheetName As String, rowIndex As Integer, columnIndex As Integer)
            ClassicAssert.Ignore("Test not applicable for engine " & wb.EngineName)
            'CType(wb, ExcelOps.FreeSpireXlsDataOperations).ResetCellValueFromFormulaCell(sheetName, rowIndex, columnIndex)
        End Sub

    End Class

End Namespace