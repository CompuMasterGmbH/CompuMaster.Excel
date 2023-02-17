Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps
Imports System.Data

Namespace ExcelOpsTests.MsExcelSpecials

    ''' <summary>
    ''' Test (component design) issue: MsExcelAutoCalc doesn't work but only manual calc works -> Bug!?
    ''' </summary>
    <NonParallelizable>
    <TestFixture>
    Public Class MsExcelCalcTestFreeSpireXls
        Inherits CompuMaster.Excel.Test.ExcelOpsTests.MsExcelSpecials.MsExcelCalcTestBase

        Protected Overrides ReadOnly Property EngineName As String
            Get
                Static Result As String
                If Result Is Nothing Then Result = (New ExcelOps.FreeSpireXlsDataOperations()).EngineName
                Return Result
            End Get
        End Property

        Protected Overrides Function CreateEngineInstance(testFile As String) As ExcelOps.ExcelDataOperationsBase
            Return New ExcelOps.FreeSpireXlsDataOperations(testFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, False, String.Empty)
        End Function

        Protected Overrides Sub EngineResetCellValueFromFormulaCell(wb As ExcelOps.ExcelDataOperationsBase, sheetName As String, rowIndex As Integer, columnIndex As Integer)
            Assert.Ignore("Test not applicable for engine " & wb.EngineName)
            'CType(wb, ExcelOps.FreeSpireXlsDataOperations).ResetCellValueFromFormulaCell(sheetName, rowIndex, columnIndex)
        End Sub

    End Class

End Namespace