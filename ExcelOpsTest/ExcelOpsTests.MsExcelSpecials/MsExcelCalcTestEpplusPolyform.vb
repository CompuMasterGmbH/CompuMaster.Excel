Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps
Imports System.Data

Namespace ExcelOpsTests.MsExcelSpecials

    ''' <summary>
    ''' Test (component design) issue: MsExcelAutoCalc doesn't work but only manual calc works -> Bug!?
    ''' </summary>
    <NonParallelizable>
    <TestFixture>
    Public Class MsExcelCalcTestEpplusPolyform
        Inherits MsExcelCalcTestBase

        <OneTimeSetUp>
        Public Sub InitializeLicenseContext()
            ExcelOps.EpplusPolyformExcelDataOperations.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
        End Sub

        Protected Overrides ReadOnly Property EngineName As String
            Get
                Static Result As String
                If Result Is Nothing Then
                    Result = (New ExcelOps.EpplusPolyformExcelDataOperations(ExcelDataOperationsBase.OpenMode.Uninitialized)).EngineName
                End If
                Return Result
            End Get
        End Property

        Protected Overrides Function CreateEngineInstanceWithCreateFileMode(testFile As String) As ExcelOps.ExcelDataOperationsBase
            ExcelOpsTests.Engines.EpplusPolyformEditionOpsTest.AssignLicenseContext()
            Return New ExcelOps.EpplusPolyformExcelDataOperations(testFile, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile))
        End Function

        Protected Overrides Sub EngineResetCellValueFromFormulaCell(wb As ExcelOps.ExcelDataOperationsBase, sheetName As String, rowIndex As Integer, columnIndex As Integer)
            Assert.Ignore("Test not applicable for engine " & wb.EngineName)
            'CType(wb, ExcelOps.EpplusPolyformExcelDataOperations).ResetCellValueFromFormulaCell(sheetName, rowIndex, columnIndex)
        End Sub

    End Class

End Namespace