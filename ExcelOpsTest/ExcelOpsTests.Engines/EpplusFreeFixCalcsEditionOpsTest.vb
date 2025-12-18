Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework

Namespace ExcelOpsTests.Engines

    Public Class EpplusFreeFixCalcsEditionOpsTest
        Inherits ExcelOpsTestBase(Of ExcelOps.EpplusFreeExcelDataOperations)

        Public Overrides ReadOnly Property ExpectedEngineName As String = "Epplus 4 (LGPL)"

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.EpplusFreeExcelDataOperations
            'disableCalculationEngine not required since always disabled calc-module by engine
            Return New ExcelOps.EpplusFreeExcelDataOperations(file, mode, options)
        End Function

        Protected Overrides Function _CreateInstanceUninitialized() As ExcelOps.EpplusFreeExcelDataOperations
            Return New ExcelOps.EpplusFreeExcelDataOperations(ExcelDataOperationsBase.OpenMode.Uninitialized)
        End Function

        <Test> Public Overrides Sub CopySheetContent()
            Assert.Throws(Of NotSupportedException)(Sub() MyBase.CopySheetContent())
        End Sub

        Protected Overrides Function _CreateInstance(data() As Byte, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.EpplusFreeExcelDataOperations
            'disableCalculationEngine not required since always disabled calc-module by engine
            Return New ExcelOps.EpplusFreeExcelDataOperations(data, options)
        End Function

        Protected Overrides Function _CreateInstance(data As IO.Stream, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.EpplusFreeExcelDataOperations
            'disableCalculationEngine not required since always disabled calc-module by engine
            Return New ExcelOps.EpplusFreeExcelDataOperations(data, options)
        End Function

    End Class

End Namespace