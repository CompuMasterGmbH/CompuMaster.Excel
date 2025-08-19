Imports NUnit.Framework

Namespace ExcelOpsTests.Engines

    Public Class EpplusFreeFixCalcsEditionOpsTest
        Inherits ExcelOpsTestBase(Of ExcelOps.EpplusFreeExcelDataOperations)

        Public Overrides ReadOnly Property ExpectedEngineName As String = "Epplus 4 (LGPL)"

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String, disableInitialCalculation As Boolean) As ExcelOps.EpplusFreeExcelDataOperations
            'disableCalculationEngine not required since always disabled calc-module by engine
            Return New ExcelOps.EpplusFreeExcelDataOperations(file, mode, [readOnly], passwordForOpening)
        End Function

        Protected Overrides Function _CreateInstance() As ExcelOps.EpplusFreeExcelDataOperations
            Return New ExcelOps.EpplusFreeExcelDataOperations()
        End Function

        <Test> Public Overrides Sub CopySheetContent()
            Assert.Throws(Of NotSupportedException)(Sub() MyBase.CopySheetContent())
        End Sub

        Protected Overrides Function _CreateInstance(data() As Byte, passwordForOpening As String, disableCalculationEngine As Boolean) As ExcelOps.EpplusFreeExcelDataOperations
            'disableCalculationEngine not required since always disabled calc-module by engine
            Return New ExcelOps.EpplusFreeExcelDataOperations(data, passwordForOpening)
        End Function

        Protected Overrides Function _CreateInstance(data As IO.Stream, passwordForOpening As String, disableCalculationEngine As Boolean) As ExcelOps.EpplusFreeExcelDataOperations
            'disableCalculationEngine not required since always disabled calc-module by engine
            Return New ExcelOps.EpplusFreeExcelDataOperations(data, passwordForOpening)
        End Function

    End Class

End Namespace