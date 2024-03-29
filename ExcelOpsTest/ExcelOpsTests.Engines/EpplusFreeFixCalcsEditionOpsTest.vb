﻿Imports NUnit.Framework

Namespace ExcelOpsTests.Engines

    Public Class EpplusFreeFixCalcsEditionOpsTest
        Inherits ExcelOpsTestBase(Of ExcelOps.EpplusFreeExcelDataOperations)

        Public Overrides ReadOnly Property ExpectedEngineName As String = "Epplus 4 (LGPL)"

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String, disableInitialCalculation As Boolean) As ExcelOps.EpplusFreeExcelDataOperations
            Return New ExcelOps.EpplusFreeExcelDataOperations(file, mode, [readOnly], passwordForOpening)
        End Function

        Protected Overrides Function _CreateInstance() As ExcelOps.EpplusFreeExcelDataOperations
            Return New ExcelOps.EpplusFreeExcelDataOperations()
        End Function

        <Test> Public Overrides Sub CopySheetContent()
            Assert.Throws(Of NotSupportedException)(Sub() MyBase.CopySheetContent())
        End Sub

    End Class

End Namespace