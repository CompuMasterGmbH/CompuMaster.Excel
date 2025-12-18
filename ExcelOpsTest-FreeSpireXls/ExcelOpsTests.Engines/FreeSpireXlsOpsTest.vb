Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework

Namespace ExcelOpsTests.Engines

    Public Class FreeSpireXlsOpsTest
        Inherits CompuMaster.Excel.Test.ExcelOpsTests.Engines.ExcelOpsTestBase(Of ExcelOps.FreeSpireXlsDataOperations)

        Public Overrides ReadOnly Property ExpectedEngineName As String = "FreeSpire.Xls"

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.FreeSpireXlsDataOperations
            Return New ExcelOps.FreeSpireXlsDataOperations(file, mode, options)
        End Function

        Protected Overrides Function _CreateInstance() As ExcelOps.FreeSpireXlsDataOperations
            Return New ExcelOps.FreeSpireXlsDataOperations()
        End Function

        <Test> Public Overrides Sub CopySheetContent()
            Assert.Throws(Of NotSupportedException)(Sub() MyBase.CopySheetContent())
        End Sub

        Protected Overrides Sub TestInCultureContext_AssignCurrentThreadCulture()
            MyBase.TestInCultureContext_AssignCurrentThreadCulture()
        End Sub

        Protected Overrides Function _CreateInstance(data() As Byte, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.FreeSpireXlsDataOperations
            Return New ExcelOps.FreeSpireXlsDataOperations(data, options)
        End Function

        Protected Overrides Function _CreateInstance(data As IO.Stream, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.FreeSpireXlsDataOperations
            Return New ExcelOps.FreeSpireXlsDataOperations(data, options)
        End Function

    End Class

End Namespace