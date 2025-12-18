Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework

Namespace ExcelOpsTests.Engines

    Public Class EpplusPolyformEditionOpsTest
        Inherits ExcelOpsTestBase(Of ExcelOps.EpplusPolyformExcelDataOperations)

        Public Sub New()
            AssignLicenseContext()
        End Sub

        Friend Shared Sub AssignLicenseContext()
            ExcelOps.EpplusPolyformExcelDataOperations.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
        End Sub

        Public Overrides ReadOnly Property ExpectedEngineName As String = "Epplus (Polyform license edition)"

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations(file, mode, options)
        End Function

        Protected Overrides Function _CreateInstance() As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations()
        End Function

        <Test> Public Overrides Sub CopySheetContent()
            Assert.Throws(Of NotSupportedException)(Sub() MyBase.CopySheetContent())
        End Sub

        Protected Overrides Function _CreateInstance(data() As Byte, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations(data, options)
        End Function

        Protected Overrides Function _CreateInstance(data As IO.Stream, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations(data, options)
        End Function

    End Class

End Namespace