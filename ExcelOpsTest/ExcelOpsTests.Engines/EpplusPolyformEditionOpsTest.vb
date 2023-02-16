Namespace ExcelOpsTests.Engines

    Public Class EpplusPolyformEditionOpsTest
        Inherits ExcelOpsTestBase(Of ExcelOps.EpplusPolyformExcelDataOperations)

        Public Sub New()
            AssignLicenseContext()
        End Sub

        Friend Shared Sub AssignLicenseContext()
            ExcelOps.EpplusPolyformExcelDataOperations.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
        End Sub

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String) As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations(file, mode, [readOnly], passwordForOpening)
        End Function

        Protected Overrides Function _CreateInstance() As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations()
        End Function

    End Class

End Namespace