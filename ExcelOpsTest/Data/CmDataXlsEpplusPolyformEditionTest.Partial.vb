Option Explicit On
Option Strict On
Imports CompuMaster.Excel.ExcelOps

Namespace Data

    Partial Public Class CmDataXlsEpplusPolyformEditionTest

        Public Sub New()
            AssignLicenseContext()
        End Sub

        Friend Shared Sub AssignLicenseContext()
            ExcelOps.EpplusPolyformExcelDataOperations.LicenseContext = New EpplusPolyformExcelDataOperations.EpplusLicenseActivator(OfficeOpenXml.EPPlusLicenseType.NonCommercialPersonal, "Unit Testing ExcelDataOperations")
        End Sub

    End Class

End Namespace
