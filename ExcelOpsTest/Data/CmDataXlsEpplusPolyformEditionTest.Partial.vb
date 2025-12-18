Option Explicit On
Option Strict On

Namespace Data

    Partial Public Class CmDataXlsEpplusPolyformEditionTest

        Public Sub New()
            AssignLicenseContext()
        End Sub

        Friend Shared Sub AssignLicenseContext()
            ExcelOps.EpplusPolyformExcelDataOperations.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
        End Sub

    End Class

End Namespace
