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

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String, disableInitialCalculation As Boolean) As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations(file, mode, [readOnly], passwordForOpening, disableInitialCalculation)
        End Function

        Protected Overrides Function _CreateInstance() As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations()
        End Function

        <Test> Public Overrides Sub CopySheetContent()
            Assert.Throws(Of NotSupportedException)(Sub() MyBase.CopySheetContent())
        End Sub

        Protected Overrides Function _CreateInstance(data() As Byte, passwordForOpening As String, disableCalculationEngine As Boolean) As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations(data, passwordForOpening, disableCalculationEngine)
        End Function

        Protected Overrides Function _CreateInstance(data As IO.Stream, passwordForOpening As String, disableCalculationEngine As Boolean) As ExcelOps.EpplusPolyformExcelDataOperations
            Return New ExcelOps.EpplusPolyformExcelDataOperations(data, passwordForOpening, disableCalculationEngine)
        End Function

    End Class

End Namespace