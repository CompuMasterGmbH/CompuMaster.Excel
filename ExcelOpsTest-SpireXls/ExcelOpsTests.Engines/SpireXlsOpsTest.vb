Imports NUnit.Framework

Namespace ExcelOpsTests.Engines

    Public Class SpireXlsOpsTest
        Inherits CompuMaster.Excel.Test.ExcelOpsTests.Engines.ExcelOpsTestBase(Of ExcelOps.SpireXlsDataOperations)

        Public Overrides ReadOnly Property ExpectedEngineName As String = "Spire.Xls"

        Public Sub New()
            ExcelOps.SpireXlsDataOperations.AllowInstancingForNonLicencedContextForTestingPurposesOnly = True
        End Sub

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String, disableInitialCalculation As Boolean) As ExcelOps.SpireXlsDataOperations
            Return New ExcelOps.SpireXlsDataOperations(file, mode, [readOnly], passwordForOpening, disableInitialCalculation)
        End Function

        Protected Overrides Function _CreateInstance() As ExcelOps.SpireXlsDataOperations
            Return New ExcelOps.SpireXlsDataOperations()
        End Function

        <Test> Public Overrides Sub CopySheetContent()
            Assert.Throws(Of NotSupportedException)(Sub() MyBase.CopySheetContent())
        End Sub

        Protected Overrides Sub TestInCultureContext_AssignCurrentThreadCulture()
            MyBase.TestInCultureContext_AssignCurrentThreadCulture()
        End Sub

        <SetUp>
        Public Sub Setup()
            ExcelOps.SpireXlsDataOperations.AllowInstancingForNonLicencedContextForTestingPurposesOnly = True
        End Sub

        <Test>
        Public Sub IsLicensedContext()
            'Simulation: license assigned
            ExcelOps.SpireXlsDataOperations.AllowInstancingForNonLicencedContextForTestingPurposesOnly = True
            Assert.NotNull(Me.CreateInstance)
            Assert.NotNull(Me.CreateInstance(Nothing, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, True, Nothing))

            'No license assigned -> instancing must fail
            ExcelOps.SpireXlsDataOperations.AllowInstancingForNonLicencedContextForTestingPurposesOnly = False
            Assert.Throws(Of System.ComponentModel.LicenseException)(Sub() Me.CreateInstance())
            Assert.Throws(Of System.ComponentModel.LicenseException)(Sub() Me.CreateInstance(Nothing, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, True, Nothing))
        End Sub

        <Test>
        Public Sub Utils_IsLicensedContext()
            Assert.False(CompuMaster.Excel.ExcelOps.Utils.IsLicensedContext)
        End Sub

    End Class

End Namespace