Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework

Namespace ExcelOpsTests.Engines

    Public Class SpireXlsOpsTest
        Inherits CompuMaster.Excel.Test.ExcelOpsTests.Engines.ExcelOpsTestBase(Of ExcelOps.SpireXlsDataOperations)

        Public Overrides ReadOnly Property ExpectedEngineName As String = "Spire.Xls"

        Public Sub New()
            ExcelOps.SpireXlsDataOperations.AllowInstancingForNonLicencedContextForTestingPurposesOnly = True
        End Sub

        Protected Overrides Function _CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.SpireXlsDataOperations
            Return New ExcelOps.SpireXlsDataOperations(file, mode, options)
        End Function

        Protected Overrides Function _CreateInstance() As ExcelOps.SpireXlsDataOperations
            Return New ExcelOps.SpireXlsDataOperations(ExcelDataOperationsBase.OpenMode.Uninitialized)
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
            Assert.NotNull(Me.CreateInstance(Nothing, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelOps.ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile)))

            'No license assigned -> instancing must fail
            ExcelOps.SpireXlsDataOperations.AllowInstancingForNonLicencedContextForTestingPurposesOnly = False
            Assert.Throws(Of System.ComponentModel.LicenseException)(Sub() Me.CreateInstance())
            Assert.Throws(Of System.ComponentModel.LicenseException)(Sub() Me.CreateInstance(Nothing, ExcelOps.ExcelDataOperationsBase.OpenMode.CreateFile, New ExcelOps.ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile)))
        End Sub

        <Test>
        Public Sub Utils_IsLicensedContext()
            Assert.False(CompuMaster.Excel.ExcelOps.Utils.IsLicensedContext)
        End Sub

        Protected Overrides Function _CreateInstance(data() As Byte, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.SpireXlsDataOperations
            Return New ExcelOps.SpireXlsDataOperations(data, options)
        End Function

        Protected Overrides Function _CreateInstance(data As IO.Stream, options As ExcelOps.ExcelDataOperationsOptions) As ExcelOps.SpireXlsDataOperations
            Return New ExcelOps.SpireXlsDataOperations(data, options)
        End Function

    End Class

End Namespace