Option Explicit On
Option Strict On

Imports System.IO
Imports CompuMaster.Epplus4
Imports NUnit.Framework

Namespace Data

    Public Class ExcelPackageEncryptionCompatibilityTest

        <OneTimeSetUp>
        Public Sub ReportRuntime()
            TestContext.Progress.WriteLine("Encryption test runtime: " & System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription)
        End Sub

        ''' <summary>
        ''' Verifies that a newly encrypted workbook can be reopened by the ZIP-archive reader.
        ''' </summary>
        <TestCase(EncryptionVersion.Agile, 0)>
        <TestCase(EncryptionVersion.Agile, 128)>
        <TestCase(EncryptionVersion.Agile, 8192)>
        <TestCase(EncryptionVersion.Standard, 0)>
        <TestCase(EncryptionVersion.Standard, 128)>
        <TestCase(EncryptionVersion.Standard, 8192)>
        Public Sub EncryptedSaveAsStreamCanBeReopened(version As EncryptionVersion, payloadSize As Integer)
            Dim encrypted As Byte()
            Using output As New MemoryStream()
                Using package As New ExcelPackage()
                    PopulateWorkbook(package, payloadSize)
                    package.Encryption.Version = version
                    package.Encryption.Password = "password"
                    package.SaveAs(output)
                End Using
                encrypted = output.ToArray()
            End Using

            Using input As New MemoryStream(encrypted)
                Using package = ExcelPackage.OpenWithLoadLimits(input, ExcelPackageLoadLimits.Default, "password")
                    AssertWorkbook(package, payloadSize)
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Covers the consuming application's file-based roundtrip and multiple encrypted segments.
        ''' </summary>
        <TestCase(EncryptionVersion.Agile, 0)>
        <TestCase(EncryptionVersion.Agile, 128)>
        <TestCase(EncryptionVersion.Agile, 8192)>
        <TestCase(EncryptionVersion.Standard, 0)>
        <TestCase(EncryptionVersion.Standard, 128)>
        <TestCase(EncryptionVersion.Standard, 8192)>
        Public Sub EncryptedSaveAsFileCanBeReopened(version As EncryptionVersion, payloadSize As Integer)
            Dim path = IO.Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") & ".xlsx")
            Try
                Using package As New ExcelPackage()
                    PopulateWorkbook(package, payloadSize)
                    package.Encryption.Version = version
                    package.Encryption.IsEncrypted = True
                    package.Encryption.Password = "correct password"
                    package.SaveAs(New FileInfo(path))
                End Using

                Assert.That(BitConverter.ToUInt64(File.ReadAllBytes(path), 0), [Is].EqualTo(&HE11AB1A1E011CFD0UL), "The file must be an encrypted compound document.")
                Using reopened As New ExcelPackage(New FileInfo(path), "correct password")
                    AssertWorkbook(reopened, payloadSize)
                End Using
                Using reopened = ExcelPackage.OpenWithLoadLimits(New FileInfo(path), ExcelPackageLoadLimits.Default, "correct password")
                    AssertWorkbook(reopened, payloadSize)
                End Using
                Using fromTemplate As New ExcelPackage(New FileInfo(path), True, "correct password")
                    AssertWorkbook(fromTemplate, payloadSize)
                End Using
                Using input As New MemoryStream(File.ReadAllBytes(path)), output As New MemoryStream()
                    Using fromTemplate = ExcelPackage.CreateFromTemplateWithLoadLimits(output, input, ExcelPackageLoadLimits.Default, "correct password")
                        AssertWorkbook(fromTemplate, payloadSize)
                    End Using
                End Using
            Finally
                If File.Exists(path) Then File.Delete(path)
            End Try
        End Sub

        Private Shared Sub PopulateWorkbook(package As ExcelPackage, payloadSize As Integer)
            Dim sheet = package.Workbook.Worksheets.Add("FiBu")
            sheet.Cells(1, 1).Value = "test value"
            sheet.Cells(2, 1).Value = CreatePayload(payloadSize)
            sheet.Cells(3, 1).Formula = "1+2"
        End Sub

        Private Shared Function CreatePayload(size As Integer) As String
            Dim bytes(size - 1) As Byte
            Dim random As New Random(31)
            random.NextBytes(bytes)
            Return Convert.ToBase64String(bytes)
        End Function

        Private Shared Sub AssertWorkbook(package As ExcelPackage, payloadSize As Integer)
            Dim sheet = package.Workbook.Worksheets("FiBu")
            Assert.That(sheet.Cells(1, 1).Value, [Is].EqualTo("test value"))
            Assert.That(sheet.Cells(2, 1).Text, [Is].EqualTo(CreatePayload(payloadSize)))
            Assert.That(sheet.Cells(3, 1).Formula, [Is].EqualTo("1+2"))
        End Sub

        ''' <summary>
        ''' Verifies all supported AES key sizes on encrypted byte-array output.
        ''' </summary>
        <TestCase(EncryptionVersion.Agile, EncryptionAlgorithm.AES128)>
        <TestCase(EncryptionVersion.Agile, EncryptionAlgorithm.AES192)>
        <TestCase(EncryptionVersion.Agile, EncryptionAlgorithm.AES256)>
        <TestCase(EncryptionVersion.Standard, EncryptionAlgorithm.AES128)>
        <TestCase(EncryptionVersion.Standard, EncryptionAlgorithm.AES192)>
        <TestCase(EncryptionVersion.Standard, EncryptionAlgorithm.AES256)>
        Public Sub EncryptedByteArrayPreservesAllAesKeySizes(version As EncryptionVersion, algorithm As EncryptionAlgorithm)
            Dim bytes As Byte()
            Using package As New ExcelPackage()
                PopulateWorkbook(package, 8192)
                package.Encryption.Version = version
                package.Encryption.Algorithm = algorithm
                bytes = package.GetAsByteArray("password")
            End Using
            Using input As New MemoryStream(bytes), reopened As New ExcelPackage(input, "password")
                AssertWorkbook(reopened, 8192)
            End Using
        End Sub

        ''' <summary>
        ''' Preserves the validated legacy zero-offset repair while fixing newly generated encrypted output.
        ''' </summary>
        <TestCase(False)>
        <TestCase(True)>
        Public Sub LegacyZeroOffsetCentralDirectoryCanStillBeRead(clearCounts As Boolean)
            Dim bytes As Byte()
            Using package As New ExcelPackage()
                PopulateWorkbook(package, 0)
                bytes = package.GetAsByteArray()
            End Using
            Dim endOffset = bytes.Length - 22
            Array.Clear(bytes, endOffset + 16, 4)
            If clearCounts Then Array.Clear(bytes, endOffset + 8, 8)
            Using input As New MemoryStream(bytes), reopened As New ExcelPackage(input)
                AssertWorkbook(reopened, 0)
            End Using
        End Sub

        ''' <summary>
        ''' Verifies that central-directory repair rejects inconsistent entry counts.
        ''' </summary>
        <Test>
        Public Sub CentralDirectoryRepairRejectsInconsistentCount()
            Dim bytes As Byte()
            Using output As New MemoryStream()
                Using package As New ExcelPackage()
                    package.Workbook.Worksheets.Add("Sample")
                    package.SaveAs(output)
                End Using
                bytes = output.ToArray()
            End Using

            Dim endOffset = bytes.Length - 22
            Assert.That(BitConverter.ToUInt32(bytes, endOffset), [Is].EqualTo(&H6054B50UI))
            Array.Copy(BitConverter.GetBytes(CUShort(1)), 0, bytes, endOffset + 8, 2)
            Array.Copy(BitConverter.GetBytes(CUShort(1)), 0, bytes, endOffset + 10, 2)
            Array.Copy(BitConverter.GetBytes(CUInt(0)), 0, bytes, endOffset + 16, 4)

            Using input As New MemoryStream(bytes)
                Assert.Throws(Of InvalidDataException)(
                    Sub()
                        Using package = ExcelPackage.OpenWithLoadLimits(input, ExcelPackageLoadLimits.Default)
                        End Using
                    End Sub)
            End Using
        End Sub

    End Class

End Namespace
