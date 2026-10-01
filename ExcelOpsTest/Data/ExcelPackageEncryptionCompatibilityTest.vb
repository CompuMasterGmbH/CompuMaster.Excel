Option Explicit On
Option Strict On

Imports System.IO
Imports CompuMaster.Epplus4
Imports NUnit.Framework

Namespace Data

    Public Class ExcelPackageEncryptionCompatibilityTest

        ''' <summary>
        ''' Verifies that a newly encrypted workbook can be reopened by the ZIP-archive reader.
        ''' </summary>
        <Test>
        Public Sub EncryptedSaveAsStreamCanBeReopened()
            Dim encrypted As Byte()
            Using output As New MemoryStream()
                Using package As New ExcelPackage()
                    package.Workbook.Worksheets.Add("Secret").Cells(1, 1).Value = "Value"
                    package.Encryption.Password = "password"
                    package.SaveAs(output)
                End Using
                encrypted = output.ToArray()
            End Using

            Using input As New MemoryStream(encrypted)
                Using package = ExcelPackage.OpenWithLoadLimits(input, ExcelPackageLoadLimits.Default, "password")
                    Assert.That(package.Workbook.Worksheets(0).Cells(1, 1).Value, [Is].EqualTo("Value"))
                End Using
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
