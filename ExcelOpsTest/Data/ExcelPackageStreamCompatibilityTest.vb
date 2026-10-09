Option Explicit On
Option Strict On

Imports System.IO
Imports CompuMaster.Epplus4
Imports CompuMaster.Excel.ExcelOps
Imports NUnit.Framework

Namespace Data

    Public Class ExcelPackageStreamCompatibilityTest

        ''' <summary>
        ''' Verifies that external streams returning short reads do not silently corrupt byte-array output.
        ''' </summary>
        <TestCase(False)>
        <TestCase(True)>
        Public Sub GetAsByteArrayPreservesShortReadStream(encrypted As Boolean)
            Dim bytes As Byte()
            Using output As New ShortReadMemoryStream(), package As New ExcelPackage(output)
                package.Workbook.Worksheets.Add("Sample").Cells(1, 1).Value = "complete value"
                If encrypted Then package.Encryption.Password = "password"
                bytes = package.GetAsByteArray()
            End Using
            Using input As New MemoryStream(bytes), reopened As New ExcelPackage(input, If(encrypted, "password", Nothing))
                Assert.That(reopened.Workbook.Worksheets(0).Cells(1, 1).Value, [Is].EqualTo("complete value"))
            End Using
        End Sub

        ''' <summary>
        ''' Rejects a premature stream end instead of returning a partially zero-filled workbook.
        ''' </summary>
        <Test>
        Public Sub GetAsByteArrayRejectsPrematureStreamEnd()
            Using output As New ShortReadMemoryStream(maximumReadableBytes:=3), package As New ExcelPackage(output)
                package.Workbook.Worksheets.Add("Sample").Cells(1, 1).Value = "complete value"
                Dim exception = Assert.Throws(Of InvalidDataException)(Sub() package.GetAsByteArray())
                Assert.That(exception.Message, Does.Contain("package stream ended"))
            End Using
        End Sub

        ''' <summary>
        ''' Verifies that OLE header detection handles short reads in both synchronized EPPlus engines.
        ''' </summary>
        <TestCase(False, "WorkbookNoPassword.xls")>
        <TestCase(True, "WorkbookNoPassword.xls")>
        <TestCase(False, "WorkbookPasswordProtected.xlsx")>
        <TestCase(True, "WorkbookPasswordProtected.xlsx")>
        Public Sub ShortReadInputPreservesFormatAndPasswordExceptions(polyform As Boolean, fileName As String)
            Dim data = File.ReadAllBytes(TestEnvironment.FullPathOfExistingTestFile("test_data", fileName))
            Using input As New ShortReadMemoryStream(data)
                Dim action As TestDelegate =
                    Sub()
                        Dim workbook As ExcelDataOperationsBase = Nothing
                        Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly, "wrong password")
                        Try
                            If polyform Then
                                workbook = New EpplusPolyformExcelDataOperations(input, options)
                            Else
                                workbook = New EpplusFreeExcelDataOperations(input, options)
                            End If
                        Finally
                            If workbook IsNot Nothing Then workbook.Close()
                        End Try
                    End Sub
                If fileName.EndsWith(".xls", StringComparison.Ordinal) Then
                    Assert.Throws(Of BinaryXlsFileNotSupportedException)(action)
                    Assert.That(input.Position, [Is].EqualTo(0), "Format detection must restore the caller's stream position.")
                Else
                    Assert.Throws(Of FilePasswordProtectedMismatchException)(action)
                End If
            End Using
        End Sub

        Private NotInheritable Class ShortReadMemoryStream
            Inherits MemoryStream

            Private ReadOnly _MaximumReadableBytes As Long

            Public Sub New(Optional data As Byte() = Nothing, Optional maximumReadableBytes As Long = Long.MaxValue)
                _MaximumReadableBytes = maximumReadableBytes
                If data IsNot Nothing Then
                    Write(data, 0, data.Length)
                    Position = 0
                End If
            End Sub

            ''' <inheritdoc/>
            Public Overrides Function Read(buffer As Byte(), offset As Integer, count As Integer) As Integer
                If Position >= _MaximumReadableBytes Then Return 0
                Return MyBase.Read(buffer, offset, CInt(Math.Min(Math.Min(count, 3), _MaximumReadableBytes - Position)))
            End Function
        End Class

    End Class

End Namespace
