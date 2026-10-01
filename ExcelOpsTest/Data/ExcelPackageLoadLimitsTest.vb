Option Explicit On
Option Strict On

Imports System.IO
Imports System.IO.Compression
Imports System.Data
Imports System.Text
Imports ZipCompressionLevel = System.IO.Compression.CompressionLevel
Imports CompuMaster.Epplus4
Imports NUnit.Framework

Namespace Data

    Public Class ExcelPackageLoadLimitsTest

        ''' <summary>
        ''' Verifies that oversized compressed inputs are rejected before loading.
        ''' </summary>
        <Test>
        Public Sub RejectsCompressedInputAboveConfiguredLimit()
            Dim inputPath = TestEnvironment.FullPathOfExistingTestFile("test_data", "SampleTable01.xlsx")
            Dim limits As New ExcelPackageLoadLimits(maxInputBytes:=1)

            Dim exception = Assert.Throws(Of InvalidDataException)(
                Sub()
                    Using workbook = ExcelPackage.OpenWithLoadLimits(New FileInfo(inputPath), limits)
                    End Using
                End Sub)

            Assert.That(exception.Message, Does.Contain("compressed input size"))
        End Sub

        ''' <summary>
        ''' Verifies that every ZIP entry is counted before extraction.
        ''' </summary>
        <Test>
        Public Sub RejectsTooManyZipEntries()
            Dim exception = Assert.Throws(Of InvalidDataException)(
                Sub() OpenWorkbook(CreateSampleWorkbook(), New ExcelPackageLoadLimits(maxZipEntries:=1)))

            Assert.That(exception.Message, Does.Contain("entry count"))
        End Sub

        ''' <summary>
        ''' Verifies that per-entry, total, and XML limits are enforced independently.
        ''' </summary>
        <Test>
        Public Sub RejectsOversizedEntriesAndXml()
            Dim workbook = CreateSampleWorkbook()

            Assert.That(Assert.Throws(Of InvalidDataException)(
                Sub() OpenWorkbook(workbook, New ExcelPackageLoadLimits(maxEntryBytes:=1))).Message,
                Does.Contain("entry size"))
            Assert.That(Assert.Throws(Of InvalidDataException)(
                Sub() OpenWorkbook(workbook, New ExcelPackageLoadLimits(maxTotalUncompressedBytes:=1))).Message,
                Does.Contain("total uncompressed size"))
            Assert.That(Assert.Throws(Of InvalidDataException)(
                Sub() OpenWorkbook(workbook, New ExcelPackageLoadLimits(maxXmlBytes:=1))).Message,
                Does.Contain("entry size"))
        End Sub

        ''' <summary>
        ''' Verifies that a highly compressible, unreferenced ZIP entry cannot bypass the ratio limit.
        ''' </summary>
        <Test>
        Public Sub RejectsZipBombEntry()
            Dim workbook = AddEntry(CreateSampleWorkbook(), "xl/media/padding.bin", New Byte(1024 * 1024 - 1) {}, ZipCompressionLevel.Optimal)

            Dim exception = Assert.Throws(Of InvalidDataException)(Sub() OpenWorkbook(workbook, ExcelPackageLoadLimits.Default))

            Assert.That(exception.Message, Does.Contain("compression ratio"))
        End Sub

        ''' <summary>
        ''' Verifies that the data facade passes explicit load limits to the ZIP reader.
        ''' </summary>
        <Test>
        Public Sub DataFacadeReadOptionsApplyPackageLimits()
            Dim inputPath = TestEnvironment.FullPathOfExistingTestFile("test_data", "SampleTable01.xlsx")
            Dim options As New CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadOptions(True, 0, New ExcelPackageLoadLimits(maxInputBytes:=1))

            Dim exception = Assert.Throws(Of InvalidDataException)(
                Sub() CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFileWithOptions(inputPath, options))

            Assert.That(exception.Message, Does.Contain("compressed input size"))
            Assert.That(Assert.Throws(Of InvalidDataException)(
                Sub() CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataSetFromXlsFileWithOptions(inputPath, options)).Message,
                Does.Contain("compressed input size"))
            Assert.That(Assert.Throws(Of InvalidDataException)(
                Sub() CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFileWithOptions(inputPath, "Sample", options)).Message,
                Does.Contain("compressed input size"))
            Assert.That(Assert.Throws(Of InvalidDataException)(
                Sub() CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFileWithOptions(inputPath, "Sample", options, New DataTable())).Message,
                Does.Contain("compressed input size"))
        End Sub

        ''' <summary>
        ''' Verifies that relationship-targeted XML remains limited when it has a non-XML file extension.
        ''' </summary>
        <Test>
        Public Sub RejectsRenamedXmlPartAboveLimit()
            Dim workbook = CreateSampleWorkbook()
            Dim randomBytes(24 * 1024 - 1) As Byte
            Dim random As New Random(1234)
            random.NextBytes(randomBytes)
            Using output As New MemoryStream()
                output.Write(workbook, 0, workbook.Length)
                output.Position = 0
                Using archive As New ZipArchive(output, ZipArchiveMode.Update, True)
                    Dim original = archive.Entries.FirstOrDefault(Function(e) e.FullName.EndsWith("sheet1.xml", StringComparison.OrdinalIgnoreCase))
                    Assert.That(original, [Is].Not.Null, String.Join(", ", archive.Entries.Select(Function(e) e.FullName)))
                    Dim xml As String
                    Using reader As New StreamReader(original.Open(), Encoding.UTF8)
                        xml = reader.ReadToEnd()
                    End Using
                    Dim renamedPath = original.FullName.Replace("sheet1.xml", "sheet1.bin")
                    original.Delete()
                    Using writer As New StreamWriter(archive.CreateEntry(renamedPath, ZipCompressionLevel.Optimal).Open(), New UTF8Encoding(False))
                        writer.Write(xml)
                        writer.Write("<!--")
                        writer.Write(Convert.ToBase64String(randomBytes))
                        writer.Write("-->")
                    End Using
                    RenameWorksheetReference(archive, "[Content_Types].xml")
                    RenameWorksheetReference(archive, "xl/_rels/workbook.xml.rels")
                End Using
                workbook = output.ToArray()
            End Using

            Dim exception = Assert.Throws(Of InvalidDataException)(
                Sub()
                    Using input As New MemoryStream(workbook)
                        Using package = ExcelPackage.OpenWithLoadLimits(input, New ExcelPackageLoadLimits(maxXmlBytes:=16 * 1024))
                            Dim value = package.Workbook.Worksheets(0).Cells(1, 1).Value
                        End Using
                    End Using
                End Sub)
            Assert.That(exception.Message, Does.Contain("XML part"))
        End Sub

        ''' <summary>
        ''' Verifies that a template stream is read from its beginning and is subject to input limits.
        ''' </summary>
        <Test>
        Public Sub StreamTemplateAtEndHonorsInputLimits()
            Dim source = CreateSampleWorkbook()
            Using template As New MemoryStream(source), output As New MemoryStream()
                template.Position = template.Length
                Dim exception = Assert.Throws(Of InvalidDataException)(
                    Sub()
                        Using package = ExcelPackage.CreateFromTemplateWithLoadLimits(output, template, New ExcelPackageLoadLimits(maxInputBytes:=1))
                        End Using
                    End Sub)
                Assert.That(exception.Message, Does.Contain("compressed input size"))
            End Using
        End Sub

        ''' <summary>
        ''' Verifies that an encrypted workbook is checked before decryption and still opens with valid limits.
        ''' </summary>
        <Test>
        Public Sub EncryptedWorkbookHonorsInputLimits()
            Dim path = TestEnvironment.FullPathOfExistingTestFile("test_data", "WorkbookPasswordProtected.xlsx")
            Dim encrypted = File.ReadAllBytes(path)

            Using input As New MemoryStream(encrypted)
                Assert.That(Assert.Throws(Of InvalidDataException)(
                    Sub()
                        Using package = ExcelPackage.OpenWithLoadLimits(input, New ExcelPackageLoadLimits(maxInputBytes:=1), "test")
                        End Using
                    End Sub).Message, Does.Contain("compressed input size"))
            End Using
            Using input As New MemoryStream(encrypted)
                Using package = ExcelPackage.OpenWithLoadLimits(input, ExcelPackageLoadLimits.Default, "test")
                    Assert.That(package.Workbook.Worksheets.Count, [Is].GreaterThan(0))
                End Using
            End Using
            Using input As New MemoryStream(encrypted)
                Assert.That(Assert.Throws(Of InvalidDataException)(
                    Sub()
                        Using package = ExcelPackage.OpenWithLoadLimits(input, New ExcelPackageLoadLimits(maxPasswordHashIterations:=1), "test")
                        End Using
                    End Sub).Message, Does.Contain("password-hash iteration"))
            End Using
            Using input As New MemoryStream(encrypted)
                Assert.That(Assert.Throws(Of InvalidDataException)(
                    Sub()
                        Using package = ExcelPackage.OpenWithLoadLimits(input, New ExcelPackageLoadLimits(maxEncryptionMetadataBytes:=1), "test")
                        End Using
                    End Sub).Message, Does.Contain("encryption metadata"))
            End Using
        End Sub

#If NET8_0 Then
        ''' <summary>
        ''' Verifies that a valid XLSX larger than 50 MiB remains below the default limits.
        ''' </summary>
        <Test>
        Public Sub DefaultLimitsAcceptFiftyMegabyteWorkbook()
            Dim payload(50 * 1024 * 1024 - 1) As Byte
            Dim random As New Random(1234)
            random.NextBytes(payload)
            Dim workbook = AddEntry(CreateSampleWorkbook(), "xl/media/large.bin", payload, ZipCompressionLevel.NoCompression)

            Assert.That(workbook.Length, [Is].GreaterThan(50 * 1024 * 1024))
            OpenWorkbook(workbook, ExcelPackageLoadLimits.Default)
        End Sub
#End If

        Private Shared Function CreateSampleWorkbook() As Byte()
            Using output As New MemoryStream()
                Using workbook As New ExcelPackage()
                    workbook.Workbook.Worksheets.Add("Sample").Cells(1, 1).Value = "Safe"
                    workbook.SaveAs(output)
                End Using
                Return output.ToArray()
            End Using
        End Function

        Private Shared Function AddEntry(workbook As Byte(), name As String, payload As Byte(), compression As ZipCompressionLevel) As Byte()
            Using output As New MemoryStream()
                output.Write(workbook, 0, workbook.Length)
                output.Position = 0
                Using archive As New ZipArchive(output, ZipArchiveMode.Update, True)
                    Using target = archive.CreateEntry(name, compression).Open()
                        target.Write(payload, 0, payload.Length)
                    End Using
                End Using
                Return output.ToArray()
            End Using
        End Function

        Private Shared Sub RenameWorksheetReference(archive As ZipArchive, entryName As String)
            Dim original = archive.Entries.FirstOrDefault(Function(e) e.FullName.Replace("\", "/").TrimStart("/"c).Equals(entryName, StringComparison.OrdinalIgnoreCase))
            Assert.That(original, [Is].Not.Null, String.Join(", ", archive.Entries.Select(Function(e) e.FullName)))
            Dim xml As String
            Using reader As New StreamReader(original.Open(), Encoding.UTF8)
                xml = reader.ReadToEnd()
            End Using
            Assert.That(xml, Does.Contain("sheet1.xml"))
            Dim originalPath = original.FullName
            original.Delete()
            Using writer As New StreamWriter(archive.CreateEntry(originalPath).Open(), New UTF8Encoding(False))
                writer.Write(xml.Replace("sheet1.xml", "sheet1.bin"))
            End Using
        End Sub

        Private Shared Sub OpenWorkbook(workbook As Byte(), limits As ExcelPackageLoadLimits)
            Using input As New MemoryStream(workbook)
                Using package = ExcelPackage.OpenWithLoadLimits(input, limits)
                    Assert.That(package.Workbook.Worksheets.Count, [Is].EqualTo(1))
                End Using
            End Using
        End Sub

    End Class

End Namespace
