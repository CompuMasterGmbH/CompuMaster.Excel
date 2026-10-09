Option Explicit On
Option Strict On

Imports System.IO
Imports System.IO.Compression
Imports System.Xml.Linq
Imports CompuMaster.Epplus4
Imports NUnit.Framework

Namespace Data

    Public Class ExcelPackageZipCompatibilityTest

        ''' <summary>
        ''' Verifies OPC entry names independently of the fork's tolerant reader, for all save paths.
        ''' </summary>
        <TestCase("File")>
        <TestCase("Stream")>
        <TestCase("ByteArray")>
        Public Sub SavedWorkbookUsesRelativeZipEntryNames(savePath As String)
            Dim path = IO.Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") & ".xlsx")
            Try
                Dim bytes As Byte()
                Using package As New ExcelPackage()
                    Dim sheet = package.Workbook.Worksheets.Add("Sample")
                    sheet.Cells(1, 1).Value = "Label"
                    sheet.Cells(2, 1).Value = 42
                    sheet.Cells(3, 1).Formula = "A2+1"
                    sheet.Cells(2, 1).Style.Numberformat.Format = "0.00"
                    sheet.Cells(1, 2).Hyperlink = New Uri("https://example.com/")
                    sheet.Cells(2, 1).AddComment("Synthetic comment", "Test")
                    sheet.Tables.Add(sheet.Cells(1, 1, 2, 1), "SampleTable")
                    Select Case savePath
                        Case "File"
                            package.SaveAs(New FileInfo(path))
                            bytes = File.ReadAllBytes(path)
                        Case "Stream"
                            Using output As New MemoryStream()
                                package.SaveAs(output)
                                bytes = output.ToArray()
                            End Using
                        Case Else
                            bytes = package.GetAsByteArray()
                    End Select
                End Using

                AssertOpcEntryNames(bytes)
                Using input As New MemoryStream(bytes), reopened As New ExcelPackage(input)
                    Assert.That(reopened.Workbook.Worksheets(0).Cells(2, 1).Value, [Is].EqualTo(42))
                    Assert.That(reopened.Workbook.Worksheets(0).Cells(3, 1).Formula, [Is].EqualTo("A2+1"))
                    Assert.That(reopened.Workbook.Worksheets(0).Cells(1, 2).Hyperlink.AbsoluteUri, [Is].EqualTo("https://example.com/"))
                    Assert.That(reopened.Workbook.Worksheets(0).Cells(2, 1).Comment.Text, [Is].EqualTo("Synthetic comment"))
                    Assert.That(reopened.Workbook.Worksheets(0).Tables.Count, [Is].EqualTo(1))
                    Using output As New MemoryStream()
                        reopened.SaveAs(output)
                        AssertOpcEntryNames(output.ToArray())
                    End Using
                End Using
            Finally
                If File.Exists(path) Then File.Delete(path)
            End Try
        End Sub

        ''' <summary>
        ''' Verifies generated files with Microsoft Excel itself when it is available on the test host.
        ''' </summary>
        <NonParallelizable>
        <TestCase("None")>
        <TestCase("Agile")>
        <TestCase("Standard")>
        Public Sub SavedWorkbookCanBeOpenedByMicrosoftExcel(encryption As String)
            If Not CompuMaster.ComInterop.ComTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled("Excel.Application") Then
                Assert.Ignore("Microsoft Excel is not available on this test host.")
            End If
            Dim path = IO.Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") & ".xlsx")
            Dim password As String = Nothing
            Try
                Using package As New ExcelPackage()
                    Dim sheet = package.Workbook.Worksheets.Add("Sample")
                    sheet.Cells(1, 1).Value = "test value"
                    sheet.Cells(2, 1).Value = 42
                    sheet.Cells(3, 1).Formula = "A2+1"
                    If encryption <> "None" Then
                        package.Encryption.Version = If(encryption = "Agile", EncryptionVersion.Agile, EncryptionVersion.Standard)
                        password = "correct password"
                        package.Encryption.Password = password
                    End If
                    package.SaveAs(New FileInfo(path))
                End Using
                Using app As New CompuMaster.Excel.MsExcelCom.MsExcelApplicationWrapper()
                    Dim workbook = app.Workbooks.Open(path, True, password)
                    Try
                        Assert.That(workbook.Name, [Is].EqualTo(IO.Path.GetFileName(path)))
                    Finally
                        workbook.CloseAndDispose()
                    End Try
                End Using
            Finally
                If File.Exists(path) Then File.Delete(path)
                CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
            End Try
        End Sub

        Private Shared Sub AssertOpcEntryNames(bytes As Byte())
            Using input As New MemoryStream(bytes), archive As New ZipArchive(input, ZipArchiveMode.Read)
                For Each entry In archive.Entries
                    Assert.That(entry.FullName, Does.Not.StartWith("/"), entry.FullName)
                    Assert.That(entry.FullName, Does.Not.Contain("\"), entry.FullName)
                Next
                For Each required In {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels", "xl/styles.xml", "xl/sharedStrings.xml"}
                    Assert.That(archive.GetEntry(required), [Is].Not.Null, required)
                Next

                ' OPC PartName values retain their leading slash; only ZIP entry names are relative.
                Using source = archive.GetEntry("[Content_Types].xml").Open()
                    Dim types = XDocument.Load(source)
                    Dim ns As XNamespace = "http://schemas.openxmlformats.org/package/2006/content-types"
                    For Each part In types.Root.Elements(ns + "Override")
                        Dim name = part.Attribute("PartName").Value
                        Assert.That(name, Does.StartWith("/"))
                        Assert.That(archive.GetEntry(name.Substring(1)), [Is].Not.Null, name)
                    Next
                End Using

                Dim relNs As XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
                For Each entry In archive.Entries.Where(Function(e) e.FullName.EndsWith(".rels", StringComparison.Ordinal))
                    Dim folder = entry.FullName.Substring(0, entry.FullName.LastIndexOf("_rels/", StringComparison.Ordinal))
                    Dim baseUri As New Uri("https://opc.example/" & folder)
                    Using source = entry.Open()
                        For Each relationship In XDocument.Load(source).Root.Elements(relNs + "Relationship")
                            If CStr(relationship.Attribute("TargetMode")) = "External" Then Continue For
                            Dim target As New Uri(baseUri, relationship.Attribute("Target").Value)
                            Assert.That(archive.GetEntry(target.AbsolutePath.TrimStart("/"c)), [Is].Not.Null, entry.FullName & " -> " & target.AbsolutePath)
                        Next
                    End Using
                Next
            End Using
        End Sub

    End Class

End Namespace
