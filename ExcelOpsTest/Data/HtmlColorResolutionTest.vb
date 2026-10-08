Option Explicit On
Option Strict On

Imports System.IO
Imports System.IO.Compression
Imports System.Xml
Imports CompuMaster.Epplus4
Imports NUnit.Framework

Namespace Data

    Public Class HtmlColorResolutionTest

        Private Class ColorHelperAccess
            Inherits ExcelOps.EpplusFreeExcelDataOperations

            Private Sub New()
                MyBase.New(ExcelOps.ExcelDataOperationsBase.OpenMode.Uninitialized)
            End Sub

            Friend Shared Function TintColor(hex As String, tint As Double) As String
                Return ApplyTint(hex, tint)
            End Function

            Friend Shared Function Palette(xml As XmlDocument) As String()
                Return ReadThemeColorPalette(xml)
            End Function
        End Class

        ''' <summary>
        ''' Checks additional tint rounding boundaries captured directly from Excel 16.0, build 20430.
        ''' </summary>
        <TestCase("#4F81BD", 0.12338633381145665, "#6591C5")>
        <TestCase("#4F81BD", 0.33332316049684135, "#89ABD3")>
        <TestCase("#4F81BD", 0.40119632557145907, "#96B4D8")>
        <TestCase("#4F81BD", -0.12338633381145665, "#4070AA")>
        <TestCase("#4F81BD", -0.33332316049684135, "#315582")>
        <TestCase("#4F81BD", -0.7510910367137669, "#122030")>
        <TestCase("#A5A5A5", 0.59999389629810485, "#DBDBDB")>
        <TestCase("#FF0000", 0.12338633381145665, "#FF2020")>
        <TestCase("#FF0000", 0.33332316049684135, "#FF5555")>
        <TestCase("#4F81BD", 0.0, "#4F81BD")>
        <TestCase("#4F81BD", 1.0, "#FFFFFF")>
        <TestCase("#4F81BD", -1.0, "#000000")>
        Public Sub TintMatchesExcelRounding(hex As String, tint As Double, expected As String)
            Assert.That(ColorHelperAccess.TintColor(hex, tint), [Is].EqualTo(expected))
        End Sub

        ''' <summary>
        ''' Rejects invalid tint inputs instead of emitting invalid CSS or arithmetic overflows.
        ''' </summary>
        <TestCase(-1.01)>
        <TestCase(1.01)>
        <TestCase(Double.NaN)>
        <TestCase(Double.PositiveInfinity)>
        <TestCase(Double.NegativeInfinity)>
        Public Sub InvalidTintIsRejected(tint As Double)
            Assert.Throws(Of ArgumentOutOfRangeException)(Sub() ColorHelperAccess.TintColor("#FF0000", tint))
        End Sub

        ''' <summary>
        ''' Retains the legacy palette only when no workbook theme is present.
        ''' </summary>
        <Test>
        Public Sub MissingThemeUsesCompatibleFallback()
            Dim colors = ColorHelperAccess.Palette(Nothing)
            Assert.That(colors.Length, [Is].EqualTo(12))
            Assert.That(colors(0), [Is].EqualTo("#FFFFFF"))
            Assert.That(colors(1), [Is].EqualTo("#000000"))
            Assert.That(colors(5), [Is].EqualTo("#C0504D"))
            Using package As New ExcelPackage()
                Assert.That(package.Workbook.ThemeXml, [Is].Null)
            End Using
        End Sub

        ''' <summary>
        ''' Ensures consumers cannot modify the stored theme through the detached XML accessor.
        ''' </summary>
        <Test>
        Public Sub ThemeXmlReturnsDetachedCopy()
            Dim filePath = TestEnvironment.FullPathOfExistingTestFile("test_data", "HtmlExportThemeExcelRed.xlsx")
            Using package As New ExcelPackage(New FileInfo(filePath))
                Dim xml = package.Workbook.ThemeXml
                Assert.That(ColorHelperAccess.Palette(xml)(5), [Is].EqualTo("#FF0000"))
                xml.DocumentElement.RemoveAll()
                Assert.That(ColorHelperAccess.Palette(package.Workbook.ThemeXml)(5), [Is].EqualTo("#FF0000"))
            End Using
        End Sub

        ''' <summary>
        ''' Resolves the embedded theme via the workbook relationship rather than a hardcoded ZIP path.
        ''' </summary>
        <Test>
        Public Sub ThemePartCanHaveNonDefaultName()
            Dim filePath = TestEnvironment.FullPathOfExistingTestFile("test_data", "HtmlExportThemeExcelRed.xlsx")
            Using modified As New MemoryStream()
                Dim bytes = File.ReadAllBytes(filePath)
                modified.Write(bytes, 0, bytes.Length)
                modified.Position = 0
                Using archive As New ZipArchive(modified, ZipArchiveMode.Update, True)
                    Dim theme = archive.GetEntry("xl/theme/theme1.xml")
                    Using original = theme.Open(), renamed = archive.CreateEntry("xl/theme/custom-colors.xml").Open()
                        original.CopyTo(renamed)
                    End Using
                    theme.Delete()
                    Dim relationships = archive.GetEntry("xl/_rels/workbook.xml.rels")
                    Dim xml As New XmlDocument()
                    Using input = relationships.Open()
                        xml.Load(input)
                    End Using
                    For Each relationship As XmlElement In xml.DocumentElement.ChildNodes
                        If relationship.GetAttribute("Type").EndsWith("/theme", StringComparison.Ordinal) Then
                            relationship.SetAttribute("Target", "theme/custom-colors.xml")
                        End If
                    Next
                    relationships.Delete()
                    Using output = archive.CreateEntry("xl/_rels/workbook.xml.rels").Open()
                        xml.Save(output)
                    End Using
                End Using
                modified.Position = 0
                Using package As New ExcelPackage(modified)
                    Assert.That(ColorHelperAccess.Palette(package.Workbook.ThemeXml)(5), [Is].EqualTo("#FF0000"))
                End Using
            End Using
        End Sub

    End Class

End Namespace
