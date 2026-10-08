Option Explicit On
Option Strict On

Imports System.Linq
Imports System.Xml.Linq
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Epplus4
Imports NUnit.Framework

Namespace Data

    <NonParallelizable>
    Public Class HtmlTableHeaderIndexesTest

        Private Shared ReadOnly LegacyHeaderProperty As System.Reflection.PropertyInfo =
            GetType(HtmlSheetExportOptions).GetProperty("ConsiderRowIndexesAsTableHeader")

        ''' <summary>
        ''' Ensures recompilation rejects the ambiguous legacy name and explains the migration.
        ''' </summary>
        <Test>
        Public Sub LegacyHeaderPropertyIsObsoleteWithCompilerError()
            Dim attribute = DirectCast(System.Attribute.GetCustomAttribute(LegacyHeaderProperty, GetType(ObsoleteAttribute)), ObsoleteAttribute)
            Assert.That(attribute, [Is].Not.Null)
            Assert.That(attribute.IsError, [Is].True)
            Assert.That(attribute.Message, Does.Contain(NameOf(HtmlSheetExportOptions.TableHeaderRowIndexes)))
            Assert.That(attribute.Message, Does.Contain("subtract 1"))
            Assert.That(LegacyHeaderProperty.PropertyType, [Is].EqualTo(GetType(List(Of Integer))))
            Assert.That(LegacyHeaderProperty.CanRead AndAlso LegacyHeaderProperty.CanWrite, [Is].True)
        End Sub

        ''' <summary>
        ''' Keeps the prohibited legacy property out of editor completion lists.
        ''' </summary>
        <Test>
        Public Sub LegacyHeaderPropertyIsHiddenFromIntelliSense()
            Dim attribute = DirectCast(System.Attribute.GetCustomAttribute(LegacyHeaderProperty, GetType(System.ComponentModel.EditorBrowsableAttribute)), System.ComponentModel.EditorBrowsableAttribute)
            Assert.That(attribute, [Is].Not.Null)
            Assert.That(attribute.State, [Is].EqualTo(System.ComponentModel.EditorBrowsableState.Never))
        End Sub

        <Test>
        Public Sub DefaultOptionsHaveNoHeaderRows()
            Assert.That(New HtmlSheetExportOptions().EffectiveTableHeaderRowIndexes(), [Is].Null)
        End Sub

        <Test>
        Public Sub EffectiveHeaderIndexesAreDetached()
            Dim options As New HtmlSheetExportOptions With {.TableHeaderRowIndexes = New List(Of Integer) From {0, 2}}
            options.EffectiveTableHeaderRowIndexes().Clear()
            Assert.That(options.TableHeaderRowIndexes, [Is].EqualTo(New Integer() {0, 2}))
        End Sub

        <TestCase(-1)>
        <TestCase(Integer.MinValue)>
        Public Sub NegativeHeaderIndexIsRejected(rowIndex As Integer)
            Dim options As New HtmlSheetExportOptions With {.TableHeaderRowIndexes = New List(Of Integer) From {rowIndex}}
            Dim exception = Assert.Throws(Of ArgumentOutOfRangeException)(Sub() options.EffectiveTableHeaderRowIndexes())
            Assert.That(exception.ParamName, [Is].EqualTo(NameOf(HtmlSheetExportOptions.TableHeaderRowIndexes)))
        End Sub

        ''' <summary>
        ''' Simulates existing binaries without compiling a new call to the prohibited property.
        ''' </summary>
        <Test>
        Public Sub LegacyHeaderListRetainsOneBasedAndMutableBehavior()
            Dim options As New HtmlSheetExportOptions()
            Dim legacyRows As New List(Of Integer) From {Integer.MinValue, 0, 1}
            LegacyHeaderProperty.SetValue(options, legacyRows)
            legacyRows.Add(3)
            Assert.That(LegacyHeaderProperty.GetValue(options), [Is].SameAs(legacyRows))
            Assert.That(options.EffectiveTableHeaderRowIndexes(), [Is].EqualTo(New Integer() {0, 2}))
            LegacyHeaderProperty.SetValue(options, Nothing)
            Assert.That(options.EffectiveTableHeaderRowIndexes(), [Is].Null)
        End Sub

        <TestCase(False)>
        <TestCase(True)>
        Public Sub MixingCurrentAndLegacyHeaderSettingsIsRejected(legacyFirst As Boolean)
            Dim options As New HtmlSheetExportOptions()
            If legacyFirst Then LegacyHeaderProperty.SetValue(options, New List(Of Integer) From {1})
            options.TableHeaderRowIndexes = New List(Of Integer) From {0}
            If Not legacyFirst Then LegacyHeaderProperty.SetValue(options, New List(Of Integer) From {1})
            Dim exception = Assert.Throws(Of ArgumentException)(Sub() options.EffectiveTableHeaderRowIndexes())
            Assert.That(exception.ParamName, [Is].EqualTo(NameOf(HtmlSheetExportOptions.TableHeaderRowIndexes)))
        End Sub

        ''' <summary>
        ''' Checks both sheet and workbook output with each EPPlus engine.
        ''' </summary>
        <TestCase("epplus4", False)>
        <TestCase("epplus4", True)>
        <TestCase("epplus8", False)>
        <TestCase("epplus8", True)>
        Public Sub HtmlTableHeadersUseZeroBasedWorksheetIndexes(engine As String, wholeWorkbook As Boolean)
            Dim options As New HtmlWorkbookExportOptions With {.TableHeaderRowIndexes = New List(Of Integer) From {0, 2}}
            AssertRowTags(ExportHeaders(engine, options, wholeWorkbook), "th", "td", "th")
        End Sub

        <TestCase("epplus4")>
        <TestCase("epplus8")>
        Public Sub HtmlTableHeadersAreAbsoluteWhenUsedRangeStartsLater(engine As String)
            Dim options As New HtmlSheetExportOptions With {.TableHeaderRowIndexes = New List(Of Integer) From {4, 6}}
            AssertRowTags(ExportHeaders(engine, options, False, 5), "th", "td", "th")
        End Sub

        <TestCase("epplus4")>
        <TestCase("epplus8")>
        Public Sub HtmlTableHeadersRetainLegacyOneBasedOutput(engine As String)
            Dim options As New HtmlSheetExportOptions()
            LegacyHeaderProperty.SetValue(options, New List(Of Integer) From {1, 3})
            AssertRowTags(ExportHeaders(engine, options, False), "th", "td", "th")
        End Sub

        <TestCase("epplus4", False)>
        <TestCase("epplus4", True)>
        <TestCase("epplus8", False)>
        <TestCase("epplus8", True)>
        Public Sub HtmlWithoutConfiguredHeadersRetainsDataCells(engine As String, explicitlyEmpty As Boolean)
            Dim options As New HtmlSheetExportOptions()
            If explicitlyEmpty Then options.TableHeaderRowIndexes = New List(Of Integer)()
            AssertRowTags(ExportHeaders(engine, options, False), "td", "td", "td")
        End Sub

        <TestCase("epplus4")>
        <TestCase("epplus8")>
        Public Sub HtmlExportRejectsMixedHeaderSettings(engine As String)
            Dim options As New HtmlSheetExportOptions With {.TableHeaderRowIndexes = New List(Of Integer) From {0}}
            LegacyHeaderProperty.SetValue(options, New List(Of Integer) From {1})
            Assert.Throws(Of ArgumentException)(Sub() ExportHeaders(engine, options, False))
        End Sub

        <TestCase("epplus4")>
        <TestCase("epplus8")>
        Public Sub HtmlExportRejectsNegativeHeaderIndexes(engine As String)
            Dim options As New HtmlSheetExportOptions With {.TableHeaderRowIndexes = New List(Of Integer) From {-1}}
            Assert.Throws(Of ArgumentOutOfRangeException)(Sub() ExportHeaders(engine, options, False))
        End Sub

        Private Shared Function ExportHeaders(engine As String, options As HtmlSheetExportOptions, wholeWorkbook As Boolean, Optional firstRow As Integer = 1) As String
            Dim bytes As Byte()
            Using package As New ExcelPackage()
                Dim worksheet = package.Workbook.Worksheets.Add("Headers")
                For row As Integer = firstRow To firstRow + 2
                    worksheet.Cells(row, 1).Value = "Row " & row.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    worksheet.Cells(row, 2).Value = "Value"
                Next
                bytes = package.GetAsByteArray()
            End Using

            Dim workbook As ExcelDataOperationsBase
            Select Case engine
                Case "epplus4"
                    workbook = New EpplusFreeExcelDataOperations(bytes, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Case "epplus8"
                    ExcelOpsTests.Engines.EpplusPolyformEditionOpsTest.AssignLicenseContext()
                    workbook = New EpplusPolyformExcelDataOperations(bytes, New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly))
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(engine))
            End Select
            Try
                If wholeWorkbook Then Return workbook.ExportWorkbookToHtml(DirectCast(options, HtmlWorkbookExportOptions)).ToString()
                Return workbook.ExportSheetToHtml("Headers", options).ToString()
            Finally
                workbook.Close()
            End Try
        End Function

        Private Shared Sub AssertRowTags(html As String, ParamArray expectedTags As String())
            Dim start = html.IndexOf("<table ", StringComparison.Ordinal)
            Dim finish = html.IndexOf("</table>", start, StringComparison.Ordinal) + "</table>".Length
            Dim rows = XElement.Parse(html.Substring(start, finish - start)).Elements("tr").ToArray()
            Assert.That(rows.Length, [Is].EqualTo(expectedTags.Length))
            For index As Integer = 0 To rows.Length - 1
                Assert.That(rows(index).Elements().Count(), [Is].EqualTo(2))
                Assert.That(rows(index).Elements().Select(Function(cell) cell.Name.LocalName),
                            [Is].All.EqualTo(expectedTags(index)), "Exported row " & index.ToString(System.Globalization.CultureInfo.InvariantCulture))
            Next
        End Sub

    End Class

End Namespace
