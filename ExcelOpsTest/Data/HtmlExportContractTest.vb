Option Explicit On
Option Strict On

Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml.Linq
Imports CompuMaster.Excel.ExcelOps
Imports CompuMaster.Excel.ExcelOps.ExcelDataOperationsBase
Imports CompuMaster.Epplus4
Imports NUnit.Framework

Namespace Data

    <TestFixture("epplus4")>
    <TestFixture("epplus8")>
    <NonParallelizable>
    Public Class HtmlExportContractTest

        Private ReadOnly _engine As String
        Private _workbook As ExcelDataOperationsBase
        Private Const SheetName As String = "Grid"

        Public Sub New(engine As String)
            _engine = engine
        End Sub

        <SetUp>
        Public Sub OpenWorkbook()
            ResetWorkbook()
        End Sub

        <TearDown>
        Public Sub CloseWorkbook()
            If _workbook IsNot Nothing Then _workbook.Close()
        End Sub

        Private Sub ResetWorkbook(Optional empty As Boolean = False, Optional merge As String = Nothing, Optional mergedOnly As Boolean = False)
            If _workbook IsNot Nothing Then _workbook.Close()
            Dim bytes As Byte()
            Using package As New ExcelPackage()
                Dim sheet = package.Workbook.Worksheets.Add(SheetName)
                If Not empty AndAlso Not mergedOnly Then
                    For row As Integer = 1 To 4
                        For column As Integer = 1 To 4
                            sheet.Cells(row, column).Value = "R" & row.ToString(System.Globalization.CultureInfo.InvariantCulture) & "C" & column.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        Next
                    Next
                End If
                If mergedOnly Then sheet.Cells(1, 1).Value = "Merged master"
                If merge IsNot Nothing Then sheet.Cells(merge).Merge = True
                bytes = package.GetAsByteArray()
            End Using
            Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
            If _engine = "epplus4" Then
                _workbook = New EpplusFreeExcelDataOperations(bytes, options)
            Else
                ExcelOpsTests.Engines.EpplusPolyformEditionOpsTest.AssignLicenseContext()
                _workbook = New EpplusPolyformExcelDataOperations(bytes, options)
            End If
        End Sub

        Private Function OutputFile() As String
            Return TestEnvironment.FullPathOfDynTestFile(_workbook, Guid.NewGuid().ToString("N") & ".html")
        End Function

        Private Function ExportTable(options As HtmlSheetExportOptions) As String
            Return _workbook.ExportSheetToHtml(SheetName, options, HtmlDocumentExportParts.TableOnly).ToString()
        End Function

        Private Shared Function Rows(html As String) As XElement()
            Return XElement.Parse(html.Trim()).Elements("tr").ToArray()
        End Function

        ''' <summary>
        ''' All three output modes must be identical across builder, sync file and awaitable file exports.
        ''' </summary>
        <TestCase(HtmlDocumentExportParts.FullHtmlDocument, False)>
        <TestCase(HtmlDocumentExportParts.ContentOnly, False)>
        <TestCase(HtmlDocumentExportParts.TableOnly, False)>
        <TestCase(HtmlDocumentExportParts.FullHtmlDocument, True)>
        <TestCase(HtmlDocumentExportParts.ContentOnly, True)>
        <TestCase(HtmlDocumentExportParts.TableOnly, True)>
        Public Async Function HtmlOutputModesMatchAcrossSyncAndTaskExports(mode As HtmlDocumentExportParts, empty As Boolean) As Task
            If empty Then ResetWorkbook(empty:=True)
            Dim options As New HtmlSheetExportOptions With {.ExportSheetNameAsTitle = HtmlSheetExportOptions.SheetTitleStyles.H2}
            Dim expected = _workbook.ExportSheetToHtml(SheetName, options, mode).ToString()
            Assert.That(expected.Contains("<!doctype html>"), [Is].EqualTo(mode = HtmlDocumentExportParts.FullHtmlDocument))
            Assert.That(expected.Contains("<section "), [Is].EqualTo(mode <> HtmlDocumentExportParts.TableOnly))
            Assert.That(expected.Contains("<H2 "), [Is].EqualTo(mode <> HtmlDocumentExportParts.TableOnly))
            Assert.That(expected.Contains("<table "), [Is].EqualTo(Not empty))
            If empty Then Assert.That(expected, Does.Contain("-/-"))
            Dim syncPath = OutputFile()
            Dim asyncPath = OutputFile()
            Try
                _workbook.ExportSheetToHtml(SheetName, syncPath, options, mode)
                Await _workbook.ExportSheetToHtmlFileAsync(SheetName, asyncPath, options, mode)
                Assert.That(File.ReadAllBytes(asyncPath), [Is].EqualTo(File.ReadAllBytes(syncPath)))
                Assert.That(File.ReadAllText(asyncPath), [Is].EqualTo(expected))
                Assert.That(File.ReadAllBytes(asyncPath).Take(3), [Is].EqualTo(New Byte() {&HEF, &HBB, &HBF}))
                ' Completion includes disposing the writer, so exclusive access must already succeed.
                Using exclusive = File.Open(asyncPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                    Assert.That(exclusive.Length, [Is].GreaterThan(0))
                End Using
            Finally
                File.Delete(syncPath)
                File.Delete(asyncPath)
            End Try
        End Function

        <Test>
        Public Async Function HtmlTaskDefaultsMatchExistingFullDocumentExports() As Task
            Dim options As New HtmlSheetExportOptions()
            Dim path = OutputFile()
            Try
                Await _workbook.ExportSheetToHtmlFileAsync(SheetName, path, options)
                Assert.That(File.ReadAllText(path), [Is].EqualTo(_workbook.ExportSheetToHtml(SheetName, options).ToString()))
                _workbook.ExportSheetToHtml(SheetName, path, options)
                Assert.That(File.ReadAllText(path), [Is].EqualTo(_workbook.ExportSheetToHtml(SheetName, options).ToString()))
            Finally
                File.Delete(path)
            End Try
        End Function

        <Test>
        Public Async Function HtmlWorkbookTaskMatchesSynchronousExport() As Task
            Dim options As New HtmlWorkbookExportOptions With {.FirstRowIndex = 1, .LastRowIndex = 2, .TableHeaderRowIndexes = New List(Of Integer) From {1}}
            Dim syncPath = OutputFile()
            Dim asyncPath = OutputFile()
            Try
                _workbook.ExportWorkbookToHtml(syncPath, options)
                Await _workbook.ExportWorkbookToHtmlFileAsync(asyncPath, options)
                Assert.That(File.ReadAllBytes(asyncPath), [Is].EqualTo(File.ReadAllBytes(syncPath)))
                Assert.That(File.ReadAllText(asyncPath), Does.Contain("<th "))
                Assert.That(File.ReadAllText(asyncPath), Does.Not.Contain("R1C1"))
                Assert.That(File.ReadAllBytes(asyncPath).Take(3), [Is].EqualTo(New Byte() {&HEF, &HBB, &HBF}))
            Finally
                File.Delete(syncPath)
                File.Delete(asyncPath)
            End Try
        End Function

        <TestCase(HtmlDocumentExportParts.FullHtmlDocument)>
        <TestCase(HtmlDocumentExportParts.ContentOnly)>
        <TestCase(HtmlDocumentExportParts.TableOnly)>
        Public Sub HtmlBuilderRetainsExistingContent(mode As HtmlDocumentExportParts)
            Dim builder As New StringBuilder("prefix")
            _workbook.ExportSheetToHtml(SheetName, "grid", True, builder, New HtmlSheetExportOptions(), mode)
            Assert.That(builder.ToString(), Does.StartWith("prefix"))
            Assert.That(builder.ToString(), Does.Contain("R1C1"))
        End Sub

        <Test>
        Public Sub HtmlRangeBoundsAreInclusiveAndHeadersStayAbsolute()
            Dim options As New HtmlSheetExportOptions With {
                .FirstRowIndex = 1, .LastRowIndex = 2, .FirstColumnIndex = 1, .LastColumnIndex = 2,
                .TableHeaderRowIndexes = New List(Of Integer) From {1}}
            Dim exportedRows = Rows(ExportTable(options))
            Assert.That(exportedRows.Length, [Is].EqualTo(2))
            Assert.That(exportedRows(0).Elements().Select(Function(cell) cell.Value), [Is].EqualTo(New String() {"R2C2", "R2C3"}))
            Assert.That(exportedRows(1).Elements().Select(Function(cell) cell.Value), [Is].EqualTo(New String() {"R3C2", "R3C3"}))
            Assert.That(exportedRows(0).Elements().Select(Function(cell) cell.Name.LocalName), [Is].All.EqualTo("th"))
            Assert.That(exportedRows(1).Elements().Select(Function(cell) cell.Name.LocalName), [Is].All.EqualTo("td"))
        End Sub

        <Test>
        Public Sub HtmlUnsetEndBoundsUseLastUsedCells()
            Dim options As New HtmlSheetExportOptions With {.FirstRowIndex = 1, .FirstColumnIndex = 2}
            Dim exportedRows = Rows(ExportTable(options))
            Assert.That(exportedRows.Length, [Is].EqualTo(3))
            Assert.That(exportedRows(0).Elements().Select(Function(cell) cell.Value), [Is].EqualTo(New String() {"R2C3", "R2C4"}))
            Assert.That(exportedRows(2).Elements().Select(Function(cell) cell.Value), [Is].EqualTo(New String() {"R4C3", "R4C4"}))
        End Sub

        <Test>
        Public Sub HtmlBoundsBeyondUsedRangeDoNotAddBlankCells()
            Dim options As New HtmlSheetExportOptions With {.LastRowIndex = 1048575, .LastColumnIndex = 16383}
            Dim exportedRows = Rows(ExportTable(options))
            Assert.That(exportedRows.Length, [Is].EqualTo(4))
            Assert.That(exportedRows.All(Function(row) row.Elements().Count() = 4), [Is].True)
        End Sub

        <TestCase(True)>
        <TestCase(False)>
        Public Sub HtmlRangeOutsideUsedCellsUsesEmptyPlaceholder(rowOutside As Boolean)
            Dim options As New HtmlSheetExportOptions With {.HtmlForEmptySheet = "<p>outside</p>"}
            If rowOutside Then
                options.FirstRowIndex = 1048575
            Else
                options.FirstColumnIndex = 16383
            End If
            Assert.That(ExportTable(options).Trim(), [Is].EqualTo("<p>outside</p>"))
        End Sub

        <TestCase("FirstRowIndex", -1)>
        <TestCase("FirstColumnIndex", -1)>
        <TestCase("LastRowIndex", -1)>
        <TestCase("LastColumnIndex", -1)>
        <TestCase("FirstRowIndex", Integer.MaxValue)>
        <TestCase("FirstColumnIndex", Integer.MaxValue)>
        <TestCase("LastRowIndex", Integer.MaxValue)>
        <TestCase("LastColumnIndex", Integer.MaxValue)>
        <TestCase("FirstRowIndex", 1048576)>
        <TestCase("FirstColumnIndex", 16384)>
        <TestCase("LastRowIndex", 1048576)>
        <TestCase("LastColumnIndex", 16384)>
        Public Sub HtmlInvalidBoundsAreRejectedEvenOnEmptySheets(propertyName As String, value As Integer)
            ResetWorkbook(empty:=True)
            Dim options As New HtmlSheetExportOptions()
            GetType(HtmlSheetExportOptions).GetProperty(propertyName).SetValue(options, value)
            Dim exception = Assert.Throws(Of ArgumentOutOfRangeException)(Sub() ExportTable(options))
            Assert.That(exception.ParamName, [Is].EqualTo(propertyName))
        End Sub

        <TestCase(True)>
        <TestCase(False)>
        Public Sub HtmlReversedBoundsAreRejected(rowsReversed As Boolean)
            Dim options As New HtmlSheetExportOptions()
            If rowsReversed Then
                options.FirstRowIndex = 2
                options.LastRowIndex = 1
            Else
                options.FirstColumnIndex = 2
                options.LastColumnIndex = 1
            End If
            Assert.Throws(Of ArgumentOutOfRangeException)(Sub() ExportTable(options))
        End Sub

        ''' <summary>
        ''' A merged rectangle must never silently disappear or extend past the requested bounds.
        ''' </summary>
        <TestCase("FirstRowIndex", 2)>
        <TestCase("LastRowIndex", 1)>
        <TestCase("FirstColumnIndex", 2)>
        <TestCase("LastColumnIndex", 1)>
        Public Sub HtmlPartiallyIntersectedMergedCellsAreRejected(propertyName As String, value As Integer)
            ResetWorkbook(merge:="B2:C3")
            Dim options As New HtmlSheetExportOptions()
            GetType(HtmlSheetExportOptions).GetProperty(propertyName).SetValue(options, value)
            Dim exception = Assert.Throws(Of ArgumentException)(Sub() ExportTable(options))
            Assert.That(exception.Message, Does.Contain("B2:C3"))
            Assert.That(exception.Message, Does.Contain(SheetName))
        End Sub

        <Test>
        Public Sub HtmlFullyIncludedMergedCellsKeepTheirSpans()
            ResetWorkbook(merge:="B2:C3")
            Dim options As New HtmlSheetExportOptions With {.FirstRowIndex = 1, .LastRowIndex = 2, .FirstColumnIndex = 1, .LastColumnIndex = 2, .TableHeaderRowIndexes = New List(Of Integer) From {1}}
            Dim exportedRows = Rows(ExportTable(options))
            Assert.That(exportedRows.Length, [Is].EqualTo(2))
            Dim master = exportedRows(0).Elements().Single()
            Assert.That(master.Value, [Is].EqualTo("R2C2"))
            Assert.That(master.Name.LocalName, [Is].EqualTo("th"))
            Assert.That(master.Attribute("rowspan").Value, [Is].EqualTo("2"))
            Assert.That(master.Attribute("colspan").Value, [Is].EqualTo("2"))
            Assert.That(exportedRows(1).Elements(), [Is].Empty)
        End Sub

        <Test>
        Public Sub HtmlCompletelyExcludedMergedCellsAreIgnored()
            ResetWorkbook(merge:="B2:C3")
            Dim exportedRows = Rows(ExportTable(New HtmlSheetExportOptions With {.LastRowIndex = 0}))
            Assert.That(exportedRows.Length, [Is].EqualTo(1))
            Assert.That(exportedRows(0).Elements().Count(), [Is].EqualTo(4))
        End Sub

        <Test>
        Public Sub HtmlDefaultRangeIncludesMergedExtents()
            ResetWorkbook(merge:="A1:B2", mergedOnly:=True)
            Dim exportedRows = Rows(ExportTable(New HtmlSheetExportOptions()))
            Assert.That(exportedRows.Length, [Is].EqualTo(2))
            Assert.That(exportedRows(0).Elements().Single().Attribute("colspan").Value, [Is].EqualTo("2"))
            Assert.That(exportedRows(0).Elements().Single().Attribute("rowspan").Value, [Is].EqualTo("2"))
        End Sub

        <TestCase(Nothing, "-/-")>
        <TestCase("", "")>
        <TestCase("<p>custom</p>", "<p>custom</p>")>
        Public Sub HtmlEmptySheetsUseEffectivePlaceholder(configured As String, expected As String)
            ResetWorkbook(empty:=True)
            Dim options As New HtmlSheetExportOptions With {.HtmlForEmptySheet = configured}
            Assert.That(ExportTable(options).Trim(), [Is].EqualTo(expected))
        End Sub

        Private Class CustomEmptyOptions
            Inherits HtmlSheetExportOptions

            Public Overrides Function EffectiveHtmlForEmptySheet() As String
                Return "<p>overridden</p>"
            End Function
        End Class

        <Test>
        Public Sub HtmlEmptySheetsCallOverriddenFallback()
            ResetWorkbook(empty:=True)
            Assert.That(ExportTable(New CustomEmptyOptions()).Trim(), [Is].EqualTo("<p>overridden</p>"))
        End Sub

        <Test>
        Public Sub HtmlRenderingFailuresDoNotTruncateExistingFiles()
            ResetWorkbook(merge:="B2:C3")
            Dim options As New HtmlSheetExportOptions With {.LastRowIndex = 1}
            Dim path = OutputFile()
            Try
                File.WriteAllText(path, "retain existing file")
                Assert.Throws(Of ArgumentException)(Sub() _workbook.ExportSheetToHtml(SheetName, path, options))
                Assert.That(File.ReadAllText(path), [Is].EqualTo("retain existing file"))
                Assert.ThrowsAsync(Of ArgumentException)(Function() _workbook.ExportSheetToHtmlFileAsync(SheetName, path, options))
                Assert.That(File.ReadAllText(path), [Is].EqualTo("retain existing file"))
                Assert.ThrowsAsync(Of ArgumentException)(Function() _workbook.ExportWorkbookToHtmlFileAsync(path, New HtmlWorkbookExportOptions With {.LastRowIndex = 1}))
                Assert.That(File.ReadAllText(path), [Is].EqualTo("retain existing file"))
            Finally
                File.Delete(path)
            End Try
        End Sub

        <Test>
        Public Sub HtmlInvalidModeAndNullOptionsAreAwaitableFailures()
            Dim path = OutputFile()
            Assert.ThrowsAsync(Of ArgumentOutOfRangeException)(Function() _workbook.ExportSheetToHtmlFileAsync(SheetName, path, New HtmlSheetExportOptions(), CType(255, HtmlDocumentExportParts)))
            Assert.ThrowsAsync(Of ArgumentNullException)(Function() _workbook.ExportSheetToHtmlFileAsync(SheetName, path, Nothing))
            Assert.ThrowsAsync(Of ArgumentNullException)(Function() _workbook.ExportWorkbookToHtmlFileAsync(path, Nothing))
            Assert.That(File.Exists(path), [Is].False)
        End Sub

        <Test>
        Public Sub HtmlWriteFailuresCanBeAwaited()
            Dim directory = Path.GetDirectoryName(OutputFile())
            Assert.ThrowsAsync(Of UnauthorizedAccessException)(Function() _workbook.ExportSheetToHtmlFileAsync(SheetName, directory, New HtmlSheetExportOptions()))
        End Sub

        <Test>
        Public Sub HtmlWorkbookWithNoSelectedSheetsDoesNotReplaceFile()
            Dim path = OutputFile()
            Try
                File.WriteAllText(path, "retain existing file")
                Assert.ThrowsAsync(Of InvalidOperationException)(Function() _workbook.ExportWorkbookToHtmlFileAsync(path, New HtmlWorkbookExportOptions With {.ExportWorkSheets = False, .ExportChartSheets = False}))
                Assert.That(File.ReadAllText(path), [Is].EqualTo("retain existing file"))
            Finally
                File.Delete(path)
            End Try
        End Sub

#If Not NETFRAMEWORK Then
        Private Class LegacyAsyncCompletion
            Inherits SynchronizationContext

            Private ReadOnly _completion As New TaskCompletionSource(Of Object)(TaskCreationOptions.RunContinuationsAsynchronously)

            Friend ReadOnly Property Completion As Task
                Get
                    Return _completion.Task
                End Get
            End Property

            Public Overrides Sub OperationCompleted()
                _completion.TrySetResult(Nothing)
            End Sub

            Public Overrides Sub Post(callback As SendOrPostCallback, state As Object)
                Try
                    callback(state)
                Catch exception As Exception
                    _completion.TrySetException(exception)
                End Try
            End Sub
        End Class

        ''' <summary>
        ''' Legacy Async Sub signatures and their different output defaults remain compatible.
        ''' </summary>
        <TestCase(False)>
        <TestCase(True)>
        Public Async Function HtmlLegacyAsyncWrappersRetainOutputAndAreHidden(wholeWorkbook As Boolean) As Task
            Dim name = If(wholeWorkbook, "ExportWorkbookToHtmlAsync", "ExportSheetToHtmlAsync")
            Dim replacement = If(wholeWorkbook, "ExportWorkbookToHtmlFileAsync", "ExportSheetToHtmlFileAsync")
            Dim method = GetType(ExcelDataOperationsBase).GetMethod(name)
            Dim obsolete = DirectCast(System.Attribute.GetCustomAttribute(method, GetType(ObsoleteAttribute)), ObsoleteAttribute)
            Assert.That(obsolete.Message, Does.Contain(replacement))
            Dim browsable = DirectCast(System.Attribute.GetCustomAttribute(method, GetType(System.ComponentModel.EditorBrowsableAttribute)), System.ComponentModel.EditorBrowsableAttribute)
            Assert.That(browsable.State, [Is].EqualTo(System.ComponentModel.EditorBrowsableState.Never))
            Dim options As New HtmlWorkbookExportOptions With {.ExportSheetNameAsTitle = HtmlSheetExportOptions.SheetTitleStyles.H2}
            Dim expected = If(wholeWorkbook, _workbook.ExportWorkbookToHtml(options).ToString(), ExportTable(options))
            Dim path = OutputFile()
            Dim context As New LegacyAsyncCompletion()
            Dim previous = SynchronizationContext.Current
            Try
                Try
                    SynchronizationContext.SetSynchronizationContext(context)
                    If wholeWorkbook Then
                        method.Invoke(_workbook, New Object() {path, options})
                    Else
                        method.Invoke(_workbook, New Object() {SheetName, path, options})
                    End If
                Finally
                    SynchronizationContext.SetSynchronizationContext(previous)
                End Try
                Dim completed = Await Task.WhenAny(context.Completion, Task.Delay(10000))
                Assert.That(completed, [Is].SameAs(context.Completion), "Legacy async wrapper did not complete.")
                Await context.Completion
                Assert.That(File.ReadAllText(path), [Is].EqualTo(expected))
            Finally
                File.Delete(path)
            End Try
        End Function
#End If

    End Class

End Namespace
