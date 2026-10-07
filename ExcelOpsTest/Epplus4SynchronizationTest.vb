Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Threading.Tasks
Imports CompuMaster.Epplus4
Imports CompuMaster.Epplus4.FormulaParsing
Imports CompuMaster.Epplus4.FormulaParsing.Excel.Functions
Imports CompuMaster.Epplus4.FormulaParsing.ExpressionGraph
Imports NUnit.Framework

''' <summary>
''' Verifies internal synchronization without implying general workbook thread safety.
''' </summary>
<TestFixture>
Public Class Epplus4SynchronizationTest

    ' Fresh ranges isolate store synchronization from the mutable cached-range contract.
    <TestCase(False), TestCase(True)>
    Public Sub CellWritesSurviveColumnGrowth(parallel As Boolean)
        Using package As New ExcelPackage()
            Dim sheet = package.Workbook.Worksheets.Add("Values")
            RunWorkers(parallel, 64,
                Sub(column)
                    For row = 1 To 300
                        Dim value = column * 1000 + row
                        sheet.Cells(row, column + 1).Value = value
                        Assert.That(sheet.Cells(row, column + 1).Value, [Is].EqualTo(value))
                    Next
                End Sub)
            For column = 0 To 63
                For row = 1 To 300
                    Assert.That(sheet.Cells(row, column + 1).Value, [Is].EqualTo(column * 1000 + row))
                Next
            Next
        End Using
    End Sub

    ' Every worksheet adds distinct workbook-wide format IDs; validate saved IDs as well.
    <TestCase(False), TestCase(True)>
    Public Sub StylesAcrossWorksheetsRoundTrip(parallel As Boolean)
        Using package As New ExcelPackage()
            Dim sheets = Enumerable.Range(0, 8).Select(Function(i) package.Workbook.Worksheets.Add("Sheet" & i)).ToArray()
            RunWorkers(parallel, sheets.Length,
                Sub(i)
                    For row = 1 To 150
                        sheets(i).Cells(row, 1).Value = row
                        sheets(i).Cells(row, 1).Style.Numberformat.Format = FormatFor(i, row)
                    Next
                End Sub)
            Using reopened As New ExcelPackage(New MemoryStream(package.GetAsByteArray()))
                For i = 0 To sheets.Length - 1
                    For row = 1 To 150
                        Assert.That(reopened.Workbook.Worksheets(i).Cells(row, 1).Style.Numberformat.Format, [Is].EqualTo(FormatFor(i, row)))
                    Next
                Next
            End Using
        End Using
    End Sub

    <TestCase(False), TestCase(True)>
    Public Sub CalculationsAcrossWorksheetsKeepResults(parallel As Boolean)
        Using package As New ExcelPackage()
            Dim sheets = Enumerable.Range(0, 8).Select(Function(i) package.Workbook.Worksheets.Add("Sheet" & i)).ToArray()
            For i = 0 To sheets.Length - 1
                For row = 1 To 150
                    sheets(i).Cells(row, 1).Value = i + row
                    sheets(i).Cells(row, 2).Formula = "A" & row & "*2+1"
                Next
            Next
            RunWorkers(parallel, sheets.Length, Sub(i) CalculationExtension.Calculate(sheets(i)))
            For i = 0 To sheets.Length - 1
                For row = 1 To 150
                    Assert.That(sheets(i).Cells(row, 2).Value, [Is].EqualTo(CDbl((i + row) * 2 + 1)))
                Next
            Next
        End Using
    End Sub

    <Test>
    Public Sub IndependentPackagesRoundTripInParallel()
        RunWorkers(True, 32,
            Sub(i)
                Using package As New ExcelPackage()
                    Dim sheet = package.Workbook.Worksheets.Add("Sheet")
                    sheet.Cells(1, 1).Value = i
                    sheet.Cells(1, 2).Formula = "A1*2+1"
                    sheet.Cells(1, 1).Style.Numberformat.Format = "0.00"
                    CalculationExtension.Calculate(sheet)
                    Using reopened As New ExcelPackage(New MemoryStream(package.GetAsByteArray()))
                        Assert.That(reopened.Workbook.Worksheets(0).Cells(1, 2).Value, [Is].EqualTo(CDbl(i * 2 + 1)))
                    End Using
                End Using
            End Sub)
    End Sub

    <Test>
    Public Sub FirstUsePublishesOneWorkbookAndStyles()
        Using package As New ExcelPackage()
            Dim workbooks(31) As ExcelWorkbook
            Dim styles(31) As ExcelStyles
            Dim managers(31) As FormulaParserManager
            Dim worksheets(31) As ExcelWorksheets
            RunWorkers(True, workbooks.Length,
                Sub(i)
                    workbooks(i) = package.Workbook
                    styles(i) = workbooks(i).Styles
                    managers(i) = workbooks(i).FormulaParserManager
                    worksheets(i) = workbooks(i).Worksheets
                End Sub)
            For i = 1 To workbooks.Length - 1
                Assert.That(workbooks(i), [Is].SameAs(workbooks(0)))
                Assert.That(styles(i), [Is].SameAs(styles(0)))
                Assert.That(managers(i), [Is].SameAs(managers(0)))
                Assert.That(worksheets(i), [Is].SameAs(worksheets(0)))
            Next
        End Using
    End Sub

    <Test>
    Public Sub AddressCacheSupportsConcurrentClearAndLookup()
        Dim cache As New ExcelAddressCache()
        RunWorkers(True, 8,
            Sub(worker)
                For i = 1 To 3000
                    If worker = 0 Then
                        cache.Clear()
                    Else
                        cache.Add(i, "A" & i)
                        Dim address = cache.Get(i)
                        Assert.That(address = String.Empty OrElse address = "A" & i, [Is].True)
                        Assert.That(cache.Count, [Is].GreaterThanOrEqualTo(0))
                    End If
                Next
            End Sub)
    End Sub

    ' The shared range remains mutable for compatibility; callers must coordinate selection and use.
    <Test>
    Public Sub CachedRangeKeepsExistingSelectionSemantics()
        Using package As New ExcelPackage()
            Dim sheet = package.Workbook.Worksheets.Add("Sheet")
            Dim cells = sheet.Cells
            Dim first = cells(1, 1)
            Dim second = cells(1, 2)
            Assert.That(first, [Is].SameAs(second))
            first.Value = "selected last"
            Assert.That(sheet.Cells(1, 1).Value, [Is].Null)
            Assert.That(sheet.Cells(1, 2).Value, [Is].EqualTo("selected last"))
        End Using
    End Sub

    ' Growing beyond 32 columns used to replace the monitor while a write held it.
    <Test>
    Public Sub CellStoreMonitorSurvivesColumnIndexGrowth()
        Using package As New ExcelPackage()
            Dim sheet = package.Workbook.Worksheets.Add("Sheet")
            Dim store = GetType(ExcelWorksheet).GetField("_values", BindingFlags.Instance Or BindingFlags.NonPublic).GetValue(sheet)
            Dim rootField = store.GetType().GetField("SyncRoot", BindingFlags.Instance Or BindingFlags.NonPublic)
            Dim columnsField = store.GetType().GetField("_columnIndex", BindingFlags.Instance Or BindingFlags.NonPublic)
            Dim monitorBeforeGrowth = If(rootField Is Nothing, columnsField.GetValue(store), rootField.GetValue(store))
            For col = 1 To 70
                sheet.Cells(1, col).Value = col
            Next
            Dim monitorAfterGrowth = If(rootField Is Nothing, columnsField.GetValue(store), rootField.GetValue(store))
            Assert.That(monitorAfterGrowth, [Is].SameAs(monitorBeforeGrowth))
            Assert.That(sheet.Cells(1, 70).Value, [Is].EqualTo(70))
        End Using
    End Sub

    ' Pause inside the parser before starting a competing public operation.
    <TestCase("workbook"), TestCase("worksheet"), TestCase("range"), TestCase("formula"), TestCase("parse"), TestCase("save"), TestCase("dispose")>
    Public Sub CompetingOperationsWaitForActiveCalculation(operation As String)
        Using package As New ExcelPackage(), entered As New ManualResetEventSlim(), release As New ManualResetEventSlim(), attempting As New ManualResetEventSlim(), finished As New ManualResetEventSlim()
            Dim sheet = package.Workbook.Worksheets.Add("First")
            Dim other = package.Workbook.Worksheets.Add("Second")
            package.Workbook.FormulaParserManager.LoadFunctionModule(New CallbackModule(
                Sub()
                    entered.Set()
                    If Not release.Wait(TimeSpan.FromSeconds(10)) Then Throw New TimeoutException()
                End Sub))
            sheet.Cells(1, 1).Formula = "BLOCKING()"
            other.Cells(1, 1).Formula = "2+3"
            Dim manager = package.Workbook.FormulaParserManager
            Dim first = Task.Factory.StartNew(Sub() CalculationExtension.Calculate(sheet), CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default)
            Dim second As Task = Nothing
            Dim blocked As Boolean = False
            Try
                Assert.That(entered.Wait(TimeSpan.FromSeconds(10)), [Is].True)
                second = Task.Factory.StartNew(
                    Sub()
                        attempting.Set()
                        Try
                            Select Case operation
                                Case "workbook"
                                    CalculationExtension.Calculate(package.Workbook)
                                Case "worksheet"
                                    CalculationExtension.Calculate(other)
                                Case "range"
                                    CalculationExtension.Calculate(other.Cells(1, 1))
                                Case "formula"
                                    Assert.That(CalculationExtension.Calculate(other, "2+3"), [Is].EqualTo(5.0))
                                Case "parse"
                                    Assert.That(manager.Parse("2+3"), [Is].EqualTo(5.0))
                                Case "save"
                                    Dim bytes = package.GetAsByteArray()
                                    Using reopened As New ExcelPackage(New MemoryStream(bytes))
                                        Assert.That(reopened.Workbook.Worksheets(0).Cells(1, 1).Value, [Is].EqualTo(7.0))
                                    End Using
                                Case "dispose"
                                    package.Dispose()
                            End Select
                        Finally
                            finished.Set()
                        End Try
                    End Sub, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default)
                Assert.That(attempting.Wait(TimeSpan.FromSeconds(10)), [Is].True)
                blocked = Not finished.Wait(TimeSpan.FromMilliseconds(250))
            Finally
                release.Set()
                Dim tasks = If(second Is Nothing, New Task() {first}, New Task() {first, second})
                Assert.That(Task.WaitAll(tasks, TimeSpan.FromSeconds(10)), [Is].True)
            End Try
            Assert.That(blocked, [Is].True, "The competing " & operation & " operation ran during calculation.")
            If operation <> "dispose" Then
                Assert.That(sheet.Cells(1, 1).Value, [Is].EqualTo(7.0))
            End If
        End Using
    End Sub

    ' Independent parsers must be able to enter custom functions at the same time.
    <Test>
    Public Sub IndependentWorkbookCalculationsCanOverlap()
        Using firstPackage As New ExcelPackage(), secondPackage As New ExcelPackage(), barrier As New Barrier(2)
            Dim first = firstPackage.Workbook.Worksheets.Add("First")
            Dim second = secondPackage.Workbook.Worksheets.Add("Second")
            Dim callback As Action =
                Sub()
                    If Not barrier.SignalAndWait(TimeSpan.FromSeconds(10)) Then Throw New TimeoutException("Independent calculations were serialized.")
                End Sub
            firstPackage.Workbook.FormulaParserManager.LoadFunctionModule(New CallbackModule(callback))
            secondPackage.Workbook.FormulaParserManager.LoadFunctionModule(New CallbackModule(callback))
            first.Cells(1, 1).Formula = "BLOCKING()"
            second.Cells(1, 1).Formula = "BLOCKING()"
            Dim tasks = New Task() {
                Task.Factory.StartNew(Sub() CalculationExtension.Calculate(first), CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default),
                Task.Factory.StartNew(Sub() CalculationExtension.Calculate(second), CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default)
            }
            Assert.That(Task.WaitAll(tasks, TimeSpan.FromSeconds(15)), [Is].True)
            Assert.That(first.Cells(1, 1).Value, [Is].EqualTo(7.0))
            Assert.That(second.Cells(1, 1).Value, [Is].EqualTo(7.0))
        End Using
    End Sub

    ' Enumeration captures collection membership before a competing style update.
    <Test>
    Public Sub StyleEnumerationKeepsItsSnapshot()
        Using package As New ExcelPackage()
            Dim sheet = package.Workbook.Worksheets.Add("Sheet")
            sheet.Cells(1, 1).Style.Numberformat.Format = "0.00""snapshot-a"""
            Dim formats = package.Workbook.Styles.NumberFormats
            Dim count = formats.Count
            Using iterator = formats.GetEnumerator()
                Dim writer = Task.Run(Sub() sheet.Cells(1, 2).Style.Numberformat.Format = "0.00""snapshot-b""")
                Assert.That(writer.Wait(TimeSpan.FromSeconds(10)), [Is].True)
                Dim enumerated = 0
                While iterator.MoveNext()
                    enumerated += 1
                End While
                Assert.That(enumerated, [Is].EqualTo(count))
                Assert.That(formats.Count, [Is].EqualTo(count + 1))
            End Using
        End Using
    End Sub

    Private NotInheritable Class CallbackModule
        Inherits FunctionsModule

        Public Sub New(callback As Action)
            Functions.Add("blocking", New CallbackFunction(callback))
        End Sub
    End Class

    Private NotInheritable Class CallbackFunction
        Inherits ExcelFunction

        Private ReadOnly _callback As Action

        Public Sub New(callback As Action)
            _callback = callback
        End Sub

        Public Overrides Function Execute(arguments As IEnumerable(Of FunctionArgument), context As ParsingContext) As CompileResult
            _callback()
            Return New CompileResult(7.0, DataType.Decimal)
        End Function
    End Class

    Private Shared Function FormatFor(sheet As Integer, row As Integer) As String
        Return "0.00""s" & sheet & "r" & row & """"
    End Function

    Private Shared Sub RunWorkers(parallel As Boolean, count As Integer, action As Action(Of Integer))
        If parallel Then
            System.Threading.Tasks.Parallel.For(0, count, New ParallelOptions With {.MaxDegreeOfParallelism = 8}, action)
        Else
            For i = 0 To count - 1
                action(i)
            Next
        End If
    End Sub
End Class
