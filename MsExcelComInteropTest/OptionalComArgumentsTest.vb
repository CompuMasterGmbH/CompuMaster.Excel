Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports CompuMaster.ComInterop
Imports CompuMaster.Excel.MsExcelComInterop
Imports NUnit.Framework

Public Class OptionalComArgumentsTest

    ''' <summary>
    ''' Verifies optional printer arguments with a managed dispatch recorder, without submitting print jobs.
    ''' </summary>
    <TestCase(False, Nothing, Nothing)>
    <TestCase(True, Nothing, Nothing)>
    <TestCase(False, "Selected printer", "output.prn")>
    <TestCase(True, "Selected printer", "output.prn")>
    <TestCase(False, "", "")>
    <TestCase(True, "", "")>
    Public Sub PrintOutPreservesOmittedAndExplicitStrings(worksheet As Boolean, printer As String, fileName As String)
        Dim dispatch As New PrinterDispatch()
        Dim wrapperType = If(worksheet, GetType(ExcelSheet), GetType(ExcelWorkbook))
        ' Only method dispatch is under test. There are no native objects or parent wrappers to dispose.
        Dim wrapper = RuntimeHelpers.GetUninitializedObject(wrapperType)
#Disable Warning CA1816 ' The dispatch-only wrapper intentionally has no native lifecycle to finalize.
        GC.SuppressFinalize(wrapper)
#Enable Warning CA1816
        GetType(ComObjectBase).GetField("_ComObject", BindingFlags.Instance Or BindingFlags.NonPublic).SetValue(wrapper, dispatch)
        GetType(ComObjectBase).GetField("_ComObjectType", BindingFlags.Instance Or BindingFlags.NonPublic).SetValue(wrapper, dispatch.GetType())

        If worksheet Then
            CType(wrapper, ExcelSheet).PrintOut(activePrinter:=printer, printToFileName:=fileName)
        Else
            CType(wrapper, ExcelWorkbook).PrintOut(activePrinter:=printer, printToFileName:=fileName)
        End If

        Assert.That(dispatch.ActivePrinter, [Is].EqualTo(If(printer Is Nothing, "Excel active printer", printer)))
        Assert.That(dispatch.PrintFileName, [Is].EqualTo(If(fileName Is Nothing, "Excel default file name", fileName)))
        Assert.That(dispatch.FromPage, [Is].EqualTo(1))
        Assert.That(dispatch.ToPage, [Is].EqualTo(Int16.MaxValue))
        Assert.That(dispatch.Copies, [Is].EqualTo(1))
    End Sub

    Public NotInheritable Class PrinterDispatch
        Public Property ActivePrinter As Object
        Public Property PrintFileName As Object
        Public Property FromPage As Integer
        Public Property ToPage As Integer
        Public Property Copies As Integer

        Public Sub PrintOut(fromPage As Integer, toPage As Integer, copies As Integer, preview As Boolean,
                            Optional activePrinter As Object = "Excel active printer",
                            Optional printToFile As Boolean = False, Optional collate As Boolean = False,
                            Optional printFileName As Object = "Excel default file name", Optional ignorePrintAreas As Boolean = False)
            Me.ActivePrinter = activePrinter
            Me.PrintFileName = printFileName
            Me.FromPage = fromPage
            Me.ToPage = toPage
            Me.Copies = copies
        End Sub
    End Class

End Class
