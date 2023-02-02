Public Class ExcelSheet
    Inherits ComChildObject(Of ExcelSheetCollection, Object)

    Friend Sub New(parent As ExcelSheetCollection, sheetComObject As Object)
        MyBase.New(parent, sheetComObject)
    End Sub

    Public ReadOnly Property Name As String
        Get
            Return InvokePropertyGet(Of String)("Name")
        End Get
    End Property

    Public ReadOnly Property CodeName As String
        Get
            Return InvokePropertyGet(Of String)("CodeName")
        End Get
    End Property

    Public Sub [Select]()
        InvokeMethod("Select")
    End Sub

    Public Sub Delete()
        InvokeMethod("Delete")
    End Sub

    Private oRanges As New List(Of ExcelRange)

    Public ReadOnly Property Index As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Index") - 1
        End Get
    End Property

    Public Property Visible As Enumerations.XlSheetVisibility
        Get
            Return InvokePropertyGet(Of Enumerations.XlSheetVisibility)("Visible")
        End Get
        Set(value As Enumerations.XlSheetVisibility)
            InvokePropertySet(Of Enumerations.XlSheetVisibility)("Visible", value)
        End Set
    End Property

    Public Sub ExportAsFixedFormat(type As Enumerations.XlFixedFormatType,
                                   fileName As String,
                                   Optional quality As Enumerations.XlFixedFormatQuality = Enumerations.XlFixedFormatQuality.xlQualityStandard,
                                   Optional includeDocProperties As Boolean = False,
                                   Optional ignorePrintAreas As Boolean = False,
                                   Optional fromPageIndex As Integer = 0,
                                   Optional toPageIndex As Integer = Int16.MaxValue - 1,
                                   Optional openAfterPublish As Boolean = False)
        InvokeMethod("ExportAsFixedFormat", type, fileName, quality, includeDocProperties, ignorePrintAreas, fromPageIndex + 1, toPageIndex + 1, openAfterPublish)
    End Sub

    Public Sub PrintOut(Optional fromPageIndex As Integer = 0,
                        Optional toPageIndex As Integer = Int16.MaxValue - 1,
                        Optional copies As Integer = 1,
                        Optional preview As Boolean = False,
                        Optional activePrinter As String = Nothing,
                        Optional printToFile As Boolean = False,
                        Optional collatePages As Boolean = False,
                        Optional printToFileName As String = Nothing,
                        Optional ignorePrintAreas As Boolean = False)
        InvokeMethod("PrintOut", fromPageIndex + 1, toPageIndex + 1, copies, preview, activePrinter, printToFile, collatePages, printToFileName, ignorePrintAreas)
    End Sub

    Protected Overrides Sub OnDisposeChildren()
        For MyCounter As Integer = oRanges.Count - 1 To 0 Step -1
            oRanges(MyCounter).Dispose()
            oRanges.RemoveAt(MyCounter)
        Next
    End Sub

End Class
