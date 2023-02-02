Public Class ExcelWorkbook
    Inherits ComObjectBase

    Friend Sub New(parentItemResponsibleForDisposal As ComObjectBase, c As ExcelWorkbooksCollection, path As String)
        MyBase.New(parentItemResponsibleForDisposal, c.InvokeFunction(Of Object)("Open", New Object() {path}))
        Parent = c
        FilePath = path
        Sheets = New ExcelSheetCollection(Me, Me)
    End Sub

    Public ReadOnly Property Sheets() As ExcelSheetCollection

    Public ReadOnly FilePath As String

    Friend ReadOnly Parent As ExcelWorkbooksCollection

    Public ReadOnly Property Name As String
        Get
            Return InvokePropertyGet(Of String)("Name")
        End Get
    End Property

    Public Sub Save()
        InvokeMethod("Save")
    End Sub

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

    Public ReadOnly Property IsClosed As Boolean
        Get
            Return MyBase.IsDisposedComObject
        End Get
    End Property

    Public Sub Close()
        MyBase.CloseAndDisposeChildrenAndComObject()
    End Sub

    Protected Overrides Sub OnDisposeChildren()
        Sheets.Dispose()
    End Sub

    Protected Overrides Sub OnClosing()
        InvokeMethod("Close")
    End Sub

    Protected Overrides Sub OnClosed()
        Parent.Workbooks.Remove(Me)
    End Sub

End Class
