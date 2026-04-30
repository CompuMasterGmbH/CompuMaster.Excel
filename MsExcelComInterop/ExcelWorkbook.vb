''' <summary>
''' A wrapper for an Excel workbook.
''' </summary>
Public Class ExcelWorkbook
    Inherits ComChildObject(Of ExcelWorkbooksCollection, Object)

    Friend Sub New(parentItem As ExcelWorkbooksCollection, path As String)
        MyBase.New(parentItem, parentItem.InvokeFunction(Of Object)("Open", New Object() {path}))
        FilePath = path
        Sheets = New ExcelSheetCollection(Me)
    End Sub

    ''' <summary>
    ''' Gets the worksheet collection of this workbook.
    ''' </summary>
    Public ReadOnly Property Sheets() As ExcelSheetCollection

    ''' <summary>
    ''' Gets the workbook file path.
    ''' </summary>
    Public ReadOnly Property FilePath As String

    ''' <summary>
    ''' Gets the workbook name.
    ''' </summary>
    Public ReadOnly Property Name As String
        Get
            Return InvokePropertyGet(Of String)("Name")
        End Get
    End Property

    ''' <summary>
    ''' Saves the workbook.
    ''' </summary>
    Public Sub Save()
        InvokeMethod("Save")
    End Sub

    ''' <summary>
    ''' Exports the workbook as a fixed-format file.
    ''' </summary>
    ''' <param name="type">Fixed-format export type.</param>
    ''' <param name="fileName">Target file path.</param>
    ''' <param name="quality">Export quality.</param>
    ''' <param name="includeDocProperties">Whether document properties shall be included.</param>
    ''' <param name="ignorePrintAreas">Whether configured print areas shall be ignored.</param>
    ''' <param name="fromPageIndex">Zero-based first page index to export.</param>
    ''' <param name="toPageIndex">Zero-based last page index to export.</param>
    ''' <param name="openAfterPublish">Whether Excel shall open the exported file.</param>
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

    ''' <summary>
    ''' Prints the workbook.
    ''' </summary>
    ''' <param name="fromPageIndex">Zero-based first page index to print.</param>
    ''' <param name="toPageIndex">Zero-based last page index to print.</param>
    ''' <param name="copies">Number of copies to print.</param>
    ''' <param name="preview">Whether Excel shall show print preview.</param>
    ''' <param name="activePrinter">Printer name or <see langword="Nothing"/> for Excel's active printer.</param>
    ''' <param name="printToFile">Whether output shall be printed to a file.</param>
    ''' <param name="collatePages">Whether printed pages shall be collated.</param>
    ''' <param name="printToFileName">Target file name when printing to a file.</param>
    ''' <param name="ignorePrintAreas">Whether configured print areas shall be ignored.</param>
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

    ''' <inheritdoc/>
    Protected Overrides Sub OnDisposeChildren()
        Sheets.Dispose()
    End Sub

    ''' <inheritdoc/>
    Protected Overrides Sub OnClosing()
        InvokeMethod("Close")
    End Sub

    ''' <inheritdoc/>
    Protected Overrides Sub OnClosed()
        Me.Parent.Workbooks.Remove(Me)
    End Sub

End Class
