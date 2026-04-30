''' <summary>
''' A wrapper for an Excel worksheet collection.
''' </summary>
<CodeAnalysis.SuppressMessage("Naming", "CA1711:Bezeichner dürfen kein falsches Suffix aufweisen", Justification:="<Ausstehend>")>
Public Class ExcelSheetCollection
    Inherits ComChildObject(Of ExcelWorkbook, Object)

    ''' <summary>
    ''' Creates a wrapper for the worksheets of a workbook.
    ''' </summary>
    ''' <param name="parent">Parent workbook.</param>
    Public Sub New(parent As ExcelWorkbook)
        MyBase.New(parent, parent.InvokePropertyGet(Of Object)("Sheets"))
    End Sub

    ''' <summary>
    ''' Adds a worksheet.
    ''' </summary>
    ''' <returns>The added worksheet.</returns>
    Public Function Add() As ExcelSheet
        Return New ExcelSheet(Me, InvokeFunction(Of Object)("Add"))
    End Function

    ''' <summary>
    ''' Gets a worksheet by name.
    ''' </summary>
    ''' <param name="sheetName">Worksheet name.</param>
    ''' <returns>The matching worksheet.</returns>
    Public Function Item(sheetName As String) As ExcelSheet
        If sheetName = Nothing Then Throw New ArgumentNullException(NameOf(sheetName))
        Return New ExcelSheet(Me, InvokePropertyGet(Of Object)("Item", sheetName))
    End Function

    ''' <summary>
    ''' Gets a worksheet by zero-based index.
    ''' </summary>
    ''' <param name="index">Zero-based worksheet index.</param>
    ''' <returns>The matching worksheet.</returns>
    Public Function Item(index As Integer) As ExcelSheet
        Return New ExcelSheet(Me, InvokePropertyGet(Of Object)("Item", index + 1))
    End Function

    ''' <summary>
    ''' Gets the number of worksheets.
    ''' </summary>
    ''' <returns>The number of worksheets.</returns>
    Public Function Count() As Integer
        Return InvokePropertyGet(Of Integer)("Count")
    End Function

End Class
