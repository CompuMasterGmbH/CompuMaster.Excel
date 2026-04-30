Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type
    ''' </summary>
    Public Class MissingCalculatedWorkbookException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        ''' <summary>
        ''' Creates an exception for a workbook that still requires calculation.
        ''' </summary>
        ''' <param name="filePath">Path of the workbook.</param>
        Public Sub New(filePath As String)
            MyBase.New("Workbook must not require any re-calculations: " & filePath)
            Me.FilePath = filePath
        End Sub

        ''' <summary>
        ''' Gets or sets the workbook file path.
        ''' </summary>
        Public Property FilePath As String

    End Class
End Namespace
