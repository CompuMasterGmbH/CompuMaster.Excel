Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type.
    ''' </summary>
    Public Class FileReadOnlyException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        ''' <summary>
        ''' Creates an exception for a read-only file.
        ''' </summary>
        ''' <param name="filePath">Path of the read-only file.</param>
        Public Sub New(filePath As String)
            MyBase.New("File is write-protected (read-only): " & filePath)
            Me.FilePath = filePath
        End Sub

        ''' <summary>
        ''' Gets or sets the path of the read-only file.
        ''' </summary>
        Public Property FilePath As String

    End Class
End Namespace
