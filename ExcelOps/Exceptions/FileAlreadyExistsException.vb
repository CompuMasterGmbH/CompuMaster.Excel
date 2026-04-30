Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type
    ''' </summary>
    Public Class FileAlreadyExistsException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        ''' <summary>
        ''' Creates an exception for an existing file path.
        ''' </summary>
        ''' <param name="filePath">Path of the file that already exists.</param>
        Public Sub New(filePath As String)
            MyBase.New("File already exists: " & filePath)
            Me.FilePath = filePath
        End Sub

        ''' <summary>
        ''' Gets or sets the path of the file that already exists.
        ''' </summary>
        Public Property FilePath As String

    End Class
End Namespace
