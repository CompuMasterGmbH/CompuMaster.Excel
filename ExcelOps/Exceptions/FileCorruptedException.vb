Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type
    ''' </summary>
    Public Class FileCorruptedOrInvalidFileFormatException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        Public Sub New(filePath As String)
            MyBase.New(filePath, Nothing)
        End Sub

        Public Sub New(filePath As String, innerException As Exception)
            MyBase.New("File corrupted or invalid file format: " & filePath, innerException)
            Me.FilePath = filePath
        End Sub

        Public Sub New(file As System.IO.FileInfo)
            Me.New(file.FullName)
        End Sub

        Public Sub New(file As System.IO.FileInfo, innerException As Exception)
            Me.New(file.FullName, innerException)
        End Sub

        Public Property FilePath As String

    End Class
End Namespace