Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type
    ''' </summary>
    Public Class FileReadOnlyException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        Public Sub New(filePath As String)
            MyBase.New("File is write-protected (read-only): " & filePath)
            Me.FilePath = filePath
        End Sub

        Public Property FilePath As String

    End Class
End Namespace