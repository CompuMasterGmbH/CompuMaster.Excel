Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when a binary Excel workbook cannot be opened by the selected engine.
    ''' </summary>
    Public Class BinaryXlsFileNotSupportedException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits FileCorruptedOrInvalidFileFormatException

        ''' <summary>
        ''' Creates an exception for an unsupported binary Excel workbook.
        ''' </summary>
        ''' <param name="filePath">Path of the affected file.</param>
        Public Sub New(filePath As String)
            MyBase.New(filePath)
        End Sub

        ''' <summary>
        ''' Creates an exception for an unsupported binary Excel workbook.
        ''' </summary>
        ''' <param name="filePath">Path of the affected file.</param>
        ''' <param name="innerException">Original exception that caused this exception.</param>
        Public Sub New(filePath As String, innerException As Exception)
            MyBase.New(filePath, innerException)
        End Sub

        ''' <summary>
        ''' Creates an exception for an unsupported binary Excel workbook.
        ''' </summary>
        ''' <param name="file">Affected file.</param>
        Public Sub New(file As System.IO.FileInfo)
            Me.New(file.FullName)
        End Sub

        ''' <summary>
        ''' Creates an exception for an unsupported binary Excel workbook.
        ''' </summary>
        ''' <param name="file">Affected file.</param>
        ''' <param name="innerException">Original exception that caused this exception.</param>
        Public Sub New(file As System.IO.FileInfo, innerException As Exception)
            Me.New(file.FullName, innerException)
        End Sub

    End Class
End Namespace
