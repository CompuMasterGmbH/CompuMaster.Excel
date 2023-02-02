Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data doesn't match with expected data
    ''' </summary>
    Public Class ExpectedSheetLayoutMismatchException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        Public Sub New(targetSheetName As String)
            MyBase.New("Sheet layout at """ & targetSheetName & """ doesn't match to expected sheet layout")
            Me.TargetSheetName = targetSheetName
        End Sub

        Public Sub New(targetSheetName As String, message As String)
            MyBase.New("Sheet layout at """ & targetSheetName & """ doesn't match to expected sheet layout: " & message)
            Me.TargetSheetName = targetSheetName
        End Sub

        Public Property TargetSheetName As String

    End Class
End Namespace