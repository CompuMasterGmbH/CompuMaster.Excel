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

        ''' <summary>
        ''' Creates an exception for an unexpected sheet layout.
        ''' </summary>
        ''' <param name="targetSheetName">Name of the worksheet with the unexpected layout.</param>
        Public Sub New(targetSheetName As String)
            MyBase.New("Sheet layout at """ & targetSheetName & """ doesn't match to expected sheet layout")
            Me.TargetSheetName = targetSheetName
        End Sub

        ''' <summary>
        ''' Creates an exception for an unexpected sheet layout with additional details.
        ''' </summary>
        ''' <param name="targetSheetName">Name of the worksheet with the unexpected layout.</param>
        ''' <param name="message">Additional mismatch details.</param>
        Public Sub New(targetSheetName As String, message As String)
            MyBase.New("Sheet layout at """ & targetSheetName & """ doesn't match to expected sheet layout: " & message)
            Me.TargetSheetName = targetSheetName
        End Sub

        ''' <summary>
        ''' Gets or sets the worksheet name with the unexpected layout.
        ''' </summary>
        Public Property TargetSheetName As String

    End Class
End Namespace
