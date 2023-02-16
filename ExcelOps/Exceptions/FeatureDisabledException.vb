Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type
    ''' </summary>
    Public Class FeatureDisabledException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        Public Sub New(featureName As String)
            Me.New(featureName, "Feature has been disabled: " & featureName)
        End Sub

        Public Sub New(featureName As String, message As String)
            MyBase.New(message)
            Me.FeatureName = featureName
        End Sub

        Public Property FeatureName As String

    End Class
End Namespace