Namespace ExcelOps
#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    ''' An exception which is thrown when found data can't be converted into requests data type.
    ''' </summary>
    Public Class FeatureDisabledException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        ''' <summary>
        ''' Creates an exception for a disabled feature.
        ''' </summary>
        ''' <param name="featureName">Name of the disabled feature.</param>
        Public Sub New(featureName As String)
            Me.New(featureName, "Feature has been disabled: " & featureName)
        End Sub

        ''' <summary>
        ''' Creates an exception for a disabled feature with a custom message.
        ''' </summary>
        ''' <param name="featureName">Name of the disabled feature.</param>
        ''' <param name="message">Exception message.</param>
        Public Sub New(featureName As String, message As String)
            MyBase.New(message)
            Me.FeatureName = featureName
        End Sub

        ''' <summary>
        ''' Gets or sets the name of the disabled feature.
        ''' </summary>
        Public Property FeatureName As String

    End Class
End Namespace
