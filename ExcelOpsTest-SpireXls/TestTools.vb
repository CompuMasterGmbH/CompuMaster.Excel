Public Class TestTools

    Public Shared Function EnumValues(Of EnumBaseType As Structure)() As List(Of EnumBaseType)
        Dim Result As New List(Of EnumBaseType)
        For Each Value As EnumBaseType In [Enum].GetValues(GetType(EnumBaseType))
            Result.Add(Value)
        Next
        Return Result
    End Function

End Class
