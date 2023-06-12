Public Class TestTools

    Public Shared Function EnumValues(Of TEnumBaseType As Structure)() As List(Of TEnumBaseType)
        Dim Result As New List(Of TEnumBaseType)
        For Each Value As TEnumBaseType In [Enum].GetValues(GetType(TEnumBaseType))
            Result.Add(Value)
        Next
        Return Result
    End Function

End Class
