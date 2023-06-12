Public Class TestTools

    Public Shared Function EnumValues(Of TEnumBaseType As Structure)() As List(Of TEnumBaseType)
        Dim Result As New List(Of TEnumBaseType)
        For Each Value As TEnumBaseType In [Enum].GetValues(GetType(TEnumBaseType))
            Result.Add(Value)
        Next
        Return Result
    End Function

    Public Shared Function IsWindowsPlatform() As Boolean
        Select Case System.Environment.OSVersion.Platform
            Case PlatformID.Win32NT
                Return True
                'Case PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE
                '    Return True
            Case Else
                Return False
        End Select
    End Function
End Class
