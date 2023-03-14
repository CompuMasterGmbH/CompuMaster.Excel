Public Class TestTools

    Public Shared Function EnumValues(Of EnumBaseType As Structure)() As List(Of EnumBaseType)
        Dim Result As New List(Of EnumBaseType)
        For Each Value As EnumBaseType In [Enum].GetValues(GetType(EnumBaseType))
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
