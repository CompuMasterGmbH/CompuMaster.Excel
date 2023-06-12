Friend NotInheritable Class Utils

    ''' <summary>
    '''     Return the string which is not nothing or else String.Empty
    ''' </summary>
    ''' <param name="value">The string to be validated</param>
    ''' <returns></returns>
    <DebuggerHidden()> Public Shared Function StringNotEmptyOrNothing(ByVal value As String) As String
        If value = Nothing Then
            Return Nothing
        Else
            Return value
        End If
    End Function

End Class
