Imports System.Data.Odbc
Imports System.Reflection

Namespace ExcelOps
    Friend Class Utils

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

        ''' <summary>
        ''' Spire.Xls requires an assigned license
        ''' </summary>
        ''' <returns></returns>
        Friend Shared Function IsLicensedContext() As Boolean
            Dim MemberName As String = Nothing
            Dim MemberField As System.Reflection.FieldInfo = Nothing
            Dim Members = CompuMaster.Reflection.NonPublicStaticMembers.GetMembers(GetType(Spire.License.LicenseProvider))
            For Each Member In Members
                Select Case Member.MemberType
                    Case System.Reflection.MemberTypes.Property
                    Case System.Reflection.MemberTypes.Field
                        If CType(Member, System.Reflection.FieldInfo).FieldType.ToString.StartsWith("System.Collections.Generic.Dictionary") Then
                            MemberName = Member.Name
                            MemberField = CType(Member, System.Reflection.FieldInfo)
                            Exit For
                        End If
                    Case Else
                End Select
            Next
            If MemberName Is Nothing OrElse MemberField Is Nothing Then
                Throw New NotSupportedException("Spire.Xls version not supported (validation of license failed)")
            End If
            Dim Dict As IDictionary
            Dict = CType(GetType(Spire.License.LicenseProvider).InvokeMember(MemberName, System.Reflection.BindingFlags.GetField Or System.Reflection.BindingFlags.Static Or System.Reflection.BindingFlags.NonPublic, Nothing, Nothing, Nothing), IDictionary)
            Return Dict IsNot Nothing AndAlso Dict.Count > 0
        End Function

    End Class
End Namespace