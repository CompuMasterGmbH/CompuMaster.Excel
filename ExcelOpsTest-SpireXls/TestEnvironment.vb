Imports NUnit.Framework

Public NotInheritable Class TestEnvironment

#If CI_CD = True Then
    Public Shared Property ConsoleOutputSuppressed As Boolean = True
#Else
    Public Shared Property ConsoleOutputSuppressed As Boolean = False
#End If

    Public Shared Function DirectoryOfTestAssembly() As String
        Return System.IO.Path.GetDirectoryName(GetType(TestEnvironment).Assembly.Location)
    End Function

    Private Shared Function GetCallingMethodName() As String
        Dim stackTrace = New StackTrace()
        Return stackTrace.GetFrame(2).GetMethod().Name 'Drop 2 frames for this function and the calling function to provide the calling method name outside of TestEnvironment class
    End Function

    Public Shared Function FullPathOfExistingTestFile(ParamArray subDirsAndFile As String()) As String
        Dim Paths As New List(Of String)(subDirsAndFile)
        Paths.Insert(0, DirectoryOfTestAssembly)
        Dim Result As String = System.IO.Path.Combine(Paths.ToArray)
        If Result.Length > 250 Then 'To prevent exceeeding issues caused by the 260 chars limit on windows platforms
            'Shorten sub dir name for GetCallingMethodName by hashing it
            Paths(2) = Paths(2).GetHashCode.ToString
            Result = System.IO.Path.Combine(Paths.ToArray)
        End If
        If System.IO.File.Exists(Result) = False Then
            Throw New System.IO.FileNotFoundException(Result)
        End If
        Return Result
    End Function

    Public Shared Function FullPathOfDynTestFile(ParamArray subDirsAndFile As String()) As String
        Dim Paths As New List(Of String)(subDirsAndFile)
        Paths.Insert(0, DirectoryOfTestAssembly)
        Paths.Insert(1, "temp")
        Paths.Insert(2, GetCallingMethodName)
        Dim Result As String = System.IO.Path.Combine(Paths.ToArray)
        Dim ParentDir As String = System.IO.Path.GetDirectoryName(Result)
        If Result.Length > 250 Then 'To prevent exceeeding issues caused by the 260 chars limit on windows platforms
            'Shorten sub dir name for GetCallingMethodName by hashing it
            Paths(2) = Paths(2).GetHashCode.ToString
            Result = System.IO.Path.Combine(Paths.ToArray)
            ParentDir = System.IO.Path.GetDirectoryName(Result)
        End If
        If System.IO.Directory.Exists(ParentDir) = False Then
            System.IO.Directory.CreateDirectory(ParentDir)
        End If
        If System.IO.File.Exists(Result) Then
            System.IO.File.Delete(Result)
        End If
        Return Result
    End Function

    Public Shared Function FullPathOfDynTestFile_KeepExistingFile(ParamArray subDirsAndFile As String()) As String
        Dim Paths As New List(Of String)(subDirsAndFile)
        Paths.Insert(0, DirectoryOfTestAssembly)
        Paths.Insert(1, "temp")
        Paths.Insert(2, GetCallingMethodName)
        Dim Result As String = System.IO.Path.Combine(Paths.ToArray)
        Dim ParentDir As String = System.IO.Path.GetDirectoryName(Result)
        If System.IO.Directory.Exists(ParentDir) = False Then
            System.IO.Directory.CreateDirectory(ParentDir)
        End If
        Return Result
    End Function

End Class