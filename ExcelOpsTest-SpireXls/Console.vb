Public NotInheritable Class Console

    Public Shared Sub WriteLine()
        WriteLine("")
    End Sub

    Public Shared Sub WriteLine(text As String)
        Write(text & System.Environment.NewLine)
    End Sub

    Public Shared Sub Write()
        Write("")
    End Sub

    Public Shared Sub Write(text As String)
        If TestEnvironment.ConsoleOutputSuppressed Then
            If IsNewLogicalTestConsole = True Then
                System.Console.WriteLine("Test output available if TestEnvionment.")
                IsNewLogicalTestConsole = False
            Else
                'ignore/suppress console output
            End If
        Else
            System.Console.Write(text)
        End If
        Log.Append(text)
    End Sub

    Private Shared IsNewLogicalTestConsole As Boolean = True

    Public Shared Sub ResetConsoleForTestOutput()
        IsNewLogicalTestConsole = True
        Log.Clear()
    End Sub

    Private Shared Log As New System.Text.StringBuilder()

    Public Shared Function GetConsoleLog() As String
        Return Log.ToString
    End Function

End Class