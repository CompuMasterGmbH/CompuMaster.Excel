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
    End Sub

    Private Shared IsNewLogicalTestConsole As Boolean = True

    Public Shared Sub ResetConsoleForTestOutput()
        IsNewLogicalTestConsole = True
    End Sub

End Class