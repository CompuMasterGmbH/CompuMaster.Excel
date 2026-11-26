Imports System.Diagnostics
Imports NUnit.Framework

Friend NotInheritable Class HtmlVisualReview

    Private Const SelectionVariable As String = "COMPUMASTER_EXCEL_HTML_VISUAL_TESTS"

    Private Sub New()
    End Sub

    Friend Shared Sub OpenIfRequested(filePath As String, openByDefault As Boolean)
#If CI_CD Then
        Return
#Else
        Dim test = TestContext.CurrentContext.Test
        Dim ci = Environment.GetEnvironmentVariable("CI")
        Dim runningInCi = String.Equals(ci, "true", StringComparison.OrdinalIgnoreCase) OrElse ci = "1"
        If Not ShouldOpen(Environment.GetEnvironmentVariable(SelectionVariable), test.MethodName, test.FullName,
                          openByDefault, runningInCi) Then Return

        Using browser = Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
        End Using
#End If
    End Sub

    Friend Shared Function ShouldOpen(selection As String, methodName As String, fullName As String,
                                      openByDefault As Boolean, runningInCi As Boolean) As Boolean
        If runningInCi Then Return False
        If String.IsNullOrWhiteSpace(selection) Then Return openByDefault
        If String.Equals(selection.Trim(), "off", StringComparison.OrdinalIgnoreCase) Then Return False

        Return selection.Split(","c).Any(
            Function(name) String.Equals(name.Trim(), methodName, StringComparison.Ordinal) OrElse
                           String.Equals(name.Trim(), fullName, StringComparison.Ordinal))
    End Function

End Class
