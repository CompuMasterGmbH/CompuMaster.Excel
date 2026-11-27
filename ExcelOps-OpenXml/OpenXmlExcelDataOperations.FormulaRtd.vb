Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Threading

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations
        Private _nativeRtd As OpenXmlNativeRtdClient
        Private _nativeRtdCalculationPending As Integer

        Private Function ReadNativeRtd(progId As String, topics As String()) As ExcelFormulaValue
            If _nativeRtd Is Nothing Then _nativeRtd = New OpenXmlNativeRtdClient(AddressOf QueueNativeRtdCalculation)
            Return _nativeRtd.Read(progId, topics)
        End Function

        Private Sub QueueNativeRtdCalculation(source As OpenXmlNativeRtdClient)
            If Interlocked.Exchange(_nativeRtdCalculationPending, 1) <> 0 Then Return
            ThreadPool.QueueUserWorkItem(
                Sub(state)
                    SyncLock WorkbookSaveLock
                        Interlocked.Exchange(_nativeRtdCalculationPending, 0)
                        If IsClosed OrElse source IsNot _nativeRtd OrElse Not UsesAutomaticValueCalculation() Then Return
                        Try
                            Dim calculation As New FormulaCalculation(Me)
                            calculation.RecalculateAfterEdit()
                        Catch ex As Exception
                            Trace.TraceError("Open XML automatic RTD calculation failed: {0}", ex)
                        End Try
                    End SyncLock
                End Sub)
        End Sub

        Private Sub CloseNativeRtd()
            Dim client = _nativeRtd
            _nativeRtd = Nothing
            client?.Dispose()
        End Sub

        Private NotInheritable Partial Class FormulaCalculation
            Private Function EvaluateRtd(arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell,
                                         fallback As ExcelFormulaRegistration) As Object
                ' A session-wide custom function predates the built-in. Retain its existing argument contract.
                If fallback Is Nothing AndAlso (arguments.Count < 3 OrElse arguments.Count > 255) Then
                    Return New FormulaError("#VALUE!")
                End If
                Dim values = arguments.Select(Function(argument) ToFormulaValue(Scalar(EvaluateExpression(argument, formulaCell)))).ToArray()
                Dim progId = If(values.Length > 0 AndAlso values(0).Kind = ExcelFormulaValueKind.Text,
                                TryCast(values(0).Value, String), Nothing)
                Dim provider As ExcelFormulaRegistration = Nothing
                If Not String.IsNullOrWhiteSpace(progId) AndAlso _run.Session IsNot Nothing Then
                    _run.Session.RtdProviders.TryGetValue(progId, provider)
                End If
                If provider IsNot Nothing AndAlso (values.Length < 3 OrElse values.Length > 255) Then
                    Return New FormulaError("#VALUE!")
                End If
                provider = If(provider, fallback)
                If provider IsNot Nothing Then
                    If provider.IsVolatile Then MarkVolatile()
                    Return InvokeRegisteredFunction("RTD", provider, values, formulaCell)
                End If

                For Each value In values
                    If value.Kind = ExcelFormulaValueKind.Error Then Return New FormulaError(DirectCast(value.Value, String))
                Next
                If String.IsNullOrWhiteSpace(progId) Then Return New FormulaError("#VALUE!")
                If values(1).Kind <> ExcelFormulaValueKind.Blank AndAlso values(1).Kind <> ExcelFormulaValueKind.Text Then
                    Return New FormulaError("#VALUE!")
                End If
                Const registrationHint As String = " Register an RTD provider for this ProgID in the workbook's formula session."
                If Not RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then
                    Throw New NotSupportedException("Windows COM RTD is not supported on this platform." & registrationHint)
                End If
                If Not String.IsNullOrEmpty(TryCast(values(1).Value, String)) Then
                    Throw New NotSupportedException("Native RTD supports local COM servers only." & registrationHint)
                End If
                Dim allowed = _owner.LoadOptions.AllowedNativeRtdProgIds
                If allowed Is Nothing OrElse Not allowed.Contains(progId, StringComparer.OrdinalIgnoreCase) Then
                    Throw New NotSupportedException("Native COM RTD for " & progId & " requires an entry in AllowedNativeRtdProgIds." & registrationHint)
                End If

                MarkVolatile()
                Dim topics = values.Skip(2).Select(Function(value) TextValue(value.Value)).ToArray()
                Dim result = _owner.ReadNativeRtd(progId, topics)
                If result.Kind = ExcelFormulaValueKind.Error Then Return New FormulaError(DirectCast(result.Value, String))
                Return result.Value
            End Function
        End Class
    End Class

End Namespace
