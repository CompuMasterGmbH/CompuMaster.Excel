Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Shared ReadOnly OmittedFormulaArgument As FormulaExpression = Function(calculation, formulaCell) Nothing
        Private Shared ReadOnly MissingLambdaArgument As New Object()
        Private Shared ReadOnly NativeLambdaFunctions As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "PRODUCT", "SUMSQ", "MEDIAN", "PERCENTOF", "AND", "OR", "XOR", "ABS"
        }

        Private Shared Function IsLambdaArrayFunctionName(name As String) As Boolean
            Select Case name.ToUpperInvariant().Replace("_XLFN.", "")
                Case "BYROW", "BYCOL", "MAP", "MAKEARRAY", "REDUCE", "SCAN" : Return True
                Case Else : Return False
            End Select
        End Function

        Private NotInheritable Class FormulaLambda
            Friend ReadOnly Parameters As IReadOnlyList(Of Object)
            Friend ReadOnly Body As FormulaExpression
            Friend ReadOnly CapturedValues As Dictionary(Of Object, Lazy(Of Object))
            Friend ReadOnly NativeFunction As String

            Friend Sub New(parameters As IReadOnlyList(Of Object), body As FormulaExpression, captured As Dictionary(Of Object, Lazy(Of Object)))
                Me.Parameters = parameters
                Me.Body = body
                CapturedValues = captured
            End Sub

            Friend Sub New(nativeFunction As String)
                Me.NativeFunction = nativeFunction
            End Sub

            Friend Function Accepts(arguments As Integer) As Boolean
                Return NativeFunction IsNot Nothing OrElse Parameters.Count = arguments
            End Function
        End Class

        Private NotInheritable Partial Class FormulaCalculation

            Friend Function CaptureFormulaLambda(parameters As IReadOnlyList(Of Object), body As FormulaExpression) As FormulaLambda
                Return New FormulaLambda(parameters, body, New Dictionary(Of Object, Lazy(Of Object))(_localValues))
            End Function

            Friend Function InvokeFormulaLambda(callable As Object, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If TypeOf callable Is FormulaError Then Return callable
                Dim lambda = TryCast(callable, FormulaLambda)
                If lambda Is Nothing OrElse Not lambda.Accepts(arguments.Count) Then Return New FormulaError("#VALUE!")
                Dim values = arguments.Select(Function(argument)
                                                  Return If(argument Is OmittedFormulaArgument, MissingLambdaArgument, EvaluateExpression(argument, formulaCell))
                                              End Function).ToArray()
                Return CallFormulaLambda(lambda, values, formulaCell)
            End Function

            Private Function CallFormulaLambda(lambda As FormulaLambda, arguments As Object(), formulaCell As ExcelCell) As Object
                If lambda.NativeFunction IsNot Nothing Then
                    Dim expressions = arguments.Select(Function(argument) CType(Function(calculation, cell) argument, FormulaExpression)).ToArray()
                    Return EvaluateFunction(lambda.NativeFunction, expressions, formulaCell)
                End If

                Dim callerValues = _localValues
                _localValues = New Dictionary(Of Object, Lazy(Of Object))(lambda.CapturedValues)
                Try
                    For index = 0 To arguments.Length - 1
                        Dim value = arguments(index)
                        _localValues.Add(lambda.Parameters(index), New Lazy(Of Object)(Function() value))
                    Next
                    Return EvaluateExpression(lambda.Body, formulaCell)
                Finally
                    _localValues = callerValues
                End Try
            End Function

        End Class

    End Class

End Namespace
