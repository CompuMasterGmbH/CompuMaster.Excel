Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateRegisteredFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                        formulaCell As ExcelCell) As Object
                Dim session = _run.Session
                Dim registration As ExcelFormulaRegistration = Nothing
                If session Is Nothing OrElse Not session.Functions.TryGetValue(name, registration) Then
                    Throw New NotSupportedException("The function " & name & " is outside the supported calculation scope.")
                End If

                If registration.IsVolatile Then MarkVolatile()
                Dim values = arguments.Select(Function(argument) ToFormulaValue(Scalar(EvaluateExpression(argument, formulaCell)))).ToArray()
                Dim views As New Dictionary(Of String, ExcelFormulaWorkbookReference)(StringComparer.OrdinalIgnoreCase)
                Dim inCall = True
                For Each item In session.Workbooks
                    Dim workbook = TryCast(item.Value, OpenXmlExcelDataOperations)
                    If workbook Is Nothing OrElse workbook.IsClosed Then Continue For
                    views.Add(item.Key, New ExcelFormulaWorkbookReference(item.Key,
                        Function(address)
                            If Not inCall Then Throw New InvalidOperationException("The custom formula invocation has ended.")
                            ValidateCoordinates(address.RowIndex, address.ColumnIndex)
                            Return ToFormulaValue(ForWorkbook(workbook).ReadCell(address.SheetName, address.RowIndex, address.ColumnIndex))
                        End Function))
                Next
                Dim workbookName = If(_owner._formulaWorkbookName,
                    session.Workbooks.First(Function(item) item.Value Is _owner).Key)
                Dim context As New ExcelFormulaContext(views(workbookName), New ExcelFormulaWorkbookSession(views),
                                                       formulaCell, session.ActiveWorkbook, session.ActiveCell, nativeExcelApplication:=Nothing)
                Try
                    Dim result = registration.Implementation(context, Array.AsReadOnly(values))
                    If result.Kind = ExcelFormulaValueKind.Error Then Return New FormulaError(DirectCast(result.Value, String))
                    Return result.Value
                Catch ex As Exception When Not TypeOf ex Is CircularReferenceException AndAlso
                                           Not TypeOf ex Is FormulaComplexityNotSupportedException AndAlso
                                           Not TypeOf ex Is FeatureDisabledException
                    Throw New InvalidOperationException("The custom formula function " & name & " failed.", ex)
                Finally
                    inCall = False
                End Try
            End Function

            Private Shared Function ToFormulaValue(value As Object) As ExcelFormulaValue
                If TypeOf value Is Double Then Return ExcelFormulaValue.FromNumber(DirectCast(value, Double))
                If TypeOf value Is Boolean Then Return ExcelFormulaValue.FromBoolean(DirectCast(value, Boolean))
                If TypeOf value Is String Then Return ExcelFormulaValue.FromText(DirectCast(value, String))
                If TypeOf value Is FormulaError Then Return ExcelFormulaValue.FromError(DirectCast(value, FormulaError).Code)
                Return Nothing
            End Function

        End Class

    End Class

End Namespace
