Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Function EvaluateRegression(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                formulaCell As ExcelCell) As Object
                If arguments.Count < 1 OrElse arguments.Count > 4 Then Return New FormulaError("#VALUE!")
                Dim fitOnly = name = "LINEST" OrElse name = "LOGEST"
                Dim exponential = name = "LOGEST" OrElse name = "GROWTH"
                Dim evaluated = arguments.Select(Function(expression) EvaluateExpression(expression, formulaCell)).ToArray()
                Dim argumentError = evaluated.OfType(Of FormulaError)().FirstOrDefault()
                If argumentError IsNot Nothing Then Return argumentError

                Dim response = AsMatrix(evaluated(0))
                Dim known = If(evaluated.Length > 1 AndAlso evaluated(1) IsNot Nothing, AsMatrix(evaluated(1)),
                    New FormulaMatrix(response.Rows, response.Columns, Enumerable.Range(1, response.Values.Length).Select(Function(index) CObj(CDbl(index))).ToArray()))
                Dim addIntercept = True
                Dim full = False
                If fitOnly Then
                    If evaluated.Length > 2 AndAlso evaluated(2) IsNot Nothing Then
                        Dim optionValue = LogicalOptionValue(Scalar(evaluated(2)))
                        If TypeOf optionValue Is FormulaError Then Return optionValue
                        addIntercept = DirectCast(optionValue, Boolean)
                    End If
                    If evaluated.Length > 3 AndAlso evaluated(3) IsNot Nothing Then
                        Dim optionValue = LogicalOptionValue(Scalar(evaluated(3)))
                        If TypeOf optionValue Is FormulaError Then Return optionValue
                        full = DirectCast(optionValue, Boolean)
                    End If
                ElseIf evaluated.Length > 3 AndAlso evaluated(3) IsNot Nothing Then
                    Dim optionValue = LogicalOptionValue(Scalar(evaluated(3)))
                    If TypeOf optionValue Is FormulaError Then Return optionValue
                    addIntercept = DirectCast(optionValue, Boolean)
                End If

                Dim vertical = response.Columns = 1
                Dim sameShape = known.Rows = response.Rows AndAlso known.Columns = response.Columns
                If Not sameShape AndAlso Not (If(vertical, response.Rows = known.Rows, response.Rows = 1 AndAlso response.Columns = known.Columns)) Then
                    Return New FormulaError("#REF!")
                End If
                Dim predictors = If(sameShape, 1, If(vertical, known.Columns, known.Rows))
                If response.Values.Any(Function(value) value IsNot Nothing AndAlso Not TypeOf value Is Double) OrElse
                    known.Values.Any(Function(value) value IsNot Nothing AndAlso Not TypeOf value Is Double) Then Return New FormulaError("#VALUE!")

                Dim y = response.Values.Select(Function(value) If(value Is Nothing, 0.0, DirectCast(value, Double))).ToArray()
                If exponential Then
                    If y.Any(Function(value) value <= 0.0) Then Return New FormulaError("#NUM!")
                    For index = 0 To y.Length - 1
                        y(index) = Math.Log(y(index))
                    Next
                End If
                Dim x(predictors - 1)() As Double
                For predictor = 0 To predictors - 1
                    x(predictor) = New Double(y.Length - 1) {}
                    For observation = 0 To y.Length - 1
                        Dim value = If(sameShape, known.Values(observation), If(vertical, known.At(observation, predictor), known.At(predictor, observation)))
                        x(predictor)(observation) = If(value Is Nothing, 0.0, DirectCast(value, Double))
                    Next
                Next

                Dim fit As New RegressionFit(y, x, addIntercept, calculateStatistics:=full)
                If fitOnly Then Return RegressionCoefficients(fit, predictors, exponential, addIntercept, full)
                Dim forecast = If(evaluated.Length > 2 AndAlso evaluated(2) IsNot Nothing, AsMatrix(evaluated(2)), known)
                If predictors > 1 AndAlso If(vertical, forecast.Columns, forecast.Rows) <> predictors Then Return New FormulaError("#REF!")
                Dim outputRows = If(predictors = 1 OrElse vertical, forecast.Rows, 1)
                Dim outputColumns = If(predictors = 1 OrElse Not vertical, forecast.Columns, 1)
                Dim output = AllocateMatrixValues(outputRows, outputColumns)
                For index = 0 To output.Length - 1
                    Dim value = fit.Intercept
                    Dim failed As FormulaError = Nothing
                    For predictor = 0 To predictors - 1
                        Dim input = If(predictors = 1, forecast.Values(index), If(vertical, forecast.At(index, predictor), forecast.At(predictor, index)))
                        If input IsNot Nothing AndAlso Not TypeOf input Is Double Then
                            failed = New FormulaError("#VALUE!")
                            Exit For
                        End If
                        value += fit.Slopes(predictor) * If(input Is Nothing, 0.0, DirectCast(input, Double))
                    Next
                    output(index) = If(failed, FiniteNumber(If(exponential, Math.Exp(value), value)))
                Next
                If output.Length = 1 Then Return output(0)
                Return New FormulaMatrix(outputRows, outputColumns, output)
            End Function

            Private Shared Function RegressionCoefficients(fit As RegressionFit, predictors As Integer, exponential As Boolean,
                                                          includeIntercept As Boolean, full As Boolean) As Object
                Dim columns = predictors + 1
                Dim rows = If(full, 5, 1)
                Dim results = AllocateMatrixValues(rows, columns)
                For index = 0 To results.Length - 1
                    results(index) = New FormulaError("#N/A")
                Next
                For index = 0 To predictors - 1
                    Dim coefficient = fit.Slopes(predictors - index - 1)
                    results(index) = FiniteNumber(If(exponential, Math.Exp(coefficient), coefficient))
                Next
                results(predictors) = FiniteNumber(If(exponential, Math.Exp(fit.Intercept), fit.Intercept))
                If Not full Then Return New FormulaMatrix(rows, columns, results)
                Dim degrees = fit.Observations - fit.StatisticsRank - If(includeIntercept, 1, 0)
                Dim sigma = If(degrees > 0, Math.Sqrt(fit.ResidualSum / degrees), 0.0)
                For index = 0 To predictors - 1
                    results(columns + index) = FiniteNumber(sigma * fit.SlopeMultipliers(predictors - index - 1))
                Next
                If includeIntercept Then results(columns + predictors) = FiniteNumber(sigma * fit.InterceptMultiplier)
                results(2 * columns) = FiniteNumber(If(fit.TotalSum = 0.0, 1.0, fit.RegressionSum / fit.TotalSum))
                results(2 * columns + 1) = FiniteNumber(sigma)
                results(3 * columns) = If(fit.ResidualSum <= 0.0 OrElse degrees <= 0, CObj(New FormulaError("#NUM!")),
                                         FiniteNumber(fit.RegressionSum * degrees / (fit.ResidualSum * Math.Max(1, fit.StatisticsRank))))
                results(3 * columns + 1) = CDbl(degrees)
                results(4 * columns) = FiniteNumber(fit.RegressionSum)
                results(4 * columns + 1) = FiniteNumber(fit.ResidualSum)
                Return New FormulaMatrix(rows, columns, results)
            End Function

            Private NotInheritable Class RegressionFit
                Friend ReadOnly Observations As Integer
                Friend ReadOnly Rank As Integer
                Friend ReadOnly StatisticsRank As Integer
                Friend ReadOnly Slopes As Double()
                Friend ReadOnly SlopeMultipliers As Double()
                Friend ReadOnly Intercept As Double
                Friend ReadOnly InterceptMultiplier As Double
                Friend ReadOnly TotalSum As Double
                Friend ReadOnly RegressionSum As Double
                Friend ReadOnly ResidualSum As Double

                Friend Sub New(y As Double(), x As Double()(), includeIntercept As Boolean, calculateStatistics As Boolean)
                    Observations = y.Length
                    Slopes = New Double(x.Length - 1) {}
                    SlopeMultipliers = New Double(x.Length - 1) {}
                    Dim yMean = If(includeIntercept, y.Average(), 0.0)
                    Dim target = y.Select(Function(value) value - yMean).ToArray()
                    TotalSum = Dot(target, target)
                    Dim limit = Math.Min(x.Length, Observations - If(includeIntercept, 1, 0))
                    Dim basis As New List(Of Double())(limit)
                    Dim selected As New List(Of Integer)(limit)
                    Dim means(x.Length - 1) As Double
                    Dim scales(x.Length - 1) As Double
                    Dim triangle(limit - 1, limit - 1) As Double
                    Dim emptyPredictors As Integer
                    ' Select independent columns from right to left, as used by the coefficient output.
                    For predictor = x.Length - 1 To 0 Step -1
                        If basis.Count >= limit Then Exit For
                        Dim input = x(predictor)
                        means(predictor) = If(includeIntercept, input.Average(), 0.0)
                        Dim vector = input.Select(Function(value) value - means(predictor)).ToArray()
                        Dim scale = vector.Max(Function(value) Math.Abs(value))
                        If scale = 0.0 Then
                            emptyPredictors += 1
                            Continue For
                        End If
                        scales(predictor) = scale
                        For index = 0 To vector.Length - 1
                            vector(index) /= scale
                        Next
                        Dim originalNorm = Math.Sqrt(Dot(vector, vector))
                        For pass = 0 To 1
                            For index = 0 To basis.Count - 1
                                Dim projection = Dot(basis(index), vector)
                                If pass = 0 Then triangle(index, basis.Count) = 0.0
                                triangle(index, basis.Count) += projection
                                For observation = 0 To vector.Length - 1
                                    vector(observation) -= projection * basis(index)(observation)
                                Next
                            Next
                        Next
                        Dim norm = Math.Sqrt(Dot(vector, vector))
                        If norm <= originalNorm * 1.0E-13 Then Continue For
                        triangle(basis.Count, basis.Count) = norm
                        For index = 0 To vector.Length - 1
                            vector(index) /= norm
                        Next
                        selected.Add(predictor)
                        basis.Add(vector)
                    Next

                    Rank = basis.Count
                    Dim projected = basis.Select(Function(vector) Dot(vector, target)).ToArray()
                    Dim coefficients = SolveTriangle(triangle, projected)
                    Intercept = yMean
                    For index = 0 To Rank - 1
                        Dim predictor = selected(index)
                        Slopes(predictor) = coefficients(index) / scales(predictor)
                        Intercept -= Slopes(predictor) * means(predictor)
                    Next
                    If Not calculateStatistics Then Return
                    Dim residuals = DirectCast(target.Clone(), Double())
                    For index = 0 To Rank - 1
                        For observation = 0 To Observations - 1
                            residuals(observation) -= projected(index) * basis(index)(observation)
                        Next
                    Next
                    Dim residual = Dot(residuals, residuals)
                    Dim unusedSlots = Math.Min(emptyPredictors, Observations - Rank - If(includeIntercept, 1, 0))
                    StatisticsRank = Rank + unusedSlots
                    If unusedSlots > 0 Then
                        ' An exactly zero predictor still occupies a slot in Excel's QR display statistics.
                        Dim displayed = HouseholderCoordinates(y, basis, includeIntercept)
                        Dim first = Rank + If(includeIntercept, 1, 0)
                        For index = first To first + unusedSlots - 1
                            residual -= displayed(index) * displayed(index)
                        Next
                        residual = Math.Max(0.0, residual)
                    End If
                    ResidualSum = If(residual <= TotalSum * 1.0E-27, 0.0, residual)
                    RegressionSum = Math.Max(0.0, TotalSum - ResidualSum)
                    Dim interceptVariance = 1.0 / Observations
                    For column = 0 To Rank - 1
                        Dim unit(Rank - 1) As Double
                        unit(column) = 1.0
                        Dim inverseColumn = SolveTriangle(triangle, unit)
                        Dim interceptComponent = 0.0
                        For index = 0 To Rank - 1
                            Dim predictor = selected(index)
                            Dim coefficient = inverseColumn(index) / scales(predictor)
                            SlopeMultipliers(predictor) += coefficient * coefficient
                            interceptComponent += means(predictor) * coefficient
                        Next
                        interceptVariance += interceptComponent * interceptComponent
                    Next
                    For index = 0 To Slopes.Length - 1
                        SlopeMultipliers(index) = Math.Sqrt(SlopeMultipliers(index))
                    Next
                    InterceptMultiplier = Math.Sqrt(interceptVariance)
                End Sub

                Private Shared Function HouseholderCoordinates(y As Double(), basis As List(Of Double()), includeIntercept As Boolean) As Double()
                    Dim vectors As New List(Of Double())(basis.Count + 1)
                    If includeIntercept Then vectors.Add(Enumerable.Repeat(1.0 / Math.Sqrt(y.Length), y.Length).ToArray())
                    vectors.AddRange(basis.Select(Function(vector) DirectCast(vector.Clone(), Double())))
                    Dim result = DirectCast(y.Clone(), Double())
                    For stepIndex = 0 To vectors.Count - 1
                        Dim direction = vectors(stepIndex)
                        Dim squaredLength = 0.0
                        For row = stepIndex To direction.Length - 1
                            squaredLength += direction(row) * direction(row)
                        Next
                        Dim norm = Math.Sqrt(squaredLength)
                        Dim leading = direction(stepIndex)
                        direction(stepIndex) += If(leading < 0.0, -norm, norm)
                        Dim weight = 1.0 / (norm * (norm + Math.Abs(leading)))
                        ReflectVector(result, direction, stepIndex, weight)
                        For index = stepIndex + 1 To vectors.Count - 1
                            ReflectVector(vectors(index), direction, stepIndex, weight)
                        Next
                    Next
                    Return result
                End Function

                Private Shared Sub ReflectVector(values As Double(), direction As Double(), first As Integer, weight As Double)
                    Dim projection = 0.0
                    For index = first To values.Length - 1
                        projection += values(index) * direction(index)
                    Next
                    projection *= weight
                    For index = first To values.Length - 1
                        values(index) -= projection * direction(index)
                    Next
                End Sub

                Private Shared Function Dot(first As Double(), second As Double()) As Double
                    Dim total = 0.0
                    For index = 0 To first.Length - 1
                        total += first(index) * second(index)
                    Next
                    Return total
                End Function

                Private Shared Function SolveTriangle(triangle As Double(,), right As Double()) As Double()
                    Dim result = DirectCast(right.Clone(), Double())
                    For index = result.Length - 1 To 0 Step -1
                        For column = index + 1 To result.Length - 1
                            result(index) -= triangle(index, column) * result(column)
                        Next
                        result(index) /= triangle(index, index)
                    Next
                    Return result
                End Function
            End Class

        End Class

    End Class

End Namespace
