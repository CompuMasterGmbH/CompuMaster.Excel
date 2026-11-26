Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Const LargestConsecutiveInteger As Double = 9007199254740991.0
            Private Const LargestWorksheetTrialCount As Double = 2147483646.0

            Private Function EvaluateDiscreteDistribution(name As String, arguments As IReadOnlyList(Of FormulaExpression),
                                                          formulaCell As ExcelCell) As Object
                Dim minimum As Integer
                Dim cumulativeIndex = -1
                Dim defaults As Double() = Nothing
                Select Case name
                    Case "BINOM.DIST", "BINOMDIST", "NEGBINOM.DIST"
                        minimum = 4
                        cumulativeIndex = 3
                    Case "HYPGEOM.DIST"
                        minimum = 5
                        cumulativeIndex = 4
                    Case "HYPGEOMDIST"
                        minimum = 4
                    Case "POISSON.DIST", "POISSON"
                        minimum = 3
                        cumulativeIndex = 2
                    Case "BINOM.DIST.RANGE"
                        minimum = 3
                        defaults = {0.0, 0.0, 0.0, Double.NaN}
                    Case Else
                        minimum = 3
                End Select
                If arguments.Count < minimum OrElse arguments.Count > If(name = "BINOM.DIST.RANGE", 4, minimum) Then
                    Return New FormulaError("#VALUE!")
                End If
                Return EvaluateProbabilityArguments(arguments, formulaCell, False, cumulativeIndex,
                    Function(numbers, cumulative) DiscreteDistributionValue(name, numbers, cumulative), defaults)
            End Function

            Private Shared Function DiscreteDistributionValue(name As String, numbers As Double(), cumulative As Boolean) As Object
                Dim first = numbers(0)
                Dim second = numbers(1)
                Dim third = numbers(2)
                Select Case name
                    Case "BINOM.DIST", "BINOMDIST"
                        If Not ValidDistributionCount(first) OrElse Not ValidDistributionCount(second) OrElse third < 0.0 OrElse third > 1.0 Then
                            Return New FormulaError("#NUM!")
                        End If
                        Dim successes = Math.Truncate(first)
                        Dim trials = Math.Truncate(second)
                        If successes > trials OrElse trials > LargestWorksheetTrialCount Then Return New FormulaError("#NUM!")
                        If cumulative Then Return NormalizeWorksheetUnderflow(BinomialTail(successes, trials, third, False))
                        If third = 1.0 AndAlso successes = 0.0 AndAlso trials > 0.0 Then Return New FormulaError("#NUM!")
                        Return NormalizeWorksheetUnderflow(Math.Exp(BinomialLogMass(successes, trials, third)))
                    Case "BINOM.INV", "CRITBINOM"
                        If Not ValidDistributionCount(first) OrElse second <= 0.0 OrElse second >= 1.0 OrElse third <= 0.0 OrElse third >= 1.0 Then
                            Return New FormulaError("#NUM!")
                        End If
                        Dim trials = Math.Truncate(first)
                        If trials > LargestWorksheetTrialCount Then Return New FormulaError("#NUM!")
                        Return InverseBinomial(trials, second, third)
                    Case "BINOM.DIST.RANGE"
                        Return BinomialRange(first, second, third, numbers(3))
                    Case "NEGBINOM.DIST", "NEGBINOMDIST"
                        If Not ValidDistributionCount(first) OrElse Not ValidDistributionCount(second) OrElse second < 1.0 OrElse
                            third <= 0.0 OrElse third >= 1.0 Then Return New FormulaError("#NUM!")
                        Dim failures = Math.Truncate(first)
                        Dim successes = Math.Truncate(second)
                        If failures = 0.0 Then Return NormalizeWorksheetUnderflow(Math.Pow(third, successes))
                        Dim complement = 1.0 - third
                        Dim mass = Math.Exp(BetaLogPrefix(third, complement, successes, failures) - Math.Log(failures))
                        If Not cumulative OrElse name = "NEGBINOMDIST" Then Return NormalizeWorksheetUnderflow(mass)
                        If third = 0.5 AndAlso failures = successes Then Return 0.5 + mass
                        Return NormalizeWorksheetUnderflow(BetaProbability(third, complement, successes, failures + 1.0))
                    Case "POISSON.DIST", "POISSON"
                        If Not ValidDistributionCount(first) Then Return New FormulaError("#NUM!")
                        Dim occurrences = Math.Truncate(first)
                        If occurrences = 0.0 Then Return FiniteNumber(NormalizeWorksheetUnderflow(Math.Exp(-second)))
                        If second < 0.0 Then Return New FormulaError("#NUM!")
                        If second = 0.0 Then Return If(cumulative, 1.0, 0.0)
                        If cumulative Then Return NormalizeWorksheetUnderflow(GammaProbability(occurrences + 1.0, second, True))
                        Return NormalizeWorksheetUnderflow(Math.Exp(GammaLogPrefix(occurrences + 1.0, second) - Math.Log(second)))
                    Case Else
                        Return HypergeometricValue(first, second, third, numbers(3), cumulative AndAlso name = "HYPGEOM.DIST")
                End Select
            End Function

            Private Shared Function ValidDistributionCount(value As Double) As Boolean
                Return value >= 0.0 AndAlso value <= LargestConsecutiveInteger
            End Function

            Private Shared Function BinomialLogMass(successes As Double, trials As Double, probability As Double) As Double
                If trials = 0.0 Then Return 0.0
                If probability = 0.0 Then Return If(successes = 0.0, 0.0, Double.NegativeInfinity)
                If probability = 1.0 Then Return If(successes = trials, 0.0, Double.NegativeInfinity)
                Dim complement = 1.0 - probability
                Dim failures = trials - successes
                If successes = 0.0 Then Return trials * LogBetaSide(complement, probability)
                If failures = 0.0 Then Return trials * LogBetaSide(probability, complement)
                Return BetaLogPrefix(probability, complement, successes, failures) +
                    Math.Log(trials) - Math.Log(successes) - Math.Log(failures)
            End Function

            Private Shared Function BinomialTail(successes As Double, trials As Double, probability As Double, upperTail As Boolean) As Double
                If successes < 0.0 Then Return If(upperTail, 1.0, 0.0)
                If successes >= trials OrElse probability = 0.0 Then Return If(upperTail, 0.0, 1.0)
                If probability = 1.0 Then Return If(upperTail, 1.0, 0.0)
                If probability = 0.5 AndAlso successes * 2.0 = trials Then
                    Dim halfMass = 0.5 * Math.Exp(BinomialLogMass(successes, trials, probability))
                    Return If(upperTail, 0.5 - halfMass, 0.5 + halfMass)
                End If
                Dim complement = 1.0 - probability
                Dim failures = trials - successes
                If successes = 0.0 Then
                    If upperTail Then Return ExponentialProbability(-trials * LogBetaSide(complement, probability))
                    Return If(probability < 1.0E-8, Math.Exp(trials * LogBetaSide(complement, probability)), Math.Pow(complement, trials))
                End If
                If failures = 1.0 Then
                    If upperTail Then Return Math.Pow(probability, trials)
                    Return ExponentialProbability(-trials * LogBetaSide(probability, complement))
                End If
                If upperTail Then Return BetaProbability(probability, complement, successes + 1.0, failures)
                Return BetaProbability(1.0 - probability, probability, failures, successes + 1.0)
            End Function

            Private Shared Function InverseBinomial(trials As Double, probability As Double, criterion As Double) As Double
                Dim lower = 0.0
                Dim upper = trials
                Dim upperTail = criterion > 0.5
                While lower < upper
                    Dim middle = lower + Math.Floor((upper - lower) * 0.5)
                    Dim tail = BinomialTail(middle, trials, probability, upperTail)
                    If If(upperTail, tail <= 1.0 - criterion, tail >= criterion) Then
                        upper = middle
                    Else
                        lower = middle + 1.0
                    End If
                End While
                Return lower
            End Function

            Private Shared Function BinomialRange(rawTrials As Double, probability As Double, rawLower As Double, rawUpper As Double) As Object
                If Double.IsNaN(rawUpper) Then rawUpper = rawLower
                If Not ValidDistributionCount(rawTrials) OrElse Not ValidDistributionCount(rawLower) OrElse Not ValidDistributionCount(rawUpper) OrElse
                    probability < 0.0 OrElse probability > 1.0 Then Return New FormulaError("#NUM!")
                Dim trials = Math.Truncate(rawTrials)
                Dim lower = Math.Truncate(rawLower)
                Dim upper = Math.Truncate(rawUpper)
                If lower > upper OrElse upper > trials OrElse trials > LargestWorksheetTrialCount Then Return New FormulaError("#NUM!")
                If probability = 0.0 Then Return If(lower = 0.0, 1.0, 0.0)
                If probability = 1.0 Then Return If(upper = trials, 1.0, 0.0)
                If upper - lower <= 64.0 Then
                    Dim sum = 0.0
                    For current = lower To upper
                        sum += Math.Exp(BinomialLogMass(current, trials, probability))
                    Next
                    Return NormalizeWorksheetUnderflow(Math.Min(1.0, sum))
                End If
                Dim upperTail = lower > trials * probability
                Dim before = BinomialTail(lower - 1.0, trials, probability, upperTail)
                Dim through = BinomialTail(upper, trials, probability, upperTail)
                Return NormalizeWorksheetUnderflow(Math.Max(0.0, If(upperTail, before - through, through - before)))
            End Function

            Private Shared Function HypergeometricValue(rawSuccesses As Double, rawDraws As Double, rawPopulationSuccesses As Double,
                                                         rawPopulation As Double, cumulative As Boolean) As Object
                If Not ValidDistributionCount(rawSuccesses) OrElse Not ValidDistributionCount(rawDraws) OrElse
                    Not ValidDistributionCount(rawPopulationSuccesses) OrElse Not ValidDistributionCount(rawPopulation) Then Return New FormulaError("#NUM!")
                Dim successes = Math.Truncate(rawSuccesses)
                Dim draws = Math.Truncate(rawDraws)
                Dim populationSuccesses = Math.Truncate(rawPopulationSuccesses)
                Dim population = Math.Truncate(rawPopulation)
                If draws > population OrElse populationSuccesses > population OrElse population > LargestWorksheetTrialCount Then
                    Return New FormulaError("#NUM!")
                End If
                Dim minimum = Math.Max(0.0, draws - (population - populationSuccesses))
                Dim maximum = Math.Min(draws, populationSuccesses)
                If successes < minimum Then Return 0.0
                If successes > maximum Then Return If(cumulative, 1.0, 0.0)
                If minimum = maximum Then Return 1.0
                Dim probability = draws / population
                Dim expected = populationSuccesses * probability
                Dim mass = HypergeometricMass(successes, draws, populationSuccesses, population, probability)
                If Not cumulative Then Return NormalizeWorksheetUnderflow(mass)
                If successes >= maximum Then Return 1.0
                If successes = expected AndAlso (populationSuccesses * 2.0 = population OrElse draws * 2.0 = population) Then
                    Return 0.5 + 0.5 * mass
                End If

                ' Summing outward from the requested point keeps the recurrence in the smaller tail.
                If successes <= expected Then Return NormalizeWorksheetUnderflow(HypergeometricLower(successes, draws, populationSuccesses, population, mass))
                Dim reflected = draws - successes - 1.0
                Dim other = population - populationSuccesses
                Dim upperMass = HypergeometricMass(reflected, draws, other, population, probability)
                Return Math.Max(0.0, 1.0 - HypergeometricLower(reflected, draws, other, population, upperMass))
            End Function

            Private Shared Function HypergeometricMass(successes As Double, draws As Double, populationSuccesses As Double,
                                                        population As Double, probability As Double) As Double
                ' Conditioning two binomial counts on their sum avoids subtraction of large factorial logarithms.
                Dim result = BinomialLogMass(successes, populationSuccesses, probability) +
                    BinomialLogMass(draws - successes, population - populationSuccesses, probability) -
                    BinomialLogMass(draws, population, probability)
                Return Math.Min(1.0, Math.Exp(result))
            End Function

            Private Shared Function HypergeometricLower(successes As Double, draws As Double, populationSuccesses As Double,
                                                         population As Double, mass As Double) As Double
                Dim populationFailures = population - populationSuccesses
                Dim support = Math.Max(0.0, draws - populationFailures)
                Dim sum = mass
                For current = successes To support + 1.0 Step -1.0
                    If mass = 0.0 Then Exit For
                    Dim previousFailures = draws - current
                    mass *= (current / (populationSuccesses - current + 1.0)) *
                        ((populationFailures - previousFailures) / (previousFailures + 1.0))
                    sum += mass
                    If mass <= sum * 1.0E-16 Then Exit For
                Next
                Return Math.Min(1.0, sum)
            End Function

        End Class

    End Class

End Namespace
