Imports System.Globalization

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Partial Private NotInheritable Class FormulaCalculation

            Friend Function EvaluateUnary(operation As String, operand As Object) As Object
                If IsMultiValue(operand) Then
                    Dim matrix = AsMatrix(operand)
                    Return New FormulaMatrix(matrix.Rows, matrix.Columns,
                        matrix.Values.Select(Function(item) EvaluateUnary(operation, item)).ToArray())
                End If
                Dim value = NumberValue(Scalar(operand))
                If TypeOf value Is FormulaError Then Return value
                Dim number = DirectCast(value, Double)
                Select Case operation
                    Case "+" : Return number
                    Case "-" : Return -number
                    Case "%" : Return number / 100.0
                    Case Else : Throw New NotSupportedException("Unsupported unary operator " & operation & ".")
                End Select
            End Function

            Friend Function EvaluateBinary(operation As String, first As Object, second As Object) As Object
                If IsMultiValue(first) OrElse IsMultiValue(second) Then Return EvaluateArrayBinary(operation, AsMatrix(first), AsMatrix(second))
                first = Scalar(first)
                second = Scalar(second)
                If TypeOf first Is FormulaError Then Return first
                If TypeOf second Is FormulaError Then Return second
                Select Case operation
                    Case "&"
                        Return TextValue(first) & TextValue(second)
                    Case "=", "<>", "<", "<=", ">", ">="
                        Return EvaluateComparison(operation, first, second)
                End Select

                Dim left = NumberValue(first)
                Dim right = NumberValue(second)
                If TypeOf left Is FormulaError Then Return left
                If TypeOf right Is FormulaError Then Return right
                Dim a = DirectCast(left, Double)
                Dim b = DirectCast(right, Double)
                Select Case operation
                    Case "+" : Return FiniteNumber(a + b)
                    Case "-" : Return FiniteNumber(a - b)
                    Case "*" : Return FiniteNumber(a * b)
                    Case "/"
                        If b = 0.0 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber(a / b)
                    Case "^"
                        If a = 0.0 AndAlso b < 0.0 Then Return New FormulaError("#DIV/0!")
                        If a = 0.0 AndAlso b = 0.0 Then Return New FormulaError("#NUM!")
                        Return FiniteNumber(Math.Pow(a, b))
                    Case Else : Throw New NotSupportedException("Unsupported binary operator " & operation & ".")
                End Select
            End Function

            Friend Function EvaluateFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell,
                                             Optional localBinding As Object = Nothing) As Object
                Select Case name.ToUpperInvariant()
                    Case "IF"
                        Return EvaluateIf(arguments, formulaCell)
                    Case "ABS"
                        If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                        Dim operand = NumberValue(Scalar(EvaluateExpression(arguments(0), formulaCell)))
                        If TypeOf operand Is FormulaError Then Return operand
                        Return Math.Abs(DirectCast(operand, Double))
                    Case "ISREF"
                        If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                        Dim reference = TryCast(EvaluateExpression(arguments(0), formulaCell), FormulaArea)
                        Return reference IsNot Nothing AndAlso ResolveReference(reference)._worksheets.ContainsKey(reference.Sheet)
                    Case "_XLFN.ANCHORARRAY"
                        Return EvaluateSpillReference(arguments, formulaCell)
                    Case "ROW", "COLUMN", "ROWS", "COLUMNS"
                        Return EvaluateReferenceFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "DAYS", "_XLFN.DAYS", "WEEKDAY", "EDATE", "EOMONTH", "WEEKNUM", "ISOWEEKNUM", "_XLFN.ISOWEEKNUM"
                        Return EvaluateCalendarFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "WORKDAY", "WORKDAY.INTL", "NETWORKDAYS", "NETWORKDAYS.INTL", "_XLFN.WORKDAY.INTL", "_XLFN.NETWORKDAYS.INTL"
                        Return EvaluateWorkdayFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "YEARFRAC"
                        Return EvaluateYearFraction(arguments, formulaCell)
                    Case "DATEDIF", "DAYS360"
                        Return EvaluateDateDifferenceFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "DATEVALUE", "TIMEVALUE"
                        Return EvaluateDateTextFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "DATE", "TIME", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "TODAY", "NOW"
                        Return EvaluateDateFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "EXP", "LN", "LOG", "LOG10", "PI", "DEGREES", "RADIANS", "SIN", "COS", "TAN",
                         "COT", "_XLFN.COT", "SEC", "_XLFN.SEC", "CSC", "_XLFN.CSC",
                         "SINH", "COSH", "TANH", "COTH", "_XLFN.COTH", "SECH", "_XLFN.SECH", "CSCH", "_XLFN.CSCH",
                         "ASIN", "ACOS", "ATAN", "ATAN2", "ACOT", "_XLFN.ACOT",
                         "ASINH", "ACOSH", "ATANH", "ACOTH", "_XLFN.ACOTH"
                        Return EvaluateTranscendental(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "FACT", "FACTDOUBLE", "COMBIN", "COMBINA", "_XLFN.COMBINA", "PERMUT", "PERMUTATIONA", "_XLFN.PERMUTATIONA",
                         "MULTINOMIAL", "GCD", "LCM", "SUMSQ"
                        Return EvaluateCombinatorialFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "BITAND", "_XLFN.BITAND", "BITOR", "_XLFN.BITOR", "BITXOR", "_XLFN.BITXOR",
                         "BITLSHIFT", "_XLFN.BITLSHIFT", "BITRSHIFT", "_XLFN.BITRSHIFT",
                         "FLOOR", "FLOOR.MATH", "_XLFN.FLOOR.MATH", "FLOOR.PRECISE", "_XLFN.FLOOR.PRECISE",
                         "CEILING", "CEILING.MATH", "_XLFN.CEILING.MATH", "CEILING.PRECISE", "_XLFN.CEILING.PRECISE",
                         "ISO.CEILING", "_XLFN.ISO.CEILING", "MROUND", "EVEN", "ODD", "QUOTIENT"
                        Return EvaluateIntegerMath(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "DELTA", "GESTEP"
                        Return EvaluateNumericThreshold(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC", "INT", "MOD", "SIGN", "SQRT", "POWER"
                        Return EvaluateNumericFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "FIND", "SEARCH", "FINDB", "SEARCHB"
                        Return EvaluateTextSearch(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "VALUE", "NUMBERVALUE", "_XLFN.NUMBERVALUE", "CHAR", "CODE", "UNICHAR", "_XLFN.UNICHAR",
                         "UNICODE", "_XLFN.UNICODE", "CLEAN", "PROPER", "ENCODEURL", "_XLFN.ENCODEURL"
                        Return EvaluateTextConversion(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "TEXTBEFORE", "_XLFN.TEXTBEFORE", "TEXTAFTER", "_XLFN.TEXTAFTER", "TEXTSPLIT", "_XLFN.TEXTSPLIT"
                        Return EvaluateTextDelimiterFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "CONCAT", "_XLFN.CONCAT", "TEXTJOIN", "_XLFN.TEXTJOIN"
                        Return EvaluateTextJoin(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "REGEXTEST", "_XLFN.REGEXTEST", "REGEXEXTRACT", "_XLFN.REGEXEXTRACT", "REGEXREPLACE", "_XLFN.REGEXREPLACE"
                        Return EvaluateRegexFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "EXACT", "REPT", "REPLACE", "SUBSTITUTE"
                        Return EvaluateTextTransform(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "TRIM", "UPPER", "LOWER", "LEN", "LEFT", "RIGHT", "MID", "CONCATENATE"
                        Return EvaluateTextFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "MODE", "MODE.SNGL", "_XLFN.MODE.SNGL", "MODE.MULT", "_XLFN.MODE.MULT", "AVEDEV", "DEVSQ",
                         "VAR", "VAR.S", "_XLFN.VAR.S", "VARP", "VAR.P", "_XLFN.VAR.P", "VARA", "VARPA",
                         "STDEV", "STDEV.S", "_XLFN.STDEV.S", "STDEVP", "STDEV.P", "_XLFN.STDEV.P", "STDEVA", "STDEVPA",
                         "AVERAGEA", "MINA", "MAXA"
                        Return EvaluateSummaryStatistic(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "MEDIAN", "SMALL", "LARGE", "RANK", "RANK.EQ", "_XLFN.RANK.EQ", "RANK.AVG", "_XLFN.RANK.AVG",
                         "PERCENTILE", "PERCENTILE.INC", "_XLFN.PERCENTILE.INC", "PERCENTILE.EXC", "_XLFN.PERCENTILE.EXC",
                         "QUARTILE", "QUARTILE.INC", "_XLFN.QUARTILE.INC", "QUARTILE.EXC", "_XLFN.QUARTILE.EXC"
                        Return EvaluateOrderStatistic(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "SUBTOTAL", "AGGREGATE", "_XLFN.AGGREGATE"
                        Return EvaluateRowAggregate(name.Equals("SUBTOTAL", StringComparison.OrdinalIgnoreCase), arguments, formulaCell)
                    Case "RAND", "RANDBETWEEN", "RANDARRAY", "_XLFN.RANDARRAY"
                        Return EvaluateRandomFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "DOLLAR", "FIXED", "TEXT", "ARRAYTOTEXT", "_XLFN.ARRAYTOTEXT", "VALUETOTEXT", "_XLFN.VALUETOTEXT"
                        Return EvaluateDisplayFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "HYPERLINK", "FILTERXML", "_XLFN.FILTERXML"
                        Return EvaluateLinkOrXml(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "ASC", "DBCS", "_XLFN.DBCS", "JIS", "PHONETIC"
                        Return EvaluateWidthFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "LENB", "LEFTB", "RIGHTB", "MIDB", "REPLACEB"
                        Return EvaluateByteFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "BESSELI", "BESSELJ", "BESSELK", "BESSELY", "CONVERT"
                        Return EvaluateEngineeringFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "BAHTTEXT"
                        Return EvaluateBahtText(arguments, formulaCell)
                    Case "EUROCONVERT", "_XLFN.EUROCONVERT"
                        Return EvaluateEuroConvert(arguments, formulaCell)
                    Case "LINEST", "LOGEST", "TREND", "GROWTH"
                        Return EvaluateRegression(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "FREQUENCY"
                        Return EvaluateFrequency(arguments, formulaCell)
                    Case "PROB"
                        Return EvaluateDiscreteProbability(arguments, formulaCell)
                    Case "PERCENTRANK", "PERCENTRANK.INC", "_XLFN.PERCENTRANK.INC", "PERCENTRANK.EXC", "_XLFN.PERCENTRANK.EXC"
                        Return EvaluatePercentRank(name.EndsWith(".EXC", StringComparison.OrdinalIgnoreCase), arguments, formulaCell)
                    Case "BYROW", "_XLFN.BYROW", "BYCOL", "_XLFN.BYCOL", "MAP", "_XLFN.MAP",
                         "MAKEARRAY", "_XLFN.MAKEARRAY", "REDUCE", "_XLFN.REDUCE", "SCAN", "_XLFN.SCAN"
                        Return EvaluateLambdaArray(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "ISOMITTED", "_XLFN.ISOMITTED"
                        If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")
                        Return EvaluateExpression(arguments(0), formulaCell) Is MissingLambdaArgument
                    Case "ISERR", "ISNONTEXT", "ISEVEN", "ISODD", "N", "T", "TYPE", "ERROR.TYPE"
                        Return EvaluateInformationFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "IFS", "_XLFN.IFS", "SWITCH", "_XLFN.SWITCH"
                        Return EvaluateConditionalChoice(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "TRUE", "FALSE", "IFERROR", "IFNA", "_XLFN.IFNA", "NOT", "AND", "OR", "XOR", "_XLFN.XOR",
                         "ISBLANK", "ISNUMBER", "ISTEXT", "ISLOGICAL", "ISERROR", "ISNA", "NA"
                        Return EvaluateLogicalFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS", "AVERAGEIF", "AVERAGEIFS", "MINIFS", "MAXIFS", "_XLFN.MINIFS", "_XLFN.MAXIFS"
                        Return EvaluateCriteriaFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP"
                        Return EvaluateDatabaseFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "SUMPRODUCT"
                        Return EvaluateSumProduct(arguments, formulaCell)
                    Case "SERIESSUM"
                        Return EvaluateSeriesSum(arguments, formulaCell)
                    Case "SUMXMY2", "SUMX2MY2", "SUMX2PY2"
                        Return EvaluatePairwiseSquares(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "MDETERM", "MINVERSE", "MMULT", "MUNIT", "_XLFN.MUNIT"
                        Return EvaluateMatrixAlgebra(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "BASE", "_XLFN.BASE", "DECIMAL", "_XLFN.DECIMAL",
                         "BIN2DEC", "BIN2HEX", "BIN2OCT", "OCT2BIN", "OCT2DEC", "OCT2HEX",
                         "HEX2BIN", "HEX2DEC", "HEX2OCT", "DEC2BIN", "DEC2OCT", "DEC2HEX"
                        Return EvaluateRadixConversion(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "ROMAN", "ARABIC", "_XLFN.ARABIC"
                        Return EvaluateRomanNumerals(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "COMPLEX", "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMDIV", "IMREAL", "IMSUB", "IMSUM", "IMPRODUCT",
                         "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMSQRT", "IMSIN", "IMCOS",
                         "IMSINH", "_XLFN.IMSINH", "IMCOSH", "_XLFN.IMCOSH", "IMTAN", "_XLFN.IMTAN", "IMCOT", "_XLFN.IMCOT",
                         "IMSEC", "_XLFN.IMSEC", "IMCSC", "_XLFN.IMCSC", "IMSECH", "_XLFN.IMSECH", "IMCSCH", "_XLFN.IMCSCH"
                        Return EvaluateComplexFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "FV", "PV", "PMT", "NPER", "ISPMT", "IPMT", "PPMT", "CUMIPMT", "CUMPRINC", "RATE",
                         "RRI", "_XLFN.RRI", "PDURATION", "_XLFN.PDURATION"
                        Return EvaluateLoanFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "FVSCHEDULE"
                        Return EvaluateGrowthSchedule(arguments, formulaCell)
                    Case "SLN", "SYD", "DB", "DDB", "EFFECT", "NOMINAL", "DOLLARDE", "DOLLARFR"
                        Return EvaluateDepreciationOrRate(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "ERF", "ERFC", "ERF.PRECISE", "_XLFN.ERF.PRECISE", "ERFC.PRECISE", "_XLFN.ERFC.PRECISE",
                         "PHI", "_XLFN.PHI", "GAUSS", "_XLFN.GAUSS", "NORMSDIST", "NORM.S.DIST", "_XLFN.NORM.S.DIST",
                         "NORM.DIST", "_XLFN.NORM.DIST", "NORMDIST", "LOGNORM.DIST", "_XLFN.LOGNORM.DIST", "LOGNORMDIST", "STANDARDIZE",
                         "NORM.S.INV", "_XLFN.NORM.S.INV", "NORMSINV", "NORM.INV", "_XLFN.NORM.INV", "NORMINV",
                         "LOGNORM.INV", "_XLFN.LOGNORM.INV", "LOGINV", "CONFIDENCE", "CONFIDENCE.NORM", "_XLFN.CONFIDENCE.NORM"
                        Return EvaluateNormalFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "GAMMA", "_XLFN.GAMMA", "GAMMALN", "GAMMALN.PRECISE", "_XLFN.GAMMALN.PRECISE",
                         "GAMMA.DIST", "_XLFN.GAMMA.DIST", "GAMMADIST", "GAMMA.INV", "_XLFN.GAMMA.INV", "GAMMAINV",
                         "CHISQ.DIST", "_XLFN.CHISQ.DIST", "CHISQ.DIST.RT", "_XLFN.CHISQ.DIST.RT", "CHIDIST",
                         "CHISQ.INV", "_XLFN.CHISQ.INV", "CHISQ.INV.RT", "_XLFN.CHISQ.INV.RT", "CHIINV",
                         "EXPON.DIST", "_XLFN.EXPON.DIST", "EXPONDIST", "WEIBULL.DIST", "_XLFN.WEIBULL.DIST", "WEIBULL"
                        Return EvaluateGammaFunction(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "BETA.DIST", "_XLFN.BETA.DIST", "BETADIST", "BETA.INV", "_XLFN.BETA.INV", "BETAINV",
                         "T.DIST", "_XLFN.T.DIST", "T.DIST.RT", "_XLFN.T.DIST.RT", "T.DIST.2T", "_XLFN.T.DIST.2T", "TDIST",
                         "T.INV", "_XLFN.T.INV", "T.INV.2T", "_XLFN.T.INV.2T", "TINV", "CONFIDENCE.T", "_XLFN.CONFIDENCE.T",
                         "F.DIST", "_XLFN.F.DIST", "F.DIST.RT", "_XLFN.F.DIST.RT", "FDIST",
                         "F.INV", "_XLFN.F.INV", "F.INV.RT", "_XLFN.F.INV.RT", "FINV"
                        Return EvaluateBetaFamily(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "BINOM.DIST", "_XLFN.BINOM.DIST", "BINOMDIST", "BINOM.INV", "_XLFN.BINOM.INV", "CRITBINOM",
                         "BINOM.DIST.RANGE", "_XLFN.BINOM.DIST.RANGE", "NEGBINOM.DIST", "_XLFN.NEGBINOM.DIST", "NEGBINOMDIST",
                         "HYPGEOM.DIST", "_XLFN.HYPGEOM.DIST", "HYPGEOMDIST", "POISSON.DIST", "_XLFN.POISSON.DIST", "POISSON"
                        Return EvaluateDiscreteDistribution(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "ADDRESS", "INDIRECT", "OFFSET", "CHOOSE", "AREAS",
                         "SHEET", "_XLFN.SHEET", "SHEETS", "_XLFN.SHEETS",
                         "ISFORMULA", "_XLFN.ISFORMULA", "FORMULATEXT", "_XLFN.FORMULATEXT"
                        Return EvaluateReferenceUtility(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "EXPAND", "_XLFN.EXPAND", "TOCOL", "_XLFN.TOCOL", "TOROW", "_XLFN.TOROW",
                         "WRAPCOLS", "_XLFN.WRAPCOLS", "WRAPROWS", "_XLFN.WRAPROWS", "TRIMRANGE", "_XLFN.TRIMRANGE"
                        Return EvaluateReshape(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "PERCENTOF", "_XLFN.PERCENTOF"
                        Return EvaluatePercentOf(arguments, formulaCell)
                    Case "LOOKUP"
                        Return EvaluateClassicLookup(arguments, formulaCell)
                    Case "CORREL", "PEARSON", "COVAR", "COVARIANCE.P", "_XLFN.COVARIANCE.P", "COVARIANCE.S", "_XLFN.COVARIANCE.S",
                         "SLOPE", "INTERCEPT", "RSQ", "STEYX", "FORECAST", "FORECAST.LINEAR", "_XLFN.FORECAST.LINEAR"
                        Return EvaluatePairedStatistic(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "GEOMEAN", "HARMEAN", "SKEW", "SKEW.P", "_XLFN.SKEW.P", "KURT", "TRIMMEAN"
                        Return EvaluateShapeStatistic(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "FISHER", "FISHERINV", "SQRTPI"
                        Return EvaluateFisherTransform(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "CHISQ.TEST", "_XLFN.CHISQ.TEST", "CHITEST", "F.TEST", "_XLFN.F.TEST", "FTEST",
                         "T.TEST", "_XLFN.T.TEST", "TTEST", "Z.TEST", "_XLFN.Z.TEST", "ZTEST"
                        Return EvaluateHypothesisTest(name.ToUpperInvariant().Replace("_XLFN.", ""), arguments, formulaCell)
                    Case "DISC", "INTRATE", "RECEIVED", "PRICEDISC", "YIELDDISC", "TBILLEQ", "TBILLPRICE", "TBILLYIELD"
                        Return EvaluateDiscountedSecurity(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "NPV", "IRR", "MIRR"
                        Return EvaluateCashFlowFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "XNPV", "XIRR"
                        Return EvaluateDatedCashFlow(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "MATCH", "INDEX", "VLOOKUP", "HLOOKUP", "XLOOKUP", "_XLFN.XLOOKUP", "XMATCH", "_XLFN.XMATCH"
                        Return EvaluateLookupFunction(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "SORT", "_XLFN.SORT", "_XLFN._XLWS.SORT", "SORTBY", "_XLFN.SORTBY", "UNIQUE", "_XLFN.UNIQUE"
                        Return EvaluateSortFunction(name.ToUpperInvariant().Replace("_XLFN.", "").Replace("_XLWS.", ""), arguments, formulaCell)
                    Case "SEQUENCE", "_XLFN.SEQUENCE", "TRANSPOSE", "FILTER", "_XLFN.FILTER", "_XLFN._XLWS.FILTER",
                         "TAKE", "_XLFN.TAKE", "DROP", "_XLFN.DROP", "CHOOSECOLS", "_XLFN.CHOOSECOLS",
                         "CHOOSEROWS", "_XLFN.CHOOSEROWS", "HSTACK", "_XLFN.HSTACK", "VSTACK", "_XLFN.VSTACK"
                        Return EvaluateArrayFunction(name.ToUpperInvariant().Replace("_XLFN.", "").Replace("_XLWS.", ""), arguments, formulaCell)
                    Case "COUNTA", "COUNTBLANK"
                        Return CountValues(name.ToUpperInvariant(), arguments, formulaCell)
                    Case "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "PRODUCT"
                        Return Aggregate(name.ToUpperInvariant(), arguments, formulaCell)
                    Case Else
                        If localBinding IsNot Nothing Then Return InvokeFormulaLambda(ReadLocalName(localBinding), arguments, formulaCell)
                        Dim definedResult As Object = Nothing
                        If TryInvokeDefinedLambda(name, arguments, formulaCell, definedResult) Then Return definedResult
                        Return EvaluateRegisteredFunction(name, arguments, formulaCell)
                End Select
            End Function

            Private Function EvaluateReferenceFunction(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count = 0 Then
                    If name = "ROW" Then Return CDbl(formulaCell.RowIndex + 1)
                    If name = "COLUMN" Then Return CDbl(formulaCell.ColumnIndex + 1)
                End If
                If arguments.Count <> 1 Then Return New FormulaError("#VALUE!")

                ' Keep the reference as an area so address queries do not read cell values.
                Dim argument = EvaluateExpression(arguments(0), formulaCell)
                If TypeOf argument Is FormulaError Then Return argument
                Dim area = TryCast(argument, FormulaArea)
                If area Is Nothing Then
                    Dim matrix = TryCast(argument, FormulaMatrix)
                    If name = "ROWS" Then Return If(matrix Is Nothing, 1.0, CDbl(matrix.Rows))
                    If name = "COLUMNS" Then Return If(matrix Is Nothing, 1.0, CDbl(matrix.Columns))
                    Return New FormulaError("#VALUE!")
                End If
                If Not ResolveReference(area)._worksheets.ContainsKey(area.Sheet) Then Return New FormulaError("#REF!")

                Select Case name
                    Case "ROW" : Return CDbl(area.FirstRow + 1)
                    Case "COLUMN" : Return CDbl(area.FirstColumn + 1)
                    Case "ROWS" : Return CDbl(area.LastRow - area.FirstRow + 1)
                    Case Else : Return CDbl(area.LastColumn - area.FirstColumn + 1)
                End Select
            End Function

            Private Function Aggregate(name As String, arguments As IReadOnlyList(Of FormulaExpression), formulaCell As ExcelCell) As Object
                If arguments.Count = 0 Then Return New FormulaError("#VALUE!")
                Dim sum As Double
                Dim product = 1.0
                Dim minimum As Double = Double.PositiveInfinity
                Dim maximum As Double = Double.NegativeInfinity
                Dim count As Integer
                Dim firstError As FormulaError = Nothing
                For Each argument In arguments
                    Dim value = EvaluateExpression(argument, formulaCell)
                    Dim collection = TypeOf value Is FormulaArea OrElse TypeOf value Is FormulaMatrix OrElse TypeOf value Is FormulaReferenceUnion
                    For Each item In ReferenceArgumentValues(value)
                        If collection AndAlso
                            Not TypeOf item Is Double AndAlso Not TypeOf item Is FormulaError Then Continue For
                        If item Is Nothing Then Continue For
                        Dim number = NumberValue(item)
                        If TypeOf number Is FormulaError Then
                            If name <> "COUNT" AndAlso firstError Is Nothing Then firstError = DirectCast(number, FormulaError)
                            Continue For
                        End If
                        Dim numeric = DirectCast(number, Double)
                        count += 1
                        sum += numeric
                        product *= numeric
                        minimum = Math.Min(minimum, numeric)
                        maximum = Math.Max(maximum, numeric)
                    Next
                Next

                If firstError IsNot Nothing Then Return firstError
                Select Case name
                    Case "SUM" : Return FiniteNumber(sum)
                    Case "PRODUCT" : Return If(count = 0, 0.0, FiniteNumber(product))
                    Case "AVERAGE"
                        If count = 0 Then Return New FormulaError("#DIV/0!")
                        Return FiniteNumber(sum / count)
                    Case "MIN" : Return If(count = 0, 0.0, minimum)
                    Case "MAX" : Return If(count = 0, 0.0, maximum)
                    Case Else : Return CDbl(count)
                End Select
            End Function

            Private Shared Function NormalizeWorksheetUnderflow(number As Double) As Double
                Return If(Math.Abs(number) < SmallestNormalDouble, 0.0, number)
            End Function

            Friend Shared Function FiniteNumber(number As Double) As Object
                If Double.IsInfinity(number) OrElse Double.IsNaN(number) Then Return New FormulaError("#NUM!")
                Return number
            End Function

            Private ReadOnly Property ConversionCulture As CultureInfo
                Get
                    Return If(_owner.LoadOptions.FormulaCulture, CultureInfo.InvariantCulture)
                End Get
            End Property

            Private Function NumberValue(value As Object) As Object
                If value Is Nothing OrElse value Is MissingLambdaArgument Then Return 0.0
                If TypeOf value Is FormulaError OrElse TypeOf value Is Double Then Return value
                If TypeOf value Is Boolean Then Return If(DirectCast(value, Boolean), 1.0, 0.0)
                If TypeOf value Is String Then
                    Dim parsed As Double
                    If Double.TryParse(DirectCast(value, String), NumberStyles.Float, ConversionCulture, parsed) Then Return FiniteNumber(parsed)
                    Return New FormulaError("#VALUE!")
                End If
                Throw New NotSupportedException("The referenced value type is outside the supported calculation scope.")
            End Function

            Private Function TextValue(value As Object) As String
                If value Is Nothing OrElse value Is MissingLambdaArgument Then Return String.Empty
                If TypeOf value Is Boolean Then
                    If ConversionCulture.TwoLetterISOLanguageName = "de" Then Return If(DirectCast(value, Boolean), "WAHR", "FALSCH")
                    Return If(DirectCast(value, Boolean), "TRUE", "FALSE")
                End If
                If TypeOf value Is Double Then Return DirectCast(value, Double).ToString("G15", ConversionCulture)
                Return Convert.ToString(value, ConversionCulture)
            End Function

            Private Shared Function EvaluateComparison(operation As String, first As Object, second As Object) As Boolean
                If first Is Nothing Then first = ComparableBlank(second)
                If second Is Nothing Then second = ComparableBlank(first)
                Dim firstType = ComparisonRank(first)
                Dim secondType = ComparisonRank(second)
                Dim comparison As Integer
                If firstType <> secondType Then
                    comparison = firstType.CompareTo(secondType)
                ElseIf firstType = 0 Then
                    comparison = DirectCast(first, Double).CompareTo(DirectCast(second, Double))
                ElseIf firstType = 1 Then
                    comparison = String.Compare(DirectCast(first, String), DirectCast(second, String), StringComparison.OrdinalIgnoreCase)
                Else
                    comparison = DirectCast(first, Boolean).CompareTo(DirectCast(second, Boolean))
                End If
                Select Case operation
                    Case "=" : Return comparison = 0
                    Case "<>" : Return comparison <> 0
                    Case "<" : Return comparison < 0
                    Case "<=" : Return comparison <= 0
                    Case ">" : Return comparison > 0
                    Case Else : Return comparison >= 0
                End Select
            End Function

            Private Shared Function ComparableBlank(other As Object) As Object
                If TypeOf other Is String Then Return String.Empty
                If TypeOf other Is Boolean Then Return False
                Return 0.0
            End Function

            Private Shared Function ComparisonRank(value As Object) As Integer
                If TypeOf value Is Double Then Return 0
                If TypeOf value Is String Then Return 1
                If TypeOf value Is Boolean Then Return 2
                Throw New NotSupportedException("This value type is not supported in formula comparisons.")
            End Function

        End Class

    End Class

End Namespace
