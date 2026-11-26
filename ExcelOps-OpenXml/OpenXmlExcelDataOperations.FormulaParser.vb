Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Delegate Function FormulaExpression(calculation As FormulaCalculation, formulaCell As ExcelCell) As Object

        Private NotInheritable Class FormulaParser
            Private Shared ReadOnly TokenPattern As New Regex(
                "\G\s*(""(?:[^""]|"""")*""|'(?:[^']|'')*'|#[A-Z0-9/?!._]+|(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?|\[[0-9]+\][A-Z_][A-Z0-9_.$]*|\$?[A-Z_][A-Z0-9_.$]*|<=|>=|<>|[=<>+\-*/^&%(),:!{};])",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly CellAddressPattern As New Regex("^\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}$",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly UnsupportedSheetCharacters As Char() = {"["c, "]"c, ":"c}
            Private ReadOnly _tokens As New List(Of String)
            Private ReadOnly _tokenStarts As List(Of Integer)
            Private ReadOnly _storagePrefixes As SortedDictionary(Of Integer, String)
            Private _position As Integer
            Private ReadOnly _requireAbsoluteReferences As Boolean
            Private _localNames As Dictionary(Of String, Object)

            Private Sub New(formula As String, requireAbsoluteReferences As Boolean, Optional collectStoragePrefixes As Boolean = False)
                _requireAbsoluteReferences = requireAbsoluteReferences
                If collectStoragePrefixes Then
                    _tokenStarts = New List(Of Integer)()
                    _storagePrefixes = New SortedDictionary(Of Integer, String)()
                End If
                formula = formula.Trim()
                If formula.Length > 0 AndAlso formula(0) = "="c Then formula = formula.Substring(1).Trim()
                Dim offset = 0
                While offset < formula.Length
                    Dim token = TokenPattern.Match(formula, offset)
                    If Not token.Success Then Throw New NotSupportedException("Unsupported formula syntax at character " & (offset + 1).ToString(CultureInfo.InvariantCulture) & ".")
                    _tokens.Add(token.Groups(1).Value)
                    If _storagePrefixes IsNot Nothing Then _tokenStarts.Add(token.Groups(1).Index)
                    If _tokens.Count > 512 Then Throw New NotSupportedException("This formula exceeds the supported expression size.")
                    offset += token.Length
                End While
            End Sub

            Friend Shared Function Parse(formula As String, Optional requireAbsoluteReferences As Boolean = False) As FormulaExpression
                Dim parser As New FormulaParser(formula, requireAbsoluteReferences)
                Dim expression = parser.ParseExpression(1)
                If parser.Current IsNot Nothing Then Throw New NotSupportedException("Unexpected formula token " & parser.Current & ".")
                Return expression
            End Function

            Friend Shared Function NormalizeLambdaStorage(formula As String) As String
                Try
                    Dim parser As New FormulaParser(formula, requireAbsoluteReferences:=False, collectStoragePrefixes:=True)
                    parser.ParseExpression(1)
                    If parser.Current IsNot Nothing OrElse parser._storagePrefixes.Count = 0 Then Return formula
                    Dim normalized As New StringBuilder(formula)
                    For Each prefix In parser._storagePrefixes.Reverse()
                        normalized.Insert(prefix.Key, prefix.Value)
                    Next
                    Return normalized.ToString()
                Catch ex As NotSupportedException
                    ' Deferred formulas can carry syntax outside the portable evaluator.
                    Return formula
                End Try
            End Function

            Private Sub AddStoragePrefix(index As Integer, prefix As String)
                If _storagePrefixes Is Nothing OrElse _tokens(index).StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Return
                _storagePrefixes(_tokenStarts(index)) = prefix
            End Sub

            Private ReadOnly Property Current As String
                Get
                    Return If(_position < _tokens.Count, _tokens(_position), Nothing)
                End Get
            End Property

            Private Function Consume() As String
                Dim token = Current
                If token Is Nothing Then Throw New NotSupportedException("The formula ends unexpectedly.")
                _position += 1
                Return token
            End Function

            Private Function Match(token As String) As Boolean
                If Current <> token Then Return False
                _position += 1
                Return True
            End Function

            Private Sub Expect(token As String)
                If Not Match(token) Then Throw New NotSupportedException("Expected formula token " & token & ".")
            End Sub

            Private Function ParseExpression(minimumPrecedence As Integer) As FormulaExpression
                Dim left = ParsePrimary()
                While Match("(")
                    Dim callable = left
                    Dim arguments = ParseArguments()
                    left = Function(calculation, formulaCell) calculation.InvokeFormulaLambda(
                        calculation.EvaluateExpression(callable, formulaCell), arguments, formulaCell)
                End While
                Do
                    Dim operation = Current
                    Dim precedence = OperatorPrecedence(operation)
                    If precedence < minimumPrecedence Then Exit Do
                    Consume()
                    If operation = "%" Then
                        left = UnaryExpression(operation, left)
                    Else
                        Dim right = ParseExpression(precedence + 1)
                        left = BinaryExpression(operation, left, right)
                    End If
                Loop
                Return left
            End Function

            Private Function ParsePrimary() As FormulaExpression
                Dim tokenIndex = _position
                Dim token = Consume()
                If token = "+" OrElse token = "-" Then Return UnaryExpression(token, ParseExpression(7))
                If token = "(" Then
                    Dim inner = ParseExpression(1)
                    If Not Match(",") Then
                        Expect(")")
                        Return inner
                    End If
                    Dim areas As New List(Of FormulaExpression) From {inner}
                    Do
                        areas.Add(ParseExpression(1))
                    Loop While Match(",")
                    Expect(")")
                    Return Function(calculation, formulaCell) calculation.EvaluateReferenceUnion(areas, formulaCell)
                End If
                If token = "{" Then Return ParseArrayConstant()
                If token(0) = """"c Then
                    Dim contents = ReadTextLiteral(token)
                    Return Function(calculation, formulaCell) contents
                End If
                If token(0) = "#"c Then
                    Dim excelError = ReadErrorLiteral(token)
                    Return Function(calculation, formulaCell) excelError
                End If
                If Match("!") Then
                    Dim sheetName = token
                    If token(0) = "'"c Then
                        sheetName = token.Substring(1, token.Length - 2).Replace("''", "'")
                    End If
                    Dim workbookIndex As Integer? = Nothing
                    If sheetName.StartsWith("[", StringComparison.Ordinal) Then
                        Dim close = sheetName.IndexOf("]"c)
                        Dim index As Integer
                        If close < 2 OrElse Not Integer.TryParse(sheetName.Substring(1, close - 1), NumberStyles.None, CultureInfo.InvariantCulture, index) OrElse index < 1 Then
                            Throw New NotSupportedException("The external reference must use a workbook index from the XLSX package.")
                        End If
                        workbookIndex = index
                        sheetName = sheetName.Substring(close + 1)
                    End If
                    If sheetName.Length = 0 OrElse sheetName.IndexOfAny(UnsupportedSheetCharacters) >= 0 Then
                        Throw New NotSupportedException("Unsupported worksheet name or 3-D reference.")
                    End If
                    Return ParseReferenceOrName(sheetName, Consume(), workbookIndex)
                End If
                If Match("(") Then
                    If token.Equals("LET", StringComparison.OrdinalIgnoreCase) OrElse
                        token.Equals("_xlfn.LET", StringComparison.OrdinalIgnoreCase) Then Return ParseLet()
                    If token.Equals("LAMBDA", StringComparison.OrdinalIgnoreCase) OrElse
                        token.Equals("_xlfn.LAMBDA", StringComparison.OrdinalIgnoreCase) Then Return ParseLambda()
                    Return ParseFunction(token)
                End If

                If Char.IsDigit(token(0)) OrElse token(0) = "."c Then
                    Dim value = ReadNumberLiteral(token)
                    Return Function(calculation, formulaCell) value
                End If
                If String.Equals(token, "TRUE", StringComparison.OrdinalIgnoreCase) Then Return Function(calculation, formulaCell) True
                If String.Equals(token, "FALSE", StringComparison.OrdinalIgnoreCase) Then Return Function(calculation, formulaCell) False
                Dim binding As Object = Nothing
                If _localNames IsNot Nothing AndAlso _localNames.TryGetValue(LocalName(token), binding) Then
                    AddStoragePrefix(tokenIndex, "_xlpm.")
                    Return Function(calculation, formulaCell) calculation.ReadLocalName(binding)
                End If
                Return ParseReferenceOrName(Nothing, token, Nothing)
            End Function

            Private Function ParseArrayConstant() As FormulaExpression
                Dim values As New List(Of Object)
                Dim rows = 0
                Dim columns = 0
                Do
                    Dim rowLength = 0
                    Do
                        values.Add(ReadArrayConstantElement())
                        rowLength += 1
                    Loop While Match(",")
                    If rows = 0 Then columns = rowLength
                    If rowLength <> columns Then Throw New NotSupportedException("Array constant rows must have the same number of elements.")
                    rows += 1
                Loop While Match(";")
                Expect("}")
                Dim matrix As New FormulaMatrix(rows, columns, values.ToArray())
                Return Function(calculation, formulaCell) matrix
            End Function

            Private Function ReadArrayConstantElement() As Object
                Dim token = Consume()
                If token = "+" OrElse token = "-" Then
                    Dim number = Consume()
                    If Not Char.IsDigit(number(0)) AndAlso number(0) <> "."c Then
                        Throw New NotSupportedException("A signed array constant element must be a number.")
                    End If
                    Return ReadNumberLiteral(token & number)
                End If
                If Char.IsDigit(token(0)) OrElse token(0) = "."c Then Return ReadNumberLiteral(token)
                If token(0) = """"c Then Return ReadTextLiteral(token)
                If token(0) = "#"c Then Return ReadErrorLiteral(token)
                If token.Equals("TRUE", StringComparison.OrdinalIgnoreCase) Then Return True
                If token.Equals("FALSE", StringComparison.OrdinalIgnoreCase) Then Return False
                Throw New NotSupportedException("An array constant must contain literal numbers, text, logical values or Excel errors.")
            End Function

            Private Shared Function ReadTextLiteral(token As String) As String
                Return token.Substring(1, token.Length - 2).Replace("""""", """")
            End Function

            Private Shared Function ReadNumberLiteral(token As String) As Object
                Dim number As Double
                Return If(Double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, number),
                    FormulaCalculation.FiniteNumber(number), New FormulaError("#NUM!"))
            End Function

            Private Shared Function ReadErrorLiteral(token As String) As FormulaError
                Dim code = token.ToUpperInvariant()
                Select Case code
                    Case "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"
                        Return New FormulaError(code)
                    Case Else
                        Throw New NotSupportedException("Unsupported Excel error literal " & token & ".")
                End Select
            End Function

            Private Function ParseLet() As FormulaExpression
                Dim outerNames = _localNames
                _localNames = If(outerNames Is Nothing,
                    New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                    New Dictionary(Of String, Object)(outerNames, StringComparer.OrdinalIgnoreCase))
                Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim bindings As New List(Of (Key As Object, Value As FormulaExpression))
                Try
                    Do
                        Dim nameIndex = _position
                        Dim name = LocalName(Consume())
                        If Not IsValidDefinedNameToken(name) OrElse CellAddressPattern.IsMatch(name) OrElse
                            name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) OrElse
                            name.Equals("FALSE", StringComparison.OrdinalIgnoreCase) OrElse Not names.Add(name) Then
                            Throw New NotSupportedException("Unsupported or repeated LET name " & name & ".")
                        End If
                        AddStoragePrefix(nameIndex, "_xlpm.")
                        Expect(",")
                        Dim value As FormulaExpression
                        If Current = "," Then
                            value = Function(calculation, formulaCell) Nothing
                        Else
                            value = ParseExpression(1)
                        End If
                        Expect(",")
                        Dim key As New Object()
                        bindings.Add((key, value))
                        _localNames(name) = key
                    Loop While _position + 1 < _tokens.Count AndAlso _tokens(_position + 1) = ","

                    Dim result = ParseExpression(1)
                    Expect(")")
                    Return Function(calculation, formulaCell) calculation.EvaluateLet(bindings, result, formulaCell)
                Finally
                    _localNames = outerNames
                End Try
            End Function

            Private Shared Function LocalName(token As String) As String
                Const localPrefix As String = "_xlpm."
                Return If(token.StartsWith(localPrefix, StringComparison.OrdinalIgnoreCase), token.Substring(localPrefix.Length), token)
            End Function

            Private Function ParseReferenceOrName(sheetName As String, token As String, workbookIndex As Integer?) As FormulaExpression
                If CellAddressPattern.IsMatch(token) Then Return ParseReference(sheetName, token, workbookIndex)
                If workbookIndex.HasValue Then Throw New NotSupportedException("External defined names are outside the supported calculation scope.")
                Return Function(calculation, formulaCell) calculation.EvaluateDefinedName(token, sheetName, formulaCell)
            End Function

            Private Function ParseReference(sheetName As String, first As String, workbookIndex As Integer?) As FormulaExpression
                Dim start = ParseAddress(first)
                Dim finish = If(Match(":"), ParseAddress(Consume()), start)
                Return Function(calculation, formulaCell) New FormulaArea(If(sheetName, formulaCell.SheetName), start, finish, workbookIndex)
            End Function

            Private Function ParseAddress(token As String) As ExcelCell
                If Not CellAddressPattern.IsMatch(token) Then Throw New NotSupportedException("Unsupported worksheet reference or name " & token & ".")
                If _requireAbsoluteReferences AndAlso token.Count(Function(character) character = "$"c) <> 2 Then
                    Throw New NotSupportedException("Relative references inside a defined name are outside the supported calculation scope.")
                End If
                Dim cell = New ExcelCell(Nothing, token.Replace("$", ""), ExcelCell.ValueTypes.All)
                ValidateCoordinates(cell.RowIndex, cell.ColumnIndex)
                Return cell
            End Function

            Private Function ParseFunction(name As String) As FormulaExpression
                Dim nameIndex = _position - 2
                Dim arguments = ParseArguments(allowNativeCallback:=IsLambdaArrayFunctionName(name))
                Dim binding As Object = Nothing
                If _localNames IsNot Nothing AndAlso _localNames.TryGetValue(LocalName(name), binding) AndAlso
                    Not NativeLambdaFunctions.Contains(name) Then AddStoragePrefix(nameIndex, "_xlpm.")
                Return Function(calculation, formulaCell) calculation.EvaluateFunction(name, arguments, formulaCell, binding)
            End Function

            Private Function ParseArguments(Optional allowNativeCallback As Boolean = False) As List(Of FormulaExpression)
                Dim arguments As New List(Of FormulaExpression)
                If Not Match(")") Then
                    Do
                        If Current = "," OrElse Current = ")" Then
                            arguments.Add(OmittedFormulaArgument)
                        ElseIf allowNativeCallback AndAlso _position + 1 < _tokens.Count AndAlso _tokens(_position + 1) = ")" AndAlso
                            NativeLambdaFunctions.Contains(NativeLambdaName(Current)) AndAlso
                            (_localNames Is Nothing OrElse Not _localNames.ContainsKey(LocalName(Current))) Then
                            AddStoragePrefix(_position, "_xleta.")
                            Dim callback As New FormulaLambda(NativeLambdaName(Consume()))
                            arguments.Add(Function(calculation, formulaCell) callback)
                        Else
                            arguments.Add(ParseExpression(1))
                        End If
                    Loop While Match(",")
                    Expect(")")
                End If
                Return arguments
            End Function

            Private Shared Function NativeLambdaName(token As String) As String
                Const prefix As String = "_xleta."
                Return If(token.StartsWith(prefix, StringComparison.OrdinalIgnoreCase), token.Substring(prefix.Length), token)
            End Function

            Private Function ParseLambda() As FormulaExpression
                Dim outer = _localNames
                _localNames = If(outer Is Nothing, New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                                 New Dictionary(Of String, Object)(outer, StringComparer.OrdinalIgnoreCase))
                Dim parameters As New List(Of Object)
                Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Try
                    While _position + 1 < _tokens.Count AndAlso _tokens(_position + 1) = ","
                        Dim nameIndex = _position
                        Dim name = LocalName(Consume())
                        If Not IsValidDefinedNameToken(name) OrElse CellAddressPattern.IsMatch(name) OrElse
                            name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) OrElse name.Equals("FALSE", StringComparison.OrdinalIgnoreCase) OrElse
                            Not names.Add(name) Then Throw New NotSupportedException("Unsupported or repeated LAMBDA parameter " & name & ".")
                        AddStoragePrefix(nameIndex, "_xlpm.")
                        Expect(",")
                        Dim key As New Object()
                        parameters.Add(key)
                        _localNames(name) = key
                    End While
                    Dim body = ParseExpression(1)
                    Expect(")")
                    Return Function(calculation, formulaCell) calculation.CaptureFormulaLambda(parameters, body)
                Finally
                    _localNames = outer
                End Try
            End Function

            Private Shared Function UnaryExpression(operation As String, argument As FormulaExpression) As FormulaExpression
                Return Function(calculation, formulaCell) calculation.EvaluateUnary(operation, calculation.EvaluateExpression(argument, formulaCell))
            End Function

            Private Shared Function BinaryExpression(operation As String, left As FormulaExpression, right As FormulaExpression) As FormulaExpression
                Return Function(calculation, formulaCell) calculation.EvaluateBinary(operation, calculation.EvaluateExpression(left, formulaCell), calculation.EvaluateExpression(right, formulaCell))
            End Function

            Private Shared Function OperatorPrecedence(operation As String) As Integer
                Select Case operation
                    Case "=", "<>", "<", "<=", ">", ">=" : Return 1
                    Case "&" : Return 2
                    Case "+", "-" : Return 3
                    Case "*", "/" : Return 4
                    Case "^" : Return 5
                    Case "%" : Return 6
                    Case Else : Return 0
                End Select
            End Function
        End Class

    End Class

End Namespace
