Imports System.Globalization
Imports System.Text.RegularExpressions

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Delegate Function FormulaExpression(calculation As FormulaCalculation, worksheet As String) As Object

        Private NotInheritable Class FormulaParser
            Private Shared ReadOnly TokenPattern As New Regex(
                "\G\s*(""(?:[^""]|"""")*""|'(?:[^']|'')*'|#[A-Z0-9/?!._]+|(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?|\$?[A-Z_][A-Z0-9_.$]*|<=|>=|<>|[=<>+\-*/^&%(),:!])",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly CellAddressPattern As New Regex("^\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}$",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            Private Shared ReadOnly UnsupportedSheetCharacters As Char() = {"["c, "]"c, ":"c}
            Private ReadOnly _tokens As New List(Of String)
            Private _position As Integer

            Private Sub New(formula As String)
                formula = formula.Trim()
                If formula.Length > 0 AndAlso formula(0) = "="c Then formula = formula.Substring(1).Trim()
                Dim offset = 0
                While offset < formula.Length
                    Dim token = TokenPattern.Match(formula, offset)
                    If Not token.Success Then Throw New NotSupportedException("Unsupported formula syntax at character " & (offset + 1).ToString(CultureInfo.InvariantCulture) & ".")
                    _tokens.Add(token.Groups(1).Value)
                    If _tokens.Count > 512 Then Throw New NotSupportedException("This formula exceeds the supported expression size.")
                    offset += token.Length
                End While
            End Sub

            Friend Shared Function Parse(formula As String) As FormulaExpression
                Dim parser As New FormulaParser(formula)
                Dim expression = parser.ParseExpression(1)
                If parser.Current IsNot Nothing Then Throw New NotSupportedException("Unexpected formula token " & parser.Current & ".")
                Return expression
            End Function

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
                Dim token = Consume()
                If token = "+" OrElse token = "-" Then Return UnaryExpression(token, ParseExpression(7))
                If token = "(" Then
                    Dim inner = ParseExpression(1)
                    Expect(")")
                    Return inner
                End If
                If token(0) = """"c Then
                    Dim contents = token.Substring(1, token.Length - 2).Replace("""""", """")
                    Return Function(calculation, worksheet) contents
                End If
                If token(0) = "#"c Then
                    Dim code = token.ToUpperInvariant()
                    Select Case code
                        Case "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"
                            Dim excelError = New FormulaError(code)
                            Return Function(calculation, worksheet) excelError
                        Case Else
                            Throw New NotSupportedException("Unsupported Excel error literal " & token & ".")
                    End Select
                End If
                If Match("!") Then
                    Dim sheetName = token
                    If token(0) = "'"c Then
                        sheetName = token.Substring(1, token.Length - 2).Replace("''", "'")
                    End If
                    If sheetName.IndexOfAny(UnsupportedSheetCharacters) >= 0 Then Throw New NotSupportedException("External workbook and 3-D references are outside the supported calculation scope.")
                    Return ParseReference(sheetName, Consume())
                End If
                If Match("(") Then Return ParseFunction(token)

                If Char.IsDigit(token(0)) OrElse token(0) = "."c Then
                    Dim number As Double
                    Dim value = If(Double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, number),
                        FormulaCalculation.FiniteNumber(number), New FormulaError("#NUM!"))
                    Return Function(calculation, worksheet) value
                End If
                If String.Equals(token, "TRUE", StringComparison.OrdinalIgnoreCase) Then Return Function(calculation, worksheet) True
                If String.Equals(token, "FALSE", StringComparison.OrdinalIgnoreCase) Then Return Function(calculation, worksheet) False
                Return ParseReference(Nothing, token)
            End Function

            Private Function ParseReference(sheetName As String, first As String) As FormulaExpression
                Dim start = ParseAddress(first)
                Dim finish = If(Match(":"), ParseAddress(Consume()), start)
                Return Function(calculation, worksheet) New FormulaArea(If(sheetName, worksheet), start, finish)
            End Function

            Private Shared Function ParseAddress(token As String) As ExcelCell
                If Not CellAddressPattern.IsMatch(token) Then Throw New NotSupportedException("Unsupported worksheet reference or name " & token & ".")
                Dim cell = New ExcelCell(Nothing, token.Replace("$", ""), ExcelCell.ValueTypes.All)
                ValidateCoordinates(cell.RowIndex, cell.ColumnIndex)
                Return cell
            End Function

            Private Function ParseFunction(name As String) As FormulaExpression
                Dim arguments As New List(Of FormulaExpression)
                If Not Match(")") Then
                    Do
                        If Current = "," OrElse Current = ")" Then
                            arguments.Add(Function(calculation, worksheet) Nothing)
                        Else
                            arguments.Add(ParseExpression(1))
                        End If
                    Loop While Match(",")
                    Expect(")")
                End If
                Return Function(calculation, worksheet) calculation.EvaluateFunction(name, arguments, worksheet)
            End Function

            Private Shared Function UnaryExpression(operation As String, argument As FormulaExpression) As FormulaExpression
                Return Function(calculation, worksheet) calculation.EvaluateUnary(operation, calculation.EvaluateExpression(argument, worksheet))
            End Function

            Private Shared Function BinaryExpression(operation As String, left As FormulaExpression, right As FormulaExpression) As FormulaExpression
                Return Function(calculation, worksheet) calculation.EvaluateBinary(operation, calculation.EvaluateExpression(left, worksheet), calculation.EvaluateExpression(right, worksheet))
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
