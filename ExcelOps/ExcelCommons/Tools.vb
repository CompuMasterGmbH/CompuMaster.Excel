Option Explicit On
Option Strict On
Imports System.IO
Imports System.IO.Compression
Imports System.Xml

Namespace ExcelOps
    Public NotInheritable Class Tools

        Public Enum CellAddressCombineMode As Byte
            LeftUpperCorner = 0
            RightLowerCorner = 1
        End Enum

        ''' <summary>
        ''' Find the corner cell of 2 cells
        ''' </summary>
        ''' <param name="cell1"></param>
        ''' <param name="cell2"></param>
        ''' <param name="mode"></param>
        ''' <returns></returns>
        Public Shared Function CombineCellAddresses(cell1 As ExcelCell, cell2 As ExcelCell, mode As CellAddressCombineMode) As ExcelCell
            If cell1.SheetName <> cell2.SheetName Then Throw New ArgumentException("Cell must be member of the same sheet as cell1", NameOf(cell2))
            Select Case mode
                Case CellAddressCombineMode.LeftUpperCorner
                    Return New ExcelCell(cell1.SheetName, System.Math.Min(cell1.RowIndex, cell2.RowIndex), System.Math.Min(cell1.ColumnIndex, cell2.ColumnIndex), ExcelCell.ValueTypes.All)
                Case CellAddressCombineMode.RightLowerCorner
                    Return New ExcelCell(cell1.SheetName, System.Math.Max(cell1.RowIndex, cell2.RowIndex), System.Math.Max(cell1.ColumnIndex, cell2.ColumnIndex), ExcelCell.ValueTypes.All)
                Case Else
                    Throw New NotImplementedException
            End Select
        End Function

        ''' <summary>
        ''' Resolve range addresses (e.g. "A1:C3" or "A1") to cell addresses (e.g. "A1" or "C3")
        ''' </summary>
        ''' <param name="range">Range address (e.g. "A1:C3" or "A1")</param>
        ''' <param name="index">0 for 1st cell address, 1 for 2nd cell address</param>
        ''' <returns>Cell address (e.g. "A1" or "C3")</returns>
        Public Shared Function LookupCellAddresFromRange(range As String, index As Integer) As String
            Dim Cells As String() = range.Split(":"c)
            If Cells.Length = 0 OrElse Cells.Length > 2 Then Throw New ArgumentException("Invalid range", NameOf(range))
            Select Case index
                Case 0
                    Return Cells(0)
                Case 1
                    Return Cells(If(Cells.Length = 2, 1, 0))
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(index))
            End Select
        End Function

        Public Shared Function ReplaceWholeValue(value As String, searchValue As String, replacementValue As String) As String
            If value = searchValue Then
                Return replacementValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        ''' Formula without cell references contain just simple mathematic operations such as *, /, +, -, ^, (, )
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaWithoutCellReferences(formula As String) As Boolean
            Dim Parts As String() = formula.Split(IsFormulaWithoutCellReferences_Separators)
            For Each Part As String In Parts
                If ExcelCell.IsValidAddress(Part, True) = True Then
                    Return False
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' Formula without cell references contain just simple mathematic operations such as *, /, +, -, ^, (, )
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaWithoutCellReferencesOrCellReferenceInSameRow(formula As String, rowIndex As Integer) As Boolean
            Dim Parts As String() = formula.Split(IsFormulaWithoutCellReferences_Separators)
            For Each Part As String In Parts
                If ExcelCell.IsValidAddress(Part, True) = True Then
                    Dim Cell As New ExcelCell(Part, ExcelCell.ValueTypes.All)
                    If Cell.SheetName <> Nothing Then
                        Return False
                    ElseIf Cell.RowIndex <> rowIndex Then
                        Return False
                    End If
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' Formula is just a simple cell reference like "Grunddaten!B4"
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaSimpleCellReference(formula As String) As Boolean
            Return ExcelCell.IsValidAddress(formula, True)
        End Function

        Private Shared ReadOnly IsFormulaSimpleSumFunction_Separators As Char() = New Char() {":"c, "("c, ")"c}
        Private Shared ReadOnly IsFormulaWithoutCellReferences_Separators As Char() = New Char() {":"c, "+"c, "-"c, "*"c, "/"c, "^"c, "("c, ")"c, ","c}

        ''' <summary>
        ''' Formula is just a simple SUM function with a cell range
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaSimpleSumFunction(formula As String) As Boolean
            If formula.StartsWith("SUM(", StringComparison.Ordinal) = False Then
                Return False
            End If
            Dim Parts As String() = formula.Substring(3).Split(IsFormulaSimpleSumFunction_Separators, StringSplitOptions.RemoveEmptyEntries)
            If Parts.Length = 0 OrElse Parts.Length > 2 Then
                Return False
            End If
            For Each Part As String In Parts
                If ExcelCell.IsValidAddress(Part, True) = False Then
                    Return False
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formula"></param>
        ''' <returns></returns>
        Public Shared Function IsFormulaWithSheetReference(formula As String, checkForSheetName As String) As Boolean
            If formula = Nothing Then
                Return False
            Else
                Return formula.Contains(checkForSheetName & "!") OrElse formula.Contains("'" & checkForSheetName & "'!") OrElse formula.Contains("]" & checkForSheetName & "'!")
            End If
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formulas"></param>
        ''' <returns></returns>
        Public Shared Function FormulasWithSheetReferencesToSheet(formulas As List(Of TextTableCell), checkForSheetNameInFormula As String, dontSearchInThisSheetName As String) As List(Of TextTableCell)
            Dim Result As New List(Of TextTableCell)
            For MyFormulaCounter As Integer = 0 To formulas.Count - 1
                If dontSearchInThisSheetName = Nothing OrElse dontSearchInThisSheetName <> formulas(MyFormulaCounter).SheetName Then
                    If ExcelOps.Tools.IsFormulaWithSheetReference(formulas(MyFormulaCounter).CellContent, checkForSheetNameInFormula) = True Then
                        Result.Add(formulas(MyFormulaCounter))
                    End If
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formulas"></param>
        ''' <returns></returns>
        Public Shared Function ContainsFormulasWithSheetReferencesToSheet(formulas As List(Of TextTableCell), checkForSheetNameInFormula As String, dontSearchInThisSheetName As String) As Boolean
            For MyFormulaCounter As Integer = 0 To formulas.Count - 1
                If dontSearchInThisSheetName = Nothing OrElse dontSearchInThisSheetName <> formulas(MyFormulaCounter).SheetName Then
                    If ExcelOps.Tools.IsFormulaWithSheetReference(formulas(MyFormulaCounter).CellContent, checkForSheetNameInFormula) = True Then
                        Return True
                    End If
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formulas"></param>
        ''' <returns></returns>
        Public Shared Function FormulasWithSheetReferencesFromSheet(searchInThisSheetName As String, formulas As List(Of TextTableCell), checkForSheetNamesInFormula As String()) As List(Of TextTableCell)
            Dim Result As New List(Of TextTableCell)
            If searchInThisSheetName = Nothing Then Throw New ArgumentNullException(NameOf(searchInThisSheetName))
            For MyFormulaCounter As Integer = 0 To formulas.Count - 1
                If searchInThisSheetName = formulas(MyFormulaCounter).SheetName Then
                    For SearchForSheetNamesCounter As Integer = 0 To checkForSheetNamesInFormula.Length - 1
                        If ExcelOps.Tools.IsFormulaWithSheetReference(formulas(MyFormulaCounter).CellContent, checkForSheetNamesInFormula(SearchForSheetNamesCounter)) = True Then
                            Result.Add(formulas(MyFormulaCounter))
                        End If
                    Next
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Formula contains a reference to specified sheet
        ''' </summary>
        ''' <param name="formulas"></param>
        ''' <returns></returns>
        Public Shared Function ContainsFormulasWithSheetReferencesFromSheet(searchInThisSheetName As String, formulas As List(Of TextTableCell), checkForSheetNamesInFormula As String()) As Boolean
            If searchInThisSheetName = Nothing Then Throw New ArgumentNullException(NameOf(searchInThisSheetName))
            For MyFormulaCounter As Integer = 0 To formulas.Count - 1
                If searchInThisSheetName = formulas(MyFormulaCounter).SheetName Then
                    For SearchForSheetNamesCounter As Integer = 0 To checkForSheetNamesInFormula.Length - 1
                        If ExcelOps.Tools.IsFormulaWithSheetReference(formulas(MyFormulaCounter).CellContent, checkForSheetNamesInFormula(SearchForSheetNamesCounter)) = True Then
                            Return True
                        End If
                    Next
                End If
            Next
            Return False
        End Function

        Public Shared Function ParseToDoubleCultureSafe(s As String) As Double
            Dim Result As Double
            If s = Nothing Then
                Throw New ArgumentNullException(NameOf(s))
            ElseIf TryParseToDoubleCultureSafe(s, Result) = True Then
                Return Result
            Else
                Throw New FormatException("Not a parsable value: " & s)
            End If
        End Function

        Public Shared Function TryParseToDoubleCultureSafe(s As String, ByRef result As Double) As Boolean
            Dim sb As New System.Text.StringBuilder(s.Trim)

            'Find separators and shortcut exit with failure if invalid chars are present
            Dim SeparatorsFromRightDirection As List(Of KeyValuePair(Of Char, Integer))
            SeparatorsFromRightDirection = FindSeparatorsFromRightDirection(sb)
            If SeparatorsFromRightDirection Is Nothing Then
                'FindSeparatorsFromRightDirection failed
                Return False
            End If

            'Analyse findings
            If SeparatorsFromRightDirection.Count = 0 Then
                'Integer value -> just continue
            ElseIf SeparatorsFromRightDirection.Count = 1 Then
                'Either a decimal or a bigger integer between 1000 and 999999
                Select Case SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehind(sb, SeparatorsFromRightDirection(0))
                    Case SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator
                    'continue with next steps
                    Case SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.UnsafeDecision
                        Return False
                    Case Else
                        Throw New NotImplementedException
                End Select
            ElseIf SeparatorsFromRightDirection.Count = 2 Then
                If SeparatorsFromRightDirection(0).Key = SeparatorsFromRightDirection(1).Key Then
                    If SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehind(sb, SeparatorsFromRightDirection(0)) = SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator Then
                        'there can't be multiple decimal separators
                        Return False
                    ElseIf HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(sb, SeparatorsFromRightDirection(1)) = False Then
                        'Throw New FormatException("Rule mismatch: HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator: " & s)
                        Return False
                    Else
                        'same char twice -> both are thousands separator -> just remove all of them
                        For Each Finding In SeparatorsFromRightDirection
                            sb.Remove(Finding.Value, 1)
                        Next
                    End If
                Else
                    'different chars -> first is decimal separator, 2nd is thousands separator
                    sb.Remove(SeparatorsFromRightDirection(1).Value, 1)
                End If
            Else 'If SeparatorsFromRightDirection.Count >= 2 Then
                Dim FindingsStartIndex As Integer
                If SeparatorsFromRightDirection(0).Key = SeparatorsFromRightDirection(1).Key Then
                    'same char twice -> both are thousands separator -> just remove all of them
                    If SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehind(sb, SeparatorsFromRightDirection(0)) = SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator Then
                        'there can't be multiple decimal separators
                        Return False
                    Else
                        FindingsStartIndex = 0
                    End If
                Else
                    'different chars -> first is decimal separator, 2nd is thousands separator
                    FindingsStartIndex = 1
                End If
                For FindingsCounter As Integer = FindingsStartIndex To SeparatorsFromRightDirection.Count - 1
                    If SeparatorsFromRightDirection(FindingsCounter).Key <> SeparatorsFromRightDirection(FindingsStartIndex).Key Then
                        'Throw New FormatException("Too many decimal separators: " & s)
                        Return False
                    ElseIf HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(sb, SeparatorsFromRightDirection(FindingsCounter)) = False Then
                        'Throw New FormatException("Rule mismatch: HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator: " & s)
                        Return False
                    End If
                Next
                For FindingsCounter As Integer = FindingsStartIndex To SeparatorsFromRightDirection.Count - 1
                    sb.Remove(SeparatorsFromRightDirection(FindingsCounter).Value, 1)
                Next
            End If

            'Refresh separators list
            SeparatorsFromRightDirection = FindSeparatorsFromRightDirection(sb)
            If SeparatorsFromRightDirection Is Nothing Then
                'FindSeparatorsFromRightDirection failed
                Return False
            End If

            'Re-check again and convert decimal separator "," to "." to match InvariantCulture
            If SeparatorsFromRightDirection.Count = 0 Then
                'Ok - nothing to do, here
            ElseIf SeparatorsFromRightDirection.Count = 1 Then
                'Ok - just convert decimal separator "," to "." to match InvariantCulture (if applicable)
                If SeparatorsFromRightDirection(0).Key = ","c Then
                    sb.Chars(SeparatorsFromRightDirection(0).Value) = "."c
                End If
            Else 'If SeparatorsFromRightDirection.Count > 1 Then
                'Failure: only 1 separator allowed
                Return False
            End If

            'Try parsing safe value
            Return Double.TryParse(sb.ToString, Globalization.NumberStyles.AllowDecimalPoint Or Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, result)
        End Function

        Friend Shared Function HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(s As String, foundSeparator As KeyValuePair(Of Char, Integer)) As Boolean
            Return HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(New System.Text.StringBuilder(s), foundSeparator)
        End Function

        Friend Shared Function HasLeadingDigitsAsWellAs3FollowingDigitsBeforeEndOrNextSeparator(sb As System.Text.StringBuilder, foundSeparator As KeyValuePair(Of Char, Integer)) As Boolean
            If foundSeparator.Value = 0 Then
                'no char before separator
                Return False
            ElseIf IsDigitChar(sb.Chars(foundSeparator.Value - 1)) = False Then
                'not a digit
                Return False
            ElseIf sb.Length - 1 < foundSeparator.Value + 3 Then
                'no 3 further chars
                Return False
            ElseIf IsDigitChar(sb.Chars(foundSeparator.Value + 1)) = False Then
                'not a digit
                Return False
            ElseIf IsDigitChar(sb.Chars(foundSeparator.Value + 2)) = False Then
                'not a digit
                Return False
            ElseIf IsDigitChar(sb.Chars(foundSeparator.Value + 3)) = False Then
                'not a digit
                Return False
            ElseIf sb.Length - 1 >= foundSeparator.Value + 4 AndAlso IsSeparatorChar(sb.Chars(foundSeparator.Value + 4)) = False Then
                'not a digit
                Return False
            Else
                'all checks successful
                Return True
            End If
        End Function

        Private Shared Function IsSeparatorChar(value As Char) As Boolean
            Select Case value
                Case ","c, "."c
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsDigitChar(value As Char) As Boolean
            Select Case value
                Case "0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Enum SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult As Byte
            MustBeDecimalSeparator = 1
            UnsafeDecision = 2
        End Enum

        Private Shared Function SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehind(sb As System.Text.StringBuilder, foundSeparator As KeyValuePair(Of Char, Integer)) As SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult
            Dim NumberOfDigitsBehind As Integer = sb.Length - 1 - foundSeparator.Value
            Select Case NumberOfDigitsBehind
                Case 0, 1, 2
                    'it's a decimal separator!
                    Return SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator
                Case 3
                    'it might be a decimal separator or a thousands separator -> method must fail for unsafety reasons
                    Return SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.UnsafeDecision
                Case Else
                    '4 or more digits -> must be a decimal separator
                    Return SeparatorMustBeDecimalSeparatorBasedOnNumberOfDigitsBehindResult.MustBeDecimalSeparator
            End Select
        End Function

        Private Shared Function FindSeparatorsFromRightDirection(sb As System.Text.StringBuilder) As List(Of KeyValuePair(Of Char, Integer))
            Dim SeparatorsFromRightDirection As New List(Of KeyValuePair(Of Char, Integer))
            For MyCounter As Integer = sb.Length - 1 To 0 Step -1
                Dim CurrentChar As Char = sb.Chars(MyCounter)
                Select Case CurrentChar
                    Case "0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c, "-"c
                    'OK
                    Case ","c, "."c
                        SeparatorsFromRightDirection.Add(New KeyValuePair(Of Char, Integer)(CurrentChar, MyCounter))
                    Case " "c, "+"c
                        'Not allowed
                        Return Nothing
                    Case Else
                        'Throw New FormatException("Invalid number value for conversion to Double: " & s)
                        Return Nothing
                End Select
            Next
            Return SeparatorsFromRightDirection
        End Function

        ''' <summary>
        ''' Check if a value is member of an array of values
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="value"></param>
        ''' <param name="allowedValues"></param>
        ''' <returns></returns>
        Friend Shared Function IsOneOf(Of T)(value As T, ParamArray allowedValues As T()) As Boolean
            Dim GType As Type = GetType(T)
            If GType.IsArray Then
                Throw New NotSupportedException("Arrays as generic type not supported")
            ElseIf GType.IsInterface Then
                Throw New NotSupportedException("Interfaces as generic type not supported")
            ElseIf GType.IsClass AndAlso GType Is GetType(String) Then
                If allowedValues Is Nothing OrElse allowedValues.Length = 0 Then
                    Return False
                Else
                    Return allowedValues.Contains(value)
                End If
            ElseIf GType.IsClass Then
                If allowedValues Is Nothing OrElse allowedValues.Length = 0 Then
                    Return False
                Else
                    Return allowedValues.Contains(value)
                End If
            ElseIf GType.IsValueType Then
                If allowedValues Is Nothing OrElse allowedValues.Length = 0 Then
                    Return False
                Else
                    Return allowedValues.Contains(value)
                End If
            Else
                Throw New NotSupportedException("Unsupported generic type " & GType.FullName & "/" & GType.GetGenericTypeDefinition.FullName)
            End If
        End Function

        ''' <summary>
        ''' Convert a Double value from Excel to a DateTime value
        ''' </summary>
        ''' <param name="excelDate"></param>
        ''' <param name="baseDateValue"></param>
        ''' <returns></returns>
        Public Shared Function ConvertExcelDateToDateTime(excelDate As Double, baseDateValue As XlsxDateSystem) As DateTime
            Dim baseDate As DateTime
            Select Case baseDateValue
                Case XlsxDateSystem.Date1900
                    ' Basisdatum für Windows Excel (30. Dezember 1899)
                    baseDate = New DateTime(1899, 12, 30)
                Case XlsxDateSystem.Date1904
                    ' Basisdatum für Mac Excel (1. Januar 1904)
                    baseDate = New DateTime(1904, 1, 1)
                Case Else
                    Throw New NotSupportedException("Unsupported base date value: " & baseDateValue)
            End Select

            ' Excel-Datumswerte als Double umwandeln in DateTime
            Dim dateTimeValue As DateTime = baseDate.AddDays(excelDate)
            Return dateTimeValue
        End Function

        ''' <summary>
        ''' The base date for Excel date values
        ''' </summary>
        Public Enum XlsxDateSystem
            ''' <summary>
            ''' The 1900 date system is used
            ''' </summary>
            Date1900 = 1
            ''' <summary>
            ''' The 1904 date system is used
            ''' </summary>
            ''' <remarks>
            ''' This flag is often used in Excel for Macintosh versions 2004 and earlier or when the 1904 date system is enabled for the workbook in Excel for Windows
            ''' </remarks>
            Date1904 = 2
        End Enum

        ''' <summary>
        ''' Determines whether the 1904 date system is used in an XLSX file
        ''' </summary>
        ''' <param name="xlsxFilePath">The path to the XLSX file</param>
        ''' <returns>A value of <see cref="XlsxDateSystem">XlsxDateSystem</see></returns>
        Public Shared Function DetectXlsxDateSystem(xlsxFilePath As String) As XlsxDateSystem
            ' Überprüfen, ob die XLSX-Datei existiert
            If File.Exists(xlsxFilePath) Then
                ' XLSX-Datei als ZIP-Archiv öffnen
                Using archive As ZipArchive = ZipFile.OpenRead(xlsxFilePath)
                    ' nach der Datei "workbook.xml" im ZIP-Archiv suchen
                    Dim workbookEntry As ZipArchiveEntry = archive.GetEntry("xl/workbook.xml")
                    If workbookEntry IsNot Nothing Then
                        ' XML-Dokument laden
                        Dim xmlDoc As New XmlDocument()
                        Using reader As Stream = workbookEntry.Open()
                            xmlDoc.Load(reader)
                        End Using

                        ' Namespace-Manager für XPath-Abfrage
                        Dim nsmgr As New XmlNamespaceManager(xmlDoc.NameTable)
                        nsmgr.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

                        ' Überprüfen, ob das Datumssystem auf 1904 gesetzt ist
                        Dim node As XmlNode = xmlDoc.SelectSingleNode("//d:workbookPr", nsmgr)
                        If node IsNot Nothing AndAlso node.Attributes("date1904") IsNot Nothing Then
                            Dim Uses1904DateSystem = (node.Attributes("date1904").Value = "1")
                            If Uses1904DateSystem Then
                                Return XlsxDateSystem.Date1904
                            Else
                                Return XlsxDateSystem.Date1900
                            End If
                        End If
                    End If
                End Using
                ' Standardwert zurückgeben
                Return XlsxDateSystem.Date1900
            Else
                Throw New System.IO.FileNotFoundException("The file does not exist: " & xlsxFilePath, xlsxFilePath)
            End If
        End Function

    End Class

End Namespace