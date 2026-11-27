Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DynamicArray = DocumentFormat.OpenXml.Office2019.Excel.DynamicArray

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private _authoredDynamicMetadata As Metadata

        Private Shared ReadOnly DynamicArrayFormulaTokens As New Regex(
            """(?:[^""]|"""")*""|(?<![\p{L}\p{M}\p{Nd}_.$!\]])(?<anchor>(?:(?:'(?:[^']|'')+'|(?:\[[1-9][0-9]*\])?[\p{L}\p{M}\p{Nd}_.]+)!)?\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6})#(?![\p{L}\p{M}\p{Nd}_.?])|" &
            "'(?:[^']|'')*'|(?<![\p{L}\p{M}\p{Nd}_.])(?<name>SEQUENCE|RANDARRAY|ARRAYTOTEXT|VALUETOTEXT|FILTERXML|DBCS|FILTER|SORT|SORTBY|UNIQUE|TAKE|DROP|CHOOSECOLS|CHOOSEROWS|HSTACK|VSTACK|XLOOKUP|XMATCH|CONCAT|TEXTJOIN|LET|LAMBDA|ISOMITTED|BYROW|BYCOL|MAP|MAKEARRAY|REDUCE|SCAN|MINIFS|MAXIFS|IFNA|REGEXTEST|REGEXEXTRACT|REGEXREPLACE|DAYS|ISOWEEKNUM|IFS|SWITCH|XOR|RANK\.EQ|RANK\.AVG|PERCENTILE\.INC|PERCENTILE\.EXC|QUARTILE\.INC|QUARTILE\.EXC|MODE\.SNGL|MODE\.MULT|VAR\.S|VAR\.P|STDEV\.S|STDEV\.P|COT|SEC|CSC|COTH|SECH|CSCH|ACOT|ACOTH|BITAND|BITOR|BITXOR|BITLSHIFT|BITRSHIFT|FLOOR\.MATH|FLOOR\.PRECISE|CEILING\.MATH|CEILING\.PRECISE|ISO\.CEILING|COMBINA|PERMUTATIONA|TEXTBEFORE|TEXTAFTER|TEXTSPLIT|NUMBERVALUE|UNICHAR|UNICODE|MUNIT|BASE|DECIMAL|ARABIC|IMSINH|IMCOSH|IMTAN|IMCOT|IMSEC|IMCSC|IMSECH|IMCSCH|RRI|PDURATION|ERF\.PRECISE|ERFC\.PRECISE|PHI|GAUSS|NORM\.S\.DIST|NORM\.DIST|LOGNORM\.DIST|NORM\.S\.INV|NORM\.INV|LOGNORM\.INV|CONFIDENCE\.NORM|GAMMA|GAMMA\.DIST|GAMMA\.INV|GAMMALN\.PRECISE|CHISQ\.DIST|CHISQ\.DIST\.RT|CHISQ\.INV|CHISQ\.INV\.RT|EXPON\.DIST|WEIBULL\.DIST|BETA\.DIST|BETA\.INV|T\.DIST|T\.DIST\.RT|T\.DIST\.2T|T\.INV|T\.INV\.2T|CONFIDENCE\.T|F\.DIST|F\.DIST\.RT|F\.INV|F\.INV\.RT|BINOM\.DIST|BINOM\.DIST\.RANGE|BINOM\.INV|HYPGEOM\.DIST|NEGBINOM\.DIST|POISSON\.DIST|SHEET|SHEETS|ISFORMULA|FORMULATEXT|EXPAND|TOCOL|TOROW|WRAPCOLS|WRAPROWS|TRIMRANGE|GROUPBY|PIVOTBY|PERCENTOF|AGGREGATE|COVARIANCE\.P|COVARIANCE\.S|FORECAST\.LINEAR|SKEW\.P|CHISQ\.TEST|F\.TEST|T\.TEST|Z\.TEST)(?=\s*\()",
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

        ''' <summary>Writes a formula that can return a dynamic array to the worksheet.</summary>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="rowIndex">The zero-based row index of the anchor.</param>
        ''' <param name="columnIndex">The zero-based column index of the anchor.</param>
        ''' <param name="formula">The invariant Excel formula with an optional leading equals sign.</param>
        ''' <param name="immediatelyCalculateCellValue">Whether to calculate and store the output immediately.</param>
        ''' <remarks>Replaces an existing dynamic array when the address is its anchor. Excel calculation errors are stored as results. A deferred replacement clears the previous output.</remarks>
        ''' <exception cref="ArgumentException">The formula is empty.</exception>
        ''' <exception cref="NotSupportedException">The address is covered by another grouped formula or unrelated cell metadata.</exception>
        ''' <exception cref="FeatureDisabledException">Immediate calculation was requested while the calculation engine is disabled.</exception>
        ''' <exception cref="FormulaComplexityNotSupportedException">Immediate calculation cannot evaluate the formula.</exception>
        Public Sub WriteDynamicArrayFormula(sheetName As String, rowIndex As Integer, columnIndex As Integer, formula As String,
                                            immediatelyCalculateCellValue As Boolean)
            SyncLock WorkbookSaveLock
                ValidateCoordinates(rowIndex, columnIndex)
                Dim expression = If(formula, String.Empty).Trim()
                If expression.StartsWith("=", StringComparison.Ordinal) Then expression = expression.Substring(1).Trim()
                If expression.Length = 0 Then Throw New ArgumentException("A formula is required.", NameOf(formula))
                If immediatelyCalculateCellValue Then EnsureCalculationEnabled()

                Dim part = GetWorksheetPart(sheetName)
                Dim existing = GetCell(sheetName, rowIndex, columnIndex)
                Dim metadata = FormulaCellMetadata
                Dim existingSpill = If(IsDynamicFormulaAnchor(metadata, existing), New SpillArea(existing, Nothing), Nothing)
                If existingSpill Is Nothing AndAlso (IsGroupedFormulaCell(part, sheetName, rowIndex, columnIndex, existing?.CellFormula) OrElse
                    existing?.CellMetaIndex IsNot Nothing) Then
                    Throw UnsupportedOperation("Replacing a grouped formula or cell metadata with a dynamic array formula")
                End If

                Dim metadataIndex = If(existingSpill Is Nothing, FindReusableDynamicMetadata(metadata), existing.CellMetaIndex.Value)
                Dim newMetadata As Metadata = Nothing
                If metadataIndex = 0UI Then
                    newMetadata = If(metadata Is Nothing, New Metadata(), DirectCast(metadata.CloneNode(True), Metadata))
                    metadataIndex = AppendDynamicMetadata(newMetadata)
                End If

                Dim address = ExcelCell.LocalCellAddress(rowIndex, columnIndex)
                Dim replacement As New Cell With {
                    .CellReference = address,
                    .CellMetaIndex = metadataIndex,
                    .CellFormula = New CellFormula(NormalizeDynamicArrayFormula(expression)) With {
                        .FormulaType = CellFormulaValues.Array,
                        .Reference = If(immediatelyCalculateCellValue AndAlso existingSpill IsNot Nothing, existing.CellFormula.Reference.Value, address)
                    }
                }
                Dim calculation As FormulaCalculation = Nothing
                If immediatelyCalculateCellValue Then
                    calculation = New FormulaCalculation(Me)
                    calculation.EvaluateReplacement(sheetName, rowIndex, columnIndex, replacement)
                End If
                EditCell(sheetName, rowIndex, columnIndex, replacement, calculateOnEdit:=False, replaceSpillAnchor:=existingSpill IsNot Nothing)
                If newMetadata IsNot Nothing Then _authoredDynamicMetadata = newMetadata
                If calculation IsNot Nothing Then
                    calculation.AttachDynamicReplacement(sheetName, rowIndex, columnIndex, GetCell(sheetName, rowIndex, columnIndex))
                    calculation.Commit()
                ElseIf existingSpill IsNot Nothing Then
                    ClearSpillChildResults(GetSheetData(part).Descendants(Of Cell)(), {existingSpill})
                End If
            End SyncLock
        End Sub

        Private Shared Function NormalizeDynamicArrayFormula(expression As String) As String
            Dim normalized = DynamicArrayFormulaTokens.Replace(expression,
                Function(token)
                    Dim anchor = token.Groups("anchor")
                    If anchor.Success Then Return "_xlfn.ANCHORARRAY(" & anchor.Value & ")"
                    Dim name = token.Groups("name")
                    If Not name.Success Then Return token.Value
                    Dim functionName = name.Value.ToUpperInvariant()
                    Return If(functionName = "FILTER" OrElse functionName = "SORT", "_xlfn._xlws.", "_xlfn.") & functionName
                End Function)
            If Regex.IsMatch(normalized, "_xlfn\.(?:LAMBDA|MAP|BYROW|BYCOL|MAKEARRAY|REDUCE|SCAN|GROUPBY|PIVOTBY)\s*\(", RegexOptions.IgnoreCase) Then
                Return FormulaParser.NormalizeLambdaStorage(normalized)
            End If
            Return normalized
        End Function

        Private Shared Function FindReusableDynamicMetadata(metadata As Metadata) As UInteger
            Dim blocks = metadata?.GetFirstChild(Of CellMetadata)()
            If blocks Is Nothing Then Return 0UI
            Dim index As UInteger
            For Each block In blocks.Elements(Of MetadataBlock)()
                index += 1UI
                If block.ChildElements.Count = 1 AndAlso IsDynamicMetadataRecord(metadata, index) Then Return index
            Next
            Return 0UI
        End Function

        Private Shared Function AppendDynamicMetadata(metadata As Metadata) As UInteger
            Dim types = metadata.MetadataTypes
            If types Is Nothing Then
                types = New MetadataTypes()
                metadata.MetadataTypes = types
            End If
            Dim entries = types.Elements(Of MetadataType)().ToList()
            Dim typeIndex = entries.FindIndex(Function(entry) entry.Name?.Value = DynamicArrayMetadataName)
            If typeIndex = -1 Then
                typeIndex = entries.Count
                types.AppendChild(New MetadataType With {
                    .Name = DynamicArrayMetadataName, .MinSupportedVersion = 120000UI, .Copy = True, .PasteAll = True,
                    .PasteValues = True, .Merge = True, .SplitFirst = True, .RowColumnShift = True,
                    .ClearFormats = True, .ClearComments = True, .Assign = True, .Coerce = True, .CellMeta = True
                })
                types.Count = CUInt(typeIndex + 1)
            End If

            Dim future = GetOrAddFutureMetadata(metadata, DynamicArrayMetadataName)
            Dim futureIndex = CUInt(future.Elements(Of FutureMetadataBlock)().Count())
            Dim extension As New Extension With {.Uri = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"}
            extension.AppendChild(New DynamicArray.DynamicArrayProperties With {.FDynamic = True, .FCollapsed = False})
            future.AppendChild(New FutureMetadataBlock(New ExtensionList(extension)))
            future.Count = futureIndex + 1UI

            Dim cells = metadata.GetFirstChild(Of CellMetadata)()
            If cells Is Nothing Then
                cells = New CellMetadata()
                metadata.AddChild(cells)
            End If
            cells.AppendChild(New MetadataBlock(New MetadataRecord With {.TypeIndex = CUInt(typeIndex + 1), .Val = futureIndex}))
            cells.Count = CUInt(cells.Elements(Of MetadataBlock)().Count())
            Return cells.Count.Value
        End Function

    End Class

End Namespace
