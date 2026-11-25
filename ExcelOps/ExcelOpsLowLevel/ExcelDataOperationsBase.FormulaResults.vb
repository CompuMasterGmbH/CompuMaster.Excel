Namespace ExcelOps

    ' Keeps virtual dispatch to the original signatures for existing engine assemblies.
    Friend Interface ILegacyCellValueReader
        Function ReadCellValue(Of T)(cell As ExcelCell) As T
        Function ReadCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T
        Function ReadCellValueAsObject(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Object
    End Interface

    ''' <summary>Specifies whether reading a formula result may trigger calculation.</summary>
    Public Enum FormulaCalculationMode
        ''' <summary>Uses only the result currently stored for the formula.</summary>
        UseCachedValueOnly = 0
        ''' <summary>Attempts to calculate the formula if no result is stored.</summary>
        CalculateIfMissing = 1
        ''' <summary>Uses the engine's original cell-reading behavior, including any automatic calculation it performs.</summary>
        EngineDefault = 2
    End Enum

    Partial Public MustInherit Class ExcelDataOperationsBase

        ''' <summary>Determines whether a formula cell has no stored result.</summary>
        ''' <param name="cell">The cell to inspect.</param>
        ''' <returns>Whether the formula result is missing.</returns>
        ''' <remarks>This method does not calculate the cell.</remarks>
        Public Function HasMissingCachedFormulaResult(cell As ExcelCell) As Boolean
            ValidateFormulaResultAddress(cell)
            Return HasMissingCachedFormulaResult(cell.SheetName, cell.RowIndex, cell.ColumnIndex)
        End Function

        ''' <summary>Determines whether a formula cell has no stored result.</summary>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="rowIndex">The zero-based row index.</param>
        ''' <param name="columnIndex">The zero-based column index.</param>
        ''' <returns>Whether the formula result is missing.</returns>
        ''' <remarks>Implementations must inspect the cache without calculating the cell.</remarks>
        ''' <exception cref="NotSupportedException">The engine does not expose formula result caches.</exception>
        Public Overridable Function HasMissingCachedFormulaResult(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Boolean
            Throw New NotSupportedException(EngineName & " does not expose formula result caches.")
        End Function

        ''' <summary>Reads a cell value with an explicit formula calculation policy.</summary>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="rowIndex">The zero-based row index.</param>
        ''' <param name="columnIndex">The zero-based column index.</param>
        ''' <param name="calculationMode">Whether a missing formula result may be calculated.</param>
        ''' <returns>The cell value.</returns>
        ''' <exception cref="MissingCalculatedCellValueException">The formula has no result after applying the requested policy.</exception>
        ''' <exception cref="FeatureDisabledException">The requested calculation module is disabled.</exception>
        ''' <exception cref="NotSupportedException">The engine does not support the requested calculation.</exception>
        ''' <exception cref="FormulaComplexityNotSupportedException">The requested scope cannot calculate this formula.</exception>
        Public Function LookupCellValueAsObject(sheetName As String, rowIndex As Integer, columnIndex As Integer, calculationMode As FormulaCalculationMode) As Object
            RequireFormulaResult(sheetName, rowIndex, columnIndex, calculationMode)
            Return ReadLegacyCellValueAsObject(sheetName, rowIndex, columnIndex)
        End Function

        ''' <summary>Reads and converts a cell value with an explicit formula calculation policy.</summary>
        ''' <typeparam name="T">The requested value type.</typeparam>
        ''' <param name="cell">The cell to read.</param>
        ''' <param name="calculationMode">Whether a missing formula result may be calculated.</param>
        ''' <returns>The converted cell value.</returns>
        ''' <exception cref="MissingCalculatedCellValueException">The formula has no result after applying the requested policy.</exception>
        ''' <exception cref="FeatureDisabledException">The requested calculation module is disabled.</exception>
        ''' <exception cref="NotSupportedException">The engine does not support the requested calculation.</exception>
        ''' <exception cref="FormulaComplexityNotSupportedException">The requested scope cannot calculate this formula.</exception>
        Public Function LookupCellValue(Of T)(cell As ExcelCell, calculationMode As FormulaCalculationMode) As T
            ValidateFormulaResultAddress(cell)
            RequireFormulaResult(cell.SheetName, cell.RowIndex, cell.ColumnIndex, calculationMode)
            Return ReadLegacyCellValue(Of T)(cell)
        End Function

        ''' <summary>Reads and converts a cell value with an explicit formula calculation policy.</summary>
        ''' <typeparam name="T">The requested value type.</typeparam>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="rowIndex">The zero-based row index.</param>
        ''' <param name="columnIndex">The zero-based column index.</param>
        ''' <param name="calculationMode">Whether a missing formula result may be calculated.</param>
        ''' <returns>The converted cell value.</returns>
        ''' <exception cref="MissingCalculatedCellValueException">The formula has no result after applying the requested policy.</exception>
        ''' <exception cref="FeatureDisabledException">The requested calculation module is disabled.</exception>
        ''' <exception cref="NotSupportedException">The engine does not support the requested calculation.</exception>
        ''' <exception cref="FormulaComplexityNotSupportedException">The requested scope cannot calculate this formula.</exception>
        Public Function LookupCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer, calculationMode As FormulaCalculationMode) As T
            RequireFormulaResult(sheetName, rowIndex, columnIndex, calculationMode)
            Return ReadLegacyCellValue(Of T)(sheetName, rowIndex, columnIndex)
        End Function

        ''' <inheritdoc cref="LookupCellValueAsObject(String, Integer, Integer, FormulaCalculationMode)"/>
        ''' <param name="cell">The cell to read.</param>
        ''' <param name="calculationMode">The explicit formula calculation policy.</param>
        Public Function LookupCellValueAsObject(cell As ExcelCell, calculationMode As FormulaCalculationMode) As Object
            ValidateFormulaResultAddress(cell)
            Return LookupCellValueAsObject(cell.SheetName, cell.RowIndex, cell.ColumnIndex, calculationMode)
        End Function

        ''' <summary>Reads the cells in a range with an explicit formula calculation policy.</summary>
        ''' <typeparam name="T">The requested value type.</typeparam>
        ''' <param name="range">The cells to read.</param>
        ''' <param name="calculationMode">The explicit formula calculation policy for each cell.</param>
        ''' <returns>The converted cell values in range order.</returns>
        Public Function LookupCellValue(Of T)(range As ExcelRange, calculationMode As FormulaCalculationMode) As List(Of T)
            If range Is Nothing Then Throw New ArgumentNullException(NameOf(range))
            If calculationMode <> FormulaCalculationMode.EngineDefault Then ValidateFormulaCalculationMode(calculationMode)
            Dim result As New List(Of T)
            For Each cell In range
                result.Add(LookupCellValue(Of T)(cell, calculationMode))
            Next
            Return result
        End Function

        ''' <summary>Reads a cell using the engine's original lookup implementation.</summary>
        ''' <typeparam name="T">The requested value type.</typeparam>
        ''' <param name="cell">The cell to read.</param>
        ''' <returns>The converted cell value.</returns>
        ''' <remarks>Used by pre-existing engine APIs to preserve their lookup behavior.</remarks>
        Protected Function ReadLegacyCellValue(Of T)(cell As ExcelCell) As T
            Return DirectCast(Me, ILegacyCellValueReader).ReadCellValue(Of T)(cell)
        End Function

        ''' <summary>Reads a cell using the engine's original lookup implementation.</summary>
        ''' <typeparam name="T">The requested value type.</typeparam>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="rowIndex">The zero-based row index.</param>
        ''' <param name="columnIndex">The zero-based column index.</param>
        ''' <returns>The converted cell value.</returns>
        ''' <remarks>Used by pre-existing engine APIs to preserve their lookup behavior.</remarks>
        Protected Function ReadLegacyCellValue(Of T)(sheetName As String, rowIndex As Integer, columnIndex As Integer) As T
            Return DirectCast(Me, ILegacyCellValueReader).ReadCellValue(Of T)(sheetName, rowIndex, columnIndex)
        End Function

        ''' <summary>Reads a cell as an object using the engine's original lookup implementation.</summary>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="rowIndex">The zero-based row index.</param>
        ''' <param name="columnIndex">The zero-based column index.</param>
        ''' <returns>The cell value.</returns>
        ''' <remarks>Used by pre-existing engine APIs to preserve their lookup behavior.</remarks>
        Protected Function ReadLegacyCellValueAsObject(sheetName As String, rowIndex As Integer, columnIndex As Integer) As Object
            Return DirectCast(Me, ILegacyCellValueReader).ReadCellValueAsObject(sheetName, rowIndex, columnIndex)
        End Function

        Private Shared Sub ValidateFormulaResultAddress(cell As ExcelCell)
            If cell Is Nothing Then Throw New ArgumentNullException(NameOf(cell))
            If Not cell.ValidateFullCellAddressInclSheetName() Then Throw New InvalidCellAddressException(cell)
        End Sub

        Private Sub RequireFormulaResult(sheetName As String, rowIndex As Integer, columnIndex As Integer, calculationMode As FormulaCalculationMode)
            If calculationMode = FormulaCalculationMode.EngineDefault Then Return
            If FormulaResultIsStillMissing(sheetName, rowIndex, columnIndex, calculationMode) Then
                Throw MissingFormulaResult(sheetName, rowIndex, columnIndex)
            End If
        End Sub

        Private Function FormulaResultIsStillMissing(sheetName As String, rowIndex As Integer, columnIndex As Integer, calculationMode As FormulaCalculationMode) As Boolean
            ValidateFormulaCalculationMode(calculationMode)
            If Not HasMissingCachedFormulaResult(sheetName, rowIndex, columnIndex) Then Return False
            If calculationMode = FormulaCalculationMode.UseCachedValueOnly Then Return True
            CalculateMissingFormulaResult(sheetName, rowIndex, columnIndex)
            Return HasMissingCachedFormulaResult(sheetName, rowIndex, columnIndex)
        End Function

        Private Shared Sub ValidateFormulaCalculationMode(calculationMode As FormulaCalculationMode)
            If calculationMode <> FormulaCalculationMode.UseCachedValueOnly AndAlso calculationMode <> FormulaCalculationMode.CalculateIfMissing Then
                Throw New ArgumentOutOfRangeException(NameOf(calculationMode))
            End If
        End Sub

        Private Function MissingFormulaResult(sheetName As String, rowIndex As Integer, columnIndex As Integer) As MissingCalculatedCellValueException
            Return New MissingCalculatedCellValueException(FilePath, sheetName, rowIndex, columnIndex, LookupCellFormula(sheetName, rowIndex, columnIndex))
        End Function

        ''' <summary>Attempts to calculate a formula whose result is missing.</summary>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="rowIndex">The zero-based row index.</param>
        ''' <param name="columnIndex">The zero-based column index.</param>
        ''' <remarks>Excel error values returned by an evaluation count as results.</remarks>
        ''' <exception cref="FeatureDisabledException">The calculation module is disabled.</exception>
        Protected Overridable Sub CalculateMissingFormulaResult(sheetName As String, rowIndex As Integer, columnIndex As Integer)
            If CalculationModuleDisabled Then Throw New FeatureDisabledException("Calculation engine")
            RecalculateCell(sheetName, rowIndex, columnIndex, throwExceptionOnCalculationError:=False)
        End Sub

        ''' <summary>Resolves the text to render for a cell according to the HTML formula options.</summary>
        ''' <param name="sheetName">The worksheet name.</param>
        ''' <param name="rowIndex">The zero-based row index.</param>
        ''' <param name="columnIndex">The zero-based column index.</param>
        ''' <param name="options">The export options.</param>
        ''' <param name="readFormattedText">Reads the formatted cell text when a replacement is unnecessary.</param>
        ''' <returns>The cell text before HTML encoding.</returns>
        Protected Function ResolveHtmlCellText(sheetName As String, rowIndex As Integer, columnIndex As Integer, options As HtmlSheetExportOptions,
                                               readFormattedText As Func(Of String)) As String
            If Not FormulaResultIsStillMissing(sheetName, rowIndex, columnIndex, options.FormulaCalculationMode) Then Return readFormattedText()
            If options.MissingFormulaResultText IsNot Nothing Then Return options.MissingFormulaResultText
            Throw MissingFormulaResult(sheetName, rowIndex, columnIndex)
        End Function

    End Class

End Namespace
