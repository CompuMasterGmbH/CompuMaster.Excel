Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations
        Private NotInheritable Class ChartAnchorGeometry
            Friend ReadOnly Original As Long()
            Friend ReadOnly Edited As Long()

            Friend Sub New(original As Long(), edited As Long())
                Me.Original = original
                Me.Edited = edited
            End Sub

            Friend Function MarkerAt(position As Long) As (Index As Integer, Offset As Long)
                If position < 0 OrElse position >= Edited(Edited.Length - 1) Then Throw UnsupportedReferenceSyntax()
                Dim low = 0
                Dim high = Edited.Length - 1
                While high - low > 1
                    Dim middle = low + (high - low) \ 2
                    If Edited(middle) <= position Then
                        low = middle
                    Else
                        high = middle
                    End If
                End While
                Return (low, position - Edited(low))
            End Function
        End Class

        Private Function CreateChartAnchorGeometry(worksheet As WorksheetPart, shift As StructuralShift, editedData As SheetData,
                                                    editedColumns As Columns) As ChartAnchorGeometry
            Dim format = worksheet.Worksheet.GetFirstChild(Of SheetFormatProperties)()
            If shift.Vertical Then
                Return New ChartAnchorGeometry(ReadChartRowPositions(GetSheetData(worksheet), format), ReadChartRowPositions(editedData, format))
            End If

            Dim styles = worksheet.GetParentParts().OfType(Of WorkbookPart)().Single().WorkbookStylesPart?.Stylesheet
            Dim normal = styles?.CellStyles?.Elements(Of CellStyle)().FirstOrDefault(Function(style) If(style.BuiltinId?.Value, UInteger.MaxValue) = 0)
            Dim styleFormat = styles?.CellStyleFormats?.Elements(Of CellFormat)().ElementAtOrDefault(CInt(If(normal?.FormatId?.Value, 0UI)))
            Dim font = styles?.Fonts?.Elements(Of Font)().ElementAtOrDefault(CInt(If(styleFormat?.FontId?.Value, 0UI)))
            Dim fontName = If(font?.FontName?.Val?.Value, "Calibri")
            Dim fontSize = If(font?.FontSize?.Val?.Value, 11.0)
            If (fontSize <> 11 OrElse (Not String.Equals(fontName, "Calibri", StringComparison.OrdinalIgnoreCase) AndAlso
                                       Not String.Equals(fontName, "Aptos Narrow", StringComparison.OrdinalIgnoreCase))) AndAlso
               Not LoadOptions.AllowKnownStructuralReferenceDifferences Then
                Throw New StructuralReferenceCompatibilityException(EngineName,
                    "keeping the physical chart size requires the Excel column width of the workbook's Normal font")
            End If
            Const digitWidth As Integer = 7
            Dim defaultWidth = If(format?.DefaultColumnWidth IsNot Nothing, ColumnWidthInPixels(format.DefaultColumnWidth.Value, digitWidth),
                CLng(Math.Ceiling((If(format?.BaseColumnWidth?.Value, 8UI) * digitWidth + 5) / 8.0)) * 8)
            Return New ChartAnchorGeometry(
                ReadChartColumnPositions(worksheet.Worksheet.GetFirstChild(Of Columns)(), defaultWidth, digitWidth),
                ReadChartColumnPositions(editedColumns, defaultWidth, digitWidth))
        End Function

        Private Shared Function ReadChartRowPositions(data As SheetData, format As SheetFormatProperties) As Long()
            Const emusPerPixel As Long = 9525
            Dim defaultPixels = If(If(format?.ZeroHeight?.Value, False), 0L, RowHeightInPixels(If(format?.DefaultRowHeight?.Value, 15.0)))
            Dim positions(RowCount) As Long
            For i = 1 To RowCount
                positions(i) = defaultPixels * emusPerPixel
            Next
            Dim rowIndex = 0
            For Each row In data.Elements(Of Row)()
                rowIndex = CInt(If(row.RowIndex?.Value, CUInt(rowIndex + 1)))
                If If(row.Hidden?.Value, False) Then
                    positions(rowIndex) = 0
                ElseIf row.Height IsNot Nothing Then
                    positions(rowIndex) = RowHeightInPixels(row.Height.Value) * emusPerPixel
                End If
            Next
            SumChartPositions(positions)
            Return positions
        End Function

        Private Shared Function ReadChartColumnPositions(columns As Columns, defaultPixels As Long, digitWidth As Integer) As Long()
            Const emusPerPixel As Long = 9525
            Dim positions(ColumnCount) As Long
            For i = 1 To ColumnCount
                positions(i) = defaultPixels * emusPerPixel
            Next
            If columns IsNot Nothing Then
                For Each column In columns.Elements(Of Column)()
                    Dim pixels = If(If(column.Hidden?.Value, False), 0L,
                                   If(column.Width Is Nothing, defaultPixels, ColumnWidthInPixels(column.Width.Value, digitWidth)))
                    For i = CInt(column.Min.Value) To CInt(column.Max.Value)
                        positions(i) = pixels * emusPerPixel
                    Next
                Next
            End If
            SumChartPositions(positions)
            Return positions
        End Function

        Private Shared Function RowHeightInPixels(points As Double) As Long
            Return CLng(Math.Round(points / 0.75, MidpointRounding.AwayFromZero))
        End Function

        Private Shared Function ColumnWidthInPixels(width As Double, digitWidth As Integer) As Long
            ' Convert the stored character width to pixels using the Normal font's maximum digit width.
            Return CLng(Math.Truncate((256 * width + (128 \ digitWidth)) / 256 * digitWidth))
        End Function

        Private Shared Sub SumChartPositions(positions As Long())
            For i = 1 To positions.Length - 1
                positions(i) += positions(i - 1)
            Next
        End Sub
    End Class

End Namespace
