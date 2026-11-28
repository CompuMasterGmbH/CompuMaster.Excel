Imports System.Drawing

Namespace ExcelOps

    ''' <summary>Provides optional formatting and layout operations for a worksheet.</summary>
    Public Interface IExcelWorksheetFormatting
        ''' <summary>Applies selected formatting to every cell in a range.</summary>
        ''' <param name="range">Worksheet and rectangular range to format, including blank cells.</param>
        ''' <param name="formatting">Properties to change. Unspecified properties retain their existing formatting.</param>
        Sub ApplyCellFormatting(range As ExcelRange, formatting As ExcelCellFormatting)

        ''' <summary>Sets an explicit width for a worksheet column.</summary>
        ''' <param name="sheetName">Name of the worksheet.</param>
        ''' <param name="columnIndex">Zero-based column index.</param>
        ''' <param name="width">Width in Excel character units from 0 through 255.</param>
        Sub SetColumnWidth(sheetName As String, columnIndex As Integer, width As Double)

        ''' <summary>Sets an explicit height for a worksheet row.</summary>
        ''' <param name="sheetName">Name of the worksheet.</param>
        ''' <param name="rowIndex">Zero-based row index.</param>
        ''' <param name="height">Height in points from 0 through 409.</param>
        Sub SetRowHeight(sheetName As String, rowIndex As Integer, height As Double)
    End Interface

    ''' <summary>Describes the formatting properties to change on a range.</summary>
    Public NotInheritable Class ExcelCellFormatting
        ''' <summary>Gets or sets the Excel number format code, such as "0.00%" or "General".</summary>
        Public Property NumberFormat As String

        ''' <summary>Gets or sets the font family name.</summary>
        Public Property FontName As String

        ''' <summary>Gets or sets the font size in points.</summary>
        Public Property FontSize As Double?

        ''' <summary>Gets or sets whether the font is bold.</summary>
        Public Property Bold As Boolean?

        ''' <summary>Gets or sets whether the font is italic.</summary>
        Public Property Italic As Boolean?

        ''' <summary>Gets or sets the font RGB color.</summary>
        Public Property FontColor As Color?

        ''' <summary>Gets or sets a solid background color. Use Color.Empty to clear the background.</summary>
        Public Property FillColor As Color?

        ''' <summary>Gets or sets horizontal alignment.</summary>
        Public Property HorizontalAlignment As ExcelHorizontalAlignment?

        ''' <summary>Gets or sets vertical alignment.</summary>
        Public Property VerticalAlignment As ExcelVerticalAlignment?

        ''' <summary>Gets or sets whether cell contents wrap.</summary>
        Public Property WrapText As Boolean?

        ''' <summary>Gets or sets the border along the outer edges of the range.</summary>
        Public Property OutsideBorder As ExcelCellBorder

        ''' <summary>Gets or sets the border between cells inside the range.</summary>
        Public Property InsideBorder As ExcelCellBorder
    End Class

    ''' <summary>Specifies the horizontal placement of cell contents.</summary>
    Public Enum ExcelHorizontalAlignment
        ''' <summary>Uses Excel's automatic alignment for the cell value.</summary>
        General
        ''' <summary>Aligns contents to the left.</summary>
        Left
        ''' <summary>Centers contents horizontally.</summary>
        Center
        ''' <summary>Aligns contents to the right.</summary>
        Right
        ''' <summary>Justifies contents across the cell width.</summary>
        Justify
    End Enum

    ''' <summary>Specifies the vertical placement of cell contents.</summary>
    Public Enum ExcelVerticalAlignment
        ''' <summary>Aligns contents to the top.</summary>
        Top
        ''' <summary>Centers contents vertically.</summary>
        Center
        ''' <summary>Aligns contents to the bottom.</summary>
        Bottom
        ''' <summary>Justifies contents across the cell height.</summary>
        Justify
    End Enum

    ''' <summary>Describes an edge of a worksheet cell.</summary>
    Public NotInheritable Class ExcelCellBorder
        ''' <summary>Creates a border with a line style and an optional RGB color.</summary>
        ''' <param name="style">Line style. Specify None to clear an edge.</param>
        ''' <param name="color">RGB color, or nothing to use the automatic color.</param>
        Public Sub New(style As ExcelBorderLineStyle, Optional color As Color? = Nothing)
            Me.Style = style
            Me.Color = color
        End Sub

        ''' <summary>Gets the line style of the border.</summary>
        Public ReadOnly Property Style As ExcelBorderLineStyle

        ''' <summary>Gets the RGB color of the border, if specified.</summary>
        Public ReadOnly Property Color As Color?
    End Class

    ''' <summary>Specifies the line style of a cell border.</summary>
    Public Enum ExcelBorderLineStyle
        ''' <summary>Clears the border.</summary>
        None
        ''' <summary>Uses a thin solid line.</summary>
        Thin
        ''' <summary>Uses a medium solid line.</summary>
        Medium
        ''' <summary>Uses a thick solid line.</summary>
        Thick
        ''' <summary>Uses a double line.</summary>
        [Double]
        ''' <summary>Uses a dashed line.</summary>
        Dashed
        ''' <summary>Uses a dotted line.</summary>
        Dotted
    End Enum

End Namespace
