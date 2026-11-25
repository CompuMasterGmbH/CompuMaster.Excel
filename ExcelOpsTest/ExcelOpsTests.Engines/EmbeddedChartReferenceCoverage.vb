Imports System.Globalization
Imports System.Xml.Linq
Imports CompuMaster.Excel.ExcelOps
Imports DocumentFormat.OpenXml.Packaging

Namespace ExcelOpsTests.Engines
    Friend NotInheritable Class EmbeddedChartReferenceCoverage
        Friend Const FixtureName As String = "EmbeddedChartPlacement.xlsx"
        Private Const Dashboard As String = "Dashboard"

        Friend Shared Function CompareWithExcel(workbook As ExcelDataOperationsBase, path As String, vertical As Boolean,
                                                operation As String) As List(Of String)
            Dim differences As New List(Of String)
            Try
                ApplyEdit(workbook, vertical, operation)
            Catch ex As NotSupportedException
                differences.Add(ex.Message)
                Return differences
            End Try
            workbook.SaveCopyAs(path, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
            workbook.Close()
            Dim expected = ExpectedPlacements(vertical, operation).OrderBy(Function(value) value, StringComparer.Ordinal).ToArray()
            Dim actual = ReadPlacements(path).OrderBy(Function(value) value, StringComparer.Ordinal).ToArray()
            If Not expected.SequenceEqual(actual) Then
                differences.Add("Excel chart placement: " & String.Join(" | ", expected))
                differences.Add("Saved chart placement: " & String.Join(" | ", actual))
            End If
            Return differences
        End Function

        Friend Shared Sub ApplyEdit(workbook As ExcelDataOperationsBase, vertical As Boolean, operation As String)
            Dim first = If(vertical, 2, 1)
            Dim last = If(vertical, 7, 6)
            Dim index = If(operation.EndsWith("Before", StringComparison.Ordinal), first - 1,
                If(operation.EndsWith("After", StringComparison.Ordinal), last + 1,
                   If(operation = "DeleteAll", first, first + 2)))
            Dim count = If(operation = "DeleteAll", last - first + 1, 1)
            If operation.StartsWith("Insert", StringComparison.Ordinal) Then
                If vertical Then
                    workbook.AddRow(Dashboard, index, count, True)
                Else
                    workbook.AddColumn(Dashboard, index, count, True)
                End If
            ElseIf vertical Then
                workbook.RemoveRows(Dashboard, index, count, True)
            Else
                workbook.RemoveColumns(Dashboard, index, count, True)
            End If
        End Sub

        Private Shared Iterator Function ExpectedPlacements(vertical As Boolean, operation As String) As IEnumerable(Of String)
            Dim size = If(vertical, 15.0, 60.0)
            For Each mode In {"twoCell", "oneCell", "absolute"}
                Dim movement = 0.0
                Dim resize = 0.0
                If mode <> "absolute" Then
                    If operation = "DeleteBefore" Then movement = -size
                    If operation = "InsertBefore" Then movement = size
                    If operation = "DeleteAll" Then movement = If(vertical, -4.0, -5.0)
                End If
                If mode = "twoCell" Then
                    If operation = "DeleteInside" Then resize = -size
                    If operation = "InsertInside" Then resize = size
                    If operation = "DeleteAll" Then resize = If(vertical, -79, -307)
                End If
                Yield FormatPlacement(mode, 65 + If(vertical, 0, movement), 34 + If(vertical, movement, 0),
                                      307 + If(vertical, 0, resize), 79 + If(vertical, resize, 0))
            Next
        End Function

        Private Shared Iterator Function ReadPlacements(path As String) As IEnumerable(Of String)
            Using document = SpreadsheetDocument.Open(path, False)
                Dim drawing = StructuralReferenceCoverage.WorksheetFor(document.WorkbookPart, Dashboard).DrawingsPart?.WorksheetDrawing
                If drawing Is Nothing Then Return
                Dim root = XElement.Parse(drawing.OuterXml)
                Dim ns = root.Name.Namespace
                For Each anchor In root.Elements().Where(Function(item) item.Descendants().Any(Function(element) element.Name.LocalName = "chart"))
                    Dim first = anchor.Element(ns + "from")
                    Dim last = anchor.Element(ns + "to")
                    Dim origin = anchor.Element(ns + "pos")
                    Dim size = anchor.Element(ns + "ext")
                    Dim mode = If(CStr(anchor.Attribute("editAs")), anchor.Name.LocalName.Replace("Anchor", ""))
                    Dim left = If(first Is Nothing, CDbl(origin.Attribute("x")) / 12700.0, MeasureMarker(first, True))
                    Dim top = If(first Is Nothing, CDbl(origin.Attribute("y")) / 12700.0, MeasureMarker(first, False))
                    Dim width = If(last Is Nothing, CDbl(size.Attribute("cx")) / 12700.0, MeasureMarker(last, True) - left)
                    Dim height = If(last Is Nothing, CDbl(size.Attribute("cy")) / 12700.0, MeasureMarker(last, False) - top)
                    Yield FormatPlacement(mode, left, top, width, height)
                Next
            End Using
        End Function

        Private Shared Function MeasureMarker(marker As XElement, horizontal As Boolean) As Double
            Dim ns = marker.Name.Namespace
            Dim axis = If(horizontal, "col", "row")
            Return CDbl(marker.Element(ns + axis)) * If(horizontal, 60, 15) + CDbl(marker.Element(ns + (axis & "Off"))) / 12700.0
        End Function

        Private Shared Function FormatPlacement(mode As String, x As Double, y As Double, width As Double, height As Double) As String
            Return String.Format(CultureInfo.InvariantCulture, "{0} at ({1:0.###}, {2:0.###}), {3:0.###} x {4:0.###} points", mode, x, y, width, height)
        End Function
    End Class
End Namespace
