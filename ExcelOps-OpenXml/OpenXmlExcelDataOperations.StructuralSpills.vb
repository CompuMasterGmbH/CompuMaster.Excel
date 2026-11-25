Imports System.IO
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Shared Function HasSupportedSpillMetadata(metadata As Metadata) As Boolean
            If metadata Is Nothing OrElse metadata.ChildElements.Count <> 3 Then Return False
            Dim types = metadata.GetFirstChild(Of MetadataTypes)()
            Dim category = metadata.GetFirstChild(Of FutureMetadata)()
            Dim records = metadata.GetFirstChild(Of CellMetadata)()
            If types Is Nothing OrElse types.ChildElements.Count <> 1 OrElse types.GetFirstChild(Of MetadataType)()?.Name?.Value <> "XLDAPR" OrElse
               category Is Nothing OrElse category.ChildElements.Count <> 1 OrElse category.Name?.Value <> "XLDAPR" OrElse
               records Is Nothing OrElse records.ChildElements.Count <> 1 OrElse records.FirstChild.ChildElements.Count <> 1 Then Return False
            Dim record = TryCast(records.FirstChild.FirstChild, MetadataRecord)
            Dim extension = category.Descendants().SingleOrDefault(Function(item) item.LocalName = "dynamicArrayProperties" AndAlso
                item.NamespaceUri = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray")
            If extension Is Nothing OrElse extension.HasChildren Then Return False
            Dim dynamic = extension.GetAttributes().FirstOrDefault(Function(item) item.LocalName = "fDynamic").Value
            Dim collapsed = extension.GetAttributes().FirstOrDefault(Function(item) item.LocalName = "fCollapsed").Value
            Return record IsNot Nothing AndAlso If(record.TypeIndex?.Value, 0UI) = 1UI AndAlso If(record.Val?.Value, 1UI) = 0UI AndAlso
                (dynamic = "1" OrElse dynamic = "true") AndAlso (collapsed Is Nothing OrElse collapsed = "0" OrElse collapsed = "false")
        End Function

        Private NotInheritable Class SpillArea
            Friend ReadOnly First As ReferenceEndpoint
            Friend ReadOnly Last As ReferenceEndpoint
            Friend ReadOnly TargetRow As Integer
            Friend ReadOnly TargetColumn As Integer
            Friend Obstructed As Boolean

            Friend Sub New(anchor As Cell, shift As StructuralShift)
                Dim area = ReadStructuralRange(anchor.CellFormula.Reference?.Value)
                First = area.First
                Last = area.Last
                Dim cell = ReadSharedCellAddress(anchor)
                If First.Row <> cell.Row OrElse First.Column <> cell.Column Then
                    Throw New InvalidDataException("The dynamic array reference does not start at its anchor.")
                End If
                TargetRow = If(shift IsNot Nothing AndAlso shift.Vertical AndAlso shift.Includes(cell.Column), shift.MoveIndex(cell.Row), cell.Row)
                TargetColumn = If(shift IsNot Nothing AndAlso Not shift.Vertical AndAlso shift.Includes(cell.Row), shift.MoveIndex(cell.Column), cell.Column)
                If TargetRow >= RowCount - Height + 1 OrElse TargetColumn >= ColumnCount - Width + 1 Then
                    Throw New NotSupportedException("Shifting a dynamic spill range beyond the worksheet is unsupported.")
                End If
            End Sub

            Friend ReadOnly Property Height As Integer
                Get
                    Return Last.Row - First.Row + 1
                End Get
            End Property

            Friend ReadOnly Property Width As Integer
                Get
                    Return Last.Column - First.Column + 1
                End Get
            End Property

            Friend Function Contains(row As Integer, column As Integer) As Boolean
                Return row >= First.Row AndAlso row <= Last.Row AndAlso column >= First.Column AndAlso column <= Last.Column
            End Function

            Friend Function ContainsTarget(row As Integer, column As Integer) As Boolean
                Return TargetRow >= 0 AndAlso TargetColumn >= 0 AndAlso
                    row >= TargetRow AndAlso row < TargetRow + Height AndAlso column >= TargetColumn AndAlso column < TargetColumn + Width
            End Function

            Friend Function TargetReference() As String
                Dim start = ExcelCell.LocalCellAddress(TargetRow, TargetColumn)
                If Obstructed OrElse (Height = 1 AndAlso Width = 1) Then Return start
                Return start & ":" & ExcelCell.LocalCellAddress(TargetRow + Height - 1, TargetColumn + Width - 1)
            End Function

            Friend Function Overlaps(other As SpillArea) As Boolean
                Return TargetRow >= 0 AndAlso TargetColumn >= 0 AndAlso other.TargetRow >= 0 AndAlso other.TargetColumn >= 0 AndAlso
                    TargetRow < other.TargetRow + other.Height AndAlso other.TargetRow < TargetRow + Height AndAlso
                    TargetColumn < other.TargetColumn + other.Width AndAlso other.TargetColumn < TargetColumn + Width
            End Function
        End Class

        Private Shared Function PrepareStructuralSpills(source As SheetData, shift As StructuralShift) As SheetData
            Dim spills = source.Descendants(Of Cell)().Where(Function(cell) cell.CellMetaIndex IsNot Nothing).
                Select(Function(anchor) New SpillArea(anchor, shift)).ToArray()
            If spills.Length = 0 Then Return source
            Dim data = DirectCast(source.CloneNode(True), SheetData)
            Dim anchors As New List(Of (Formula As CellFormula, Area As SpillArea))
            For Each cell In data.Descendants(Of Cell)()
                Dim address = ReadSharedCellAddress(cell)
                Dim owner = spills.FirstOrDefault(Function(spill) spill.Contains(address.Row, address.Column))
                If owner IsNot Nothing AndAlso cell.CellMetaIndex IsNot Nothing AndAlso
                   (owner.First.Row <> address.Row OrElse owner.First.Column <> address.Column) Then
                    Throw New InvalidDataException("Dynamic spill ranges overlap.")
                End If
                If owner IsNot Nothing AndAlso cell.CellMetaIndex Is Nothing Then
                    If cell.CellFormula IsNot Nothing Then Throw New InvalidDataException("A dynamic spill range covers another formula.")
                    ' Remove generated values before shifting. Excel calculates the new result from the anchor on reopen.
                    cell.CellValue = Nothing
                    cell.InlineString = Nothing
                    cell.DataType = Nothing
                    Continue For
                End If

                Dim row = If(shift IsNot Nothing AndAlso shift.Vertical AndAlso shift.Includes(address.Column),
                             shift.MoveIndex(address.Row), address.Row)
                Dim column = If(shift IsNot Nothing AndAlso Not shift.Vertical AndAlso shift.Includes(address.Row),
                                shift.MoveIndex(address.Column), address.Column)
                If row < 0 OrElse column < 0 Then Continue For
                If cell.CellFormula IsNot Nothing OrElse cell.CellValue IsNot Nothing OrElse cell.InlineString IsNot Nothing Then
                    For Each spill In spills
                        If spill IsNot owner AndAlso spill.ContainsTarget(row, column) Then spill.Obstructed = True
                    Next
                End If
                If owner IsNot Nothing Then anchors.Add((cell.CellFormula, owner))
            Next
            Dim unobstructed = spills.Where(Function(spill) Not spill.Obstructed).ToArray()
            For index = 0 To unobstructed.Length - 1
                For other = index + 1 To unobstructed.Length - 1
                    If unobstructed(index).Overlaps(unobstructed(other)) Then
                        Throw New NotSupportedException("The structural edit would overlap two dynamic spill ranges.")
                    End If
                Next
            Next
            For Each item In anchors
                item.Formula.Reference = item.Area.TargetReference()
            Next
            Return data
        End Function

    End Class

End Namespace
