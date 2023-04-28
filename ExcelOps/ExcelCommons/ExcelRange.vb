Namespace ExcelOps

    Public Class ExcelRange
        Implements IEnumerable(Of ExcelCell), ICloneable, IEqualityComparer, IComparable

        ''' <summary>
        ''' Create a range from 2 cells including all cells within this rectangle
        ''' </summary>
        ''' <param name="addressStart"></param>
        ''' <param name="addressEnd"></param>
        Public Sub New(addressStart As ExcelCell, addressEnd As ExcelCell)
            If addressStart.SheetName <> addressEnd.SheetName Then Throw New ArgumentException("Cells must be member of the same sheet")
            If addressStart.DataType <> addressEnd.DataType Then
                Throw New ArgumentException("Cells must use the same data type (addressStart: " & addressStart.DataType.ToString & ", addressEnd: " & addressEnd.DataType.ToString & ")")
            End If
            Me.AddressStart = addressStart
            Me.AddressEnd = addressEnd
            If addressStart.ColumnIndex > addressEnd.ColumnIndex OrElse addressStart.RowIndex > addressEnd.RowIndex Then
                Throw New ArgumentException("AddressStart (" & addressStart.Address(True) & ") must be located before AddressEnd (" & addressEnd.Address(True) & ")")
            End If
        End Sub

        ''' <summary>
        ''' Create a range from a single cell
        ''' </summary>
        ''' <param name="singleCell"></param>
        Public Sub New(singleCell As ExcelCell)
            Me.New(singleCell, singleCell)
        End Sub

        ''' <summary>
        ''' Create a range from an address string
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="range"></param>
        Public Sub New(sheetName As String, range As String)
            Me.New(
                New ExcelOps.ExcelCell(sheetName, Tools.LookupCellAddresFromRange(range, 0), ExcelCell.ValueTypes.All),
                New ExcelOps.ExcelCell(sheetName, Tools.LookupCellAddresFromRange(range, 1), ExcelCell.ValueTypes.All)
                )
        End Sub

        ''' <summary>
        ''' First cell of range
        ''' </summary>
        ''' <returns></returns>
        Public Property AddressStart As ExcelCell

        ''' <summary>
        ''' Last cell of range
        ''' </summary>
        ''' <returns></returns>
        Public Property AddressEnd As ExcelCell

        ''' <summary>
        ''' Name of sheet
        ''' </summary>
        ''' <returns></returns>
        Public Property SheetName As String
            Get
                Return Me.AddressStart.SheetName
            End Get
            Set(value As String)
                Me.AddressStart.SheetName = value
                Me.AddressEnd.SheetName = value
            End Set
        End Property

        ''' <summary>
        ''' An address like "A1:B2"
        ''' </summary>
        ''' <returns></returns>
        Public Function LocalAddress() As String
            Return Me.ToString(False)
        End Function

        ''' <summary>
        ''' An address like "Sheetname!A1:B2"
        ''' </summary>
        ''' <returns></returns>
        Public Function FullAddress() As String
            Return Me.ToString(True)
        End Function

        ''' <summary>
        ''' A string representation of the address like "Sheetname!A1:B2"
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function ToString() As String
            Return Me.ToString(True)
        End Function

        ''' <summary>
        ''' A string representation of the address
        ''' </summary>
        ''' <param name="inclusiveSheetName"></param>
        ''' <returns></returns>
        Public Overloads Function ToString(inclusiveSheetName As Boolean) As String
            Return Me.AddressStart.ToString(inclusiveSheetName) & ":" & Me.AddressEnd.ToString(False)
        End Function

        ''' <summary>
        ''' Number of cells in range
        ''' </summary>
        ''' <returns></returns>
        Public Function CellCount() As Integer
            Dim Cols As Integer = Me.AddressEnd.ColumnIndex - Me.AddressStart.ColumnIndex + 1
            Dim Rows As Integer = Me.AddressEnd.RowIndex - Me.AddressStart.RowIndex + 1
            Return Cols * Rows
        End Function

        ''' <summary>
        ''' An enumerator for a cells in this range
        ''' </summary>
        ''' <returns></returns>
        Public Function GetEnumerator() As IEnumerator(Of ExcelCell) Implements IEnumerable(Of ExcelCell).GetEnumerator
            Return New ExcelRangeEnum(Me.AddressStart, Me.AddressEnd)
        End Function

        ''' <summary>
        ''' An independent clone of this ExcelRange
        ''' </summary>
        ''' <returns></returns>
        Private Function ICloneable_Clone() As Object Implements ICloneable.Clone
            Return New ExcelRange(Me.AddressStart.Clone, Me.AddressEnd.Clone)
        End Function

        ''' <summary>
        ''' An independent clone of this ExcelRange
        ''' </summary>
        ''' <returns></returns>
        Public Function Clone() As ExcelRange
            Return New ExcelRange(Me.AddressStart.Clone, Me.AddressEnd.Clone)
        End Function

        ''' <summary>
        ''' Create a clone but override the sheet name to the specified name
        ''' </summary>
        ''' <param name="overrideSheetName"></param>
        ''' <returns></returns>
        Public Function Clone(overrideSheetName As String) As ExcelRange
            Dim Result As ExcelRange = Me.Clone
            Result.SheetName = overrideSheetName
            Return Result
        End Function

        ''' <summary>
        ''' A cell of this range
        ''' </summary>
        ''' <param name="index"></param>
        ''' <returns></returns>
        Public ReadOnly Property Cell(index As Integer) As ExcelCell
            Get
                Return Me.Cell(index, CellAccessDirection.AllCellsOfARowThenNextRow)
            End Get
        End Property

        ''' <summary>
        ''' A cell of this range
        ''' </summary>
        ''' <param name="index"></param>
        ''' <param name="accessDirection"></param>
        ''' <returns></returns>
        Public ReadOnly Property Cell(index As Integer, accessDirection As CellAccessDirection) As ExcelCell
            Get
                Dim RangeRectangleRowsCount As Integer = AddressEnd.RowIndex - AddressStart.RowIndex + 1
                Dim RangeRectangleColumnsCount As Integer = AddressEnd.ColumnIndex - AddressStart.ColumnIndex + 1
                Dim RowIndexWithinRangeRectangle As Integer
                Dim ColumnIndexWithinRangeRectangle As Integer
                Select Case accessDirection
                    Case CellAccessDirection.AllCellsOfARowThenNextRow
                        RowIndexWithinRangeRectangle = System.Math.DivRem(index, RangeRectangleRowsCount, ColumnIndexWithinRangeRectangle)
                    Case CellAccessDirection.AllCellsOfAColumnThenNextColumn
                        ColumnIndexWithinRangeRectangle = System.Math.DivRem(index, RangeRectangleColumnsCount, RowIndexWithinRangeRectangle)
                    Case Else
                        Throw New ArgumentOutOfRangeException(NameOf(accessDirection))
                End Select
                Return New ExcelCell(AddressStart.SheetName,
                                     AddressStart.RowIndex + RowIndexWithinRangeRectangle,
                                     AddressStart.ColumnIndex + ColumnIndexWithinRangeRectangle,
                                     AddressStart.DataType)
            End Get
        End Property

        ''' <summary>
        ''' An direction to access the cells of a range
        ''' </summary>
        Public Enum CellAccessDirection As Byte
            ''' <summary>
            ''' Row by row, column by column
            ''' </summary>
            ''' <remarks>In a Range with 3x3 cells, the cell at index 3 is located in the 2nd row and 1st column</remarks>
            AllCellsOfARowThenNextRow = 0
            ''' <summary>
            ''' Column by column, row by row
            ''' </summary>
            ''' <remarks>In a Range with 3x3 cells, the cell at index 3 is located in the 2nd column and 1st row</remarks>
            AllCellsOfAColumnThenNextColumn = 1
        End Enum

        ''' <summary>
        ''' An enumerator for a cells in this range
        ''' </summary>
        ''' <returns></returns>
        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return New ExcelRangeEnum(Me.AddressStart, Me.AddressEnd)
        End Function

#Region "Equality and comparison"
        Private Function IEqualityComparer_Equals(x As Object, y As Object) As Boolean Implements IEqualityComparer.Equals
            Return CType(x, ExcelRange) = CType(y, ExcelRange)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Me = CType(obj, ExcelRange)
        End Function

        Public Shared Operator =(x As ExcelRange, y As ExcelRange) As Boolean
            Return x.CompareTo(y) = 0
        End Operator

        Public Shared Operator <>(x As ExcelRange, y As ExcelRange) As Boolean
            Return x.CompareTo(y) <> 0
        End Operator

        Private Function IEqualityComparer_GetHashCode(obj As Object) As Integer Implements IEqualityComparer.GetHashCode
            If obj Is Nothing OrElse GetType(ExcelRange).IsInstanceOfType(obj) = False Then Throw New ArgumentException("Comparison requires values of type ExcelRange")
            Return CType(obj, ExcelRange).GetHashCode
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Me.ToString(True).GetHashCode
        End Function

        Public Shared Operator <(x As ExcelRange, y As ExcelRange) As Boolean
            Return x.CompareTo(y) < 0
        End Operator

        Public Shared Operator >(x As ExcelRange, y As ExcelRange) As Boolean
            Return x.CompareTo(y) > 0
        End Operator

        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            If obj Is Nothing OrElse GetType(ExcelRange).IsInstanceOfType(obj) = False Then Throw New ArgumentException("Comparison requires values of type ExcelRange")
            Dim ComparisonRange = CType(obj, ExcelRange)
            If Me.AddressStart < ComparisonRange.AddressStart Then
                Return -2
            ElseIf Me.AddressStart > ComparisonRange.AddressStart Then
                Return 2
            ElseIf Me.CellCount < ComparisonRange.CellCount Then
                Return -1
            ElseIf Me.CellCount > ComparisonRange.CellCount Then
                Return 1
            Else
                Return 0
            End If
        End Function
#End Region

        ''' <summary>
        ''' An enumerator for a cells in this range
        ''' </summary>
        Public Class ExcelRangeEnum
            Implements IEnumerator(Of ExcelCell)

            Friend Sub New(addressStart As ExcelCell, addressEnd As ExcelCell)
                AllCells = New List(Of ExcelCell)
                For MyRowCounter As Integer = addressStart.RowIndex To addressEnd.RowIndex
                    For MyColCounter As Integer = addressStart.ColumnIndex To addressEnd.ColumnIndex
                        AllCells.Add(New ExcelCell(addressStart.SheetName, MyRowCounter, MyColCounter, addressStart.DataType))
                    Next
                Next
            End Sub

            Private ReadOnly AllCells As List(Of ExcelCell)

            ' Enumerators are positioned before the first element
            ' until the first MoveNext() call.
            Dim position As Integer = -1

            Public ReadOnly Property Current As ExcelCell Implements IEnumerator(Of ExcelCell).Current
                Get
                    Return AllCells(position)
                End Get
            End Property

            Private ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
                Get
                    Return AllCells(position)
                End Get
            End Property

            Public Sub Reset() Implements IEnumerator.Reset
                position = -1
            End Sub

            Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
                position += 1
                Return (position < AllCells.Count)
            End Function

#Region "IDisposable Support"
            ' IDisposable
            Protected Overridable Sub Dispose(disposing As Boolean)
                'nothing to do
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                Dispose(True)
            End Sub
#End Region

        End Class

    End Class

End Namespace