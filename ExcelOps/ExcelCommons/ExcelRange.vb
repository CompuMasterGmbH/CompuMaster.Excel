Namespace ExcelOps

    Public Class ExcelRange
        Implements IEnumerable(Of ExcelCell), ICloneable

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
        ''' An enumerator for a cells in this range
        ''' </summary>
        ''' <returns></returns>
        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return New ExcelRangeEnum(Me.AddressStart, Me.AddressEnd)
        End Function

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