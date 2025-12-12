Namespace ExcelOps

#If NETFRAMEWORK Then
    <CodeAnalysis.SuppressMessage("Naming", "CA1708:Bezeichner dürfen sich nicht nur durch die Groß-/Kleinschreibung unterscheiden", Justification:=".NET 8 doesn't implement this rule any more, so might be applicable for .NET Framework only, but .NET 4.8 seems to handle everything correctly")>
    Public Class ExcelCell
#Else
    Public Class ExcelCell
#End If
        Implements ICloneable, IComparable, IEqualityComparer

        Public Sub New(addressWithSheetName As String, dataType As ValueTypes)
            Me.New(SheetNamePart(addressWithSheetName), LocalAddressPart(addressWithSheetName), dataType)
        End Sub

        Public Sub New(sheetName As String, addressWithoutSheetName As String, dataType As ValueTypes)
            Me.SheetName = sheetName
            Me.Address = addressWithoutSheetName
            Me.DataType = dataType
        End Sub

        Public Sub New(sheetName As String, rowIndex As Integer, columnIndex As Integer, dataType As ValueTypes)
            Me.New(sheetName, LocalCellAddress(rowIndex, columnIndex), dataType)
        End Sub

        Private Shared Function LocalAddressPart(addressWithSheetName As String) As String
            If addressWithSheetName.IndexOf("!"c) >= 0 Then
                Return addressWithSheetName.Substring(addressWithSheetName.IndexOf("!"c) + 1).Replace("$", "")
            Else
                Return addressWithSheetName.Replace("$", "")
            End If
        End Function

        Private Shared Function SheetNamePart(addressWithSheetName As String) As String
            If addressWithSheetName.IndexOf("!"c) >= 0 Then
                Dim Result As String = addressWithSheetName.Substring(0, addressWithSheetName.IndexOf("!"c))
                If Result.StartsWith("'", StringComparison.InvariantCulture) AndAlso Result.EndsWith("'", StringComparison.InvariantCulture) Then
                    Result = Result.Substring(1, Result.Length - 2)
                End If
                Return Result
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Translate row/column index to a local Excel sheet address without sheetname (e.g. 'A1')
        ''' </summary>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        <Obsolete("Use LocalCellAddress instead"), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared ReadOnly Property Address(rowIndex As Integer, columnIndex As Integer) As String
            Get
                Return LocalCellAddress(rowIndex, columnIndex)
            End Get
        End Property

        ''' <summary>
        ''' Translate row/column index to a local Excel sheet address without sheetname (e.g. 'A1')
        ''' </summary>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Public Shared Function LocalCellAddress(rowIndex As Integer, columnIndex As Integer) As String
            If rowIndex < 0 Then Throw New ArgumentOutOfRangeException(NameOf(rowIndex), "Must be positive or zero")
            If columnIndex < 0 Then Throw New ArgumentOutOfRangeException(NameOf(columnIndex), "Must be positive or zero")
            Return ExcelColumnName(columnIndex) & (rowIndex + 1).ToString
        End Function

        ''' <summary>
        ''' Validate that the sheet name and cell address are assigned and valid
        ''' </summary>
        ''' <returns></returns>
        Public Function ValidateFullCellAddressInclSheetName() As Boolean
            Return Me.SheetName <> Nothing AndAlso Me.Address <> Nothing And IsValidAddress(Me.Address, True)
        End Function

        ''' <summary>
        ''' Cell value types
        ''' </summary>
        Public Enum ValueTypes
            All = -1
            Text = 0
            Number = 1
            DateTime = 2
            Formula = 3
            Bool = 4
        End Enum

        ''' <summary>
        ''' Name of sheet
        ''' </summary>
        ''' <returns></returns>
        Public Property SheetName As String

        Private _Address As String
        ''' <summary>
        ''' An Excel cell address like A1 (without absolute $-addressing like $A$1)
        ''' </summary>
        ''' <returns></returns>
        Public Property Address As String
            Get
                Return _Address
            End Get
            Set(value As String)
                If IsValidAddress(value, False) = False Then Throw New ArgumentException("Address """ & value & """ must be a valid cell address e.g. like A1", NameOf(value))
                _Address = value
            End Set
        End Property

        ''' <summary>
        ''' An Excel cell address like A1 (without absolute $-addressing like $A$1), optionally inclusive sheet name
        ''' </summary>
        ''' <param name="inclusiveSheetName"></param>
        ''' <returns></returns>
        Public ReadOnly Property Address(inclusiveSheetName As Boolean) As String
            Get
                Dim Result As String = Nothing
                If inclusiveSheetName AndAlso Me.SheetName <> Nothing Then
                    Result = "'" & Me.SheetName & "'!"
                End If
                Result &= _Address
                Return Result
            End Get
        End Property

        ''' <summary>
        ''' An Excel cell address like A1 or $A$1, optionally inclusive sheet name
        ''' </summary>
        ''' <param name="inclusiveSheetName">Add sheetname to address</param>
        ''' <param name="useAbsoluteAddressingForColumn">Use $-addressing for column like $A</param>
        ''' <param name="useAbsoluteAddressingForRow">Use $-addressing for row like $1</param>
        ''' <returns></returns>
        Public ReadOnly Property Address(inclusiveSheetName As Boolean, useAbsoluteAddressingForColumn As Boolean, useAbsoluteAddressingForRow As Boolean) As String
            Get
                Dim Result As String = Nothing
                If inclusiveSheetName AndAlso Me.SheetName <> Nothing Then
                    Result = "'" & Me.SheetName & "'!"
                End If
                Result &= If(useAbsoluteAddressingForColumn, "$", "") & ExcelColumnName(Me.ColumnIndex)
                Result &= If(useAbsoluteAddressingForRow, "$", "") & Me.RowNumber
                Return Result
            End Get
        End Property

        ''' <summary>
        ''' An Excel cell address like R1C1, optionally inclusive sheet name
        ''' </summary>
        ''' <param name="inclusiveSheetName"></param>
        ''' <returns></returns>
        Public ReadOnly Property AddressR1C1(inclusiveSheetName As Boolean) As String
            Get
                Dim Result As String = Nothing
                If inclusiveSheetName AndAlso Me.SheetName <> Nothing Then
                    Result = "'" & Me.SheetName & "'!"
                End If
                Result &= "R" & Me.RowNumber & "C" & Me.ColumnNumber
                Return Result
            End Get
        End Property

        ''' <summary>
        ''' An address like "A1"
        ''' </summary>
        ''' <returns></returns>
        Public Function LocalAddress() As String
            Return Me.Address(False)
        End Function

        ''' <summary>
        ''' An address like "R1C1"
        ''' </summary>
        ''' <returns></returns>
        Public Function LocalAddressR1C1() As String
            Return Me.AddressR1C1(False)
        End Function

        ''' <summary>
        ''' An address like "Sheetname!A1"
        ''' </summary>
        ''' <returns></returns>
        Public Function FullAddress() As String
            Return Me.Address(True)
        End Function

        ''' <summary>
        ''' An address like "Sheetname!R1C1"
        ''' </summary>
        ''' <returns></returns>
        Public Function FullAddressR1C1() As String
            Return Me.AddressR1C1(True)
        End Function

        ''' <summary>
        ''' Expected cell value type
        ''' </summary>
        ''' <returns></returns>
        Public Property DataType As ValueTypes

        Friend Shared ReadOnly Property ExcelColumnName(columnIndex As Integer) As String
            Get
                If columnIndex < 0 Then Throw New ArgumentOutOfRangeException(NameOf(columnIndex), "Must be a positive value")
                Dim x As Integer = columnIndex + 1
                If x <= 26 Then
                    Return Chr(x + 64)
                End If
                Dim quotient As Integer = x \ 26
                Dim remainder As Integer = x Mod 26
                If remainder = 0 Then
                    remainder = 26
                    quotient -= 1
                End If
                Return ExcelColumnName(quotient - 1) & Chr(remainder + 64)
            End Get
        End Property

        ''' <summary>
        ''' 0-based index
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property RowIndex() As Integer
            Get
                Return Me.RowNumber - 1
            End Get
        End Property

        ''' <summary>
        ''' 0-based index
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ColumnIndex() As Integer
            Get
                Return CalculateColumnIndex(Me.ColumnName)
            End Get
        End Property

        ''' <summary>
        ''' 1-based index
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ColumnNumber() As Integer
            Get
                Return CalculateColumnIndex(Me.ColumnName) + 1
            End Get
        End Property

        ''' <summary>
        ''' 0-based index
        ''' </summary>
        ''' <param name="columnAddressPart">Column letters, e.g. "A", "B", ..., "AA", "AB", ...</param>
        ''' <returns></returns>
        Public Shared Function CalculateColumnIndex(columnAddressPart As String) As Integer
            Dim Result As Integer = 0
            Dim ColName As String = columnAddressPart.ToUpperInvariant
            For MyCounter As Integer = 0 To ColName.Length - 1
                Dim CharValue As Integer = AscW(ColName(ColName.Length - MyCounter - 1)) - AscW("A"c) + 1
                If CharValue > 26 Or CharValue < 0 Then Throw New ArgumentException("Column name must contain letters A-Z only", NameOf(columnAddressPart))
                Dim Exponent As Integer = MyCounter
                Result += CType(CharValue * 26 ^ Exponent, Integer)
            Next
            Return Result - 1
        End Function

        ''' <summary>
        ''' Validate if a potential address is a valid Excel cell address
        ''' </summary>
        ''' <param name="cellAddress"></param>
        ''' <param name="allowAbsoluteAddressing">Allow absolute addresses like $A$1</param>
        ''' <returns></returns>
        Public Shared Function IsValidAddress(ByVal cellAddress As String, allowAbsoluteAddressing As Boolean) As Boolean
            If cellAddress = Nothing Then Return False
            If cellAddress.StartsWith("""", StringComparison.InvariantCulture) Then Return False 'invalid address - is a string
            If cellAddress.EndsWith("""", StringComparison.InvariantCulture) Then Return False 'invalid address - is a string
            If cellAddress.IndexOf("!"c) >= 0 Then
                Dim SheetNamePart As String = cellAddress.Substring(0, cellAddress.IndexOf("!"c))
                If SheetNamePart.StartsWith("'", StringComparison.InvariantCulture) Xor SheetNamePart.EndsWith("'", StringComparison.InvariantCulture) Then
                    Return False 'invalid address - invalid sheet name, must be either with "'" at start and end or without any "'"
                End If
                'drop sheet name part
                cellAddress = cellAddress.Substring(cellAddress.IndexOf("!"c) + 1)
            End If
            If allowAbsoluteAddressing Then
                If cellAddress.StartsWith("$", StringComparison.InvariantCulture) Then
                    cellAddress = cellAddress.Substring(1)
                End If
            End If
            Dim FirstDigit As Integer = AddressRowNumberStartIndex(cellAddress)
            If FirstDigit < 0 Then Return False 'no digits -> invalid
            If FirstDigit = 0 Then Return False 'no letters -> invalid
            If allowAbsoluteAddressing Then
                If cellAddress(FirstDigit - 1) = "$"c Then
                    cellAddress = cellAddress.Remove(FirstDigit - 1, 1)
                    FirstDigit = AddressRowNumberStartIndex(cellAddress)
                End If
            End If
            If FirstDigit > 3 Then Return False 'more than 3 letters (=non-digits) -> invalid because max column address part is "XFD"
            Dim ColPart As String = cellAddress.Substring(0, FirstDigit)
            Dim RowPart As String = cellAddress.Substring(FirstDigit)
            Try
                If Integer.TryParse(RowPart, Nothing) = False Then
                    Return False 'not a number
                ElseIf Integer.Parse(RowPart) < 1 Then
                    Return False 'not a valid row number
                End If
                If CalculateColumnIndex(ColPart) > 16383 Then
                    Return False 'column XFD is last possible column with column index 16383
                End If
            Catch
                Return False
            End Try
            Return True
        End Function

        ''' <summary>
        ''' Column letters, e.g. "A", "B", ..., "AA", "AB", ...
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ColumnName As String
            Get
                Return Me.Address.Substring(0, Me.AddressRowNumberStartIndex)
            End Get
        End Property

        ''' <summary>
        ''' 1-based index
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property RowNumber As Integer
            Get
                Return Integer.Parse(Me.Address.Substring(Me.AddressRowNumberStartIndex))
            End Get
        End Property

        Private Shared ReadOnly AddressRowNumberStartIndex_AnyOf As Char() = New Char() {"1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c, "0"c}

        Private Shared Function AddressRowNumberStartIndex(address As String) As Integer
            Return address.IndexOfAny(AddressRowNumberStartIndex_AnyOf)
        End Function

        Private Function AddressRowNumberStartIndex() As Integer
            Return AddressRowNumberStartIndex(Me.Address)
        End Function

        ''' <summary>
        ''' A string representation of the address like "Sheetname!A1:B2"
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function ToString() As String
            Return Me.Address(True)
        End Function

        ''' <summary>
        ''' A string representation of the address
        ''' </summary>
        ''' <param name="inclusiveSheetName"></param>
        ''' <returns></returns>
        Public Overloads Function ToString(inclusiveSheetName As Boolean) As String
            Return Me.Address(inclusiveSheetName)
        End Function

        ''' <summary>
        ''' An independent clone of this ExcelCell
        ''' </summary>
        ''' <returns></returns>
        Private Function ICloneable_Clone() As Object Implements ICloneable.Clone
            Return New ExcelCell(Me.SheetName, Me.Address, Me.DataType)
        End Function

        ''' <summary>
        ''' An independent clone of this ExcelCell
        ''' </summary>
        ''' <returns></returns>
        Public Function Clone() As ExcelCell
            Return New ExcelCell(Me.SheetName, Me.Address, Me.DataType)
        End Function

        ''' <summary>
        ''' Create a clone but override the sheet name to the specified name
        ''' </summary>
        ''' <param name="overrideSheetName"></param>
        ''' <returns></returns>
        Public Function Clone(overrideSheetName As String) As ExcelCell
            Return New ExcelCell(overrideSheetName, Me.Address, Me.DataType)
        End Function

        ''' <summary>
        ''' Create a clone but override the sheet name to the specified name as well as the data type
        ''' </summary>
        ''' <param name="overrideSheetName"></param>
        ''' <returns></returns>
        Public Function Clone(overrideSheetName As String, dataType As ExcelCell.ValueTypes) As ExcelCell
            Return New ExcelCell(overrideSheetName, Me.Address, dataType)
        End Function

        ''' <summary>
        ''' Create a clone but override the data type to the specified one
        ''' </summary>
        ''' <param name="dataType"></param>
        ''' <returns></returns>
        Public Function Clone(dataType As ExcelCell.ValueTypes) As ExcelCell
            Return New ExcelCell(Me.SheetName, Me.Address, dataType)
        End Function

        ''' <summary>
        ''' Create a clone but override the data type to the specified one
        ''' </summary>
        ''' <returns></returns>
        Public Function Clone(rowIndex As Integer, columnIndex As Integer) As ExcelCell
            Return New ExcelCell(Me.SheetName, rowIndex, columnIndex, Me.DataType)
        End Function

        ''' <summary>
        ''' Create a clone pointing to a new cell position relative to the current cell
        ''' </summary>
        ''' <param name="addRows"></param>
        ''' <param name="addColumns"></param>
        ''' <returns></returns>
        Public Function GoToRelativePosition(addRows As Integer, addColumns As Integer) As ExcelCell
            Return New ExcelCell(Me.SheetName, Me.RowIndex + addRows, Me.ColumnIndex + addColumns, Me.DataType)
        End Function

#Region "Equality and comparison"
        Private Function IEqualityComparer_Equals(x As Object, y As Object) As Boolean Implements IEqualityComparer.Equals
            Return CType(x, ExcelCell) = CType(y, ExcelCell)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Me = CType(obj, ExcelCell)
        End Function

        Public Shared Operator =(x As ExcelCell, y As ExcelCell) As Boolean
            Return x.CompareTo(y) = 0
        End Operator

        Public Shared Operator <>(x As ExcelCell, y As ExcelCell) As Boolean
            Return x.CompareTo(y) <> 0
        End Operator

        Private Function IEqualityComparer_GetHashCode(obj As Object) As Integer Implements IEqualityComparer.GetHashCode
            If obj Is Nothing OrElse GetType(ExcelCell).IsInstanceOfType(obj) = False Then Throw New ArgumentException("Comparison requires values of type ExcelCell")
            Return CType(obj, ExcelCell).GetHashCode
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Me.ToString(True).GetHashCode
        End Function

        Public Shared Operator <(x As ExcelCell, y As ExcelCell) As Boolean
            Return x.CompareTo(y) < 0
        End Operator

        Public Shared Operator >(x As ExcelCell, y As ExcelCell) As Boolean
            Return x.CompareTo(y) > 0
        End Operator

        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            If obj Is Nothing OrElse GetType(ExcelCell).IsInstanceOfType(obj) = False Then Throw New ArgumentException("Comparison requires values of type ExcelCell")
            Dim ComparisonRange = CType(obj, ExcelCell)
            If Me.RowIndex < ComparisonRange.RowIndex Then
                Return -2
            ElseIf Me.RowIndex > ComparisonRange.RowIndex Then
                Return 2
            ElseIf Me.ColumnIndex < ComparisonRange.ColumnIndex Then
                Return -1
            ElseIf Me.ColumnIndex > ComparisonRange.ColumnIndex Then
                Return 1
            Else
                Return 0
            End If
        End Function
#End Region

    End Class

End Namespace