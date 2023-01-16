﻿Namespace ExcelOps

    Public Class ExcelCell
        Implements ICloneable

        Public Sub New(addressWithSheetName As String, dataType As ValueTypes)
            Me.New(SheetNamePart(addressWithSheetName), LocalAddressPart(addressWithSheetName), dataType)
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
                If Result.StartsWith("'") AndAlso Result.EndsWith("'") Then
                    Result = Result.Substring(1, Result.Length - 2)
                End If
                Return Result
            Else
                Return Nothing
            End If
        End Function

        Public Sub New(sheetName As String, addressWithoutSheetName As String, dataType As ValueTypes)
            Me.SheetName = sheetName
            Me.Address = addressWithoutSheetName
            Me.DataType = dataType
        End Sub

        Public Sub New(sheetName As String, rowIndex As Integer, columnIndex As Integer, dataType As ValueTypes)
            Me.New(sheetName, LocalCellAddress(rowIndex, columnIndex), dataType)
        End Sub

        ''' <summary>
        ''' Validate that the sheet name and cell address are assigned and valid
        ''' </summary>
        ''' <returns></returns>
        Public Function ValidateFullCellAddressInclSheetName() As Boolean
            Return Me.SheetName <> Nothing AndAlso Me.Address <> Nothing And IsValidAddress(Me.Address, True)
        End Function

        Public Shared Function LocalCellAddress(rowIndex As Integer, columnIndex As Integer) As String
            If rowIndex < 0 Then Throw New ArgumentOutOfRangeException(NameOf(rowIndex), "Must be positive or zero")
            If columnIndex < 0 Then Throw New ArgumentOutOfRangeException(NameOf(columnIndex), "Must be positive or zero")
            Return ExcelColumnName(columnIndex) & (rowIndex + 1).ToString
        End Function

        Public Enum ValueTypes
            All = -1
            Text = 0
            Number = 1
            DateTime = 2
            Formula = 3
            Bool = 4
        End Enum

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
                If IsValidAddress(value, False) = False Then Throw New ArgumentException("Must be a valid cell address e.g. like A1", NameOf(value))
                _Address = value
            End Set
        End Property

        ''' <summary>
        ''' An Excel cell address like A1 (without absolute $-addressing like $A$1) inclusive sheet name
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

        Public Property DataType As ValueTypes

        Friend Shared ReadOnly Property ExcelColumnName(columnIndex As Integer) As String
            Get
                If columnIndex < 0 Then Throw New ArgumentOutOfRangeException(NameOf(columnIndex), "Must be a positive value")
                Dim x As Integer = columnIndex + 1
                If x >= 1 And x <= 26 Then
                    Return Chr(x + 64)
                Else
                    Return ExcelColumnName(CType(((x - x Mod 26) / 26), Integer) - 1) & Chr((x Mod 26) + 64)
                End If
            End Get
        End Property

        ''' <summary>
        ''' Translate row/column index to MS Excel sheet address (e.g. 'A1')
        ''' </summary>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Public Shared ReadOnly Property Address(rowIndex As Integer, columnIndex As Integer) As String
            Get
                Return ExcelColumnName(columnIndex) & (rowIndex + 1).ToString
            End Get
        End Property

        ''' <summary>
        ''' 0-based index
        ''' </summary>
        ''' <returns></returns>
        Public Function RowIndex() As Integer
            Return Me.RowNumber - 1
        End Function

        ''' <summary>
        ''' 0-based index
        ''' </summary>
        ''' <returns></returns>
        Public Function ColumnIndex() As Integer
            Return ColumnIndex(Me.ColumnName)
        End Function

        ''' <summary>
        ''' 1-based index
        ''' </summary>
        ''' <returns></returns>
        Public Function ColumnNumber() As Integer
            Return ColumnIndex(Me.ColumnName) + 1
        End Function

        ''' <summary>
        ''' 0-based index
        ''' </summary>
        ''' <param name="columnAddressPart">Column letters, e.g. "A", "B", ..., "AA", "AB", ...</param>
        ''' <returns></returns>
        Public Shared Function ColumnIndex(columnAddressPart As String) As Integer
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
            If cellAddress.StartsWith("""") Then Return False 'invalid address - is a string
            If cellAddress.EndsWith("""") Then Return False 'invalid address - is a string
            If cellAddress.IndexOf("!"c) >= 0 Then
                Dim SheetNamePart As String = cellAddress.Substring(0, cellAddress.IndexOf("!"c))
                If SheetNamePart.StartsWith("'") Xor SheetNamePart.EndsWith("'") Then
                    Return False 'invalid address - invalid sheet name, must be either with "'" at start and end or without any "'"
                End If
                'drop sheet name part
                cellAddress = cellAddress.Substring(cellAddress.IndexOf("!"c) + 1)
            End If
            If allowAbsoluteAddressing Then
                If cellAddress.StartsWith("$") Then
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
                If ColumnIndex(ColPart) > 16383 Then
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

        Private Shared Function AddressRowNumberStartIndex(address As String) As Integer
            Return address.IndexOfAny(New Char() {"1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c, "0"c})
        End Function

        Private Function AddressRowNumberStartIndex() As Integer
            Return AddressRowNumberStartIndex(Me.Address)
        End Function

        Public Overrides Function ToString() As String
            Return Me.Address(True)
        End Function

        Public Overloads Function ToString(inclusiveSheetName As Boolean) As String
            Return Me.Address(inclusiveSheetName)
        End Function

        Private Function ICloneable_Clone() As Object Implements ICloneable.Clone
            Return New ExcelCell(Me.SheetName, Me.Address, Me.DataType)
        End Function

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

    End Class

End Namespace