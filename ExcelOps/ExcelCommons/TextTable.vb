Imports System.Data

Namespace ExcelOps
#If NETFRAMEWORK Then
    <CodeAnalysis.SuppressMessage("Naming", "CA1708:Bezeichner dürfen sich nicht nur durch die Groß-/Kleinschreibung unterscheiden", Justification:=".NET 8 doesn't implement this rule any more, so might be applicable for .NET Framework only, but .NET 4.8 seems to handle everything correctly")>
    Public Class TextTable
#Else
    Public Class TextTable
#End If
        Implements ICloneable, IDisposable

        ''' <summary>
        ''' A new instance of TextTable
        ''' </summary>
        Public Sub New()
            Me.Table = New DataTable
        End Sub

        Private Sub New(table As TextTable)
            Me.Table = CompuMaster.Data.DataTables.CreateDataTableClone(table.Table)
        End Sub

        ''' <summary>
        ''' A new instance of TextTable based on values of System.DataTable
        ''' </summary>
        ''' <param name="table"></param>
        Public Sub New(table As DataTable)
            Me.New
            LoadFromDataTable(table)
        End Sub

        Private ReadOnly Table As DataTable
        Private disposedValue As Boolean

        Public Property Cell(rowIndex As Integer, columnIndex As Integer) As String
            Get
                Return CompuMaster.Data.Utils.NoDBNull(Me.Table.Rows(rowIndex)(columnIndex), CType(Nothing, String))
            End Get
            Set(value As String)
                Me.Table.Rows(rowIndex)(columnIndex) = CompuMaster.Data.Utils.StringNotNothingOrDBNull(value)
            End Set
        End Property

        Public Property Cell(rowIndex As Integer, columnName As String) As String
            Get
                Return CompuMaster.Data.Utils.NoDBNull(Me.Table.Rows(rowIndex)(columnName), CType(Nothing, String))
            End Get
            Set(value As String)
                Me.Table.Rows(rowIndex)(columnName) = CompuMaster.Data.Utils.StringNotNothingOrDBNull(value)
            End Set
        End Property

        Public Property ColumnName(columnIndex As Integer) As String
            Get
                Return Me.Table.Columns(columnIndex).ColumnName
            End Get
            Set(value As String)
                Me.Table.Columns(columnIndex).ColumnName = value
            End Set
        End Property

        Public Sub AddColumns(number As Integer)
            For MyCounter As Integer = 0 To number - 1
                Me.Table.Columns.Add(Nothing, GetType(String))
            Next
        End Sub

        Public Sub AddColumns(ParamArray columnNames As String())
            For MyCounter As Integer = 0 To columnNames.Length - 1
                Me.Table.Columns.Add(columnNames(MyCounter), GetType(String))
            Next
        End Sub

        Public Sub AddRows(number As Integer)
            For MyCounter As Integer = 0 To number - 1
                Me.Table.Rows.Add()
            Next
        End Sub

        Public Sub AddRow(ParamArray cellData As String())
            If cellData.Length > Me.Table.Columns.Count Then
                Me.AddColumns(cellData.Length - Me.Table.Columns.Count)
            End If
            Me.Table.Rows.Add(New ArrayList(cellData).ToArray)
        End Sub

        Public Sub AddRow(cellData As List(Of String))
            Me.AddRow(cellData.ToArray)
        End Sub

        Public Sub Clear()
            Me.Table.Clear()
        End Sub

        Public Sub FillColumnWithValue(columnIndex As Integer, value As String)
            For MyCounter As Integer = 0 To Me.RowCount - 1
                Me.Cell(MyCounter, columnIndex) = value
            Next
        End Sub

        Public Sub AutoTrim()
            For MyRowCounter As Integer = Me.Table.Rows.Count - 1 To Me.LastContentRowIndex + 1 Step -1
                Me.Table.Rows.RemoveAt(MyRowCounter)
            Next
            For MyColCounter As Integer = Me.Table.Columns.Count - 1 To Me.LastContentColumnIndex + 1 Step -1
                Me.Table.Columns.RemoveAt(MyColCounter)
            Next
        End Sub

        Public Function LastContentRowIndex() As Integer
            For MyRowCounter As Integer = Me.Table.Rows.Count - 1 To 0 Step -1
                For MyColCounter As Integer = Me.Table.Columns.Count - 1 To 0 Step -1
                    If Me.Cell(MyRowCounter, MyColCounter) <> Nothing Then
                        Return MyRowCounter
                    End If
                Next
            Next
            Return -1
        End Function

        Public Function LastContentColumnIndex() As Integer
            For MyColCounter As Integer = Me.Table.Columns.Count - 1 To 0 Step -1
                For MyRowCounter As Integer = Me.Table.Rows.Count - 1 To 0 Step -1
                    If Me.Cell(MyRowCounter, MyColCounter) <> Nothing Then
                        Return MyColCounter
                    End If
                Next
            Next
            Return -1
        End Function

        Public Function ToUITable() As String
            Return CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.Table, New CompuMaster.Data.ConvertToPlainTextTableOptions() With {.MinimumColumnWidth = 2, .MaximumColumnWidth = 65535})
        End Function

        Public Function ToUIExcelTable() As String
            Dim UITable As DataTable = CompuMaster.Data.DataTables.CreateDataTableClone(Me.Table)
            UITable.Columns.Add("RowNo", GetType(Integer))
#Disable Warning IDE0028 ' Initialisierung der Sammlung vereinfachen
            Dim DestinationCols As New List(Of String)
#Enable Warning IDE0028 ' Initialisierung der Sammlung vereinfachen
            DestinationCols.Add("RowNo")
            For MyColCounter As Integer = 0 To UITable.Columns.Count - 1
                If UITable.Columns(MyColCounter).ColumnName <> "RowNo" Then DestinationCols.Add(UITable.Columns(MyColCounter).ColumnName)
            Next
            UITable = CompuMaster.Data.DataTables.CloneTableAndReArrangeDataColumns(UITable, DestinationCols.ToArray)
            'Setup column names in letters
            For MyCounter As Integer = 1 To UITable.Columns.Count - 1
                UITable.Columns(MyCounter).Caption = ExcelColumnName(MyCounter - 1)
            Next
            'Setup row numbers 1-based
            UITable.Columns(0).Caption = "#"
            For MyCounter As Integer = 0 To UITable.Rows.Count - 1
                UITable.Rows(MyCounter)(0) = MyCounter + 1
            Next
            Return CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(UITable, New CompuMaster.Data.ConvertToPlainTextTableOptions() With {.MinimumColumnWidth = 2, .MaximumColumnWidth = 65535})
        End Function

        Friend Shared ReadOnly Property ExcelColumnName(columnIndex As Integer) As String
            Get
                If columnIndex < 0 Then Throw New ArgumentOutOfRangeException(NameOf(columnIndex), "Must be a positive value")
                Dim x As Integer = columnIndex + 1
                If x >= 1 AndAlso x <= 26 Then
                    Return Chr(x + 64)
                Else
                    Return ExcelColumnName(CType(((x - x Mod 26) / 26) - 1, Integer)) & Chr((x Mod 26) + 64)
                End If
            End Get
        End Property

        ''' <summary>
        ''' Translate row/column index to MS Excel sheet address (e.g. 'A1')
        ''' </summary>
        ''' <param name="rowIndex"></param>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Friend Shared ReadOnly Property CellAddress(rowIndex As Integer, columnIndex As Integer) As String
            Get
                Return ExcelColumnName(columnIndex) & (rowIndex + 1).ToString
            End Get
        End Property

        Public Function ToCsvTable() As String
            Return CompuMaster.Data.Csv.WriteDataTableToCsvTextString(Me.Table, False, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf, ",", """"c, "."c)
        End Function

        Public ReadOnly Property ColumnCount As Integer
            Get
                Return Me.Table.Columns.Count
            End Get
        End Property

        Public ReadOnly Property RowCount As Integer
            Get
                Return Me.Table.Rows.Count
            End Get
        End Property

        Public Enum DiffMode As Byte
            ''' <summary>
            ''' Cells with different content (after trimming) 
            ''' </summary>
            DifferentTrimmedCells = 0
            ''' <summary>
            ''' Cells with content (after trimming) are equal in both tables
            ''' </summary>
            EqualTrimmedCellsWithContent = 1
        End Enum

        Public Enum DiffCellOutput As Byte
            ''' <summary>
            ''' Cells without difference are null/Nothing, empty cells in this table but with different value in comparison table are String.Empty, else cells contain content of this table
            ''' </summary>
            CellContentOfThisTable = 0
            ''' <summary>
            ''' 'D' for a diff-match, 'E' for an equal-match, null/Nothing for non-match
            ''' </summary>
            Bool = 1
            ''' <summary>
            ''' 'A' for added values, 'M' for modified values, 'R' for removed values
            ''' </summary>
            ChangeType = 2
        End Enum

        ''' <summary>
        ''' Compare this table with another table and create a new table with just filtered cells of this table
        ''' </summary>
        ''' <param name="comparisonTable"></param>
        ''' <param name="diffType"></param>
        ''' <returns></returns>
        Public Function CompareCells(comparisonTable As TextTable, diffType As DiffMode, outputStyle As DiffCellOutput) As TextTable
            Return Me.CompareCells(comparisonTable, diffType, outputStyle, 0, 0, -1, -1)
        End Function

        ''' <summary>
        ''' Compare this table with another table and create a new table with just filtered cells of this table
        ''' </summary>
        ''' <param name="comparisonTable"></param>
        ''' <param name="diffType"></param>
        ''' <param name="outputStyle"></param>
        ''' <param name="comparisonStartRowIndex"></param>
        ''' <param name="comparisonStartColumnIndex"></param>
        ''' <returns></returns>
        Public Function CompareCells(comparisonTable As TextTable, diffType As DiffMode, outputStyle As DiffCellOutput, comparisonStartRowIndex As Integer, comparisonStartColumnIndex As Integer) As TextTable
            Return Me.CompareCells(comparisonTable, diffType, outputStyle, comparisonStartRowIndex, comparisonStartColumnIndex, -1, -1)
        End Function

        ''' <summary>
        ''' Compare this table with another table and create a new table with just filtered cells of this table
        ''' </summary>
        ''' <param name="comparisonTable"></param>
        ''' <param name="diffType"></param>
        ''' <param name="outputStyle"></param>
        ''' <param name="comparisonStartRowIndex"></param>
        ''' <param name="comparisonStartColumnIndex"></param>
        ''' <param name="comparisonLastRowIndex">-1 for Auto</param>
        ''' <param name="comparisonLastColumnIndex">-1 for Auto</param>
        ''' <returns></returns>
        Public Function CompareCells(comparisonTable As TextTable, diffType As DiffMode, outputStyle As DiffCellOutput, comparisonStartRowIndex As Integer, comparisonStartColumnIndex As Integer, comparisonLastRowIndex As Integer, comparisonLastColumnIndex As Integer) As TextTable
            If comparisonTable Is Nothing Then
                Throw New ArgumentNullException(NameOf(comparisonTable))
            End If
            If outputStyle = DiffCellOutput.ChangeType AndAlso diffType = DiffMode.EqualTrimmedCellsWithContent Then
                Throw New ArgumentException("DiffType EqualTrimmedCellsWithContent can't be combined with outputStyle ChangeType")
            End If
            Dim Result As New TextTable
            Result.AddColumns(System.Math.Max(Me.ColumnCount, comparisonTable.ColumnCount))
            Result.AddRows(System.Math.Max(Me.RowCount, comparisonTable.RowCount))
            For MyColCounter As Integer = comparisonStartColumnIndex To If(
                comparisonLastColumnIndex = -1,
                Result.ColumnCount - 1,
                System.Math.Min(comparisonLastColumnIndex, Result.ColumnCount - 1)
                )
                For MyRowCounter As Integer = comparisonStartRowIndex To If(
                    comparisonLastRowIndex = -1,
                    Result.RowCount - 1,
                    System.Math.Min(comparisonLastRowIndex, Result.RowCount - 1)
                    )
                    Select Case diffType
                        Case DiffMode.DifferentTrimmedCells
                            If Me.CellExists(MyRowCounter, MyColCounter) AndAlso comparisonTable.CellExists(MyRowCounter, MyColCounter) Then
                                If Trim(Me.Cell(MyRowCounter, MyColCounter)) <> Trim(comparisonTable.Cell(MyRowCounter, MyColCounter)) Then
                                    If outputStyle = DiffCellOutput.Bool Then
                                        Result.Cell(MyRowCounter, MyColCounter) = "D"
                                    ElseIf outputStyle = DiffCellOutput.ChangeType Then
                                        If Trim(Me.Cell(MyRowCounter, MyColCounter)) <> Nothing AndAlso Trim(comparisonTable.Cell(MyRowCounter, MyColCounter)) = Nothing Then
                                            Result.Cell(MyRowCounter, MyColCounter) = "A"
                                        ElseIf Trim(Me.Cell(MyRowCounter, MyColCounter)) = Nothing AndAlso Trim(comparisonTable.Cell(MyRowCounter, MyColCounter)) <> Nothing Then
                                            Result.Cell(MyRowCounter, MyColCounter) = "R"
                                        Else
                                            Result.Cell(MyRowCounter, MyColCounter) = "M"
                                        End If
                                    Else
                                        Result.Cell(MyRowCounter, MyColCounter) = CompuMaster.Data.Utils.StringNotNothingOrEmpty(Me.Cell(MyRowCounter, MyColCounter))
                                    End If
                                End If
                            ElseIf Me.CellExists(MyRowCounter, MyColCounter) Then 'AndAlso NOT comparisonTable.CellExists(MyRowCounter, MyColCounter)
                                If Trim(Me.Cell(MyRowCounter, MyColCounter)) <> Nothing Then
                                    If outputStyle = DiffCellOutput.Bool Then
                                        Result.Cell(MyRowCounter, MyColCounter) = "D"
                                    ElseIf outputStyle = DiffCellOutput.ChangeType Then
                                        Result.Cell(MyRowCounter, MyColCounter) = "AC"
                                    Else
                                        Result.Cell(MyRowCounter, MyColCounter) = CompuMaster.Data.Utils.StringNotNothingOrEmpty(Me.Cell(MyRowCounter, MyColCounter))
                                    End If
                                End If
                            ElseIf comparisonTable.CellExists(MyRowCounter, MyColCounter) Then 'AndAlso NOT Me.CellExists(MyRowCounter, MyColCounter) 
                                If Trim(comparisonTable.Cell(MyRowCounter, MyColCounter)) <> Nothing Then
                                    If outputStyle = DiffCellOutput.Bool Then
                                        Result.Cell(MyRowCounter, MyColCounter) = "D"
                                    ElseIf outputStyle = DiffCellOutput.ChangeType Then
                                        Result.Cell(MyRowCounter, MyColCounter) = "RC"
                                    Else
                                        Result.Cell(MyRowCounter, MyColCounter) = CompuMaster.Data.Utils.StringNotNothingOrEmpty(comparisonTable.Cell(MyRowCounter, MyColCounter))
                                    End If
                                End If
                            Else 'NOT Me.CellExists(MyRowCounter, MyColCounter) AndAlso NOT comparisonTable.CellExists(MyRowCounter, MyColCounter)
                                'Table1           Table2        DiffTable:CellNotExistingInBothTables
                                '# |A             # |A |B       # |A |B 
                                '--+--            --+--+--      --+--+--
                                '1 |X             1 |X |X       1 |OK|OK 
                                '2 |X                           2 |OK|XX
                            End If
                        Case DiffMode.EqualTrimmedCellsWithContent
                            If Me.CellExists(MyRowCounter, MyColCounter) AndAlso comparisonTable.CellExists(MyRowCounter, MyColCounter) Then
                                If Trim(Me.Cell(MyRowCounter, MyColCounter)) = Trim(comparisonTable.Cell(MyRowCounter, MyColCounter)) Then
                                    If outputStyle = DiffCellOutput.Bool Then
                                        Result.Cell(MyRowCounter, MyColCounter) = "E"
                                    Else
                                        Result.Cell(MyRowCounter, MyColCounter) = CompuMaster.Data.Utils.StringNotNothingOrEmpty(Me.Cell(MyRowCounter, MyColCounter))
                                    End If
                                End If
                            ElseIf Me.CellExists(MyRowCounter, MyColCounter) Then 'AndAlso NOT comparisonTable.CellExists(MyRowCounter, MyColCounter)
                                If Trim(Me.Cell(MyRowCounter, MyColCounter)) = Nothing Then
                                    If outputStyle = DiffCellOutput.Bool Then
                                        Result.Cell(MyRowCounter, MyColCounter) = "E"
                                    Else
                                        Result.Cell(MyRowCounter, MyColCounter) = CompuMaster.Data.Utils.StringNotNothingOrEmpty(Me.Cell(MyRowCounter, MyColCounter))
                                    End If
                                End If
                            Else 'NOT Me.CellExists(MyRowCounter, MyColCounter) AndAlso comparisonTable.CellExists(MyRowCounter, MyColCounter)
                                If Trim(comparisonTable.Cell(MyRowCounter, MyColCounter)) = Nothing Then
                                    If outputStyle = DiffCellOutput.Bool Then
                                        Result.Cell(MyRowCounter, MyColCounter) = "E"
                                    Else
                                        Result.Cell(MyRowCounter, MyColCounter) = CompuMaster.Data.Utils.StringNotNothingOrEmpty(comparisonTable.Cell(MyRowCounter, MyColCounter))
                                    End If
                                End If
                            End If
                        Case Else
                            Throw New NotImplementedException
                    End Select
                Next
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Compare this table with another table and create a new table with just filtered cells of this table
        ''' </summary>
        ''' <param name="comparisonTable"></param>
        ''' <param name="diffType"></param>
        ''' <returns></returns>
        Public Function CompareCells(comparisonTable As TextTable, diffType As DiffMode) As List(Of String)
            Return Me.CompareCells(comparisonTable, diffType, 0, 0, -1, -1)
        End Function

        ''' <summary>
        ''' Compare this table with another table and create a new table with just filtered cells of this table
        ''' </summary>
        ''' <param name="comparisonTable"></param>
        ''' <param name="diffType"></param>
        ''' <param name="comparisonStartRowIndex"></param>
        ''' <param name="comparisonStartColumnIndex"></param>
        ''' <returns></returns>
        Public Function CompareCells(comparisonTable As TextTable, diffType As DiffMode, comparisonStartRowIndex As Integer, comparisonStartColumnIndex As Integer) As List(Of String)
            Return Me.CompareCells(comparisonTable, diffType, comparisonStartRowIndex, comparisonStartColumnIndex, -1, -1)
        End Function

        ''' <summary>
        ''' Compare this table with another table and create a new table with just filtered cells of this table
        ''' </summary>
        ''' <param name="comparisonTable"></param>
        ''' <param name="diffType"></param>
        ''' <param name="comparisonStartRowIndex"></param>
        ''' <param name="comparisonStartColumnIndex"></param>
        ''' <param name="comparisonLastRowIndex">-1 for Auto</param>
        ''' <param name="comparisonLastColumnIndex">-1 for Auto</param>
        ''' <returns></returns>
        Public Function CompareCells(comparisonTable As TextTable, diffType As DiffMode, comparisonStartRowIndex As Integer, comparisonStartColumnIndex As Integer, comparisonLastRowIndex As Integer, comparisonLastColumnIndex As Integer) As List(Of String)
            Dim Result As New List(Of String)
            For MyColCounter As Integer = comparisonStartColumnIndex To If(
                    comparisonLastColumnIndex = -1,
                    System.Math.Max(Me.ColumnCount, comparisonTable.ColumnCount) - 1,
                    System.Math.Min(comparisonLastColumnIndex, System.Math.Max(Me.ColumnCount, comparisonTable.ColumnCount)) - 1
                    )
                For MyRowCounter As Integer = comparisonStartRowIndex To If(
                    comparisonLastRowIndex = -1,
                    System.Math.Max(Me.RowCount, comparisonTable.RowCount) - 1,
                    System.Math.Min(comparisonStartRowIndex, System.Math.Max(Me.RowCount, comparisonTable.RowCount)) - 1
                    )
                    Select Case diffType
                        Case DiffMode.DifferentTrimmedCells
                            If Me.CellExists(MyRowCounter, MyColCounter) AndAlso comparisonTable.CellExists(MyRowCounter, MyColCounter) Then
                                If Me.Cell(MyRowCounter, MyColCounter) <> comparisonTable.Cell(MyRowCounter, MyColCounter) Then
                                    Result.Add(CellAddress(MyRowCounter, MyColCounter))
                                End If
                            ElseIf Me.CellExists(MyRowCounter, MyColCounter) Then 'AndAlso NOT comparisonTable.CellExists(MyRowCounter, MyColCounter)
                                If Me.Cell(MyRowCounter, MyColCounter) <> Nothing Then
                                    Result.Add(CellAddress(MyRowCounter, MyColCounter))
                                End If
                            Else 'NOT Me.CellExists(MyRowCounter, MyColCounter) AndAlso comparisonTable.CellExists(MyRowCounter, MyColCounter)
                                If comparisonTable.Cell(MyRowCounter, MyColCounter) <> Nothing Then
                                    Result.Add(CellAddress(MyRowCounter, MyColCounter))
                                End If
                            End If
                        Case DiffMode.EqualTrimmedCellsWithContent
                            If Me.CellExists(MyRowCounter, MyColCounter) AndAlso comparisonTable.CellExists(MyRowCounter, MyColCounter) Then
                                If Me.Cell(MyRowCounter, MyColCounter) = comparisonTable.Cell(MyRowCounter, MyColCounter) Then
                                    Result.Add(CellAddress(MyRowCounter, MyColCounter))
                                End If
                            ElseIf Me.CellExists(MyRowCounter, MyColCounter) Then 'AndAlso NOT comparisonTable.CellExists(MyRowCounter, MyColCounter)
                                If Me.Cell(MyRowCounter, MyColCounter) = Nothing Then
                                    Result.Add(CellAddress(MyRowCounter, MyColCounter))
                                End If
                            Else 'NOT Me.CellExists(MyRowCounter, MyColCounter) AndAlso comparisonTable.CellExists(MyRowCounter, MyColCounter)
                                If comparisonTable.Cell(MyRowCounter, MyColCounter) = Nothing Then
                                    Result.Add(CellAddress(MyRowCounter, MyColCounter))
                                End If
                            End If
                        Case Else
                            Throw New NotImplementedException
                    End Select
                Next
            Next
            Return Result
        End Function

        Public Function CellExists(rowIndex As Integer, columnIndex As Integer) As Boolean
            Return rowIndex < Me.RowCount AndAlso columnIndex < Me.ColumnCount
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse obj.GetType() IsNot GetType(TextTable) Then
                Return False
            Else
                Return Me.GetHashCode = CType(obj, TextTable).GetHashCode
            End If
        End Function

        Public Overrides Function GetHashCode() As Integer
            Dim AutoTrimClone As TextTable = CType(Me.Clone(), TextTable)
            AutoTrimClone.AutoTrim()
            Return AutoTrimClone.ToCsvTable.GetHashCode()
        End Function

        Public Function Clone() As Object Implements ICloneable.Clone
            Return New TextTable(Me)
        End Function

        Public Shared Operator =(obj1 As TextTable, obj2 As TextTable) As Boolean
            Return obj1.Equals(obj2)
        End Operator

        Public Shared Operator <>(obj1 As TextTable, obj2 As TextTable) As Boolean
            Return Not obj1.Equals(obj2)
        End Operator

        ''' <summary>
        ''' Create a list with values of all filled cells
        ''' </summary>
        ''' <returns></returns>
        Public Function ToCellValuesList(basedOnSheetName As String) As List(Of TextTableCell)
            Dim Result As New List(Of TextTableCell)
            For MyRowCounter As Integer = 0 To Me.RowCount - 1
                For MyColCounter As Integer = 0 To Me.ColumnCount - 1
                    If Me.Cell(MyRowCounter, MyColCounter) <> Nothing Then
                        Result.Add(New TextTableCell(basedOnSheetName, TextTable.ExcelColumnName(MyColCounter) & MyRowCounter + 1, Me.Cell(MyRowCounter, MyColCounter)))
                    End If
                Next
            Next
            Return Result
        End Function

        Private Sub LoadFromDataTable(table As System.Data.DataTable)
            Me.AddColumns(table.Columns.Count)
            Me.AddRows(table.Rows.Count)
            For RowIndex As Integer = 0 To table.Rows.Count - 1
                For ColIndex As Integer = 0 To table.Columns.Count - 1
                    Me.Cell(RowIndex, ColIndex) = CompuMaster.Data.Utils.NoDBNull(table.Rows(RowIndex)(ColIndex), CType(Nothing, String))
                Next
            Next
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    Me.Table.Dispose()
                End If

                ' TODO: Nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalizer überschreiben
                ' TODO: Große Felder auf NULL setzen
                disposedValue = True
            End If
        End Sub

        ' ' TODO: Finalizer nur überschreiben, wenn "Dispose(disposing As Boolean)" Code für die Freigabe nicht verwalteter Ressourcen enthält
        ' Protected Overrides Sub Finalize()
        '     ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
        '     Dispose(disposing:=False)
        '     MyBase.Finalize()
        ' End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace