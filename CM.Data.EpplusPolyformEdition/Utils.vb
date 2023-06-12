Option Explicit On
Option Strict On

Imports System.Data

Friend Class Utils

    ''' <summary>
    '''     Lookup a new unique column name for a data table
    ''' </summary>
    ''' <param name="dataTable">The data table which shall get a new data column</param>
    ''' <param name="suggestedColumnName">A column name suggestion</param>
    ''' <returns>The suggested column name as it is or modified column name to be unique</returns>
    Friend Shared Function LookupUniqueColumnName(ByVal dataTable As DataTable, ByVal suggestedColumnName As String) As String

        Dim ColumnNameAlreadyExistant As Boolean = False
        For MyCounter As Integer = 0 To dataTable.Columns.Count - 1
            If String.Compare(suggestedColumnName, dataTable.Columns(MyCounter).ColumnName, True) = 0 Then
                ColumnNameAlreadyExistant = True
            End If
        Next

        If ColumnNameAlreadyExistant = False Then
            'Exit function
            Return suggestedColumnName
        Else
            'Find the position range of an already existing counter at the end of the string - if there is a number
            Dim NumberPositionIndex As Integer = -1
            For NumberPartCounter As Integer = suggestedColumnName.Length - 1 To 0 Step -1
                If Char.IsNumber(suggestedColumnName.Chars(NumberPartCounter)) = False Then
                    NumberPositionIndex = NumberPartCounter + 1 'Next char behind the current char
                    Exit For
                End If
            Next
            'Read out the value of the counter
            Dim NumberCounterValue As Integer = 0
            If NumberPositionIndex = -1 OrElse NumberPositionIndex + 1 > suggestedColumnName.Length Then
                'Attach a new counter value
                NumberCounterValue = 1
                suggestedColumnName = suggestedColumnName & NumberCounterValue.ToString
            Else
                'Update the counter value
                NumberCounterValue = CType(suggestedColumnName.Substring(NumberPositionIndex), Integer) + 1
#Disable Warning CA1845 ' Use span-based 'string.Concat'
                suggestedColumnName = suggestedColumnName.Substring(0, NumberPositionIndex) & NumberCounterValue.ToString
#Enable Warning CA1845 ' Use span-based 'string.Concat'
            End If

            'Revalidate uniqueness by running recursively
            suggestedColumnName = LookupUniqueColumnName(dataTable, suggestedColumnName)
        End If

        Return suggestedColumnName

    End Function

End Class
