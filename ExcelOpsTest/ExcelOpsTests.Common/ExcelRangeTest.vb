Imports NUnit.Framework
Imports NUnit.Framework.Legacy
Imports CompuMaster.Excel.ExcelOps

Namespace ExcelOpsEngineTests
    <TestFixture> Public Class ExcelRangeTest

        <SetUp> Public Sub ResetConsoleForTestOutput()
            Test.Console.ResetConsoleForTestOutput()
        End Sub

        <Test> Public Sub ExcelRange()
            Dim Cell1 As New ExcelOps.ExcelCell("Grunddaten", "A1", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Cell2 As New ExcelOps.ExcelCell("Grunddaten", "B3", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Range As New ExcelOps.ExcelRange(Cell1, Cell2)
            ClassicAssert.AreEqual("A1:B3", Range.ToString(False))
            ClassicAssert.AreEqual("'Grunddaten'!A1:B3", Range.ToString(True))
            ClassicAssert.AreEqual("'Grunddaten'!A1:B3", Range.ToString())
            ClassicAssert.AreEqual(Range, New ExcelRange("Grunddaten", "A1:B3"))
            ClassicAssert.AreEqual(1, New ExcelOps.ExcelRange(Cell2, Cell2).CellCount)
            ClassicAssert.AreEqual(6, New ExcelOps.ExcelRange(Cell1, Cell2).CellCount)
            ClassicAssert.Catch(Of ArgumentException)(Sub()
                                                          ClassicAssert.AreEqual(Range, New ExcelRange("Grunddaten", "B3:A1"))
                                                      End Sub)
            ClassicAssert.Catch(Of ArgumentException)(Sub()
                                                          Dim InvalidRange As New ExcelOps.ExcelRange(Cell2, Cell1)
                                                      End Sub)
            ClassicAssert.Catch(Of ArgumentException)(Sub()
                                                          Dim InvalidRange As New ExcelOps.ExcelRange(
                                                        New ExcelOps.ExcelCell("Grunddaten", "B3", ExcelOps.ExcelCell.ValueTypes.All),
                                                        New ExcelOps.ExcelCell("Grunddaten", "A3", ExcelOps.ExcelCell.ValueTypes.All)
                                                        )
                                                      End Sub)
            ClassicAssert.Catch(Of ArgumentException)(Sub()
                                                          Dim InvalidRange As New ExcelOps.ExcelRange(
                                                        New ExcelOps.ExcelCell("Grunddaten", "B3", ExcelOps.ExcelCell.ValueTypes.All),
                                                        New ExcelOps.ExcelCell("Grunddaten", "B2", ExcelOps.ExcelCell.ValueTypes.All)
                                                        )
                                                      End Sub)
            ClassicAssert.AreEqual(6, Range.CellCount)
        End Sub

        <Test> Public Sub Clone()
            Dim Range As New ExcelOps.ExcelRange("Grunddaten", "A1:B3")
            Dim MyExcelCellEnumCounter As Integer = 0
            For Each MyExcelCell As ExcelCell In Range
                MyExcelCellEnumCounter += 1
                Select Case MyExcelCellEnumCounter
                    Case 1
                        ClassicAssert.AreEqual("A1", MyExcelCell.ToString(False))
                    Case 2
                        ClassicAssert.AreEqual("B1", MyExcelCell.ToString(False))
                    Case 3
                        ClassicAssert.AreEqual("A2", MyExcelCell.ToString(False))
                    Case 4
                        ClassicAssert.AreEqual("B2", MyExcelCell.ToString(False))
                    Case 5
                        ClassicAssert.AreEqual("A3", MyExcelCell.ToString(False))
                    Case 6
                        ClassicAssert.AreEqual("B3", MyExcelCell.ToString(False))
                    Case Else
                        ClassicAssert.Fail("Wrong cell count returned by ExcelRange enumerator")
                End Select
            Next
            ClassicAssert.AreEqual(6, MyExcelCellEnumCounter)
        End Sub

        <Test> Public Sub GetEnumerator()
            Dim Range As New ExcelOps.ExcelRange("Grunddaten", "A1:B3")
            Dim ClonedRange As ExcelRange = Range.Clone
            ClonedRange.SheetName = "Replaced"
            ClonedRange.AddressStart.Address = "Z5"
            ClonedRange.AddressEnd.Address = "Z7"
            ClassicAssert.AreEqual(3, ClonedRange.CellCount)
            ClassicAssert.AreEqual("Replaced", ClonedRange.SheetName)
            ClassicAssert.AreEqual(6, Range.CellCount)
            ClassicAssert.AreEqual("Grunddaten", Range.SheetName)
        End Sub


        <Test> Public Sub Cell()
            Dim Range As ExcelOps.ExcelRange

            'Test case: Range at the very top
            Range = New ExcelOps.ExcelRange("Grunddaten", "A1:C3")
            ClassicAssert.AreEqual(9, Range.CellCount)

            'Access to cells row by row
            ClassicAssert.AreEqual("A1", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("A2", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("A3", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("B1", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("B2", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("B3", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("C1", Range.Cell(6, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("C2", Range.Cell(7, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("C3", Range.Cell(8, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)

            'Access to cells column by column
            ClassicAssert.AreEqual("A1", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("B1", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("C1", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("A2", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("B2", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("C2", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("A3", Range.Cell(6, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("B3", Range.Cell(7, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("C3", Range.Cell(8, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)

            'Test case: Range in middle of a sheet
            Range = New ExcelOps.ExcelRange("Grunddaten", "D7:F9")
            ClassicAssert.AreEqual(9, Range.CellCount)

            'Access to cells row by row
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("D8", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("D9", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("E7", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("E8", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("E9", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("F7", Range.Cell(6, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("F8", Range.Cell(7, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("F9", Range.Cell(8, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)

            'Access to cells column by column
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("E7", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("F7", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("D8", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("E8", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("F8", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("D9", Range.Cell(6, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("E9", Range.Cell(7, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("F9", Range.Cell(8, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)

            'Test case: Range of single cell
            Range = New ExcelOps.ExcelRange("Grunddaten", "D7:D7")
            ClassicAssert.AreEqual(1, Range.CellCount)

            'Access to cells row by row
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)

            'Access to cells column by column
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)

            'Test case: Range in middle of a sheet, in same column
            Range = New ExcelOps.ExcelRange("Grunddaten", "D7:D9")
            ClassicAssert.AreEqual(3, Range.CellCount)

            'Access to cells row by row
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("D8", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("D9", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)

            'Access to cells column by column
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("D8", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("D9", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)

            'Test case: Range in middle of a sheet, in same column
            Range = New ExcelOps.ExcelRange("Grunddaten", "D7:F7")
            ClassicAssert.AreEqual(3, Range.CellCount)

            'Access to cells row by row
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("E7", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("F7", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)

            'Access to cells column by column
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("E7", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("F7", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)

            'Test case: Range in middle of a sheet, in same column
            Range = New ExcelOps.ExcelRange("Grunddaten", "D7:F8")
            ClassicAssert.AreEqual(6, Range.CellCount)

            'Access to cells row by row
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("D8", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("E7", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("E8", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("F7", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            ClassicAssert.AreEqual("F8", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)

            'Access to cells column by column
            ClassicAssert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("E7", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("F7", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("D8", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("E8", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            ClassicAssert.AreEqual("F8", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)

            'Test case: Range in middle of a sheet, in same column
            Range = New ExcelOps.ExcelRange("Grunddaten", "D7:F8")
            ClassicAssert.AreEqual(6, Range.CellCount)
            ClassicAssert.Throws(Of IndexOutOfRangeException)(Sub()
                                                                  Dim Dummy = Range.Cell(6, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn)
                                                              End Sub)

        End Sub

        <Test> Public Sub EqualityAndComparisons()
            Dim RangeAtA1_Fields3x3 = New ExcelRange(New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "C3", ExcelOps.ExcelCell.ValueTypes.All))
            Dim RangeAtA2_Fields3x3 = New ExcelRange(New ExcelCell("", "A2", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "C4", ExcelOps.ExcelCell.ValueTypes.All))
            Dim RangeAtB1_Fields3x3 = New ExcelRange(New ExcelCell("", "B1", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "D3", ExcelOps.ExcelCell.ValueTypes.All))
            Dim RangeAtA1_Fields3x3_Dup = New ExcelRange(New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "C3", ExcelOps.ExcelCell.ValueTypes.All))
            Dim RangeAtA1_Fields4x3 = New ExcelRange(New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "C4", ExcelOps.ExcelCell.ValueTypes.All))

            ClassicAssert.IsTrue(RangeAtA1_Fields3x3 = RangeAtA1_Fields3x3_Dup)
            ClassicAssert.IsFalse(RangeAtA1_Fields3x3 <> RangeAtA1_Fields3x3_Dup)
            ClassicAssert.IsTrue(RangeAtA1_Fields3x3.Equals(RangeAtA1_Fields3x3_Dup))
            ClassicAssert.AreEqual(RangeAtA1_Fields3x3, RangeAtA1_Fields3x3_Dup)
            ClassicAssert.AreEqual(RangeAtA1_Fields3x3.GetHashCode, RangeAtA1_Fields3x3_Dup.GetHashCode)

            ClassicAssert.IsFalse(RangeAtA1_Fields3x3 = RangeAtB1_Fields3x3)
            ClassicAssert.IsTrue(RangeAtA1_Fields3x3 <> RangeAtB1_Fields3x3)
            ClassicAssert.IsFalse(RangeAtA1_Fields3x3.Equals(RangeAtB1_Fields3x3))
            ClassicAssert.AreNotEqual(RangeAtA1_Fields3x3, RangeAtB1_Fields3x3)
            ClassicAssert.AreNotEqual(RangeAtA1_Fields3x3.GetHashCode, RangeAtB1_Fields3x3.GetHashCode)

            ClassicAssert.IsFalse(RangeAtA1_Fields3x3 > RangeAtB1_Fields3x3)
            ClassicAssert.IsTrue(RangeAtA1_Fields3x3 < RangeAtB1_Fields3x3)

            ClassicAssert.IsFalse(RangeAtA1_Fields3x3 > RangeAtA2_Fields3x3)
            ClassicAssert.IsTrue(RangeAtA1_Fields3x3 < RangeAtA2_Fields3x3)

            ClassicAssert.Negative(RangeAtA1_Fields3x3.CompareTo(RangeAtB1_Fields3x3))
            ClassicAssert.Positive(RangeAtB1_Fields3x3.CompareTo(RangeAtA1_Fields3x3))

            ClassicAssert.Negative(RangeAtA1_Fields3x3.CompareTo(RangeAtA2_Fields3x3))
            ClassicAssert.Positive(RangeAtA2_Fields3x3.CompareTo(RangeAtA1_Fields3x3))

            ClassicAssert.Zero(RangeAtA1_Fields3x3.CompareTo(RangeAtA1_Fields3x3_Dup))

            ClassicAssert.Negative(RangeAtA1_Fields3x3.CompareTo(RangeAtA1_Fields4x3))
            ClassicAssert.Positive(RangeAtA2_Fields3x3.CompareTo(RangeAtA1_Fields4x3))
        End Sub

    End Class

End Namespace