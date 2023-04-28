Imports NUnit.Framework
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
            Assert.AreEqual("A1:B3", Range.ToString(False))
            Assert.AreEqual("'Grunddaten'!A1:B3", Range.ToString(True))
            Assert.AreEqual("'Grunddaten'!A1:B3", Range.ToString())
            Assert.AreEqual(Range, New ExcelRange("Grunddaten", "A1:B3"))
            Assert.AreEqual(1, New ExcelOps.ExcelRange(Cell2, Cell2).CellCount)
            Assert.AreEqual(6, New ExcelOps.ExcelRange(Cell1, Cell2).CellCount)
            Assert.Catch(Of ArgumentException)(Sub()
                                                   Assert.AreEqual(Range, New ExcelRange("Grunddaten", "B3:A1"))
                                               End Sub)
            Assert.Catch(Of ArgumentException)(Sub()
                                                   Dim InvalidRange As New ExcelOps.ExcelRange(Cell2, Cell1)
                                               End Sub)
            Assert.Catch(Of ArgumentException)(Sub()
                                                   Dim InvalidRange As New ExcelOps.ExcelRange(
                                                        New ExcelOps.ExcelCell("Grunddaten", "B3", ExcelOps.ExcelCell.ValueTypes.All),
                                                        New ExcelOps.ExcelCell("Grunddaten", "A3", ExcelOps.ExcelCell.ValueTypes.All)
                                                        )
                                               End Sub)
            Assert.Catch(Of ArgumentException)(Sub()
                                                   Dim InvalidRange As New ExcelOps.ExcelRange(
                                                        New ExcelOps.ExcelCell("Grunddaten", "B3", ExcelOps.ExcelCell.ValueTypes.All),
                                                        New ExcelOps.ExcelCell("Grunddaten", "B2", ExcelOps.ExcelCell.ValueTypes.All)
                                                        )
                                               End Sub)
            Assert.AreEqual(6, Range.CellCount)
        End Sub

        <Test> Public Sub Clone()
            Dim Range As New ExcelOps.ExcelRange("Grunddaten", "A1:B3")
            Dim MyExcelCellEnumCounter As Integer = 0
            For Each MyExcelCell As ExcelCell In Range
                MyExcelCellEnumCounter += 1
                Select Case MyExcelCellEnumCounter
                    Case 1
                        Assert.AreEqual("A1", MyExcelCell.ToString(False))
                    Case 2
                        Assert.AreEqual("B1", MyExcelCell.ToString(False))
                    Case 3
                        Assert.AreEqual("A2", MyExcelCell.ToString(False))
                    Case 4
                        Assert.AreEqual("B2", MyExcelCell.ToString(False))
                    Case 5
                        Assert.AreEqual("A3", MyExcelCell.ToString(False))
                    Case 6
                        Assert.AreEqual("B3", MyExcelCell.ToString(False))
                    Case Else
                        Assert.Fail("Wrong cell count returned by ExcelRange enumerator")
                End Select
            Next
            Assert.AreEqual(6, MyExcelCellEnumCounter)
        End Sub

        <Test> Public Sub GetEnumerator()
            Dim Range As New ExcelOps.ExcelRange("Grunddaten", "A1:B3")
            Dim ClonedRange As ExcelRange = Range.Clone
            ClonedRange.SheetName = "Replaced"
            ClonedRange.AddressStart.Address = "Z5"
            ClonedRange.AddressEnd.Address = "Z7"
            Assert.AreEqual(3, ClonedRange.CellCount)
            Assert.AreEqual("Replaced", ClonedRange.SheetName)
            Assert.AreEqual(6, Range.CellCount)
            Assert.AreEqual("Grunddaten", Range.SheetName)
        End Sub


        <Test> Public Sub Cell()
            Dim Range As ExcelOps.ExcelRange

            'Test case: Range at the very top
            Range = New ExcelOps.ExcelRange("Grunddaten", "A1:C3")
            Assert.AreEqual(9, Range.CellCount)

            'Access to cells row by row
            Assert.AreEqual("A1", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("A2", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("A3", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("B1", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("B2", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("B3", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("C1", Range.Cell(6, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("C2", Range.Cell(7, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("C3", Range.Cell(8, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)

            'Access to cells column by column
            Assert.AreEqual("A1", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("B1", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("C1", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("A2", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("B2", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("C2", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("A3", Range.Cell(6, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("B3", Range.Cell(7, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("C3", Range.Cell(8, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)

            'Test case: Range in middle of a sheet
            Range = New ExcelOps.ExcelRange("Grunddaten", "D7:F9")
            Assert.AreEqual(9, Range.CellCount)

            'Access to cells row by row
            Assert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("D8", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("D9", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("E7", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("E8", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("E9", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("F7", Range.Cell(6, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("F8", Range.Cell(7, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)
            Assert.AreEqual("F9", Range.Cell(8, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfAColumnThenNextColumn).LocalAddress)

            'Access to cells column by column
            Assert.AreEqual("D7", Range.Cell(0, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("E7", Range.Cell(1, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("F7", Range.Cell(2, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("D8", Range.Cell(3, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("E8", Range.Cell(4, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("F8", Range.Cell(5, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("D9", Range.Cell(6, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("E9", Range.Cell(7, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
            Assert.AreEqual("F9", Range.Cell(8, ExcelOps.ExcelRange.CellAccessDirection.AllCellsOfARowThenNextRow).LocalAddress)
        End Sub

        <Test> Public Sub EqualityAndComparisons()
            Dim RangeAtA1_Fields3x3 = New ExcelRange(New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "C3", ExcelOps.ExcelCell.ValueTypes.All))
            Dim RangeAtA2_Fields3x3 = New ExcelRange(New ExcelCell("", "A2", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "C4", ExcelOps.ExcelCell.ValueTypes.All))
            Dim RangeAtB1_Fields3x3 = New ExcelRange(New ExcelCell("", "B1", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "D3", ExcelOps.ExcelCell.ValueTypes.All))
            Dim RangeAtA1_Fields3x3_Dup = New ExcelRange(New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "C3", ExcelOps.ExcelCell.ValueTypes.All))
            Dim RangeAtA1_Fields4x3 = New ExcelRange(New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All), New ExcelCell("", "C4", ExcelOps.ExcelCell.ValueTypes.All))

            Assert.IsTrue(RangeAtA1_Fields3x3 = RangeAtA1_Fields3x3_Dup)
            Assert.IsFalse(RangeAtA1_Fields3x3 <> RangeAtA1_Fields3x3_Dup)
            Assert.IsTrue(RangeAtA1_Fields3x3.Equals(RangeAtA1_Fields3x3_Dup))
            Assert.AreEqual(RangeAtA1_Fields3x3, RangeAtA1_Fields3x3_Dup)
            Assert.AreEqual(RangeAtA1_Fields3x3.GetHashCode, RangeAtA1_Fields3x3_Dup.GetHashCode)

            Assert.IsFalse(RangeAtA1_Fields3x3 = RangeAtB1_Fields3x3)
            Assert.IsTrue(RangeAtA1_Fields3x3 <> RangeAtB1_Fields3x3)
            Assert.IsFalse(RangeAtA1_Fields3x3.Equals(RangeAtB1_Fields3x3))
            Assert.AreNotEqual(RangeAtA1_Fields3x3, RangeAtB1_Fields3x3)
            Assert.AreNotEqual(RangeAtA1_Fields3x3.GetHashCode, RangeAtB1_Fields3x3.GetHashCode)

            Assert.IsFalse(RangeAtA1_Fields3x3 > RangeAtB1_Fields3x3)
            Assert.IsTrue(RangeAtA1_Fields3x3 < RangeAtB1_Fields3x3)

            Assert.IsFalse(RangeAtA1_Fields3x3 > RangeAtA2_Fields3x3)
            Assert.IsTrue(RangeAtA1_Fields3x3 < RangeAtA2_Fields3x3)

            Assert.Negative(RangeAtA1_Fields3x3.CompareTo(RangeAtB1_Fields3x3))
            Assert.Positive(RangeAtB1_Fields3x3.CompareTo(RangeAtA1_Fields3x3))

            Assert.Negative(RangeAtA1_Fields3x3.CompareTo(RangeAtA2_Fields3x3))
            Assert.Positive(RangeAtA2_Fields3x3.CompareTo(RangeAtA1_Fields3x3))

            Assert.Zero(RangeAtA1_Fields3x3.CompareTo(RangeAtA1_Fields3x3_Dup))

            Assert.Negative(RangeAtA1_Fields3x3.CompareTo(RangeAtA1_Fields4x3))
            Assert.Positive(RangeAtA2_Fields3x3.CompareTo(RangeAtA1_Fields4x3))
        End Sub

    End Class

End Namespace