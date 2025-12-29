Imports NUnit.Framework
Imports NUnit.Framework.Legacy
Imports CompuMaster.Excel.ExcelOps

Namespace ExcelOpsEngineTests
    <TestFixture> Public Class ExcelCellTest

        <SetUp> Public Sub ResetConsoleForTestOutput()
            Test.Console.ResetConsoleForTestOutput()
        End Sub

        <Test> Public Sub ExcelCell()
            Dim Cell1 As New ExcelOps.ExcelCell("Grunddaten", "A1", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Cell2 As New ExcelOps.ExcelCell("Grunddaten", "B3", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Cell3 As New ExcelOps.ExcelCell("Grunddaten", "AB200", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Cell4 As New ExcelOps.ExcelCell("Grunddaten", "XFD14", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Cell5 As New ExcelOps.ExcelCell("Grund Daten", "XFD14", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Cell6 As New ExcelOps.ExcelCell("Grund Daten!XFD14", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Cell7 As New ExcelOps.ExcelCell("'Grund Daten'!XFD14", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Cell8 As New ExcelOps.ExcelCell("XFD14", ExcelOps.ExcelCell.ValueTypes.All)
            ClassicAssert.AreEqual("A1", Cell1.ToString(False))
            ClassicAssert.AreEqual("'Grunddaten'!A1", Cell1.ToString(True))
            ClassicAssert.AreEqual("'Grunddaten'!A1", Cell1.ToString)
            ClassicAssert.AreEqual("B3", Cell2.ToString(False))
            ClassicAssert.AreEqual("'Grunddaten'!B3", Cell2.ToString(True))
            ClassicAssert.AreEqual("'Grunddaten'!B3", Cell2.ToString)
            ClassicAssert.AreEqual("'Grunddaten'!AB200", Cell3.ToString)
            ClassicAssert.AreEqual("'Grunddaten'!XFD14", Cell4.ToString)
            ClassicAssert.AreEqual(1, Cell1.RowNumber)
            ClassicAssert.AreEqual(0, Cell1.RowIndex)
            ClassicAssert.AreEqual("A", Cell1.ColumnName)
            ClassicAssert.AreEqual(0, Cell1.ColumnIndex)
            ClassicAssert.AreEqual(3, Cell2.RowNumber)
            ClassicAssert.AreEqual(2, Cell2.RowIndex)
            ClassicAssert.AreEqual("B", Cell2.ColumnName)
            ClassicAssert.AreEqual(1, Cell2.ColumnIndex)
            ClassicAssert.AreEqual(200, Cell3.RowNumber)
            ClassicAssert.AreEqual(199, Cell3.RowIndex)
            ClassicAssert.AreEqual("AB", Cell3.ColumnName)
            ClassicAssert.AreEqual(27, Cell3.ColumnIndex)
            ClassicAssert.AreEqual(14, Cell4.RowNumber)
            ClassicAssert.AreEqual(13, Cell4.RowIndex)
            ClassicAssert.AreEqual("XFD", Cell4.ColumnName)
            ClassicAssert.AreEqual(16383, Cell4.ColumnIndex)
            ClassicAssert.AreEqual("Grunddaten", Cell4.SheetName)
            ClassicAssert.AreEqual("XFD14", Cell4.Address)
            ClassicAssert.AreEqual("Grund Daten", Cell5.SheetName)
            ClassicAssert.AreEqual("XFD14", Cell5.Address)
            ClassicAssert.AreEqual("Grund Daten", Cell6.SheetName)
            ClassicAssert.AreEqual("XFD14", Cell6.Address)
            ClassicAssert.AreEqual("Grund Daten", Cell7.SheetName)
            ClassicAssert.AreEqual("XFD14", Cell7.Address)
            ClassicAssert.AreEqual(Nothing, Cell8.SheetName)
            ClassicAssert.AreEqual("XFD14", Cell8.Address)
        End Sub

        <Test> Public Sub ExcelCellIsValidAddress()
            ClassicAssert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("A1", False))
            ClassicAssert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("a1", False))
            ClassicAssert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("XFD1", False))
            ClassicAssert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("B3", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A0", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A" & Long.MaxValue, False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("*A1", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A1*", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A!1", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A,1", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("Aä1", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("AAAA1", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A1A1", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("$A$1", False))
            ClassicAssert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("$A$1", True))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("52", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("!52", False))
            ClassicAssert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("=52", False))
            ClassicAssert.Catch(Of ArgumentException)(Sub()
                                                          Dim Cell1 As New ExcelOps.ExcelCell("Grunddaten", "A1A1", ExcelOps.ExcelCell.ValueTypes.All)
                                                      End Sub)
        End Sub

        <Test> Public Sub EqualityAndComparisons()
            Dim CellA1 = New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All)
            Dim CellA2 = New ExcelCell("", "A2", ExcelOps.ExcelCell.ValueTypes.All)
            Dim CellB1 = New ExcelCell("", "B1", ExcelOps.ExcelCell.ValueTypes.All)
            Dim CellA1Dup = New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All)

            ClassicAssert.IsTrue(CellA1 = CellA1Dup)
            ClassicAssert.IsFalse(CellA1 <> CellA1Dup)
            ClassicAssert.IsTrue(CellA1.Equals(CellA1Dup))
            ClassicAssert.AreEqual(CellA1, CellA1Dup)
            ClassicAssert.AreEqual(CellA1.GetHashCode, CellA1Dup.GetHashCode)

            ClassicAssert.IsFalse(CellA1 = CellB1)
            ClassicAssert.IsTrue(CellA1 <> CellB1)
            ClassicAssert.IsFalse(CellA1.Equals(CellB1))
            ClassicAssert.AreNotEqual(CellA1, CellB1)
            ClassicAssert.AreNotEqual(CellA1.GetHashCode, CellB1.GetHashCode)

            ClassicAssert.IsFalse(CellA1 > CellB1)
            ClassicAssert.IsTrue(CellA1 < CellB1)

            ClassicAssert.IsFalse(CellA1 > CellA2)
            ClassicAssert.IsTrue(CellA1 < CellA2)

            ClassicAssert.Negative(CellA1.CompareTo(CellB1))
            ClassicAssert.Positive(CellB1.CompareTo(CellA1))

            ClassicAssert.Negative(CellA1.CompareTo(CellA2))
            ClassicAssert.Positive(CellA2.CompareTo(CellA1))

            ClassicAssert.Zero(CellA1.CompareTo(CellA1Dup))
            ClassicAssert.Positive(CellA2.CompareTo(CellA1Dup))
        End Sub

    End Class

End Namespace