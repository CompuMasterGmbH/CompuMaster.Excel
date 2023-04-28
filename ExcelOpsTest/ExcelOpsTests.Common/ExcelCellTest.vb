Imports NUnit.Framework
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
            Assert.AreEqual("A1", Cell1.ToString(False))
            Assert.AreEqual("'Grunddaten'!A1", Cell1.ToString(True))
            Assert.AreEqual("'Grunddaten'!A1", Cell1.ToString)
            Assert.AreEqual("B3", Cell2.ToString(False))
            Assert.AreEqual("'Grunddaten'!B3", Cell2.ToString(True))
            Assert.AreEqual("'Grunddaten'!B3", Cell2.ToString)
            Assert.AreEqual("'Grunddaten'!AB200", Cell3.ToString)
            Assert.AreEqual("'Grunddaten'!XFD14", Cell4.ToString)
            Assert.AreEqual(1, Cell1.RowNumber)
            Assert.AreEqual(0, Cell1.RowIndex)
            Assert.AreEqual("A", Cell1.ColumnName)
            Assert.AreEqual(0, Cell1.ColumnIndex)
            Assert.AreEqual(3, Cell2.RowNumber)
            Assert.AreEqual(2, Cell2.RowIndex)
            Assert.AreEqual("B", Cell2.ColumnName)
            Assert.AreEqual(1, Cell2.ColumnIndex)
            Assert.AreEqual(200, Cell3.RowNumber)
            Assert.AreEqual(199, Cell3.RowIndex)
            Assert.AreEqual("AB", Cell3.ColumnName)
            Assert.AreEqual(27, Cell3.ColumnIndex)
            Assert.AreEqual(14, Cell4.RowNumber)
            Assert.AreEqual(13, Cell4.RowIndex)
            Assert.AreEqual("XFD", Cell4.ColumnName)
            Assert.AreEqual(16383, Cell4.ColumnIndex)
            Assert.AreEqual("Grunddaten", Cell4.SheetName)
            Assert.AreEqual("XFD14", Cell4.Address)
            Assert.AreEqual("Grund Daten", Cell5.SheetName)
            Assert.AreEqual("XFD14", Cell5.Address)
            Assert.AreEqual("Grund Daten", Cell6.SheetName)
            Assert.AreEqual("XFD14", Cell6.Address)
            Assert.AreEqual("Grund Daten", Cell7.SheetName)
            Assert.AreEqual("XFD14", Cell7.Address)
            Assert.AreEqual(Nothing, Cell8.SheetName)
            Assert.AreEqual("XFD14", Cell8.Address)
        End Sub

        <Test> Public Sub ExcelCellIsValidAddress()
            Assert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("A1", False))
            Assert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("a1", False))
            Assert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("XFD1", False))
            Assert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("B3", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A0", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A" & Long.MaxValue, False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("*A1", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A1*", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A!1", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A,1", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("Aä1", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("AAAA1", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("A1A1", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("$A$1", False))
            Assert.IsTrue(ExcelOps.ExcelCell.IsValidAddress("$A$1", True))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("52", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("!52", False))
            Assert.IsFalse(ExcelOps.ExcelCell.IsValidAddress("=52", False))
            Assert.Catch(Of ArgumentException)(Sub()
                                                   Dim Cell1 As New ExcelOps.ExcelCell("Grunddaten", "A1A1", ExcelOps.ExcelCell.ValueTypes.All)
                                               End Sub)
        End Sub

        <Test> Public Sub EqualityAndComparisons()
            Dim CellA1 = New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All)
            Dim CellA2 = New ExcelCell("", "A2", ExcelOps.ExcelCell.ValueTypes.All)
            Dim CellB1 = New ExcelCell("", "B1", ExcelOps.ExcelCell.ValueTypes.All)
            Dim CellA1Dup = New ExcelCell("", "A1", ExcelOps.ExcelCell.ValueTypes.All)

            Assert.IsTrue(CellA1 = CellA1Dup)
            Assert.IsFalse(CellA1 <> CellA1Dup)
            Assert.IsTrue(CellA1.Equals(CellA1Dup))
            Assert.AreEqual(CellA1, CellA1Dup)
            Assert.AreEqual(CellA1.GetHashCode, CellA1Dup.GetHashCode)

            Assert.IsFalse(CellA1 = CellB1)
            Assert.IsTrue(CellA1 <> CellB1)
            Assert.IsFalse(CellA1.Equals(CellB1))
            Assert.AreNotEqual(CellA1, CellB1)
            Assert.AreNotEqual(CellA1.GetHashCode, CellB1.GetHashCode)

            Assert.IsFalse(CellA1 > CellB1)
            Assert.IsTrue(CellA1 < CellB1)

            Assert.IsFalse(CellA1 > CellA2)
            Assert.IsTrue(CellA1 < CellA2)

            Assert.Negative(CellA1.CompareTo(CellB1))
            Assert.Positive(CellB1.CompareTo(CellA1))

            Assert.Negative(CellA1.CompareTo(CellA2))
            Assert.Positive(CellA2.CompareTo(CellA1))

            Assert.Zero(CellA1.CompareTo(CellA1Dup))
            Assert.Positive(CellA2.CompareTo(CellA1Dup))
        End Sub

    End Class

End Namespace